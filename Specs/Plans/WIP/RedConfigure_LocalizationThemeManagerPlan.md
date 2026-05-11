# RedConfigure Localization And Theme Manager Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a new first-party Windows application project, `RedConfigure`, for editing RedSalamander localization resources and theme files with guided examples, live theme previews, validation, and export.

**Architecture:** `RedConfigure.exe` is a standalone DxUI-based tool in the solution. It scans first-party project/resource inputs from the selected repo root, builds editable in-memory localization and theme models, shows realistic preview surfaces, and writes explicit output files only when the user exports. Localization output is `.rc`; theme output is `.theme.json5` using the existing RedSalamander theme schema.

**Tech Stack:** MSBuild `.vcxproj`, C++ latest MSVC mode, Win32, Direct2D/DirectWrite, shared DxUI controls, WIL RAII, yyjson/JSON5, existing Common settings/theme helpers, resource compiler validation, deterministic self-tests.

---

Last updated: 2026-05-10

Status: WIP - first usable DxUI slice landed. RedConfigure can scan a workspace, switch resource owners, edit `STRINGTABLE` translations with placeholder validation, edit theme colors with a live preview, and export `.rc` / `.theme.json5`; richer resource forms, full validation/export review, perf evidence, and closeout specs remain.

Progress notes:

- 2026-05-10: Phase 1 foundation landed with `RedConfigure` and `RedConfigureTests` in the solution, plus a localized first shell page and page model tests.
- 2026-05-10: Phase 2 shared theme JSON5 I/O landed in `Common/ThemeDefinitionIo.h` and `Common/Common/ThemeDefinitionIo.cpp`; `RedSalamander/Preferences.Themes.cpp` now imports/exports themes through the shared helper.
- 2026-05-10: Phase 3 workspace discovery landed with structured `.vcxproj` XML scanning, `.build` pruning, resource owner discovery, satellite `.rc` discovery, theme file discovery, and surfaced scan errors.
- 2026-05-10: First usable DxUI version landed in `RedConfigure/Main.cpp` with six navigation pages, owner/theme selectors, translation grid/editor, theme color editor, live preview, default export paths, and explicit `.rc` / `.theme.json5` export buttons.
- 2026-05-10: Localization v1 model landed for `STRINGTABLE` parsing, placeholder validation, source/satellite merge, deterministic satellite output, and session-level owner/theme switching.
- 2026-05-10: Parser inventory expanded to `MENU`/`POPUP`/`MENUITEM` captions and dialog captions/static/control text; Validation And Export now shows pending `.rc` and `.theme.json5` text previews before writing.
- 2026-05-10: Localization Inventory now uses the generic inventory model, so the grid shows string, menu, and dialog localizable text instead of only editable string-table rows.
- 2026-05-10 validation: `RedConfigureTests` passed; `RedConfigure`, `RedConfigureTests`, and `RedSalamander` Debug x64 builds passed with zero warnings; focused Preferences theme load/save self-tests passed and archived results under `Specs/TestRuns/4cb089111a23/Commands/2026-05-10_123156/` and `Specs/TestRuns/4cb089111a23/Commands/2026-05-10_123201/`.
- 2026-05-10 validation: first usable RedConfigure slice verified with `RedConfigureTests`, `.\build.ps1 -ProjectName RedConfigure`, `.\build.ps1 -ProjectName RedSalamander`, and a native app startup smoke check.
- 2026-05-10 validation: latest iteration verified with `.\build.ps1 -ProjectName RedConfigure`, `.\build.ps1 -ProjectName RedConfigureTests`, `.\.build\x64\Debug\RedConfigureTests.exe`, and a native `RedConfigure.exe` startup smoke check.
- 2026-05-10: UX redesign pass added. The next iteration should evolve the first usable six-page shell into task workbenches for parallel localization editing and live theme authoring, with clearer navigation, stronger previews, color expressions, batch edits, and a review-first export flow.
- 2026-05-10: First UX-redesign implementation slice landed in code: task-mode navigation, a first-pass scope bar, combined Localization and Theme workbenches, expression text support in the theme color field, flattened expression export, and a RedConfigure application icon.
- 2026-05-10 validation: UX-redesign implementation slice verified with `.\build.ps1 -ProjectName RedConfigureTests`, `.\.build\x64\Debug\RedConfigureTests.exe`, `.\build.ps1 -ProjectName RedConfigure`, and a native `RedConfigure.exe` startup smoke check.
- 2026-05-10: Screenshot-driven UX follow-up landed: `Start` now auto-normalizes build output paths to the repo root and focuses on scan health, culture and active owner moved to `Localization`, culture selection uses discovered cultures plus Windows locale names, DxUi repeated `Ctrl+Backspace` path editing was fixed, and theme previews gained richer samples, click-to-select regions, active-theme authored color keys, and first group batch edit actions.
- 2026-05-10: Localization table UX follow-up landed: the confusing two-table layout was replaced by one primary resources table, column headers now sort, search and ID/status filters narrow the table, filtered counts show visible versus total rows, and culture choices now include readable locale names next to culture codes.
- 2026-05-10: Compact first-open layout fix landed for Localization: the first window is larger, narrow content stacks scope/filter controls, and the grid/editor/status bands reserve vertical space so controls do not overlap at high DPI.
- 2026-05-10: Theme first-open UX follow-up landed: the Theme workbench now uses a two-column editor/preview layout when space allows, keeps the preview in a vertical scroll panel, offers auto-opening value/expression suggestions while typing, shows the previously selected token with a swatch, and maps the visible accent stripe back to `app.accent`.
- 2026-05-10 validation: Theme first-open UX follow-up verified with `.\build.ps1 -ProjectName RedConfigure`, `.\build.ps1 -ProjectName RedConfigureTests`, `.\.build\x64\Debug\RedConfigureTests.exe`, `.\.build\x64\Debug\DxUiTests.exe TextField`, `git diff --check`, and a native `RedConfigure.exe` startup smoke check.
- 2026-05-10: Theme color-key filter landed: the Theme workbench now has a key filter above the color key picker so users can quickly narrow token groups or substrings such as `menu`, `accent`, or `background`.
- 2026-05-10: Theme token grid and preview focus landed: the Theme workbench now shows a grid of filtered keys with effective swatches and authored values, preview hit-testing prefers the smallest clicked sample region, repeated clicks cycle through containing regions, and the selected/edited key is highlighted in the live preview.

## Live Implementation Checklist

Update this checklist during implementation whenever a new required action is discovered.

- [x] Replace six implementation pages with task modes: `Start`, `Localization`, `Themes`, and `Review & Export`.
- [x] Merge localization inventory, translation grid, focused editor, and validation into one usable Localization workbench.
- [x] Merge theme library, token editing, color status, live preview, and export preparation into one usable Theme workbench.
- [x] Add a first-pass global scope/status bar so workspace, culture, owner, theme, and validation state are visible without page-hopping.
- [x] Add app icon resources and use the icon for the window class and executable.
- [x] Add theme expression support in the preview model for references and derived colors.
- [x] Allow the theme color field to accept direct colors plus expression text such as `ref(app.accent)`, `darken(app.accent,20%)`, and `blend(menu.background,app.accent,16%)`.
- [x] Export expression-authored theme colors as flattened `.theme.json5` direct colors until durable `colorExpressions` runtime support lands.
- [x] Auto-normalize a launched/debug workspace path such as `.build\x64\Debug` back to the repository root.
- [x] Remove culture and active-owner editing from `Start`; keep `Start` focused on workspace loading and scan health.
- [x] Move culture selection and active owner selection into `Localization`.
- [x] Replace the free-form culture field with an existing-culture picker plus an official Windows culture list for creating a new target culture.
- [x] Fix the overlapping status/count labels visible on the first page.
- [x] Fix repeated `Ctrl+Backspace` in DxUi text fields so workspace/path editing keeps deleting meaningful path segments.
- [x] Expand the theme token list and sample preview so more settings show immediate impact.
- [x] Make theme sample regions selectable: clicking a previewed surface selects the matching theme token.
- [x] Add first batch edit controls for theme color groups.
- [x] Replace the two-table Localization layout with one primary resources table.
- [x] Add Localization search across resource ID, source, and target text.
- [x] Add Localization column filters for ID text and validation status.
- [x] Add sortable Localization table columns with visible sort glyph updates.
- [x] Show culture names in words next to culture codes in the culture picker.
- [x] Fix compact first-open Localization layout so filter controls, grid, editor labels, and edit fields cannot overlap.
- [x] Fix first-open Theme layout so editor controls and preview examples do not overlap the status line.
- [x] Put the Theme preview examples inside a vertical scroll panel so all samples are reachable at the initial window size.
- [x] Add guided Theme color/expression suggestions that open while typing in the value editor.
- [x] Add a Theme color-key filter so the token picker can be narrowed by group or substring.
- [x] Replace the visible Theme color-key combo with a grid of key, effective value, and authored value.
- [x] Highlight the selected Theme key in the live preview while editing.
- [x] Improve preview click behavior so nested sample regions win first and repeated clicks cycle through overlapping/containing regions.
- [x] Show the previously selected Theme token with a swatch beside its name.
- [x] Correct Theme preview click mapping so the visible accent stripe selects `app.accent`.
- [x] Add DxUi editable-combo opt-in auto-open behavior with focused regression coverage.
- [x] Add or update tests before each behavior change.
- [x] Verify `RedConfigureTests`.
- [x] Verify `.\build.ps1 -ProjectName RedConfigure`.
- [x] Update authoritative specs after implementation behavior is stable.
- [ ] Move this WIP plan to `Specs/Plans/Done/` only after implementation, verification, perf evidence, and authoritative specs are complete.

## User-Facing Requirements

- Create a new project named `RedConfigure`.
- The tool manages both localization and themes.
- The tool has several pages, not a single flat editor.
- The tool shows examples so users understand what each setting changes.
- Localization output is `.rc`.
- Theme output is `.theme.json5`.
- Theme options show examples on screen.
- Theme edits update the examples in real time before export.
- The plan must stay in `Specs/Plans/WIP/` until implementation, verification, perf evidence, and authoritative specs are complete.

## Product Shape

`RedConfigure` is a configuration authoring tool, not the runtime Preferences dialog.

It should help a contributor or translator make durable project files:

- Translation resource scripts under owner language folders, for example `RedSalamander/Lang/fr-FR/RedSalamander-fr-FR.rc`.
- Theme files under `Specs/Themes/` for shipped examples or under a user-selected `Themes/` folder next to an executable.

It should not silently modify the live user settings file. The user edits a working model, sees the preview, validates it, then explicitly exports files.

## UX Redesign Direction

The first usable version proves the data flow, but the user experience is still too page-oriented. The next iteration should make the workflow obvious by turning RedConfigure into two main workbenches:

- `Localization` for editing all selected languages in parallel.
- `Themes` for editing every visible theme token with live impact preview.

The app still has several pages/modes, but they should map to user jobs instead of implementation modules:

1. `Start` - choose workspace, owners, language set, and theme set.
2. `Localization` - side-by-side translation matrix, focused editor, and resource preview.
3. `Themes` - token browser, color/expression editor, mass edits, and live preview scenes.
4. `Review & Export` - validation, diffs, output paths, and final writes.

All user-facing strings in `RedConfigure` must live in `RedConfigure/RedConfigure.rc`. Static menus must be `.rc` resources. No hardcoded UI labels in C++.

## UX Problems To Fix

- The current six-page flow separates discovery, inventory, editing, preview, and export, so the user must remember state across pages.
- Localization is centered on one target culture at a time, but real translation review needs multiple languages visible next to the English source.
- The current inventory/editor split makes editing feel indirect. Users should be able to edit directly in the grid, then use the side editor only for longer text, placeholders, or warnings.
- Theme keys are technical and flat. Users need grouped tokens, usage descriptions, visible samples, and "what changes if I touch this?" feedback.
- The theme editor only handles direct color values. Users need references, derived colors, and reversible mass operations for coherent theme families.
- Validation/export feels like a late page. Validation status should be visible throughout the app, with the final export page acting as a review basket.

## Global Interaction Model

Use a dense, professional utility layout that is easy to scan for repeated editing work:

- Left mode rail: `Start`, `Localization`, `Themes`, `Review & Export`.
- Top scope bar: workspace root, owner scope, selected language set, selected theme, dirty change count, validation status.
- Context command bar below the scope bar:
  - `Validate`
  - `Export`
  - `Undo`
  - `Redo`
  - `Find`
  - mode-specific primary action such as `Add Language`, `Batch Edit`, or `Apply Recipe`
- Main work area for the active workbench.
- Right inspector/preview rail that updates from the current selection.
- Bottom validation drawer that can be collapsed, but never hides blocking errors after validation runs.
- Nonblocking in-window alerts for scan, validation, and export results.

Important interaction rules:

- Every edit updates the in-memory model immediately and marks the changed row/key.
- Disk writes happen only from `Review & Export`.
- `Esc` cancels an active inline edit; `Ctrl+Z` and `Ctrl+Y` apply to localization text edits, theme color edits, color expression edits, and batch operations.
- `Ctrl+K` opens a command/search palette for resource ID, English text, translated text, theme key, color group, owner, culture, and output file.
- Expensive validation is debounced while typing and can also be run explicitly.
- Preview updates are immediate for the selected row/key; full app-wide validation can run in the background.

## Start Mode

Purpose: make the editable scope explicit before the user starts.

User flow:

1. RedConfigure opens and auto-detects the repo root from the executable/build tree when possible.
2. The page shows discovered resource owners, existing satellite cultures, and theme folders.
3. The user chooses:
   - owner scope: one owner, several owners, or all owners
   - language set: one or more cultures shown as columns in the localization matrix
   - theme set: one editable theme plus optional compare themes
4. The page shows the exact output paths that will be written if the user exports.
5. The user continues to `Localization` or `Themes`.

Content:

- Repo root selector.
- Owner checklist discovered from first-party `.vcxproj` files with `ResourceCompile` items.
- Language set editor:
  - existing cultures from `Lang\<culture>\`
  - add new culture by BCP-47 tag
  - duplicate from another culture as a starting point
  - fallback chain preview, for example `fr-CA -> fr -> embedded English`
- Theme source selector:
  - `Specs/Themes/*.theme.json5`
  - `.build\<Platform>\<Configuration>\Themes/*.theme.json5`
  - user-selected folder
- Scan summary:
  - number of resource owners
  - number of embedded `.rc` files
  - number of satellite `.rc` files
  - number of localizable entries
  - number of theme files
  - validation status

## Localization Workbench

Purpose: translate and review resources efficiently across every selected language.

Layout:

- Left navigator:
  - owner groups
  - resource type groups: strings, menus, dialogs, accelerators, unsupported/read-only inventory
  - status buckets: missing, needs review, placeholder error, accelerator warning, source changed, reviewed
  - saved filters such as `Missing in fr-FR` or `Warnings in menus`
- Center translation matrix:
  - pinned columns: owner, type, resource ID, English source
  - one editable column per selected culture
  - status and validation badges inside each language cell
  - inline editing for short values
  - row expansion for long values, comments, and source location
- Right inspector/preview:
  - focused large editor for the selected cell
  - English source, selected target, and optional neighboring cultures
  - placeholder inspector
  - accelerator inspector
  - rendered resource preview
  - file/source location
  - change history for the current row in this session

Translation matrix behavior:

- Users can add, remove, reorder, and pin language columns without changing files on disk.
- Editing a cell updates the preview immediately.
- `Enter` starts editing, `Ctrl+Enter` commits, `Esc` cancels the active edit.
- `Alt+Down` jumps to the next problem in the same language.
- `Alt+Right` moves to the next language cell for the same resource.
- `Ctrl+Shift+F` searches all source and target text.
- Copy/paste supports rectangular cell ranges for spreadsheet-style review.
- Multi-select supports batch operations on rows and languages.

Localization batch operations:

- Copy English source to selected missing cells.
- Copy one culture into another culture for selected rows.
- Clear selected translations.
- Find/replace within selected cultures.
- Preserve or repair accelerator markers when safe.
- Mark selected rows as reviewed.
- Normalize whitespace around placeholders.
- Show a before/after grid before applying any batch operation.

Localization validation:

- Bare `{}` and unindexed specs such as `{:08X}` are errors.
- `%s`, `%d`, and other printf-style placeholders are errors in resource strings.
- `std::format` positional placeholders in the target must match the source argument set unless the source has no placeholders.
- Target `.rc` output must stay UTF-16-compatible and resource-compiler-safe.
- Accelerator markers such as `&File` are preserved when possible and warned when lost in menu groups.
- Duplicate accelerators inside the same menu popup are warnings.
- Empty target text is a missing translation unless the source is intentionally empty.

Localization preview examples:

- Strings render as labels, button text, message text, or formatted messages depending on usage hints.
- Formatted strings show sample arguments applied to positional placeholders.
- Menus render as a small menu bar and popup with accelerator underlines.
- Dialog captions and static/control text render in a compact dialog mock that can show overflow stress.
- Unsupported resource blocks are visible as inventory, but cannot be edited until their parser/writer support lands.

## Theme Workbench

Purpose: make theme editing visual, explain every setting, and keep coherent color systems easy to maintain.

Layout:

- Left token navigator:
  - token groups: app, window, menu, title bar, navigation, folder view, monitor text, file operations, viewer diff, status colors
  - filters: overridden, inherited, expression-based, invalid, low contrast, used in current preview scene
  - search by key, description, current color, referenced key, or usage
- Center token table:
  - key
  - description
  - effective swatch
  - authored value
  - source: base, direct override, reference, expression, fallback
  - contrast status when the foreground/background pair is known
  - usage count
- Right live preview stage:
  - scene tabs: App Shell, Folder View, Menu Popup, Dialogs, File Operations, Monitor Log, Viewer Diff
  - click any preview region to select the relevant theme token
  - highlight all regions affected by the selected token
  - optional compare mode showing base theme vs edited theme

Theme editor modes:

- `Solid`: hex field supporting `#RRGGBB` and `#AARRGGBB`, swatch, alpha slider, recent colors, copy/paste color.
- `Reference`: select another token and reuse its effective color.
- `Expression`: derive a color from one or more tokens using supported transforms.

Color expression examples:

```json5
{
  "id": "user/example",
  "name": "Example",
  "baseThemeId": "builtin/dark",
  "colors": {
    "app.accent": "#2ECC71"
  },
  "colorExpressions": {
    "folderView.itemBackgroundSelected": { "ref": "app.accent", "darken": 0.20 },
    "folderView.itemForegroundSelected": { "ref": "menu.text" },
    "menu.border": { "blend": ["menu.background", "app.accent", 0.16] },
    "selection.inactiveFill": { "ref": "folderView.itemBackgroundSelected", "alpha": 0.45 }
  }
}
```

Supported expression operations for the first schema extension:

- `ref`: reuse another token's effective color.
- `lighten`: mix toward white by `0.0` to `1.0`.
- `darken`: mix toward black by `0.0` to `1.0`.
- `alpha`: replace alpha with `0.0` to `1.0`.
- `blend`: mix two token colors by `0.0` to `1.0`.
- `contrast`: choose readable light/dark foreground against a referenced background.

Expression safety rules:

- Dependencies are resolved from the selected base theme plus authored direct colors plus authored expressions.
- Cycles are errors and keep the last valid preview active.
- Missing references are errors.
- Invalid expression values keep the previous valid preview color while the editor shows the exact problem.
- The dependency inspector shows:
  - "this token depends on"
  - "changing this token affects"
  - cycle path when a cycle is detected

Theme mass modification:

- Multi-select tokens by group, filter, search result, current color, or usage.
- Batch actions:
  - lighten/darken selected tokens
  - blend selected tokens with accent
  - set alpha on selected tokens
  - replace references from one token to another
  - convert matching solid colors to references
  - remove overrides for selected tokens
  - apply a color harmony recipe
- Recipes:
  - create a dark variant from a light theme
  - create a light variant from a dark theme
  - recolor theme around a new accent
  - soften selection colors
  - increase contrast for foreground/background pairs
  - derive warning/error/info colors from semantic base colors
- Every batch operation opens a before/after preview with affected token count, changed swatches, validation changes, and affected preview scenes before it is applied.

Theme preview examples:

- App shell with title bar, command strip, navigation, and status area.
- Folder pane with normal, hovered, selected, inactive selected, disabled, warning, info, and error rows.
- Menu bar and popup menu with disabled text, selection, border, separators, and accelerator text.
- Dialog and buttons with default action, secondary action, destructive/warning action, text field, checkbox, and focus ring.
- File operations popup with total progress, item progress, graph line, graph grid, and scrollbar.
- Monitor log with text view foreground, background, caret, selection, gutter, and metadata colors.
- Viewer diff using added, removed, context, header, banner, placeholder, and divider colors.

## Review And Export

Purpose: make final output explicit, safe, and confidence-building.

The bottom validation drawer should make errors visible throughout the app, but `Review & Export` is the final checklist before disk writes.

Content:

- Validation summary grouped by localization, theme, output path, and build/resource-compiler checks.
- Export basket with one card per output file.
- Diff-like preview for every changed `.rc` and `.theme.json5` file.
- Checkboxes for export targets.
- Exact output paths.
- "After export" summary showing what will exist on disk.

Rules:

- Export is blocked by errors.
- Warnings are allowed after the user sees them.
- Existing files are overwritten only after the diff preview is visible.
- Writers must use deterministic ordering so diffs stay reviewable.
- No external side effects such as commits or publishing are performed by the tool.

## File Map

Create:

- `RedConfigure/RedConfigure.vcxproj` - first-party application project.
- `RedConfigure/RedConfigure.vcxproj.filters` - Visual Studio grouping.
- `RedConfigure/Main.cpp` - process entry point and startup.
- `RedConfigure/RedConfigureApp.h`
- `RedConfigure/RedConfigureApp.cpp` - app lifetime, command-line parsing, workspace defaults.
- `RedConfigure/RedConfigureWindow.h`
- `RedConfigure/RedConfigureWindow.cpp` - top-level DxUI window, page navigation, validation/export command routing.
- `RedConfigure/resource.h` - RedConfigure resource IDs.
- `RedConfigure/RedConfigure.rc` - all RedConfigure strings, menus, dialogs, icons, and version resource includes.
- `RedConfigure/res/exe.manifest` - application manifest.
- `RedConfigure/Workspace/WorkspaceDiscovery.h`
- `RedConfigure/Workspace/WorkspaceDiscovery.cpp` - repo scan, `.vcxproj` resource discovery, theme folder discovery.
- `RedConfigure/Localization/RcResourceModel.h`
- `RedConfigure/Localization/RcResourceModel.cpp` - normalized resource-owner, string, menu, dialog, and satellite translation model.
- `RedConfigure/Localization/RcParser.h`
- `RedConfigure/Localization/RcParser.cpp` - resource script parsing for supported localization resource forms.
- `RedConfigure/Localization/RcWriter.h`
- `RedConfigure/Localization/RcWriter.cpp` - deterministic `.rc` output writer.
- `RedConfigure/Localization/PlaceholderValidation.h`
- `RedConfigure/Localization/PlaceholderValidation.cpp` - positional `std::format` placeholder checks.
- `RedConfigure/Themes/ThemeCatalog.h`
- `RedConfigure/Themes/ThemeCatalog.cpp` - theme file discovery, load, duplicate, import, and export model.
- `RedConfigure/Themes/ThemePreviewModel.h`
- `RedConfigure/Themes/ThemePreviewModel.cpp` - base theme plus override resolution for live previews.
- `RedConfigure/Themes/ThemeValidation.h`
- `RedConfigure/Themes/ThemeValidation.cpp` - theme schema/key/color/contrast validation.
- `RedConfigure/Pages/WorkspacePage.h`
- `RedConfigure/Pages/WorkspacePage.cpp`
- `RedConfigure/Pages/LocalizationInventoryPage.h`
- `RedConfigure/Pages/LocalizationInventoryPage.cpp`
- `RedConfigure/Pages/TranslationEditorPage.h`
- `RedConfigure/Pages/TranslationEditorPage.cpp`
- `RedConfigure/Pages/ThemeLibraryPage.h`
- `RedConfigure/Pages/ThemeLibraryPage.cpp`
- `RedConfigure/Pages/ThemeDesignerPage.h`
- `RedConfigure/Pages/ThemeDesignerPage.cpp`
- `RedConfigure/Pages/ValidationExportPage.h`
- `RedConfigure/Pages/ValidationExportPage.cpp`
- `RedConfigure/Preview/LocalizationExampleControl.h`
- `RedConfigure/Preview/LocalizationExampleControl.cpp`
- `RedConfigure/Preview/ThemePreviewControl.h`
- `RedConfigure/Preview/ThemePreviewControl.cpp`
- `Tests/RedConfigureTests/RedConfigureTests.vcxproj`
- `Tests/RedConfigureTests/RedConfigureTests.cpp`
- `Specs/Core/Core_RedConfigure.md` - authoritative product/spec contract after implementation begins.
- `Specs/UI/UI_RedConfigure.md` - UI behavior and page contracts.

Modify:

- `RedSalamander.sln` - add `RedConfigure` and `RedConfigureTests`.
- `build.ps1` - allow `-ProjectName RedConfigure` and `-ProjectName RedConfigureTests`.
- `Directory.Build.props` and `Directory.Build.targets` only if new shared project defaults are needed.
- `Common/Common.vcxproj` and filters if theme import/export helpers are extracted into Common.
- `Common/ThemeDefinitionIo.h` - shared theme JSON5 import/export helper extracted from Preferences theme code.
- `Common/Common/ThemeDefinitionIo.cpp` - implementation using yyjson and existing color helpers.
- `RedSalamander/Preferences.Themes.cpp` - use shared theme import/export helpers instead of owning duplicate parser/writer code.
- `Specs/Core/Core_Localization.md` - document RedConfigure-generated satellite `.rc` expectations.
- `Specs/Core/Core_SettingsStore.md` - document `.theme.json5` authoring expectations if new validation behavior becomes normative.
- `Specs/Testing/Testing_TestCoverage.md` - add RedConfigure self-test coverage summary after tests land.

## Data Model Contracts

### Localization

Resource owners are discovered from `.vcxproj` files by reading `ResourceCompile` items and project names. This avoids hardcoding RedSalamander-only paths and keeps plugins visible as they are added.

`RcResourceModel` stores:

- owner name, for example `RedSalamander`, `RedSalamanderMonitor`, `ViewerText`
- owner project path
- embedded English `.rc` path
- target culture
- target satellite `.rc` path
- resource entries keyed by owner, type, and ID
- source text
- translated text
- comments when safely recoverable
- source order for deterministic output
- validation state

Supported v1 resource output:

- `STRINGTABLE` entries.
- Static `MENU`/`POPUP`/`MENUITEM` captions.
- Dialog captions and static control text when the parser can identify them safely.
- Unsupported resource blocks remain visible as inventory and are preserved by not rewriting them.

The writer emits a satellite `.rc` for the selected owner and culture. It includes the required owner `resource.h`, language metadata, translated string/menu/dialog resources, and stable ordering. It does not rewrite the embedded English `.rc` file.

### Themes

Theme files use the existing `ThemeDefinition` shape:

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

`ThemePreviewModel` resolves the effective theme from:

1. selected base theme
2. current color overrides
3. derived fallbacks already used by RedSalamander

The editor stores only overrides in output, not every effective color. This keeps theme files small and matches the current settings/theme contract.

For the next UX iteration, RedConfigure should use a richer internal authored-theme model:

- direct color overrides
- references to other theme tokens
- color expressions
- resolved effective colors
- dependency graph
- validation messages attached to each authored token

The initial implementation can flatten expressions into direct `colors` on export so current RedSalamander builds can consume the output. The preferred durable schema extension is an optional `colorExpressions` object next to `colors`:

```json5
{
  "id": "user/example",
  "name": "Example",
  "baseThemeId": "builtin/dark",
  "colors": {
    "app.accent": "#2ECC71"
  },
  "colorExpressions": {
    "folderView.itemBackgroundSelected": { "ref": "app.accent", "darken": 0.20 },
    "folderView.itemForegroundSelected": { "ref": "menu.text" },
    "menu.border": { "blend": ["menu.background", "app.accent", 0.16] },
    "selection.inactiveFill": { "ref": "folderView.itemBackgroundSelected", "alpha": 0.45 }
  }
}
```

Schema-extension rules:

- Existing `colors` string values stay valid and keep the current meaning.
- `colorExpressions` is optional.
- If the same key exists in both `colors` and `colorExpressions`, the direct `colors` value wins unless the UI explicitly marks the expression as active.
- Runtime theme resolution must detect missing references and cycles before applying a theme.
- RedConfigure must be able to export either flattened current-schema themes or extended themes once RedSalamander runtime support lands.
- Extended-theme import/export belongs in shared Common code, not only in RedConfigure.

## Implementation Checklist

### Phase 1: Foundation

- [x] Add `RedConfigure` application project with `Debug`, `Release`, and `ASan Debug` for `x64` and `ARM64`.
- [x] Add `RedConfigureTests` console test project.
- [x] Add `RedConfigure` and `RedConfigureTests` to `RedSalamander.sln`.
- [x] Update `build.ps1` so `.\build.ps1 -ProjectName RedConfigure` builds only the tool and its dependencies.
- [x] Add `RedConfigure/RedConfigure.rc` and `RedConfigure/resource.h` with localized strings for the first shell page.
- [x] Verify `.\build.ps1 -ProjectName RedConfigure` produces `.build\x64\Debug\RedConfigure.exe`.
- [x] Verify `.\build.ps1 -ProjectName RedConfigureTests` produces `.build\x64\Debug\RedConfigureTests.exe`.

### Phase 2: Shared Theme I/O

- [x] Write failing tests for parsing valid `.theme.json5` files with comments and trailing commas.
- [x] Write failing tests for rejecting missing `id`, missing `name`, missing `baseThemeId`, missing `colors`, invalid theme ID, invalid base theme, and invalid color value.
- [x] Extract the existing theme import/export logic from `RedSalamander/Preferences.Themes.cpp` into `Common/ThemeDefinitionIo.h` and `Common/Common/ThemeDefinitionIo.cpp`.
- [x] Keep yyjson ownership under WIL RAII wrappers.
- [x] Use copy APIs for dynamic yyjson strings.
- [x] Update `RedSalamander/Preferences.Themes.cpp` to call the shared helper.
- [x] Add deterministic export ordering by color key.
- [x] Run existing Preferences theme self-tests after extraction.

### Phase 3: Workspace Discovery

- [x] Write tests using a temp fixture repo with one app `.vcxproj`, one plugin `.vcxproj`, one embedded `.rc`, one `Lang\fr-FR` satellite `.rc`, and two theme files.
- [x] Implement `.vcxproj` XML scanning for `ResourceCompile` items using a structured XML reader.
- [x] Ignore build output folders such as `.build`.
- [x] Discover owner name, project path, embedded resource path, and satellite resource paths.
- [x] Discover theme folders and theme files.
- [x] Surface scan errors without crashing the app.

### Phase 4: Localization Parser And Writer

- [x] Write parser tests for `STRINGTABLE` blocks with escaped quotes, comments, duplicate IDs, and multiple blocks.
- [x] Write parser tests for `MENU`, `POPUP`, and `MENUITEM` captions.
- [x] Write parser tests for dialog captions and static text that RedConfigure supports in v1.
- [x] Write placeholder validation tests for indexed placeholders, missing placeholders, extra placeholders, bare `{}`, unindexed `{:08X}`, and printf-style `%s`.
- [x] Implement `RcParser` for supported resource forms.
- [ ] Implement `RcResourceModel` merge behavior for English source plus selected satellite translation.
- [ ] Implement `RcWriter` with deterministic owner/culture output.
- [ ] Validate generated `.rc` files by invoking the resource compiler in tests when the Windows SDK is available.
- [ ] Make unsupported blocks visible as inventory but keep them out of rewritten output.

### Phase 5: Application Shell

- [x] Build the top-level DxUI window with navigation pages.
- [x] Evolve the shell navigation from implementation pages into task modes: `Start`, `Localization`, `Themes`, and `Review & Export`.
- [x] Add the first-pass persistent scope bar with culture, owner, theme, and validation context.
- [ ] Extend the scope bar with workspace root, owner scope, language set editing, dirty count, and richer validation status.
- [ ] Add the context command bar with validate/export/undo/redo/find plus mode-specific primary actions.
- [ ] Add a collapsible validation drawer that is visible from every mode.
- [ ] Add page lifetime management so inactive pages do not keep heavy preview surfaces alive unnecessarily.
- [x] Load all shell labels and commands from `RedConfigure.rc`.
- [ ] Add dirty-state tracking for localization and theme changes.
- [ ] Add explicit export confirmation through the Validation And Export page.
- [ ] Add error reporting through modeless in-window alerts, not blocking message boxes for normal validation findings.

### Phase 6: Localization Workbench

- [x] Implement Workspace page scan controls and summary.
- [x] Implement Localization Inventory grid.
- [x] Combine inventory, translation grid, focused editor, and placeholder status into one Localization workbench.
- [ ] Replace the single-target translation editor with a side-by-side translation matrix.
- [ ] Allow multiple culture columns to be added, removed, reordered, and pinned.
- [ ] Add direct inline cell editing with commit/cancel behavior.
- [ ] Add a right inspector with focused editor, source, selected target, neighboring cultures, placeholder inspector, accelerator inspector, and resource preview.
- [ ] Implement owner/type/status filters.
- [ ] Implement translation search.
- [ ] Implement global resource/language/theme search through the command palette.
- [ ] Implement Translation Editor page behavior as the matrix inspector, with source, target, placeholder inspector, and sample arguments.
- [ ] Implement `LocalizationExampleControl` for strings, formatted strings, command labels, menus, and compact dialog examples.
- [ ] Update examples immediately when the selected translation changes.
- [ ] Add keyboard navigation for previous/next missing resource.
- [ ] Add spreadsheet-style copy/paste for rectangular translation ranges.
- [ ] Add localization batch operations: copy English, copy culture to culture, clear, find/replace, normalize placeholder whitespace, preserve accelerators, and mark reviewed.
- [ ] Add before/after review for localization batch operations.
- [ ] Add duplicate accelerator warnings for sibling menu items.

### Phase 7: Theme Workbench And Live Preview

- [ ] Implement Theme Library page with built-in, file, and user theme sections.
- [ ] Implement import, duplicate, reset, and export model operations.
- [ ] Implement Theme Designer grouped color key list.
- [x] Combine active theme selection, color editing, expression examples, color status, and live preview into one Theme workbench.
- [ ] Replace the first-pass theme workbench with a full token workbench: token navigator, token table, color editor, and live preview stage.
- [ ] Add token descriptions, source type, usage count, and contrast status to the token table.
- [ ] Add preview scene tabs for App Shell, Folder View, Menu Popup, Dialogs, File Operations, Monitor Log, and Viewer Diff.
- [ ] Add click-to-select behavior from preview regions to their theme tokens.
- [ ] Add affected-region highlighting when a token is selected.
- [ ] Implement hex field, swatch, alpha slider, reset override, copy effective color, and copy override color.
- [ ] Add color picker affordances for recent colors and copied colors.
- [x] Implement `ThemePreviewModel` so each color edit recomputes the effective preview theme immediately.
- [ ] Implement `ThemePreviewControl` with menu, navigation, folder view, file operations, monitor text, diff, and dialog/button samples.
- [x] Repaint preview on every valid color edit without writing to disk.
- [x] Keep the previous valid preview visible when the active edit field contains invalid color text.
- [ ] Add contrast/readability indicators for known foreground/background pairs.
- [x] Add internal authored-theme model support for direct colors, references, expressions, effective colors, and dependency validation.
- [x] Add first-pass expression editing through the color value field for solid, reference, and expression values.
- [x] Add expression operations: `ref`, `lighten`, `darken`, `alpha`, `blend`, and `contrast`.
- [ ] Add dependency inspector for "depends on" and "affects" relationships.
- [x] Add cycle and missing-reference validation with previous-valid-preview behavior.
- [ ] Add mass modification workflows: lighten/darken, blend with accent, set alpha, replace reference, convert solids to references, remove overrides, and apply recipes.
- [ ] Add before/after review for theme mass modifications.
- [ ] Add theme recipes for dark variant, light variant, accent recolor, softened selections, increased contrast, and semantic status colors.

### Phase 8: Validation And Export

- [ ] Implement combined validation summary.
- [ ] Block export on localization parse errors, placeholder errors, invalid theme IDs, invalid colors, and output path conflicts.
- [ ] Allow export with warnings after the user sees the warning list.
- [x] Implement diff-like file preview before writes.
- [ ] Convert the Validation And Export page into a Review & Export basket with one card per output file.
- [ ] Keep validation errors visible in the global validation drawer before the user reaches Review & Export.
- [x] Write `.rc` localization output with stable ordering.
- [ ] Write `.theme.json5` theme output with stable ordering and grouped comments.
- [x] Support flattened current-schema theme export when expressions are used internally.
- [ ] Add shared Common support for extended `colorExpressions` import/export before exporting extended themes as durable output.
- [x] Preserve UTF-8/UTF-16 expectations for each output type.
- [ ] Re-run validation after export and show the written paths.

### Phase 9: Self-Tests And Manual Verification

- [x] Add `RedConfigureTests` parser, writer, discovery, validation, and theme I/O unit tests.
- [ ] Add RedSalamander command/self-test coverage for shared `ThemeDefinitionIo` extraction if Preferences behavior changes.
- [ ] Add UI smoke coverage for page creation and page switching.
- [ ] Add a test that theme preview state changes when `folderView.itemBackgroundSelected` changes.
- [x] Add a test that invalid color text does not mutate the previous valid preview color.
- [ ] Add a test that formatted translation preview rejects unindexed placeholders.
- [x] Add a test that `.rc` export contains positional placeholders unchanged.
- [ ] Build Debug x64 and ASan Debug x64.
- [x] Run `RedConfigureTests`.
- [ ] Run focused existing localization tests.
- [x] Run focused existing Preferences Themes tests after shared theme I/O extraction.

### Phase 10: Performance Validation

- [ ] Define a perf scenario for scanning a repo-sized resource set and theme folder set.
- [ ] Instrument workspace scan time, theme load time, localization parse time, validation time, and preview repaint time.
- [ ] Add deterministic perf/selftest coverage for scan plus validation.
- [ ] Archive candidate run evidence under `Specs/TestRuns/`.
- [ ] Record acceptable targets in `Specs/Core/Core_RedConfigure.md`; initial targets:
  - workspace rescan under 500 ms for current repo resource/theme set on a normal development machine
  - single theme color preview update under one frame budget for the preview control
  - validation under 250 ms for current repo resource/theme set after files are already read

### Phase 11: Specs And Closeout

- [x] Create `Specs/Core/Core_RedConfigure.md`.
- [x] Create `Specs/UI/UI_RedConfigure.md`.
- [x] Update `Specs/Core/Core_Localization.md` with RedConfigure satellite `.rc` generation rules.
- [x] Update `Specs/Core/Core_SettingsStore.md` with RedConfigure expression-flattening rules.
- [x] Update `Specs/Testing/Testing_TestCoverage.md` after tests land.
- [ ] Add or update `Specs/TestRuns/` evidence references.
- [ ] Move this plan to `Specs/Plans/Done/` only after implementation, tests, perf evidence, and authoritative specs are complete.

## Validation Matrix

Build commands:

```powershell
.\build.ps1 -ProjectName RedConfigure
.\build.ps1 -ProjectName RedConfigureTests
.\build.ps1 -ProjectName LocalizationTests
.\build.ps1 -ProjectName RedSalamander
```

Required test coverage:

- `RedConfigureTests` parser/writer/discovery/theme validation.
- `LocalizationTests` to ensure resource fallback and satellite behavior remain intact.
- Focused Preferences Themes self-tests to ensure shared theme import/export behavior did not regress.
- Focused command UI smoke tests for RedConfigure page creation when available.

Required perf evidence:

- Archived scan/validation/preview metrics in `Specs/TestRuns/`.
- Before/after evidence if extracting theme I/O affects existing Preferences theme operations.

## Risks And Mitigations

- RC parsing can become fragile if it tries to rewrite arbitrary resource syntax.
  - Mitigation: support known localizable forms first, keep unsupported blocks as read-only inventory, and validate generated output with the resource compiler.
- Theme import/export already exists in Preferences and can drift if duplicated.
  - Mitigation: extract shared `ThemeDefinitionIo` helpers and keep Preferences on the same path as RedConfigure.
- Live preview can be misleading if it does not use the same fallback logic as the app.
  - Mitigation: resolve preview colors through the same `AppTheme` and theme override logic used by RedSalamander, with shared helper extraction where needed.
- Large resource sets can make every keystroke expensive.
  - Mitigation: debounce validation for text edits, update only the selected preview immediately, and keep full validation explicit or backgrounded.
- Translators can accidentally break format placeholders.
  - Mitigation: make placeholder validation page-level and export-blocking.
- Theme expressions can create hidden coupling or dependency cycles.
  - Mitigation: show the dependency graph, validate cycles immediately, keep the previous valid preview active, and block export until the expression graph is valid.
- Batch edits can make too many changes too quickly.
  - Mitigation: require a before/after review for every batch operation, record one undo unit per batch, and show affected file/key counts before applying.

## Non-Goals For The First Implementation

- Online translation services.
- Automatic commits, pushes, or pull requests.
- Editing the user's live settings file as the primary output.
- Replacing the runtime Preferences dialog.
- Shipping RedConfigure from the installer before the tool has explicit product approval.
- Full arbitrary `.rc` round-trip rewriting for unsupported resource forms.
- Machine translation, grammar checking, or external terminology services.
- Advanced theme animation/timing tokens; the first theme scope is colors and color derivation.

## Completion Criteria

- `RedConfigure.exe` builds from the solution and with `build.ps1 -ProjectName RedConfigure`.
- The app opens to task modes for Start, Localization, Themes, and Review & Export.
- Localization can load existing owner resources, edit translations for multiple languages in parallel, preview examples, validate placeholders/accelerators, batch-edit safely, and export `.rc`.
- Theme editing can load existing themes, edit direct colors, author references/expressions, mass-modify tokens, update examples in real time, validate keys/colors/dependencies, and export `.theme.json5`.
- Validation blocks broken output and shows clear reasons.
- Tests cover parser, writer, validation, theme I/O, and preview model behavior.
- Performance evidence is archived under `Specs/TestRuns/`.
- Authoritative specs are updated.
- This WIP plan is moved to `Specs/Plans/Done/`.
