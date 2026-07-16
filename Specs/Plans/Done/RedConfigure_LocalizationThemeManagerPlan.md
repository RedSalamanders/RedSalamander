# RedConfigure Localization And Theme Manager Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a new first-party Windows application project, `RedConfigure`, for editing RedSalamander localization resources and theme files with guided examples, live theme previews, validation, and export.

**Architecture:** `RedConfigure.exe` is a standalone DxUI-based tool in the solution. It scans first-party project/resource inputs from the selected repo root, builds editable in-memory localization and theme models, shows realistic preview surfaces, and writes explicit output files only when the user exports. Localization output is `.rc`; theme output is `.theme.json5` using the existing RedSalamander theme schema.

**Tech Stack:** MSBuild `.vcxproj`, C++ latest MSVC mode, Win32, Direct2D/DirectWrite, shared DxUI controls, WIL RAII, yyjson/JSON5, existing Common settings/theme helpers, resource compiler validation, deterministic self-tests.

---

Last updated: 2026-07-14

Status: Done - the RedConfigure localization/theme manager, expressive version 2 theme authoring, validation/export workflow, deterministic coverage, performance evidence, and authoritative documentation are complete.

2026-07-14 closeout: the advanced shell, configurable localization matrix, focused inspector/examples, batch and clipboard workflows, theme library/scene/recipe tooling, combined validation, per-output Review & Export basket, resource-compiler validation, UI smoke coverage, and repo-sized performance scenario are implemented. Durable behavior is specified in `Specs/Core/Core_RedConfigure.md`, `Specs/UI/UI_RedConfigure.md`, `Specs/Core/Core_Localization.md`, and `Specs/Core/Core_SettingsStore.md`. Candidate evidence is archived under `Specs/TestRuns/4cb089111a23/RedConfigure/2026-07-14_192100/`.

Progress notes (condensed 2026-07-02 folder review; the ~25 original 2026-05-10 bullets are preserved in git history):

- 2026-05-10: Phases 1-3 plus the first usable DxUI slice all landed and were verified the same day: solution/build integration, shared theme JSON5 I/O in Common, workspace discovery, task-mode shell, `STRINGTABLE`/menu/dialog parsing with placeholder validation, theme color/expression editing with live preview, and explicit `.rc` / `.theme.json5` export (validation evidence under `Specs/TestRuns/4cb089111a23/Commands/2026-05-10_123156/` and `Specs/TestRuns/4cb089111a23/Commands/2026-05-10_123201/`).
- 2026-05-10: A series of same-day UX follow-ups reshaped the six implementation pages into Start/Localization/Themes/Review & Export task modes and refined layout, filtering, sorting, culture pickers, the theme token grid, preview click mapping, and DxUi `Ctrl+Backspace` handling; each slice was verified with builds, `RedConfigureTests`, and startup smoke checks.
- 2026-07-14: completed the shared version 2 theme system and reconciled its delivered RedConfigure work: token grid/filter/editor/live composite preview, click-to-select and selected-region highlighting, palette create/rename/delete with reference safety, dependency/affected inspector, source/evaluation badges, fixed preview seeds, authored group transforms, and lossless export. Public theme documentation and all built-in/shipped-theme galleries are current. This does not complete the broader RedConfigure UX items summarized below.

### Completed scope audit (2026-07-14)

- **Shell/workflow:** command bar with undo/redo/find, global validation drawer, richer scope/dirty summary, preview lifetime policy, modeless findings, and explicit export confirmation.
- **Localization:** RC compiler validation, configurable/reorderable/pinned language columns, grid activation into the focused editor, full inspector/examples, global search, rectangular copy/paste, reviewed batch operations, and sibling accelerator diagnostics.
- **Themes:** library/import/duplicate/reset workflow, grouped token metadata, seven-scene selector, alpha/recent/copy affordances, contrast indicators, mass transforms with before/after review, and reusable recipes.
- **Review/export:** combined summary, blocking-error and acknowledged-warning rules, one preview card per output file, persistent validation visibility, grouped-comment theme output, atomic writes, reparse validation, and written-path reporting.
- **Verification/performance:** page creation/switching smoke coverage plus deterministic repo-sized scan/parse/validate/repaint scenarios, metrics, targets, and archived candidate evidence.

## Live Implementation Checklist

2026-07-02 folder review: the original 38-item live checklist was complete except the move-to-Done gate below (37/38); the completed UX-slice items are preserved in git history.

- [x] Move this WIP plan to `Specs/Plans/Done/` only after implementation, verification, perf evidence, and authoritative specs are complete.

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

Color expression examples and the durable JSON5 shape are defined by `Specs/Plans/Done/Theme_ExpressivePaletteAndReferencesPlan_2026-07-13.md` and the authoritative contract in `Specs/Core/Core_SettingsStore.md`.

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

2026-07-02 folder review: File Map superseded — ~20 planned files were never created; actual architecture: RedConfigureRoot / RedConfigureSession / RedConfigureGridModels / Localization/RcWriter / Themes/ThemeCatalog / ThemeExampleControl.

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

Export preserves the authored version 2 representation; there is no legacy flattened export or backward-compatible version 1 reader.

The durable schema and runtime work is complete under `Specs/Plans/Done/Theme_ExpressivePaletteAndReferencesPlan_2026-07-13.md`. The implemented model is a required version 2 named `palette` plus one authored literal-or-expression value per `colors` key. Shared Common code owns parsing, resolution, diagnostics, authored import/export, first-class `builtin/rainbow` inheritance, event-time system sources, perceptual transforms, and allowlisted paint-time seeded sources. RedConfigure edits that same model and uses fixed preview seeds for dynamic colors. The completed plan also owns pinned-source notices and exact upstream license artifacts for shipped Dracula/Catppuccin palettes; the authoritative lasting contracts are in `Specs/Core/Core_SettingsStore.md`, `Specs/Core/Core_RedConfigure.md`, and `Specs/UI/UI_RedConfigure.md`.

## Implementation Checklist

### Phase 1: Foundation

- [x] Complete (2026-05-10) - `RedConfigure` and `RedConfigureTests` projects, solution and `build.ps1` integration, localized shell resources, and build verification all landed; per-item detail in git history.

### Phase 2: Shared Theme I/O

- [x] Complete (2026-05-10) - theme JSON5 import/export extracted from `RedSalamander/Preferences.Themes.cpp` into `Common/ThemeDefinitionIo.h` / `Common/Common/ThemeDefinitionIo.cpp` with parse/reject tests, WIL RAII yyjson ownership, deterministic color-key ordering, and Preferences self-tests re-run; per-item detail in git history.

### Phase 3: Workspace Discovery

- [x] Complete (2026-05-10) - fixture-repo tests plus structured `.vcxproj`/`ResourceCompile` XML scanning, `.build` pruning, owner/embedded/satellite resource discovery, theme discovery, and surfaced scan errors; per-item detail in git history.

### Phase 4: Localization Parser And Writer

- [x] Write parser tests for `STRINGTABLE` blocks with escaped quotes, comments, duplicate IDs, and multiple blocks.
- [x] Write parser tests for `MENU`, `POPUP`, and `MENUITEM` captions.
- [x] Write parser tests for dialog captions and static text that RedConfigure supports in v1.
- [x] Write placeholder validation tests for indexed placeholders, missing placeholders, extra placeholders, bare `{}`, unindexed `{:08X}`, and printf-style `%s`.
- [x] Implement `RcParser` for supported resource forms.
- [x] Implement `RcResourceModel` merge behavior for English source plus selected satellite translation. (delivered by Done plan RedConfigure_LocalizationReviewWorkbenchPlan_2026-06-08.md — implemented as LocalizationReviewRow/LocalizationTargetCell in RedConfigureSession.cpp, test TestRcWriterAndMerge; evidence Specs/TestRuns/4cb089111a23/RedConfigure/LocalizationReview/2026-06-08_204926)
- [x] Implement `RcWriter` with deterministic owner/culture output. (delivered by Done plan RedConfigure_LocalizationReviewWorkbenchPlan_2026-06-08.md — RedConfigure/Localization/RcWriter.cpp + BuildSatelliteRcStringTable; evidence Specs/TestRuns/4cb089111a23/RedConfigure/LocalizationReview/2026-06-08_204926)
- [x] Validate generated `.rc` files by invoking the resource compiler in tests when the Windows SDK is available.
- [x] Make unsupported blocks visible as inventory but keep them out of rewritten output. (delivered by Done plan RedConfigure_LocalizationReviewWorkbenchPlan_2026-06-08.md — InventoryEntry/GetInventoryEntries; evidence Specs/TestRuns/4cb089111a23/RedConfigure/LocalizationReview/2026-06-08_204926)

### Phase 5: Application Shell

- [x] Build the top-level DxUI window with navigation pages.
- [x] Evolve the shell navigation from implementation pages into task modes: `Start`, `Localization`, `Themes`, and `Review & Export`.
- [x] Add the first-pass persistent scope bar with culture, owner, theme, and validation context.
- [x] Extend the scope bar with workspace root, owner scope, language set editing, dirty count, and richer validation status.
- [x] Add the context command bar with validate/export/undo/redo/find plus mode-specific primary actions.
- [x] Add a collapsible validation drawer that is visible from every mode.
- [x] Add page lifetime management so inactive pages do not keep heavy preview surfaces alive unnecessarily.
- [x] Load all shell labels and commands from `RedConfigure.rc`.
- [x] Add dirty-state tracking for localization and theme changes. (delivered by Done plan RedConfigure_LocalizationReviewWorkbenchPlan_2026-06-08.md — LocalizationTargetCell.dirty, ReviewOwnerCultureHasDirtyCell; evidence Specs/TestRuns/4cb089111a23/RedConfigure/LocalizationReview/2026-06-08_204926)
- [x] Add explicit export confirmation through the Validation And Export page.
- [x] Add error reporting through modeless in-window alerts, not blocking message boxes for normal validation findings.

### Phase 6: Localization Workbench

- [x] Implement Workspace page scan controls and summary.
- [x] Implement Localization Inventory grid.
- [x] Combine inventory, translation grid, focused editor, and placeholder status into one Localization workbench.
- [x] Replace the single-target translation editor with a side-by-side translation matrix. (delivered by Done plan RedConfigure_LocalizationReviewWorkbenchPlan_2026-06-08.md — multi-culture columns; evidence Specs/TestRuns/4cb089111a23/RedConfigure/LocalizationReview/2026-06-08_204926)
- [x] Allow multiple culture columns to be added, removed, reordered, and pinned.
- [x] Resolve direct cell editing as grid activation into the focused multi-culture editor, with edits committed through the session undo unit; this keeps editing accessible without a fragile overlay editor.
- [x] Add a right inspector with focused editor, source, selected target, neighboring cultures, placeholder inspector, accelerator inspector, and resource preview.
- [x] Implement owner/type/status filters. (delivered by Done plan RedConfigure_LocalizationReviewWorkbenchPlan_2026-06-08.md — owner/language check states; evidence Specs/TestRuns/4cb089111a23/RedConfigure/LocalizationReview/2026-06-08_204926)
- [x] Implement translation search. (delivered by Done plan RedConfigure_LocalizationReviewWorkbenchPlan_2026-06-08.md; evidence Specs/TestRuns/4cb089111a23/RedConfigure/LocalizationReview/2026-06-08_204926)
- [x] Implement global resource/language/theme search through the command palette.
- [x] Implement Translation Editor page behavior as the matrix inspector, with source, target, placeholder inspector, and sample arguments.
- [x] Implement `LocalizationExampleControl` for strings, formatted strings, command labels, menus, and compact dialog examples.
- [x] Update examples immediately when the selected translation changes.
- [x] Add keyboard navigation for previous/next missing resource.
- [x] Add spreadsheet-style copy/paste for rectangular translation ranges.
- [x] Add localization batch operations: copy English, copy culture to culture, clear, find/replace, normalize placeholder whitespace, preserve accelerators, and mark reviewed.
- [x] Add before/after review for localization batch operations.
- [x] Add duplicate accelerator warnings for sibling menu items.

### Phase 7: Theme Workbench And Live Preview

- [x] Implement Theme Library page with built-in, file, and user theme sections.
- [x] Implement import, duplicate, reset, and export model operations.
- [x] Implement Theme Designer grouped color key list.
- [x] Combine active theme selection, color editing, expression examples, color status, and live preview into one Theme workbench.
- [x] Replace the first-pass theme workbench with a token table/filter, authored color editor, and live preview stage. The separate navigator/scene-tab enrichment remains tracked below.
- [x] Add token descriptions, source type, usage count, and contrast status to the token table.
- [x] Add the compact preview scene selector for App Shell, Folder View, Menu Popup, Dialogs, File Operations, Monitor Log, and Viewer Diff; it filters/highlights the composite stage without duplicating preview controls.
- [x] Add click-to-select behavior from preview regions to their theme tokens.
- [x] Add affected-region highlighting when a token is selected.
- [x] Implement hex field, swatch, alpha slider, reset override, copy effective color, and copy override color.
- [x] Add color picker affordances for recent colors and copied colors.
- [x] Implement `ThemePreviewModel` so each color edit recomputes the effective preview theme immediately.
- [x] Implement the composite theme preview control with menu, navigation, folder view, file operations, monitor text, diff, and dialog/button samples.
- [x] Repaint preview on every valid color edit without writing to disk.
- [x] Keep the previous valid preview visible when the active edit field contains invalid color text.
- [x] Add contrast/readability indicators for known foreground/background pairs.
- [x] Add internal authored-theme model support for direct colors, references, expressions, effective colors, and dependency validation.
- [x] Add first-pass expression editing through the color value field for solid, reference, and expression values.
- [x] Add expression operations: `ref`, `lighten`, `darken`, `alpha`, `blend`, and `contrast`.
- [x] Add dependency inspector for "depends on" and "affects" relationships.
- [x] Add cycle and missing-reference validation with previous-valid-preview behavior.
- [x] Add mass modification workflows: lighten/darken, blend with accent, set alpha, replace reference, convert solids to references, remove overrides, and apply recipes.
- [x] Add before/after review for theme mass modifications.
- [x] Add theme recipes for dark variant, light variant, accent recolor, softened selections, increased contrast, and semantic status colors.

### Phase 8: Validation And Export

- [x] Implement combined validation summary.
- [x] Block export on localization parse errors, placeholder errors, invalid theme IDs, invalid colors, and output path conflicts.
- [x] Allow export with warnings after the user sees the warning list.
- [x] Implement diff-like file preview before writes.
- [x] Convert the Validation And Export page into a Review & Export basket with one card per output file.
- [x] Keep validation errors visible in the global validation drawer before the user reaches Review & Export.
- [x] Write `.rc` localization output with stable ordering.
- [x] Write `.theme.json5` theme output with stable ordering and grouped comments.
- [x] Replace flattened legacy export with lossless authored version 2 palette/expression export.
- [x] Complete `Specs/Plans/Done/Theme_ExpressivePaletteAndReferencesPlan_2026-07-13.md` for shared durable palette/expression import, Rainbow/runtime-source evaluation, licensed Dracula/Catppuccin distribution, runtime application, and authored export; do not revive the superseded two-map expression draft.
- [x] Preserve UTF-8/UTF-16 expectations for each output type.
- [x] Re-run validation after export and show the written paths.

### Phase 9: Self-Tests And Manual Verification

- [x] Add `RedConfigureTests` parser, writer, discovery, validation, and theme I/O unit tests.
- [x] Add RedSalamander command/self-test coverage for shared `ThemeDefinitionIo` extraction and lossless Preferences operations.
- [x] Add UI smoke coverage for page creation and page switching.
- [x] Add a test that theme preview state changes when `folderView.itemBackgroundSelected` changes. (delivered by Done plan RedConfigure_LocalizationReviewWorkbenchPlan_2026-06-08.md — TestThemePreviewModelKeepsLastValidColor; evidence Specs/TestRuns/4cb089111a23/RedConfigure/LocalizationReview/2026-06-08_204926)
- [x] Add a test that invalid color text does not mutate the previous valid preview color.
- [x] Add a test that formatted translation preview rejects unindexed placeholders. (delivered by Done plan RedConfigure_LocalizationReviewWorkbenchPlan_2026-06-08.md — RedConfigureTests.cpp:583-598; evidence Specs/TestRuns/4cb089111a23/RedConfigure/LocalizationReview/2026-06-08_204926)
- [x] Add a test that `.rc` export contains positional placeholders unchanged.
- [x] Build Debug x64 and ASan Debug x64.
- [x] Run `RedConfigureTests`.
- [x] Run focused existing localization tests through the green full-suite closeout run.
- [x] Run focused existing Preferences Themes tests after shared theme I/O extraction.

### Phase 10: Performance Validation

- [x] Define a perf scenario for scanning a repo-sized resource set and theme folder set.
- [x] Instrument workspace scan time, theme load time, localization parse time, validation time, and preview repaint time.
- [x] Add deterministic perf/selftest coverage for scan plus validation.
- [x] Archive candidate run evidence under `Specs/TestRuns/`.
- [x] Record acceptable targets in `Specs/Core/Core_RedConfigure.md`; initial targets:
  - workspace rescan under 500 ms for current repo resource/theme set on a normal development machine
  - single theme color preview update under one frame budget for the preview control
  - validation under 250 ms for current repo resource/theme set after files are already read

### Phase 11: Specs And Closeout

- [x] Create `Specs/Core/Core_RedConfigure.md`.
- [x] Create `Specs/UI/UI_RedConfigure.md`.
- [x] Update `Specs/Core/Core_Localization.md` with RedConfigure satellite `.rc` generation rules.
- [x] Update `Specs/Core/Core_SettingsStore.md` with lossless authored version 2 RedConfigure import/export rules and the prohibition on flattened legacy export.
- [x] Update `Specs/Testing/Testing_TestCoverage.md` after tests land.
- [x] Add or update `Specs/TestRuns/` evidence references.
- [x] Move this plan to `Specs/Plans/Done/` only after implementation, tests, perf evidence, and authoritative specs are complete.

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

### Closeout Verification (2026-07-14)

- Debug x64 builds passed with zero warnings/errors for `RedConfigure`, `RedConfigureTests`, and `LocalizationTests`.
- `RedConfigureTests` passed in Debug and ASan Debug; `LocalizationTests` passed in Debug.
- The Windows SDK resource compiler test passed, as did the four-page creation/switching smoke test.
- Tools Pester passed 261/261 runnable cases; the inventory records 262 total cases including one intentionally excluded `RequiresBuildToolchain` case.
- `Run-AllTests.ps1 -Suite Full` built the solution and passed 1,141 of 1,200 cases; `RedConfigureTests` and `LocalizationTests` both passed. The overall run was red only for two unrelated File Operations popup timing cases plus the since-fixed Pester inventory count. Both File Operations cases passed immediate isolated reruns, and Pester passed after the inventory correction.
- Repo-sized candidate metrics passed their normative limits: 1,500 source rows, 6,000 target cells, and 10 themes scanned in 320 ms; validation completed in 3 ms; preview work completed in 344 microseconds. Evidence: `Specs/TestRuns/4cb089111a23/RedConfigure/2026-07-14_192100/`.
- `Tools/Test-TestRunArchive.ps1` passed for all three archived evidence files.

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
