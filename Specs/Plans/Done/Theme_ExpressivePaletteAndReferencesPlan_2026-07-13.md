# Expressive Theme Palettes, References, And Dynamic Functions Plan

> **Executor instructions:** Follow this plan in order. Run each verification gate before proceeding. Preserve all pre-existing worktree changes. Stop on any condition in **STOP conditions** rather than inventing a different schema or compatibility policy.
>
> **Drift check (run first):**
>
> ```powershell
> git diff --stat dd4312076..HEAD -- Common RedConfigure RedSalamander RedSalamanderMonitor Specs/SettingsStore.schema.json Specs/Core/Core_SettingsStore.md Specs/Core/Core_RedConfigure.md Specs/Themes Tests/RedConfigureTests Tests/SettingsSchemaTests
> git status --short
> ```
>
> This plan was prepared from a dirty worktree. At plan time, `Common/SettingsStore.h`, `Common/Common/SettingsStore.cpp`, `Common/ThemeDefinitionIo.h`, `Common/Common/ThemeDefinitionIo.cpp`, `Tests/RedConfigureTests/RedConfigureTests.cpp`, and `Specs/Core/Core_SettingsStore.md` already contained unrelated in-progress recovery work. Reconcile and preserve those edits; do not replace them with the `HEAD` versions.

Last updated: 2026-07-14

Status: Done

## Status

- **Priority:** P1
- **Effort:** L
- **Risk:** HIGH
- **Depends on:** completed alongside `Specs/Plans/Done/RedConfigure_LocalizationThemeManagerPlan.md`
- **Category:** feature / architecture / authoring DX
- **Planned at:** commit `dd4312076`, 2026-07-13

### Completion checkpoint (2026-07-14)

The breaking version 2 cutover is complete. The shared authored model, strict JSON5 I/O, static/event/paint expression resolver, runtime dynamic-token allowlist, Rainbow precedence, atomic reload behavior, RedConfigure palette/dependency editing, lossless Preferences operations, six migrated themes, Dracula, Catppuccin Latte/Frappé/Mocha, licensing distribution, schema, mockup, authoritative specifications, golden/accessibility coverage, and deterministic performance instrumentation are implemented.

Closeout evidence:

- Debug solution/full suite: `Run-AllTests.ps1 -Suite Full`, run id `20260714T150136Z-77668-a61447b2b01640a19c32d02ed81a2c74`: 1,147 passed, 0 failed, 53 environment-dependent skips; all 17 suite classifications passed.
- Focused Preferences Themes prefix: 26 passed, 0 failed.
- Focused theme performance candidate: `Specs/TestRuns/4cb089111a23/Commands/2026-07-14_165159`; resolver p50/p95/max 7.58/8.76/11.60 ms and compiled 2,000-evaluation batch p50/p95/max 135/140/407 µs in Debug.
- Same-machine baseline: `Specs/TestRuns/4cb089111a23/Commands/2026-07-14_162005`; resolver p50/p95/max 7.74/9.02/9.90 ms and dynamic batch 147/244/270 µs. Candidate p95 improved for both protected metrics; the candidate dynamic max contains one non-gating outlier.
- ASan Debug: zero-warning application build plus passing `theme_v2_runtime_resolution_and_dynamic_perf`, Preferences save/export, and Preferences load/import cases.
- Focused build/tests: `RedConfigureTests`, `SettingsSchemaTests`, theme distribution/gallery contracts (7/7), resource localization contracts (5/5), and mockup `pnpm build` all passed.
- Public-documentation addendum: `docs/Themes.md`, `docs/Preferences.md`, `docs/UserGuide.md`, `docs/SettingsFile.md`, `docs/DeveloperGuide.md`, `docs/DxUi.md`, and `docs/res/README.md` now describe the version 2 authored model, dynamic functions, Rainbow light/dark behavior, shipped themes, and license locations.
- Gallery addendum: `DxUiTests --suite=Gallery` regenerates deterministic Rainbow Light/Dark plus every built-in and shipped-theme `theme-controls-*.png`; the distribution contract requires an exact, referenced, valid-PNG inventory so new, removed, or renamed themes cannot leave documentation stale.

## Goal

Make themes smaller, easier to maintain, and more expressive through an intentional format cutover. Every persisted or shipped theme uses the new authored palette/expression model; legacy direct-color theme definitions and flattened export are removed rather than maintained in parallel. Theme authors must be able to define a compact named palette, reference palette entries or other semantic theme keys, use a small fixed set of deterministic color functions, and opt supported semantic keys into bounded runtime color sources. RedConfigure must edit and preview the same durable representation consumed by RedSalamander and RedSalamanderMonitor.

The completed work must also:

- migrate the six existing examples under `Specs/Themes/` away from repeated literals;
- add official-palette-based Dracula and Catppuccin examples;
- ship exact upstream license texts and durable notices for every third-party-derived palette;
- preserve `builtin/rainbow` as a first-class dynamic base that version 2 themes can inherit and selectively override;
- support bounded load/event/paint-time functions, including system roles, perceptual tone, contrast repair, palette harmonization, stable seeded choice, and seeded Rainbow, without parsing, allocation, locking, or I/O during paint;
- preserve strict standalone theme-file validation and the settings store's lenient inline-theme recovery policy;
- keep system high contrast authoritative;
- add deterministic correctness, UI integration, and performance evidence;
- update authoritative specs and move this plan to `Specs/Plans/Done/` only at full closeout.

## Why This Matters

Each current example theme assigns 64 semantic color keys directly. Five of the six repeat between 29 and 40 color assignments; `RetroTerminal.theme.json5` repeats 40 assignments and one literal appears 16 times. A color-family change therefore requires synchronized edits across many unrelated-looking keys.

RedConfigure already parses and evaluates `ref`, `lighten`, `darken`, `alpha`, `blend`, and `contrast` in `RedConfigure/Themes/ThemePreviewModel.cpp:128-390`, including cycle rejection and previous-valid-preview behavior. That representation is editor-local, is cleared on theme load, and `RedConfigureSession::BuildThemeExportText` flattens it to direct hex values at `RedConfigure/RedConfigureSession.cpp:1211-1215`. The shared `ThemeDefinition` and runtime still store only `unordered_map<wstring, uint32_t>` values, so authored intent is lost.

Dracula is a useful compact-palette acceptance case: its official OSS palette has a small set of named colors and explicitly reuses the current-line color for selection. Catppuccin is a useful richer case: its four flavors each expose the same 26 named roles, and its style guide maps roles such as Base, Mantle, Surface, Overlay, Text, Green, Yellow, and Red to UI semantics. These sources must guide mappings, not introduce brand-specific branches in the engine.

Official references:

- [Dracula OSS palette and license](https://github.com/dracula/dracula-theme)
- [Dracula exact upstream MIT license text](https://raw.githubusercontent.com/dracula/dracula-theme/master/LICENSE)
- [Dracula home](https://draculatheme.com/)
- [Catppuccin palette](https://github.com/catppuccin/catppuccin)
- [Catppuccin style guide](https://github.com/catppuccin/catppuccin/blob/main/docs/style-guide.md)
- [Catppuccin palette data project](https://github.com/catppuccin/palette)
- [Catppuccin Palette exact upstream MIT license text](https://raw.githubusercontent.com/catppuccin/palette/main/LICENSE)
- [Catppuccin licensing](https://catppuccin.com/licensing/)

## Approved Mockup Contract

The interactive visual contract is `Specs/Mockups/ThemeGalleryWorkbench/` and is grounded in `docs/res/main-window.png`, `docs/res/theme-light.png`, `docs/res/theme-dark.png`, `docs/res/theme-rainbow.png`, and the UI specifications under `Specs/UI/`.

- Use one large, realistic RedSalamander application preview instead of a contact sheet of distant miniature windows.
- A compact theme catalogue switches the full application shell across Built-in Light/Dark, Rainbow Light/Dark, the six existing example themes, Dracula, and the three shipped Catppuccin flavors: Latte, Frappé, and Mocha.
- Do not multiply preview choices by plugin when plugins use the same theme anatomy. Add a separate surface only when it demonstrates a materially different semantic-token contract.
- The preview must preserve the product's title/menu rows, dual NavigationViews and FolderViews, narrow splitter, focused/unfocused pane distinction, pane status rows, and function-key bar.
- Multiple selection is the default Rainbow acceptance state. Several selected items must be visible in both panes, each with its own deterministic stable tint; the current item keeps the stronger focus border and the inactive pane remains visibly subordinate.
- Rainbow Light/Dark, focused pane, single/multiple selection, fixed preview seed, hue phase, and high-contrast fallback are interactive mockup options. Seed and phase changes must visibly and deterministically update the selection colors.
- The inspector contains an interactive function lab for the five additional functions approved in this plan: `seededChoice`, `systemColor`, `perceptualTone`, `ensureContrast`, and `harmonize`. Selecting a function and changing its bounded parameters must visibly update the real application preview; a static list of function names is not sufficient evidence.
- Catppuccin uses one explicit semantic mapping across all three shipped flavors: Lavender is the RedSalamander active/focus accent, Blue is the link/action color, selection fill is Overlay2 composited at 26% over Base, and Green/Yellow/Red retain success/warning/error meaning. Mauve remains an available palette value and may be used by authored recipes, but it is not the default global application accent.
- Dracula and Catppuccin entries must expose their palette-license status, while the shipped implementation still follows the exact license and notice artifact contract in Phase 8.

## Current State

### Data and I/O

- `Common/SettingsStore.h:237-250` defines `ThemeDefinition` as metadata plus direct resolved `colors`; it has no authored palette or expression representation.
- `Common/ThemeDefinitionIo.h:39-51` exposes the one shared parser/writer. The in-progress worktree also adds strict standalone versus lenient inline parse modes; preserve that distinction.
- `Common/Common/ThemeDefinitionIo.cpp:180-435` accepts only hex strings in `colors` and writes only direct colors in sorted key order.
- `Specs/SettingsStore.schema.json:742-856` requires a built-in base and constrains every known color key to `$defs.colorHex`.
- `Specs/Core/Core_SettingsStore.md:543-590` documents only `#RRGGBB` and `#AARRGGBB`; it explicitly says RedConfigure must flatten authoring expressions until durable runtime support lands.

### Authoring

- `RedConfigure/Themes/ThemePreviewModel.h:14-49` owns an editor-only expression enum and map.
- `RedConfigure/Themes/ThemePreviewModel.cpp:128-188` parses the six current function forms; `:295-358` resolves them recursively; `:360-390` rejects invalid edits without replacing the last valid preview.
- `RedConfigure/Themes/ThemePreviewModel.cpp:263-273` flattens expressions before export.
- `RedConfigure/Themes/ThemePreviewModel.cpp:15-48` contains only 26 preview defaults, separate from the complete runtime theme mapping.
- `Specs/Core/Core_RedConfigure.md` makes flattened export the current contract.
- During this plan's preparation, `Specs/Plans/Done/RedConfigure_LocalizationThemeManagerPlan.md` was reconciled to remove its separate `colorExpressions` schema draft and cross-link this plan as the durable owner.

### Runtime consumers

- `RedSalamander/RedSalamander.cpp:2511-2694` derives a base `AppTheme`, looks up direct overrides, and applies them.
- `RedSalamander/SettingsHotReload.cpp:803-957` has another direct-override application path that must receive identical resolved colors.
- `RedSalamanderMonitor/RedSalamanderMonitor.cpp:1623-1803` independently builds its base monitor theme and applies direct monitor overrides.
- `RedSalamander/Preferences.Themes.cpp:736-806` imports, exports, and displays shared `ThemeDefinition` values; its error mapping and editable-value UI assume direct colors.
- Both `RedSalamander.vcxproj` and `RedSalamanderMonitor.vcxproj` copy `Specs/Themes/*.theme.json5` into the build output, so every shipped expressive theme must work in both applications.

### Existing theme evidence

| Theme | Assignments | Distinct literals | Repeated assignments | Highest reuse |
|---|---:|---:|---:|---:|
| Forest Mist | 64 | 30 | 34 | 8 |
| Neon Tokyo | 64 | 35 | 29 | 7 |
| Paper And Ink | 64 | 30 | 34 | 12 |
| Retro Terminal | 64 | 24 | 40 | 16 |
| Solar Flare | 64 | 29 | 35 | 8 |
| Ugly | 64 | 53 | 11 | 3 |

Reproduce this inventory before migration with a small read-only script or test and preserve the result in the implementation notes.

## Architecture Decision

### One authored value per semantic key

Do **not** add the draft `colorExpressions` sibling map. Version 2 themes add a named `palette` object and allow either a direct hex color or an expression string as the value of each `palette` or `colors` entry.

This is the target shape:

```json5
{
  "formatVersion": 2,
  "id": "user/dracula",
  "name": "Dracula",
  "baseThemeId": "builtin/dark",
  "palette": {
    "background": "#282A36",
    "currentLine": "#44475A",
    "foreground": "#F8F8F2",
    "green": "#50FA7B",
    "purple": "#BD93F9",
    "red": "#FF5555"
  },
  "colors": {
    "app.accent": "ref(palette.purple)",
    "window.background": "ref(palette.background)",
    "menu.background": "ref(window.background)",
    "menu.selectionBg": "ref(palette.currentLine)",
    "menu.selectionText": "contrast(menu.selectionBg,palette.foreground,palette.background)",
    "folderView.itemBackgroundSelectedInactive": "alpha(folderView.itemBackgroundSelected,65%)",
    "viewer.diff.addedBackground": "alpha(palette.green,20%)",
    "viewer.diff.removedBackground": "alpha(palette.red,20%)"
  }
}
```

Rules:

- `formatVersion: 2` is required for every standalone and inline theme definition.
- A missing version, `formatVersion: 1`, or a legacy direct-color-only theme shape is rejected with a deterministic diagnostic. There is no legacy reader or flattened writer.
- Version 2 accepts direct hex sources alongside references and functions so authors can use literals where reuse would not improve the theme.
- Unknown future versions are rejected by strict file loading and retained through the existing opaque/forward-compatible inline-settings recovery path where possible.
- `colors` remains semantic RedSalamander token overrides. `palette` contains author-defined reusable values and never becomes an application token by itself.
- A key appears at most once. Duplicate JSON keys are invalid in strict files; lenient inline loading keeps a deterministic valid entry, preserves recoverable authored data, and emits one aggregate recovery diagnostic.

### Expression and runtime-source language

Keep the language deliberately small, closed, and non-general-purpose. Static sources are legal in `palette` and `colors`; runtime sources are legal only in `colors` entries whose semantic token is registered as dynamic-capable.

```text
source         := staticSource | eventSource | paintSource
staticSource   := hex | ref | unary | blend | contrast | perceptualTone
                | ensureContrast | harmonize
eventSource    := systemAccent()
                | systemColor(systemRole)
                | tone(reference, reference)
paintSource    := seededRainbow(runtime.seed, amount, amount, amount, degrees)
                | seededChoice(runtime.seed, reference, reference [, reference]...)
hex            := #RRGGBB | #AARRGGBB
ref            := ref(reference)
unary          := lighten(reference, amount)
                | darken(reference, amount)
                | alpha(reference, amount)
blend          := blend(reference, reference, amount)
contrast       := contrast(reference)
                | contrast(reference, lightCandidate, darkCandidate)
perceptualTone := perceptualTone(reference, toneNumber)
ensureContrast := ensureContrast(foregroundReference, backgroundReference, ratio)
harmonize      := harmonize(reference, targetReference, amount)
reference      := palette.<paletteName> | <semantic.theme.key>
systemRole     := accent | accentLight | accentDark | window | windowText
                | highlight | highlightText
amount         := decimal from 0.0 through 1.0 | percentage from 0% through 100%
toneNumber     := decimal from 0 through 100
ratio          := decimal from 1 through 21
degrees        := decimal from 0 through 360
```

Semantics:

- Function names are accepted ASCII case-insensitively and serialized lowercase.
- Palette names and semantic keys are case-sensitive and use the existing theme-key character policy; palette names must not contain `.` because `palette.` is the namespace separator.
- Expressions are not nested. Composition is expressed by naming another palette entry or semantic key, which keeps parsing, diagnostics, and dependency display simple.
- References are independent of declaration order and may point forward.
- `ref(palette.x)` resolves only the named palette entry. A missing palette entry is an error.
- `ref(app.accent)` first resolves an authored semantic value and then the selected base theme's effective value when no override exists.
- `lighten` and `darken` retain the existing behavior of blending toward opaque white or black.
- `alpha` replaces alpha; it does not composite against a background.
- `blend(a,b,t)` returns `a` at `0`, `b` at `1`, and interpolates all ARGB channels consistently with the current preview behavior.
- One-argument `contrast(background)` chooses opaque black or white using WCAG relative luminance, replacing the preview's current brightness shortcut.
- Three-argument `contrast(background,light,dark)` selects whichever candidate has the higher WCAG contrast ratio against the background. The names describe expected roles, but the resolver must calculate rather than trust them.
- `perceptualTone(color,tone)` replaces the OKLCH lightness with `tone / 100`, preserves alpha and hue, preserves chroma as far as the sRGB gamut permits, and uses one deterministic chroma-reduction/gamut-mapping algorithm shared by runtime and RedConfigure. It is load-time unless one of its references resolves at event time; it is never paint-time.
- `ensureContrast(foreground,background,ratio)` adjusts only the foreground along an OKLCH lightness path first,
  then reduces chroma only when needed for in-gamut output. Contrast is measured from the foreground composited at
  its authored alpha over an opaque background. The nearest candidate meeting the requested WCAG ratio preserves
  foreground alpha; a translucent background or a target unattainable at that alpha is a validation error. An
  event-time system-color change that makes the target unattainable uses the compiled highest-contrast fallback and
  reports the unmet state to RedConfigure without logging from paint.
- `harmonize(color,target,amount)` moves the source OKLCH hue toward the target over the shortest hue arc, preserves source lightness and alpha, and preserves source chroma subject to deterministic sRGB gamut mapping. `0` returns the source and `1` aligns its hue with the target. It is load-time unless a dependency is event-time; it is never paint-time.
- `systemColor(role)` resolves one closed Windows color role and is re-evaluated only on the existing system color/theme-change path. `systemAccent()` remains a supported, round-trip-preserved compatibility spelling of `systemColor(accent)` rather than being silently rewritten during serialization.
- `tone(lightReference,darkReference)` selects the first reference for an effective light base and the second for an effective dark base. The effective base follows Rainbow's current system light/dark choice and is not inferred from arbitrary color luminance.
- `seededRainbow(runtime.seed,saturation,value,alpha,phaseDegrees)` is a compiled paint-time source. It derives hue from a caller-supplied stable 32-bit seed, adds the normalized phase, and converts HSV to ARGB using the bounded authored parameters.
- `seededChoice(runtime.seed,...)` is a compiled paint-time source over 2 to 8 referenced colors. All candidates resolve before paint; paint maps the caller-supplied stable 32-bit seed to one candidate with the shared stable hash/modulo contract. Candidate order is authored and significant, the same seed and candidate list always produce the same ARGB, and changing an unrelated theme key cannot perturb the result.
- `runtime.seed` is a reserved context input, not a theme reference. Callers pass an already-computed stable hash; the evaluator never reads item text, paths, process state, or object identity.
- A paint-time source is terminal: another palette or semantic source may not reference it or derive from it. This prevents a dependency graph from entering the paint path; reuse the same authored parameters on another allowlisted token when matching dynamics are required. Event-time `systemAccent()`, `systemColor(...)`, and `tone(...)` results may be referenced normally after event-time resolution, and load/event-derived functions inherit the latest evaluation phase of their references.
- Runtime sources compile once at theme load/edit time into a closed tagged representation. Paint-time evaluation is `noexcept`, `O(1)`, and performs no parsing, allocation, locking, logging, file access, registry access, system API call, or dependency-graph traversal.
- Initially, paint-time sources are allowed only for the existing centralized Rainbow surfaces: supported folder/grid/tree/menu selection backgrounds and navigation/path accents. Both `seededRainbow(...)` and `seededChoice(...)` use this registry. Adding a dynamic-capable token requires an explicit registry entry, fallback color, runtime-context contract, tests, and performance evidence.
- Resolution is deterministic, cached per theme evaluation, and `O(nodes + dependency edges)`.
- Cycles report the full dependency path. Missing references report both the authored key and missing target.
- Enforce bounded input: maximum 128 palette entries, 512 semantic entries, 256 UTF-16 code units per expression, and dependency depth 32. Exceeding a bound is a validation error, not truncation.
- No file includes, environment variables, arbitrary arithmetic, scripts, nondeterministic random values, current time, network input, or dynamic function registration. Clock-driven functions such as `pulse`, `time`, animated hue rotation, and noise are deferred until there is one animation clock, invalidation/throttling policy, reduced-motion behavior, and measured paint evidence.

Approved additional functions, in the requested order:

1. `seededChoice(runtime.seed, ...)` — paint-time, deterministic selection from 2 to 8 pre-resolved palette references.
2. `systemColor(role)` — event-time access to one closed Windows system-color role.
3. `perceptualTone(color, tone)` — load/event-time OKLCH tone control with deterministic sRGB gamut mapping.
4. `ensureContrast(foreground, background, ratio)` — load/event-time foreground repair against a WCAG contrast target with explicit unattainable-target handling.
5. `harmonize(color, target, amount)` — load/event-time bounded hue harmonization over the shortest OKLCH hue arc.

### Authored and resolved models stay separate

Move parsing and evaluation from RedConfigure into `Common`:

- `ThemeColorSource`: variant of direct ARGB, parsed fixed expression, or one of the closed runtime-source forms.
- `ThemeDefinition`: metadata, version, authored palette sources, authored semantic sources, plus any opaque recovery data needed by the settings store.
- `ResolvedThemeColors`: resolved static semantic `key -> 0xAARRGGBB`, compiled dynamic programs keyed by supported semantic token, diagnostics, dependency edges, and reverse-dependency edges.
- `ThemeResolutionContext`: immutable event-time inputs supplied by the host: effective light/dark base, the closed `systemColor` role table, and system-high-contrast state.
- `ThemeRuntimeContext`: the minimal paint-time value context needed by compiled programs: stable `seedHash32` and system-high-contrast state. Do not put strings, handles, settings stores, clocks, or mutable host objects in this context.
- `ResolveThemeColors(...)`: accepts an immutable base semantic-color lookup plus `ThemeResolutionContext` and resolves/compiles the authored graph atomically.
- `EvaluateDynamicThemeColor(...)`: evaluates one already-compiled dynamic program against `ThemeRuntimeContext`, returning its pre-resolved fallback when high contrast suppresses dynamic color or the token/context is unavailable.

Do not store resolved results back into the authored maps. Serialization must reproduce authored literals/references/functions, while application code consumes only `ResolvedThemeColors`.

### Rainbow compatibility and precedence

`builtin/rainbow` remains a first-class dynamic base; do not flatten it to a static palette or reimplement it as a special third-party theme. A version 2 theme may declare `"baseThemeId": "builtin/rainbow"`, inherit its existing system-light/system-dark base selection, and override only selected tokens.

Resolve each semantic token in this order:

1. Windows Contrast Themes suppress authored and inherited dynamic color and remain authoritative.
2. An explicit authored source for that token wins, whether static, event-time, or paint-time dynamic.
3. If no authored source exists and the base is `builtin/rainbow`, inherit the built-in Rainbow dynamic policy for that token.
4. Otherwise use the selected built-in/static base value.

A static override suppresses Rainbow only for that semantic token; it must not turn off Rainbow globally. Existing `rainbowMode` fields remain current runtime outputs for consumers that actively use Rainbow behavior, not a legacy theme-format adapter. A non-Rainbow base may opt individual allowlisted tokens into `seededRainbow(...)` without setting those plugin-wide flags. Only inheritance from `builtin/rainbow` enables the established plugin-wide Rainbow contract.

Centralize the stable seeded color calculation and compiled-program evaluation in shared theme code, then migrate existing `AppTheme`/DxUi call sites without changing their stable-seed inputs or plugin ABI. The same seed and context must always produce the same ARGB value; changing unrelated theme keys must not perturb it.

### Failure policy

- Standalone `Themes/*.theme.json5` loading is strict: any malformed v2 source, duplicate, missing reference, cycle, unsupported function, or limit violation rejects that file and logs one useful diagnostic.
- Inline settings loading is lenient: retain the settings document, keep valid authored entries, fall back to the selected base color for invalid/affected semantic entries, preserve opaque recoverable data, and issue one aggregate recovery warning. Do not fail all settings because one custom theme is damaged.
- RedConfigure blocks v2 export while any dependency or contrast validation error remains and keeps the previous valid preview active during an invalid edit.
- A runtime source on a non-allowlisted token is a validation error. Missing runtime context at paint time is not reparsed or logged; it uses the compiled token fallback and is covered by deterministic tests.
- Runtime theme application is atomic. Never partially replace the currently displayed theme during a hot reload; resolve first, then apply the complete valid result or keep/fall back according to the strict/lenient source policy.
- Windows Contrast Themes remain authoritative. Expression results must not bypass the current system-high-contrast override path.

## Scope

### In scope

- Shared authored theme model, parser, serializer, evaluator, diagnostics, and dependency graph in `Common`.
- One required version 2 schema with no legacy theme-format compatibility path.
- RedSalamander startup, selection, Preferences, and hot-reload integration.
- RedSalamanderMonitor integration.
- RedConfigure durable import/edit/preview/export and dependency inspection.
- First-class custom-theme inheritance from `builtin/rainbow`, plus deterministic preview seeds and per-token dynamic/static inspection.
- Closed runtime sources `systemAccent()`, `systemColor(...)`, `tone(...)`, `perceptualTone(...)`, `ensureContrast(...)`, `harmonize(...)`, and allowlisted `seededRainbow(...)` / `seededChoice(...)` with shared compilation/evaluation.
- Migration of every current `Specs/Themes/*.theme.json5` example.
- New Dracula and three Catppuccin flavor examples: Latte, Frappé, and Mocha.
- Exact upstream Dracula/Catppuccin license texts, third-party notices, source snapshot metadata, and distribution-copy verification.
- Schema, specs, localization resources for new diagnostics/UI labels, correctness tests, UI selftests, and perf evidence.

### Out of scope

- General scripting or user-defined functions.
- Cross-file includes, imports, or inheritance between user themes.
- Remote palette downloads or automatic updates.
- Gradient, typography, spacing, icon, or layout tokens.
- Clock-driven animation, elapsed-time functions, nondeterministic random/noise functions, and arbitrary runtime host callbacks.
- Replacing `AppTheme` or `ColorTextView::Theme` with a fully generic runtime styling system.
- Shipping premium Dracula variants or copying assets outside the official OSS palette.
- Changing Windows high-contrast precedence.

## File Map

Expected files are listed explicitly so scope drift is visible. If implementation needs another file, update this plan before editing it.

### Shared model and persistence

- Modify `Common/SettingsStore.h` and `Common/Common/SettingsStore.cpp`.
- Modify `Common/ThemeDefinitionIo.h` and `Common/Common/ThemeDefinitionIo.cpp`.
- Create `Common/ThemeExpression.h` and `Common/Common/ThemeExpression.cpp` (names may change only to match an existing Common naming convention).
- Modify `Common/Common.vcxproj` and `Common/Common.vcxproj.filters` for new compilation units.

### Runtime consumers

- Modify `RedSalamander/AppTheme.h` and `RedSalamander/AppTheme.cpp` for shared base-color/apply adapters and migration of the existing stable Rainbow calculation.
- Modify `RedSalamander/RedSalamander.cpp` and `RedSalamander/SettingsHotReload.cpp`.
- Modify `RedSalamander/Preferences.Themes.cpp`, `RedSalamander/Preferences.Dialog.cpp`, and the smallest necessary Preferences header/state file.
- Modify `RedSalamanderMonitor/RedSalamanderMonitor.cpp`.
- Modify only the existing dynamic-theme call sites in `Common/DxUi/DxUi.Theme.cpp`, `Common/DxUi/DxUi.Grid.cpp`, `Common/DxUi/DxUi.Tree.cpp`, `Common/DxUi/DxUi.Menu.cpp`, and their smallest shared internal header when required to consume compiled dynamic colors. Do not redesign rendering controls.
- Modify `RedSalamander/RedSalamander.rc`, `RedSalamander/Resource.h`, and every existing `RedSalamander/Lang/*/RedSalamander-*.rc` satellite only when new localized diagnostics or labels are required.

### RedConfigure

- Modify `RedConfigure/Themes/ThemePreviewModel.h`, `RedConfigure/Themes/ThemePreviewModel.cpp`, `RedConfigure/Themes/ThemeCatalog.h`, and `RedConfigure/Themes/ThemeCatalog.cpp`.
- Modify `RedConfigure/RedConfigureSession.h`, `RedConfigure/RedConfigureSession.cpp`, `RedConfigure/RedConfigureRoot.h`, and `RedConfigure/RedConfigureRoot.cpp` for authored export and dependency UI.
- Modify `RedConfigure/RedConfigure.rc`, `RedConfigure/resource.h`, and every existing `RedConfigure/Lang/*/RedConfigure-*.rc` satellite for localized UI text.

### Contract, examples, and tests

- Modify `Specs/SettingsStore.schema.json`, `Specs/Core/Core_SettingsStore.md`, `Specs/Core/Core_RedConfigure.md`, `Specs/UI/UI_RedConfigure.md`, and `.github/skills/theming/SKILL.md`.
- Modify all current `Specs/Themes/*.theme.json5`; create Dracula and Catppuccin theme files.
- Create `Specs/Themes/Licenses/Dracula.LICENSE.txt` and `Specs/Themes/Licenses/Catppuccin.LICENSE.txt` from the exact upstream MIT license texts, and create `Specs/Themes/THIRD-PARTY-NOTICES.md` with source URL, pinned upstream commit/tag, copyright holder/year, palette names, local filenames, and adaptation notes.
- Modify root `LICENSE.txt` to list Dracula Theme and Catppuccin Palette in its existing third-party section, without changing the RedSalamander license grant.
- Modify `RedSalamander/RedSalamander.vcxproj` and `RedSalamanderMonitor/RedSalamanderMonitor.vcxproj` so theme license/notice artifacts are copied beside the shipped theme directory in build/distribution outputs.
- Modify `Tests/RedConfigureTests/RedConfigureTests.cpp`, `Tests/SettingsSchemaTests/SettingsSchemaTests.cpp`, and the smallest applicable `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.*.cpp` / `Commands.SelfTest.Settings.cpp` files.
- Modify the smallest applicable DxUi/selftest source to cover stable seeded runtime evaluation, high-contrast suppression, and fallback without adding a second evaluator.
- Modify `Tests/PerformanceTests2/PerformanceTests2.cpp` only if the protected resolution scenario cannot be measured deterministically in the existing command/RedConfigure harnesses.
- Create archived evidence only through the repository test/perf workflow under `Specs/TestRuns/`.
- Modify the then-WIP RedConfigure localization/theme manager plan for ownership reconciliation, then move both plans to `Specs/Plans/Done/` at their respective closeouts.

Explicitly do not touch unrelated viewer, file-system, search, or DxUi rendering code. Changes to DxUi are limited to existing Rainbow/dynamic-color consumption through the new shared boundary; controls must not be individually redesigned.

## Commands You Will Need

| Purpose | Command | Expected result |
|---|---|---|
| Build shared/runtime graph | `.\build.ps1 -ProjectName RedSalamander -Configuration Debug` | exit 0; main app, monitor, plugins, and copied themes build |
| Build RedConfigure | `.\build.ps1 -ProjectName RedConfigure -Configuration Debug` | exit 0 |
| Build focused tests | `.\build.ps1 -ProjectName RedConfigureTests -Configuration Debug` | exit 0 |
| Run focused shared/authoring tests | `.\.build\x64\Debug\RedConfigureTests.exe` | exit 0; all cases pass |
| Build schema tests | `.\build.ps1 -ProjectName SettingsSchemaTests -Configuration Debug` | exit 0 |
| Run schema tests | `.\.build\x64\Debug\SettingsSchemaTests.exe` | exit 0; real schema loads and field model passes |
| Focused Preferences theme tests | `.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_preferences_dialog_themes_ --selftest-fail-fast --selftest-timeout-multiplier=2` | exit 0; all matching cases pass |
| Full closeout | `.\Tools\Run-AllTests.ps1 -Suite Full` | exit 0; archived run created |

Do not run builds or tests concurrently; the repository serializes artifact operations.

## Git Workflow

- Suggested branch: `codex/theme-expressive-palette`.
- Preserve the dirty worktree changes called out in the drift note; do not reset, checkout, or overwrite them.
- Commit by logical phase with concise imperative messages matching current repository history.
- Do not push or open a pull request unless the operator explicitly requests it.

## Suggested Executor Toolkit

- Read and follow `.github/skills/theming/SKILL.md`, `.github/skills/yyjson/SKILL.md`, `.github/skills/perf-validation/SKILL.md`, `.github/skills/cpp-modern-style/SKILL.md`, `.github/skills/error-handling/SKILL.md`, and `.github/skills/cpp-build/SKILL.md` before implementation.
- Use WIL RAII for every yyjson document and emitted buffer.
- Use `std::optional::value()` / `value_or()`, never unary `*` on `std::optional`.
- No `catch (...)`, raw owning COM pointers, manual handle cleanup, or hardcoded user-facing C++ strings.

## Implementation Steps

### Phase 1: Lock the breaking-cutover contract with characterization tests

- [x] Run the current `RedConfigureTests` and `SettingsSchemaTests` before changing the shared model; both passed on 2026-07-14.
- [x] Replace legacy parse/write expectations with explicit rejection tests for missing `formatVersion`, version 1, and the old direct-color-only shape. Keep direct `#RRGGBB` and `#AARRGGBB` as valid version 2 source forms.
- [x] Add the version 2 fixture corpus for the target JSON5 shape, forward references, palette references, semantic references, all static functions, all closed event/paint functions, candidate-aware contrast, allowlisted/non-allowlisted runtime tokens, cycles, missing references, duplicate keys, limits, and unknown versions. Include explicit fixtures for the five additions (`seededChoice`, `systemColor`, `perceptualTone`, `ensureContrast`, and `harmonize`). Wire each fixture into executable assertions in Phases 2-3; do not commit a deliberately failing focused test binary.
- [x] Add parse -> serialize -> parse authored-losslessness assertions for the single required version 2 representation. The serializer must always emit `formatVersion: 2` and must never expose a flattened/legacy mode.
- [x] Record the existing shipped-theme repetition inventory shown above in a test/helper that can be rerun after migration.

**Verify:** build and run `RedConfigureTests`; baseline tests stay green and the new contract tests prove old theme formats are rejected rather than silently migrated.

### Phase 2: Add the shared authored model and evaluator

- [x] Add focused shared files such as `Common/ThemeExpression.h` and `Common/Common/ThemeExpression.cpp`; register them in `Common/Common.vcxproj` and `.filters`.
- [x] Move the expression enum/AST, parser, canonical formatter, ARGB/HSV operations, WCAG luminance/contrast calculation, graph resolver, closed runtime-program compiler/evaluator, diagnostics, bounds, dependency edges, and reverse edges into `Common`.
- [x] Replace `ThemeDefinition.colors` as the only source of truth with separate authored palette and semantic sources. Update all callers directly; do not add a legacy map, compatibility helper, or dual representation.
- [x] Make the evaluator accept base semantic colors from the host so references to non-overridden semantic keys resolve against the selected built-in theme.
- [x] Cache each node result within one resolution and detect cycles with explicit unvisited/visiting/resolved state rather than repeated stack scans.
- [x] Treat `std::bad_alloc` as fatal and use non-throwing validation paths at exported/noexcept boundaries.

**Verify:** `RedConfigureTests` passes every pure parse, format, graph, function, contrast, cycle, missing-reference, and bounds case.

### Phase 3: Extend shared JSON5 I/O and settings persistence

- [x] Extend `ThemeDefinitionIoError` with actionable version, palette, expression, duplicate, reference, cycle, and limit errors without exposing parser internals to UI callers.
- [x] Parse only version 2 authored sources. Use the current `JsonValue` path and preserve strict standalone versus recoverable inline-settings behavior, but treat old theme formats as invalid entries rather than converting them.
- [x] Serialize deterministic field order: `formatVersion` when needed, `id`, `name`, `baseThemeId`, `palette`, then `colors`; sort palette and semantic keys ordinally.
- [x] Use yyjson copying APIs for dynamic keys/values exactly as required by `.github/skills/yyjson/SKILL.md`.
- [x] Update settings load/save so inline v2 themes survive round trips, including the current opaque-entry recovery behavior.
- [x] Update equality/change-detection code such as `AreEquivalentThemeDefinition` so palette and expressions participate in Apply/dirty decisions.
- [x] Verify file themes and inline themes use the same parser and evaluator while retaining their different failure policies.

**Verify:** `RedConfigureTests` passes strict-file, lenient-inline, opaque-recovery, deterministic export, and authored round-trip tests.

### Phase 4: Integrate atomic resolution in every runtime path

- [x] Add one reusable app-theme adapter that exposes the selected built-in `AppTheme` as semantic base colors and applies a resolved semantic map. Reuse it from startup/selection and `SettingsHotReload.cpp` to remove their current override drift.
- [x] Resolve the custom theme before calculating the accent-dependent final `AppTheme`; ensure an expression-derived `app.accent` participates exactly as a direct accent does today.
- [x] Update RedSalamanderMonitor to resolve the same authored graph against its monitor base values before `ApplyMonitorThemeOverrides`.
- [x] Keep system high contrast above custom theme resolution in both applications.
- [x] Re-resolve event-time `systemAccent()` and `tone(...)` sources on the existing system accent/color/theme change notifications; do not poll the system or query it during paint.
- [x] Ensure file-theme reload, settings hot reload, theme cycling, and Preferences preview apply atomically and keep a valid current theme on resolution failure.
- [x] Add one aggregate `Debug::Error` for rejected strict themes and one aggregate `Debug::Warning` for recovered inline themes; do not log normal base fallback or successful resolution.

**Verify:** targeted runtime selftests cover version 2 direct sources, palette/ref, expression-derived accent, a transitive dependency, invalid hot reload preserving the live theme, Monitor consumption, and high-contrast precedence.

### Phase 5: Preserve Rainbow and add bounded runtime evaluation

- [x] Characterize current `builtin/rainbow` output before refactoring: effective light/dark base selection, stable hash inputs, per-surface colors, plugin/viewer `rainbowMode` flags, and system-high-contrast suppression.
- [x] Move the stable seeded hue/HSV calculation behind the shared compiled runtime evaluator while preserving existing seed inputs and effective ARGB outputs. Keep compatibility wrappers only where staged migration requires them.
- [x] Add one explicit registry of dynamic-capable semantic tokens with each token's required context and pre-resolved fallback. Reject `seededRainbow(...)` on every other token during validation.
- [x] Implement precedence per token: high contrast, explicit authored source, inherited built-in Rainbow source, static/base fallback. Prove that one static override suppresses Rainbow only for that token.
- [x] Allow `baseThemeId: "builtin/rainbow"` in strict v2 files and inline settings. Preserve established plugin-wide Rainbow flags only for this base; token-level dynamic sources on another base must not opt unrelated plugin/viewer surfaces into Rainbow.
- [x] Add `systemAccent()`, `systemColor(...)`, and `tone(...)` event-time re-resolution; add load/event evaluation for `perceptualTone(...)`, `ensureContrast(...)`, and `harmonize(...)`; and add `seededRainbow(...)` / `seededChoice(...)` paint-time evaluation with the no-parse/no-allocation/no-lock/no-I/O contract.
- [x] Add deterministic runtime tests: same seed/context gives the same color, representative distinct seeds vary, `seededChoice` respects authored candidate order and 2-to-8 bounds, phase is stable, light/dark tone selection is correct, every `systemColor` role refreshes only on the host event, perceptual gamut mapping is stable, contrast targets and unattainable fallbacks are correct, harmonization uses the shortest hue arc, high contrast returns fallback, and non-Rainbow themes remain unchanged.
- [x] Add focused teardown/hot-reload coverage proving compiled programs are immutable values owned by the atomically applied theme and no paint call can observe a partially replaced program.

**Verify:** focused shared/runtime tests and existing Rainbow selftests pass with output equivalence for characterized built-in cases; instrumentation confirms zero paint-time allocations and no runtime parser/system calls.

### Phase 6: Make RedConfigure use the shared durable model

- [x] Remove the private parser/evaluator from `ThemePreviewModel`; wrap the shared authored model and dependency result instead.
- [x] Stop clearing expressions during `SetTheme`; make `BuildThemeExportText` serialize the authored version 2 model directly.
- [x] Remove every flattened/legacy theme export path and associated UI/test wording.
- [x] Let the authored-value editor edit both palette entries and semantic keys. Keep one text field/value model for literal or expression sources.
- [x] Add localized source badges (`Base`, `Literal`, `Palette reference`, `Token reference`, `Function`, `Fallback`) and dependency inspector sections (`Depends on`, `Affects`, `Cycle path`).
- [x] Add `Dynamic`/`Inherited Rainbow` badges, show the selected runtime function and fallback, and provide a deterministic preview-seed chooser with several fixed representative seeds. Preview must never use wall-clock or random seed values.
- [x] Add a focused function-lab preview for `seededChoice`, `systemColor`, `perceptualTone`, `ensureContrast`, and `harmonize`, with bounded controls, authored expression, resolved swatches, evaluation phase, affected semantic token, and an immediately visible effect in the real application preview.
- [x] Keep previous-valid-preview behavior and add actionable inline errors for missing palette entry, missing semantic token, invalid amount, unsupported function, cycle, and depth/size limits.
- [x] Replace the partial 26-entry preview-default divergence with the same semantic base lookup contract used by runtime adapters, or document and test any intentionally preview-only tokens.
- [x] Add palette operations: create, rename with reference rewrite, delete with affected-key preview, and convert repeated literal to named palette entry.
- [x] Make batch edits create references/functions where appropriate instead of eagerly flattening them.

**Verify:** RedConfigure tests prove import/edit/preview/export/reopen preserves authored expressions and effective colors, no flattened export API remains, and palette rename/delete updates or blocks dependents correctly.

### Phase 7: Update Preferences theme editing

- [x] Update Preferences display/import/export to recognize v2 authored values and show effective swatches.
- [x] Do not build a second full expression designer in Preferences. For v2 entries, allow safe selection/import/export and direct authored-value editing only if the existing page can do so without ambiguity; otherwise make advanced editing explicitly RedConfigure-owned and keep Preferences lossless.
- [x] Map new shared I/O errors to localized resource strings using positional placeholders.
- [x] Ensure duplicate/reset/save/import/export paths preserve palette and expression sources.

**Verify:** existing `cmd_preferences_dialog_themes_` cases pass plus new v2 import, selection, preview, export, reopen, duplicate, and reset cases.

### Phase 8: Migrate existing themes and add Dracula/Catppuccin

- [x] Convert all six existing examples to `formatVersion: 2` with a compact named palette and references/functions. Preserve their effective ARGB value for every existing semantic key unless an intentional contrast correction is documented.
- [x] Add a golden test that resolves each migrated version 2 fixture and compares all 64 effective semantic values with the captured pre-cutover constants, without retaining a legacy parser fixture.
- [x] Add `Dracula.theme.json5` from the official OSS palette. Use palette names from the source and RedSalamander semantic mappings guided by the official accessibility/specification guidance.
- [x] Add Catppuccin Latte, Frappé, and Mocha themes. Keep the same semantic mapping across flavors and change only the official 26-value palette plus documented flavor-specific base choice.
- [x] Follow the Catppuccin style guide for Base/Mantle/Crust backgrounds, Surface/Overlay states, Text/Subtext, selection opacity, and Green/Yellow/Red status semantics.
- [x] Encode and golden-test the approved Catppuccin application mapping: Lavender -> active/focus accent and pane border, Blue -> links/actions, Overlay2 at 26% over Base -> selection fill, Base/Mantle/Crust -> content/chrome layers, Text/Subtext -> primary/secondary text, and Green/Yellow/Red -> success/warning/error. Keep Mauve in the palette but do not map it to the default global accent. Document that Mauve is a common configurable default in Catppuccin integrations, not a universal style-guide requirement.
- [x] Pin the exact upstream Dracula and Catppuccin palette commit/tag used for generation. Record it with source URLs, copyright/year, covered palette names/files, and adaptation notes in `Specs/Themes/THIRD-PARTY-NOTICES.md`.
- [x] Copy the exact upstream MIT license text into `Specs/Themes/Licenses/Dracula.LICENSE.txt` and `Specs/Themes/Licenses/Catppuccin.LICENSE.txt`; do not paraphrase or merge the two notices.
- [x] Add concise SPDX/license, source URL, and pinned snapshot comments to each third-party-derived `.theme.json5`. Update root `LICENSE.txt` third-party entries and preserve the upstream notices in distributed outputs.
- [x] Add a test/build contract that fails when a third-party-derived theme lacks its declared notice/license file or when those files are not copied with shipped themes. Do not include Dracula PRO values or non-OSS assets.
- [x] Validate every shipped theme in both main-app and Monitor base contexts.
- [x] Add a source contract ensuring migrated `colors` blocks do not repeat raw hex literals; reusable literals belong in `palette` and derived alpha/blends use functions.

**Verify:** all theme files parse strictly, resolve without diagnostics, produce nonempty required foreground/background contrasts, and are copied into both Debug output theme directories after the relevant builds; exact Dracula/Catppuccin license files and `THIRD-PARTY-NOTICES.md` are present in both outputs.

### Phase 9: Schema, documentation, and examples

- [x] Add a required `formatVersion: 2`, bounded palette-name pattern, and bounded `$defs.themeColorSource` expression shape to `Specs/SettingsStore.schema.json`.
- [x] Make known `colors` properties accept `themeColorSource`; make `palette` accept bounded named `themeColorSource` entries.
- [x] Document that schema regex validation is only a coarse editor aid; the shared parser owns exact arity, reference, graph, and bounds validation.
- [x] Reconcile `Specs/Plans/Done/RedConfigure_LocalizationThemeManagerPlan.md`: remove the conflicting `colorExpressions` schema draft, link this plan, and update checklist ownership (completed during plan preparation on 2026-07-13).
- [x] Update `Specs/Core/Core_SettingsStore.md` with versioning, static/runtime grammar, per-token precedence, Rainbow inheritance, runtime allowlisting/context, strict/lenient behavior, deterministic serialization, and compatibility.
- [x] Update `Specs/Core/Core_RedConfigure.md` and `Specs/UI/UI_RedConfigure.md` with durable authored export, palette/dependency UI, deterministic dynamic preview, and the absence of any flattened export.
- [x] Update `.github/skills/theming/SKILL.md` so future token additions include v2 schema/evaluator and shipped-theme validation.
- [x] Update public user/developer documentation and regenerate the exact `theme-controls-*.png` inventory, including both Rainbow bases and all Dracula/Catppuccin themes.

**Verify:** `SettingsSchemaTests.exe` and relevant documentation-drift Pester tests pass.

### Phase 10: Performance validation

- [x] Define two protected scenarios before optimizing:
  - runtime resolution of a Catppuccin-sized theme during startup/hot reload;
  - RedConfigure edit of a high-fan-out palette entry followed by dependency recompute and preview repaint.
- [x] Add a third protected scenario: sustained paint-time evaluation of the worst allowlisted `seededRainbow` and 8-candidate `seededChoice` surface set across fixed seeds, including high-contrast fallback and atomic theme replacement.
- [x] Reuse `App.Startup.LoadThemeDefinitions` and `redconfigure.theme_catalog.load_us`; add `theme.resolve_us`, `theme.resolve.node_count`, `theme.resolve.edge_count`, `theme.dynamic.evaluate_us`, `theme.dynamic.evaluate_count`, and `redconfigure.theme.preview_resolve_us` only if existing metrics cannot answer the scenarios.
- [x] Add deterministic selftest coverage with a fixed 128-palette/512-semantic worst-allowed graph, a fan-out graph, and a depth-32 chain.
- [x] Set budgets from a measured same-machine baseline. As an initial guard, a normal 26-palette/64-semantic theme resolution plus RedConfigure model update must remain comfortably under a 16.67 ms frame budget in Release; do not encode a tighter budget without evidence.
- [x] Set a measured per-evaluation budget for compiled dynamic colors and assert zero allocations in the measured paint loop. Report p50/p95/max or the repository-standard equivalent, not only an average.
- [x] Archive baseline and candidate evidence under `Specs/TestRuns/<MachineHash>/...` with `results.json`, `trace.txt`, and `perf_metrics.jsonl` intact.
- [x] Report same-machine baseline/candidate values, node/edge counts, and any regression honestly.

**Verify:** focused perf selftest and `.\Tools\Run-AllTests.ps1 -Suite Full` pass; the new archived run contains the required metrics.

### Phase 11: Closeout

- [x] Run Debug x64 focused builds/tests, ASan Debug x64 theme/Preferences coverage, and the full suite.
- [x] Confirm `git diff --check` is clean and no non-plan pre-existing user edits were lost.
- [x] Confirm every current and new shipped theme loads in RedSalamander and RedSalamanderMonitor and can be opened/exported/reopened in RedConfigure.
- [x] Confirm `builtin/rainbow`, a v2 theme based on Rainbow, and a non-Rainbow theme with one allowlisted `seededRainbow(...)` token all preserve the documented precedence and compatibility behavior.
- [x] Confirm exact third-party license/notice artifacts are present in the source tree and both application distribution outputs.
- [x] Confirm missing-version/version-1 themes are rejected consistently and every writer emits only the authored version 2 shape.
- [x] Merge all lasting requirements into authoritative specs and the theming skill.
- [x] Update the umbrella RedConfigure plan's remaining checklist.
- [x] Move this file to `Specs/Plans/Done/Theme_ExpressivePaletteAndReferencesPlan_2026-07-13.md` only after all correctness, accessibility, performance, and spec gates pass.

## Test Matrix

### Parser and serializer

- Missing-version/version-1/legacy-shape rejection with deterministic diagnostics.
- Version 2 direct and alpha colors, palette literals, palette expressions, semantic expressions, forward/back references, canonical ordering, and exact authored round trip.
- Event-time `systemAccent()`/`systemColor(...)`/`tone(...)`, load-or-event `perceptualTone(...)`/`ensureContrast(...)`/`harmonize(...)`, paint-time `seededRainbow(...)`/`seededChoice(...)`, canonical formatting, allowlisted-token enforcement, and compiled fallback retention.
- Percent and decimal amounts at 0 and 1 boundaries.
- All functions, candidate-aware contrast, alpha preservation, and ARGB rounding.
- Invalid syntax, arity, amount, key, palette name, duplicate, unknown version, missing ref, cycle path, depth, count, and length.
- Strict file rejection versus lenient inline recovery and opaque preservation.

### Resolver and runtime

- Authored semantic value beats base lookup.
- A semantic reference falls back to the selected base when the target is not authored.
- Palette namespace never collides with semantic keys.
- Resolution order does not depend on JSON member order.
- Expression-derived accent affects the same downstream surfaces as a direct accent.
- Invalid hot reload does not partially apply.
- System high contrast wins in main app and Monitor.
- Monitor sees the same resolved monitor keys as the main theme definition.
- A v2 theme can inherit `builtin/rainbow`; an explicit static override disables dynamic color for only that token.
- Authored dynamic source beats inherited Rainbow for the same token; inherited Rainbow beats the static base when no authored source exists.
- Same stable seed and context produce the same ARGB; representative different seeds produce the expected stable hue spread without relying on nondeterministic assertions.
- `seededChoice` accepts 2 to 8 pre-resolved candidates, rejects other arities, preserves authored candidate order, and selects the same candidate for the same seed across reloads.
- A non-Rainbow base can use one allowlisted seeded token without enabling plugin-wide `rainbowMode`; Rainbow inheritance preserves existing plugin/viewer flags.
- High contrast and missing runtime context use the compiled fallback without parsing, allocation, logging, or system calls during paint.
- System role/tone changes re-resolve on the existing host event and remain stable between events; `systemAccent()` and `systemColor(accent)` resolve equivalently while preserving their authored spelling.
- Perceptual tone and harmonization golden vectors cover hue wrap, chroma reduction, alpha preservation, and sRGB gamut edges; contrast vectors cover 3.0, 4.5, 7.0, and unattainable targets.

### RedConfigure and Preferences

- Effective swatch and authored text differ appropriately.
- Palette rename rewrites dependents as one undo unit.
- Palette delete previews/blocks affected entries.
- Invalid edit keeps the previous valid preview.
- Fixed preview seeds produce runtime colors identical to runtime evaluation and clearly distinguish inherited Rainbow, authored dynamic, and static fallback sources.
- V2 export/reopen is authored-lossless.
- Export always preserves authored version 2 sources and never flattens expressions.
- Preferences import/export/duplicate/reset remains lossless.

### Shipped themes

- Six migrated themes are effective-color-equivalent to the captured pre-cutover constants unless documented.
- Dracula uses only official OSS palette values plus deterministic derived functions.
- Catppuccin Latte, Frappé, and Mocha use official palette values and the shared semantic mapping.
- All three shipped Catppuccin flavors use Lavender for active/focus accent, Blue for links/actions, Overlay2 at 26% over Base for selection, and do not use Mauve as the default global accent.
- Required text/background pairs meet the repo's accessibility threshold; failures require mapping changes, not disabling the test.
- No raw hex repetition remains in a migrated `colors` block.
- Each third-party-derived file declares its upstream source and pinned snapshot; referenced exact license files exist and ship with `THIRD-PARTY-NOTICES.md` in both application outputs.

## Done Criteria

- [x] Missing-version/version-1 legacy themes are rejected; every shipped, inline, imported, and exported theme uses version 2.
- [x] Themes preserve palette entries and expressions through settings, file I/O, RedConfigure, and Preferences round trips.
- [x] All runtime consumers apply one shared resolved semantic map with no partial theme updates.
- [x] `ref`, `lighten`, `darken`, `alpha`, `blend`, both `contrast` forms, `systemAccent`, `systemColor`, `tone`, `perceptualTone`, `ensureContrast`, `harmonize`, `seededRainbow`, and `seededChoice` are documented and tested.
- [x] Missing references, cycles, duplicates, unsupported versions, and bounds produce deterministic actionable diagnostics.
- [x] The six existing themes are migrated without unreviewed visual drift.
- [x] `builtin/rainbow` remains behaviorally compatible; v2 Rainbow inheritance, per-token override precedence, high-contrast suppression, and non-Rainbow token-level dynamics are covered by tests.
- [x] Dracula plus Catppuccin Latte, Frappé, and Mocha are shipped with pinned-source attribution, exact upstream MIT license texts, root-license entries, and copied third-party notices.
- [x] Focused builds/tests, schema tests, Preferences selftests, ASan coverage, and `Run-AllTests.ps1 -Suite Full` pass.
- [x] Same-machine perf evidence is archived and shows resolution/preview within the accepted budget plus bounded allocation-free paint-time dynamic evaluation.
- [x] Authoritative specs and `.github/skills/theming/SKILL.md` describe the lasting contract.
- [x] The umbrella RedConfigure WIP plan no longer defines a conflicting `colorExpressions` schema.
- [x] This plan is moved from WIP to Done only after all prior criteria pass.

## STOP Conditions

Stop and report instead of improvising if:

- relevant pre-existing worktree edits cannot be reconciled without discarding user work;
- implementation requires a second expression representation outside `Common`;
- a requirement appears to need cross-file includes, remote fetching, or a general scripting engine;
- strict and lenient theme load paths cannot share one parser/evaluator without changing the documented settings recovery policy;
- RedSalamander and RedSalamanderMonitor cannot be made to consume equivalent resolved values for their supported token subsets;
- preserving Rainbow behavior would require changing plugin ABI, stable seed identity, or global high-contrast precedence without explicit approval;
- a proposed runtime function requires arbitrary callbacks, wall-clock animation, per-paint parsing/allocation/locking/I/O, or access to raw item strings/paths rather than the bounded runtime context;
- a migrated existing theme cannot preserve its effective values and the visual change has not been explicitly approved;
- official palette licensing/attribution requirements cannot be confirmed for a theme proposed for shipping;
- a performance gate fails twice after a reasonable bounded fix, or meeting it would require unrelated architecture work;
- any step requires modifying files outside the declared scope without first updating this plan and obtaining review.

## Maintenance Notes

- Future color functions must be added only to the shared parser/evaluator, schema, RedConfigure suggestions, specs, and tests together. Do not add RedConfigure-only functions.
- Future runtime functions require a closed compiled representation, explicit evaluation phase (load/event/paint), allowlisted tokens/context, fallback, accessibility behavior, invalidation policy, and measured cost. Do not expose general callback registration.
- Time-varying functions remain a separate future design. They require a single animation clock, background-window throttling, reduced-motion behavior, deterministic test clock, repaint coalescing, and archived performance/power evidence before being admitted.
- New semantic theme keys must be exposed through each applicable base-color adapter and validated in both main app and Monitor where consumed.
- Keep `builtin/rainbow` as the compatibility owner of plugin-wide Rainbow flags; token-level dynamic color in another base does not silently broaden that contract.
- When updating a third-party palette, pin and record the new upstream snapshot, refresh the copied exact license only if upstream changed it, review mapping diffs, and rerun notice/distribution tests.
- Keep authored values separate from resolved caches; flattening is not supported anywhere in the new theme system.
- Review dependency fan-out when adding palette recipes. One palette edit should cause one bounded graph recompute and one preview/theme application.
- Do not add cross-theme inheritance as a shortcut for Catppuccin variants in this plan. If later justified, design it separately with explicit file identity, cycle, precedence, security, and missing-base behavior.
