# Shared theme runtime: one theme per process for RedSalamander, RedXe and RedPrism (DxTheme)

Status: **ACTIVE** (I46), revision 3.1.
- Indexed as I46 on 2026-10-21 because the earlier I43 identifier was claimed by completed sorting work while PR #126 was open.
- The owner accepted revision 2 on 2026-10-19, then reopened its decisions the same day with new requirements (§0.1).
- The owner accepted revision 3's decisions D-1 to D-8 as recommended on 2026-10-19. On §0.3, "no duplication" means data only for now.
- Revision 3.1 (2026-10-19) folds in the PR #126 review at the owner's request (§0.4).

Owners:
- The RedSalamander theme maintainer owns DxTheme for its lifetime, plus T0–T5 and T9.
- The RedXe and RedPrism maintainers own P-1, T6 and T7 and their adapters.
- The DxUi owner reviews the DxUi pull request in T3.

This is a non-normative plan: current contracts stay in force until implementation, tests and authoritative-spec updates land.

Planned at 2026-10-19. Each row was read at the listed commit:

| Repository | Local path | Commit read | DxUi pin |
|---|---|---|---|
| RedSalamander (`DualTail/RedSalamander`) | `Z:\src\RedSalamander` | master `30c6cce0b69895fea6a0c68e8eb3e6c58f6bcdb4` | `f65e4a4b` (API revision 4) |
| DxUi (`RedSalamanders/DxUi`) | `Z:\src\DxUi` | origin/main `f65e4a4b9311370b3e3070088e2a8abe29adcb27` | — |
| RedXe (`RedSalamanders/RedXe`) | `Z:\src\RedXe` | main `f86fa016e7d950bda0123b826ad9dafa0182c19f` | `fa4da2cb` (revision 4, 29 commits behind) |
| RedPrism (`RedSalamanders/RedPrism`) | `Z:\src\RedPrism` | main `0ce207029bab75564a0bce341c7faf7499ce638d` | `271bd54` (revision 3, 232 commits behind) |

Drift check before each slice. The output must be empty, or each change must be reconciled here:

```powershell
git -C Z:\src\RedSalamander diff --stat 30c6cce0b..HEAD -- Common/ThemeExpression.h Common/Common/ThemeExpression.cpp Common/ThemeDefinitionIo.h Common/Common/ThemeDefinitionIo.cpp Common/PlugInterfaces/Viewer.h Common/PlugInterfaces/Terminal.h Common/PlugInterfaces/Host.h Common/ViewerDxUiTheme.h Common/ViewerTitleBarTheme.h RedSalamander/AppTheme.h RedSalamander/AppTheme.cpp RedSalamander/DxUiThemePalette.h RedSalamander/HostServices.cpp RedSalamander/FolderWindow.Viewers.cpp Specs/Themes Specs/SettingsStore.schema.json Specs/Core/Core_SettingsStore.md
git -C Z:\src\DxUi diff --stat f65e4a4b..origin/main -- include/DxUi/DxUi.h include/DxUi/ThemeColors.h src/Controls/DxUi.Theme.cpp src/Controls/DxUi.WindowHost.cpp src/Controls/DxUi.Grid.cpp Specs/DesignSystem Specs/Plans/WIP/ConsumerUiImprovements_S3_S7_2026-10-11.md AGENTS.md
git -C Z:\src\RedXe diff --stat f86fa016..origin/main -- Common/PlugInterfaces RedXe/Application.cpp RedXe/PluginHost.cpp RedXe/PluginManager.cpp RedXe/Renderer.cpp RedXe/HostChrome.cpp Specs/Plugins/Plugins_API.md Specs/Core Build build.ps1
git -C Z:\src\RedPrism diff --stat 0ce2070..origin/main -- RedPrism/ThemePalette.h RedPrism/ApplicationSettings.cpp RedPrism/Application.cpp RedPrism/ChromeHost.cpp RedPrism/CanvasCompositor.h RedPrism/CanvasCompositor.cpp RedPrism/CompositorBlit.hlsl RedPrism/LucideToolButton.cpp Common/SettingsStore.h Build build.ps1
```

The paths were checked with `git ls-files`: RedSalamander's engine headers are in `Common/` and their sources in
`Common/Common/`. A path that does not exist makes `git diff` report nothing, so keep these exact.

**Revisions:**
- **r1:** draft.
- **r2:** static DxTheme library, with an engine in each module and copies pushed to plugins. Red-teamed, accepted, then reopened.
- **r3:** one shared snapshot per process. Red-teamed by three reviewers; their findings are cited as RTA-n, RTB-n and RTC-n. Accepted 2026-10-19.
- **r3.1:** folds in the PR #126 review at the owner's request: §0.4 decisions D-9 to D-19 and the corrections listed in §11. It also fixes two regressions from web edits: the drift check pointed at engine source paths that do not exist, and the yyjson legs had lost their vcpkg baselines.

## Completion checklist

- [x] §0.2 decisions D-1 to D-8 accepted as recommended (2026-10-19). §0.3 answered: data only for now.
- [ ] **P-1:** RedXe restores the appearance refresh its specs require. Independent; can start now.
- [ ] **T0:** RedSalamander characterization:
  - an injected system-snapshot seam (same-machine A/B);
  - engine, product and Monitor dumps;
  - divergence baselines;
  - a delivery log;
  - performance baselines.
- [ ] **T1:** DxTheme bootstrapped (lean governance, six configurations), with the dependency-free snapshot ABI headers and conformance tests.
- [ ] **T2:** Engine imported with its seams applied first. The in-process oracle shows zero differences on each architecture.
- [ ] **T3:** Snapshot builder, table compiler, publisher and change filter. A DxUi pull request adds role ids, an equality key, a scratch-brush API and seeded tint tables. A consumer fixture has an EXE and a DLL share one snapshot.
- [ ] **T4:** RedSalamander parity switch:
  - `Common.dll` hosts the engine and publisher;
  - `ViewerTheme`/`TerminalTheme` are synthesized byte-identically;
  - delivery cadence is unchanged.
- [ ] **T5:** RedSalamander unification, with owner-approved visible changes:
  - plugins read the shared snapshot through new interfaces;
  - host and viewer roles are identical;
  - there is one Rainbow;
  - refresh is filtered;
  - DPI has its own channel.
- [ ] **T6a:** RedPrism moves to a DxUi revision that carries T3's pull request, or records that it cannot and takes the D-12 fallback.
- [ ] **T6:** RedPrism: themed chrome, canvas, overlays and shader colors, with the semantic canvas colors kept fixed; theme files; Rainbow.
- [ ] **T7:** RedXe:
  - the snapshot in Prepare;
  - `GetTheme`/observer;
  - per-instance overrides;
  - every bundled plugin themed;
  - theme files.
- [ ] **T8:** Dynamic themes beyond Rainbow: state blends everywhere, then time-based programs behind the gate (§3.7).
- [ ] **T9:** Legacy interfaces removed. Specifications, skills and docs reconciled in four repositories. Plan admitted to `Specs/Plans/Done/`.

## 0. Requirements and decisions

### 0.1 Owner requirements (2026-10-19)

| ID | Requirement |
|---|---|
| R0 | Extraction stays simple; usage in the other products stays easy (original request). |
| R1 | No specification is a constraint. Any spec in any product may change, including DxUi's "themes stay in consumer adapters". |
| R2 | The exe and every plugin share the **same theme data**, with **no duplication**. |
| R3 | RedXe uses the theme in **every plugin**. |
| R4 | RedPrism is eventually as completely themed as RedSalamander. |
| R5 | Rainbow mode and future, more dynamic themes work everywhere. |
| R6 | Consider a theme DLL and other solutions. |

### 0.2 Revision 3 decisions (accepted 2026-10-19)

The owner accepted every recommendation in this table on 2026-10-19; each is now this plan's decision. Changing one
needs a new owner decision, recorded here with its date.

| ID | Decision | Accepted choice | Reason / evidence |
|---|---|---|---|
| D-1 | Runtime data model | **A host-owned, immutable, ref-counted `ThemeSnapshot`.** It is published only when the resolved content changes. The exe, every plugin and every DxUi host share it by reference. | This is the only model that gives one runtime copy across the exe, plugin DLLs, workers and service lanes (RTA bottom line, RTB, RTC). It removes today's divergences (§2.2). |
| D-2 | Code home | **The DxTheme repository** holds the stateless engine, the snapshot builder and table compiler, and the dependency-free snapshot ABI headers. **DxUi** gains only: control-role ids and metadata (the S7 list), an equality key for `SetTheme`, a scratch-brush API, and seeded tint tables for its controls. DxUi never sees the snapshot format. | Engine-in-DxUi was rejected for three reasons. First, RedXe forbids DxUi headers in the 12 of 13 bundled plugins that do not link DxUi (`Core_DxUiIntegration.md:223-226`), and DxUi has no header-only mode (RTA-4, RTB-11). Second, each product pins a different DxUi revision (RTA-2). Third, DxUi's validators would block a verbatim import (RTA-8). |
| D-3 | Binary homes, and the DLL question | **No new DLL.**<br>• **RedSalamander:** `Common.dll`, already a DLL beside the exe that all 8 UI plugins link (zero loader or packaging work), hosts the engine and publisher (RTA-3).<br>• **RedXe and RedPrism:** the exe hosts them. | A theme DLL shares code, not data. RedSalamander already ships the engine in `Common.dll`, yet plugins still get copies and derive different colors (§2.2). In RedXe, plugins import only from `Plugins\` and System32 (`PluginHost.cpp:1226-1227`), so a DLL would need two copies and two packaging entries for no gain. RedPrism has no plugins. |
| D-4 | Dynamic model | **The host compiles theme programs into tables inside the snapshot**: seeded hues, choices, state variants and on-colors. Every consumer only indexes them (C++, HLSL, ViewerWeb JavaScript). The host implements `IThemeSnapshot::Evaluate`, reserved for continuous programs. Time-based programs stay behind the §3.7 gate. | "One evaluator" then means one binary in the host (RTA-5, RTB-19). Today there are 19 recipes, two hash families and seven HSV converters (§2.3). |
| D-5 | Rainbow semantics | **Per-token Rainbow suppression, which the current spec already requires** (`Core_SettingsStore.md:760`: "Static overrides suppress Rainbow only for those tokens"), plus one on-color rule (D-10). The visible change is approved at T5 with numeric before/after. Flag-driven Rainbow that ignores overrides is not an option, because it would delete that normative rule. | The code honors suppression only for `folderView.itemBackgroundSelected` (`AppTheme.cpp:1115-1129`). Shipped SolarFlare sets three keys statically yet still shows Rainbow there (RTC-4). T5 therefore fixes a bug against the existing spec. There are three on-color rules today (RTC-3). |
| D-6 | Vocabulary and robustness | **Layering and references.** Layers are DxUi roles ← suite keys ← product keys ← plugin keys, and references only point down that list. Plugin token specs come from static metadata and merge before the document resolves.<br>**Failure isolation.** Suite, plugin and foreign tokens fail one at a time, falling back to their declared value. RedSalamander's own keys keep whole-theme failure until the owner changes it.<br>**High Contrast and semantic colors.** Every token declares a High Contrast mapping. Semantic tokens never inherit Rainbow or accent programs.<br>**Per-plugin control.** Plugins can opt out.<br>**User colors.** A user color that is absent and one set explicitly are distinguished. | RTC-1, RTC-7, RTC-8, RTC-9, RTC-17, RTB-15. |
| D-7 | What users see | **RedSalamander:** T4 is parity; T5 changes are visible (identical host/viewer roles, one Rainbow per seed, filtered refresh) and each is owner-approved with numeric evidence.<br>**RedPrism:** the default look is unchanged at an injected appearance; adds theme files, Rainbow and a dark caption; accent following is opt-in.<br>**RedXe:** the default look is unchanged; all 13 bundled plugins become themeable (5H4D3R5's third-party art and the camera opt out); adds theme files; background becomes a per-instance override. | RTB-18, RTC-9, RTC-17. |
| D-8 | Qualification bar | Debug, Release and ASan Debug on x64 and ARM64 for DxTheme and every consuming project. Native ARM64 and ASan runs for the library tests. Parity is compared per architecture on the same machine. RedSalamander's H4 deferral stays unless lifted. | The r2 bar (D-5), plus RTC/RTB findings on per-architecture parity. |

### 0.3 Code duplication (answered 2026-10-19)

**Owner answer:** "no duplication" means **data only, for now**. This plan does not pursue a DxUi DLL. It stays a possible
separate packaging decision.

The question was: does "no duplication" (R2) also cover **code**? If it does, the next lever is a **DxUi DLL**, not a theme DLL. DxUi is a static library in about 11 RedSalamander modules: 4.38 MiB of attributed code and data, estimated 2.5–4 MiB saved on disk (`DxUi_SharedLibraryAdoptionAndReleasePlan_2026-09-09.md` §3.3). Each module also keeps its own device, timer and typography state. That needs its own packaging decision; the September decision chose static. RedXe (2 DxUi modules) and RedPrism (1) gain almost nothing. Theming does not need it, and the snapshot ABI works either way (RTA-10).

### 0.4 Decisions from the PR #126 review (accepted 2026-10-19)

The owner asked to fold the valid points of the PR #126 review, and the recommended answers to its open questions, into
the plan. These decisions are accepted on that basis. Each slice re-checks the ones it depends on when it starts.

| ID | Decision | Accepted choice | Reason / evidence |
|---|---|---|---|
| D-9 | Which DxUi role values RedSalamander's host and viewers share | **The host mapping** ("Path B", `RedSalamander/DxUiThemePalette.h`) for every RedSalamander surface. Viewers and the terminal adopt it at T5, and the host's surface variants become declared overlays.<br>DxUi's `MakeThemePalette` ("Path A") stays the derivation for themes that supply only the suite keys: RedPrism user themes and RedXe's AV Control. | Path A takes only 11 colors and 4 flags. It would drop the `menu.*`, `folderView.*` and `fileOps.*` values that RedSalamander themes author for chrome. Path B keeps them, so only viewer chrome changes (F-16). The T5 oracle is the host's Path B role table recorded in T0.<br>This revises the first answer given in review discussion (Path A), for that reason. |
| D-10 | On-color rule | DxUi's WCAG contrast rule, `ChooseContrastingTextColor` (DxUi `Theme.cpp:126-131`), applied to the composited background. It is private in DxUi today, so DxTheme carries an identical copy tested against it. | It replaces three rules (RTC-3). Where the rules disagree the change is visible; it is approved at T5. |
| D-11 | Startup and failure | The host publishes a snapshot before the first plugin open, Prepare or service creation. If the configured theme fails, it publishes the built-in base instead. A failed rebuild keeps the previous snapshot current and reports a diagnostic. | Plugins never start without a theme, and a broken theme file never blanks the UI. |
| D-12 | RedPrism's DxUi pin | Its own slice, T6a, before T6. If the pin cannot move, T6 fills today's `ThemePalette` fields from the snapshot's role values without the new DxUi API: no equality key, no scratch brush, and dynamic tints off until the pin moves. | RedPrism is on revision 3, 232 commits behind; T6 must not stall on the bump. |
| D-13 | Restore generalization | RedSalamander's `Tools/Modules/Build/DxUiDependency.psm1` is generalized only after I31 closes, or by I31's owner. RedXe and RedPrism generalize `Build/DxUiRestore.psm1` in T7 and T6. | I31 still owns that module's qualification. |
| D-14 | Who publishes | One publisher per process that shows themed UI: RedSalamander (in `HostServices`), RedSalamanderMonitor and RedConfigure each own one. RedConfigure previews build private snapshots that are never published. The search service publishes none. `Common.dll` provides the classes, never a process-wide instance. | `Common.dll` is loaded by all four executables (RedSalamander, RedSalamanderMonitor, RedConfigure and RedSalamanderSearchService reference it). |
| D-15 | P-1 message allow-list | Refresh on `WM_SETTINGCHANGE` when lParam is `"ImmersiveColorSet"` or wParam is `SPI_SETHIGHCONTRAST`, and on `WM_THEMECHANGED`, `WM_SYSCOLORCHANGE` and `WM_DWMCOLORIZATIONCOLORCHANGED`. Ignore every other broadcast. Re-probe, compare, and request one coalesced frame only on a change. | No refresh on unrelated settings broadcasts. |
| D-16 | Time-based colors and threads | Time-based colors use an allocation-free `IThemeSnapshot::Evaluate`, called on the render thread on the snapshot already pinned during Prepare. `GetTheme` stays forbidden inside `Render`, and a snapshot is never republished per frame. | Index-only tables cannot animate. Republishing every frame would break the idle rules. |
| D-17 | Theme file syntax | Theme files are a separate JSON5 document format owned by DxTheme (the v2 grammar). They are never an extension of a product's settings document. RedXe settings stay strict JSON with comments and trailing commas. | RedXe `Core_Settings.md:65-66`. |
| D-18 | Plugins that load later | Bundled plugins' token specs come from a build-time manifest in the host's bundled catalog and register before the first resolve. A plugin discovered later (RedSalamander custom paths, or re-discovery on settings apply) gets a per-plugin overlay snapshot on first bind, resolved against the current process snapshot. The process snapshot is not republished; the plugin's specs join it at the next real theme change. | Otherwise merge-before-resolve and lazy binding cannot both hold (RTB-15). |
| D-19 | RedXe background and color defaults | RedXe's default theme reproduces today's literals: `host.canvas` `#000000`; Studio Clock `#FF1616`; Desk Clock card `#FF3B43`, digit `#FFFFFF` and date `#D8D8D8`; the Matrix head and trail colors; and every other literal the T7 census finds.<br>An omitted color resolves to the theme. An explicit one wins, except under High Contrast. User files are never rewritten. Shipped templates stop authoring colors that equal the default theme, and the settings test changes with them (replacement spec text in §7.3). | `Core_Settings.md:85, 164-165` make an omitted background mean `#000000` today, and require the templates to author it (RTC-9). |

## 1. Acceptance targets

| ID | Target | Witness |
|---|---|---|
| S1 | **Same data:** one resolution and one immutable snapshot per real change per process. Every module reads the same bytes. | The T3 fixture and product tests check content-hash and pointer identity across the EXE and DLLs. |
| S2 | **Coverage:** every RedSalamander viewer and the terminal, every bundled RedXe plugin (13, with declared opt-outs), and all RedPrism chrome and canvas UI read their colors from the snapshot. | Per-product token coverage lists (T5–T7). |
| S3 | **Consistency:** in RedSalamander, host and plugin DxUi roles are identical (the host mapping, D-9); one Rainbow hue per seed in host, plugins and DxUi controls. | RTB parity gates (§8.2). |
| S4 | **Dynamic:** tables are evaluated identically in every module with no allocation. Idle rules hold: zero work while settled. | Allocation hooks, wake-up counters, DxUi's clean-frame budget. |
| E1 | **Engine parity:** imported verbatim; RedSalamander's resolution is byte-identical through T4. | T2 oracle; T0/T4 dumps. |
| E2 | **No second engine:** no module compiles a second engine. Legacy engine sources are deleted in T4; legacy interfaces in T9. | Source contract. |
| U1a | **Build glue cap:** per consumer, the lock, restore call, MSBuild wrappers and project imports stay at or under about 100 lines (RedSalamander) and 80 (RedXe, RedPrism). The STOP rule applies. | Counted in T4, T6 and T7. |
| U1b | **Governance and tests:** measured, not capped. Estimates: RedSalamander 300–450 lines, RedXe 170–300, RedPrism 160–250. Generalizing each product's existing DxUi restore and provenance code by library name cuts this to 40–80 lines each. | Reported. |
| U2 | **Sample size:** `Samples/HelloTheme` is 40 lines or fewer: publish, share with a DLL, read tokens, apply the caption, handle a change. | Consumer fixture. |
| U3 | **No imposed facility:** no logger, settings store, string table, new DLL, COM/WinRT runtime or compiled third-party dependency. ABI headers are dependency-free. | Validator. |
| U4 | **Hot-path cost:** a paint-time read is an array index. Resolution stays within RedSalamander's guards (`Core_SettingsStore.md:762-766`). | A/B/A measurements. |
| U5 | **Updates:** one reviewed lock change per product; rollback by reverting it together with its adapters. | Rehearsed in T6 or T7. |

## 2. Verified current state

### 2.1 RedSalamander engine

**Engine files**
- `Common/ThemeExpression.h/.cpp` (140/980 lines): 14 source kinds, parser and formatter, memoised resolver (depth 32), Load/Event/Paint phases, compiled paint programs, OKLab/OKLCH/WCAG math.
- `Common/ThemeDefinitionIo.h/.cpp` (61/538 lines): validators, an order-sensitive parse from `JsonValue` (`:154-349`), and a writer with group comments.
- The theme parts of `SettingsStore`: inline themes with an opaque round trip. The serializer sorts by id and inserts opaque entries at `min(originalIndex, size)` (`:6560-6588`).
- All of the above is exported from `Common.dll` (`Common/Common/Common.vcxproj:38`).

**`AppTheme`** holds:
- the product structs;
- the five built-ins as code (`AppTheme.cpp:342-774`);
- live probes (`:174-237`, `:786-848`);
- the 54-key override map;
- `ResolveAppThemeSelection` (`:1131-1175`). Its order is: accent before the base is rebuilt; High Contrast ignores custom themes; an invalid theme falls back to the base (`UI_VisualStyle.md:158-163`).

**Engine coupling**
- An include cycle through `SettingsStore.h:3` and `ThemeExpression.h:126`.
- Live Windows reads inside the engine (`ThemeExpression.cpp:708-731`; `systemColor(highlight)` is the DWM accent).
- One hard-coded paint key (`:865-868`).
- `Debug::Perf` calls (`:877-888`).
- Whole-theme failure on a missing reference or a disallowed paint source (`:514`, `:574`).

**Format facts kept exactly:**
- `formatVersion` 2;
- the id grammar, palette ≤128, colors ≤512, sources ≤256, references ≤96, numbers ≤64, 2–8 choices;
- the JSON5 flags (`SettingsStore.cpp:8783`).

**Vocabulary.** 67 keys: 54 in AppTheme, 13 in the Monitor. Nine are generic.

### 2.2 How RedSalamander shares a theme at runtime today

**Delivery**
- Plugins receive the theme by value:
  - `IViewer::SetTheme(const ViewerTheme*)`, 96 bytes (`Viewer.h:66-105`, `:131-132`);
  - `ITerminal::SetTheme(const TerminalTheme*)`, 124 bytes (`Terminal.h:133-148`, `:294`).
- Pushes happen on the UI thread with no change detection:
  - at startup;
  - on a theme pick;
  - on every settings apply;
  - on every unfiltered `WM_SETTINGCHANGE`, `WM_THEMECHANGED`, `WM_DWMCOLORIZATIONCOLORCHANGED` and `WM_SYSCOLORCHANGE` (`RedSalamander.cpp:12255-12258` → `:11729-11733` → `ApplyAppTheme` `:8853-8894` → `FolderWindow.Viewers.cpp:537-559`);
  - before each viewer opens, and when a terminal opens.
- `ViewerTheme` also carries DPI. Terminal and ViewerSpace adopt it, but `FolderWindow::OnDpiChanged` never re-pushes it (`FolderWindow.Layout.cpp:1458-1497`).

**Palette derivation**
- Each viewer derives 0–3 DxUi palettes through "Path A" (`Common/ViewerDxUiTheme.h` → DxUi `MakeThemePalette`).
- ViewerWeb duplicates the blends again in JavaScript (`ViewerWeb.cpp:1276-1297`).
- The host derives its own through "Path B" (`DxUiThemePalette.h`) and four more builders.
- About 20 sites call `ResolveConfiguredTheme()` independently.
- The FileOps popup builds a full palette for every button on every frame (`FolderWindow.FileOperations.Popup.cpp:6921-6922`).

**The same theme gives different DxUi roles in host and viewer chrome.** In built-in Dark:

| Role | Host | Viewer |
|---|---|---|
| windowBackground | .071 | .078 |
| text | .941 | .922 |
| hoverFill | accent at α .18 | a selection blend at α .16 (DxUi `Theme.cpp:1047-1049`) |
| headerHovered, border, scrollbar roles | differ | differ |
| selectionInactiveFill | differs whenever the source keys differ (e.g. `Ugly.theme.json5`) | — |

Compact mode, reduced motion and backdrop never reach plugins.

**Loading**
- `SetDefaultDllDirectories(SYSTEM32 | APPLICATION_DIR)` (`RedSalamander.cpp:8513-8521`).
- Plugins load with `LoadLibraryExW(path, LOAD_WITH_ALTERED_SEARCH_PATH)` from `<exe>\Plugins\` and user `customPluginPaths` (`ViewerPluginManager.cpp:790-817`).
- Plugins may defer unload (`RedSalamanderPluginCanUnloadNow`, `Factory.h:145-154`; `PluginModuleLifecycle.cpp:35-44`).
- Binary policy is source-tree lockstep (ATLAS-DEC-ABI-01). Incompatible changes need a new IID or method (`Viewer.h:24-27`).

**Host services**
- `IHost` is an empty `IUnknown`. Services are reached by `QueryInterface` on `HostServices`, a process-lifetime singleton that is never deleted (`HostServices.cpp:508-598`).
- Callbacks use cookies and drain through `SetCallback(nullptr, nullptr)`.

### 2.3 Rainbow today: 19 recipes in five kinds (RTC-2)

| Kind | Sites | R3 disposition |
|---|---|---|
| (a) Seeded tints | Host accent (`AppTheme.cpp:848-858`, byte-wise FNV of "RedSalamander" = 128°); FolderView selection, focus fill and focus border (`FolderView.Rendering.cpp:2334-2420`); breadcrumb (`NavigationView.Breadcrumb.cpp:569-581`); viewer title bars (`ViewerTitleBarTheme.h:28-40` plus 8 local copies, seed `"title"` = 233°); DxUi Grid, Tree and Menu tints (`DxUi.Theme.cpp:230-245`) | Theme tables (D-4) |
| (b) Ordinal and categorical palettes | Spinners +30°·i; ItemProperties index×53°; hex triad and nibble palette (`ViewerText.Hex.cpp:199-225`, active outside Rainbow too); FileOps stream slots × golden angle | Product logic reading categorical tokens (`dataviz.series.N`) and parameters |
| (c) Caller-held state | Status bar +47° per pane activation (`FolderWindow.Interaction.cpp:112-118`) | Product state passed as an ordinal seed |
| (d) Wall clock | ViewerSpace progress and spinners (`ViewerSpace.cpp:3792-3822`, `4141-4173`) | Stays product logic until the §3.7 time gate |
| (e) Data visualization | Treemap (`ViewerSpace.cpp:4040-4094`); RedXe ProcessViewer heat ramps | Product logic reading ramp-endpoint tokens |

**Counts and visible effects**
- There are two FNV families: byte-wise (`AppTheme.cpp:239-257`) and per code unit (`Helpers.h:133`, DxUi `Utilities.h:9`). There are seven HSV converters.
- In Rainbow, the host title bar is green and the viewer title bars are blue.
- Diff colors are the accent hue +118°/+248° (`AppTheme.cpp:742-748`). With a blue accent, added lines therefore turn red and removed lines green (RTC-8).

### 2.4 RedXe today

- **Theme inputs.** There is no theme system. Plugins get `RedXeFactoryOptions::backgroundColor`, which is part of the provider identity (`PluginManager.cpp:1112-1126`), so a change rebuilds the runtime. They also get `RedXeAppearance` (Windows dark/High Contrast plus 7 system colors), by value, only in Prepare. Only AV Control reads it, and it rejects unknown flag bits (`AVControl.cpp:288-289`).
- **Hard-coded colors.** 12 of the 13 bundled plugin DLLs hard-code their colors. Near-white text alone is defined at least 7 times.
- **Render technologies:**
  - D3D11 and HLSL: RotatingTriangle, MatrixRain, 5H4D3R5, StudioClock, DeskClock, Launcher.
  - Instanced SDF with DirectWrite atlases: ProcessViewer, Weather, Logicon monitor.
  - GDI DIB: GdiOrbit.
  - DxUi: AV Control.
  - CPU faces sent as JPEG to HID keys: Logicon service.
- **Constant buffers.** Every GPU plugin already maps a per-draw constant buffer, so theme colors add at most 48 bytes and no allocation.
- **Bugs:**
  - The live appearance refresh was deleted in `1c0de0a` (2026-09-07). Three specs still require it (`Plugins_API.md:660-661`, `Core_PerformanceAndResources.md:63-65`, `Plugins_AVControl.md:207`).
  - Services always get a black background (`PluginHost.cpp:2409`).
- **ABI rules:**
  - one contract;
  - `sizeBytes == sizeof`;
  - `offsetof` asserts;
  - new IIDs are direct `IUnknown` children;
  - the `IRedXeHost` vtable may grow;
  - the UI thread is the render thread.
- **Loading.** `LoadLibraryExW(..., LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR | LOAD_LIBRARY_SEARCH_SYSTEM32)` means plugin imports are never searched in the exe directory. That is why `yyjson.dll` is copied twice.
- **Settings.** The settings document must reject JSON5 extensions (`Core_Settings.md:65-66`). The shipped settings write explicit colors (`#000000`, Matrix colors).

### 2.5 RedPrism today (main)

- **Architecture.** A single process. ChromeHost copies the palette and marks dirty.
- **Theme.** `system`/`light`/`dark`, plus a High Contrast mapping (`ThemePalette.h:11-104`) that leaves the tooltip, smoke and six alert roles at their defaults. The accent is the library default. `reducedMotion` is forced off (`:19`, `:45`).
- **Refresh.** Unfiltered (`Application.cpp:5152-5156`). There is no dark caption (`:2577`).
- **Themeable surfaces:** chrome, rail SVG icons (re-parsed for each color, `LucideToolButton.cpp:293-340`), overlays' accent, the clear color, and the HLSL pasteboard checker. That checker is pinned by tests to 46/56 (`CompositorTests.cpp:1391-1392`), and goldens are never rebaselined (`AGENTS.md:90-93`).
- **Semantic colors:** curve channels, levels handles, marching-ant contrast pairs, and the document checker and grid.
- **Shaders.** `BlitConstants` (64 bytes, b0) is re-mapped on every blit.
- **Rules.** Idle work must be zero (`AGENTS.md:78-80`). There is no plugin ABI.

### 2.6 DxUi today

- **`ThemePalette`.** 46 colors plus flags, about 748 bytes, trivially copyable. Identical at revisions 3 and 4.
- **`ControlHost`** keeps one copy. Controls read it by reference.
- **`SetTheme`.** Copies the palette with no equality check, clears the brush cache and repaints everything. There is no theme generation, hash or change notification.
- **Rainbow.**
  - Hard-coded (hue = hash % 360, fixed saturation and value). ThroughputGraph keeps its own flag, which actually controls geometry (RTC-12).
  - Grid Rainbow rows go through the ARGB-keyed brush cache, which is cleared above 256 entries. Rainbow scrolling therefore already allocates (RTB-16).
  - There is no API to recolor a borrowed brush. The Selector Button calls `SetColor` on a *cached* brush (`Controls.cpp:2749`, `:1053-1056`; F-23).
- **S7 (on hold)** proposes typed role ids and metadata and a resolved-palette hash. It forbids library-owned theme files, expressions, presets and persistence.

### 2.7 Divergences and bugs (routed; each needs an owner and a witness)

**Carried from r2:**
- F-1 to F-5, F-11 to F-13 (Monitor/app differences, dialog lookup, hot reload of disk-only ids, defaults, Preferences grid, accent fallbacks, cross-application references, two FNV variants).
- F-6, documentation drift, is fixed in T9.

**New in r3:**

| ID | Finding | Evidence | Fixed by |
|---|---|---|---|
| F-16 | Host and viewer DxUi roles differ for the same theme | §2.2 | T5 |
| F-17 | Rainbow title-bar hue: host 128° vs viewers 233° | §2.3 | T5 |
| F-18 | Rainbow diff colors can invert added/removed | `AppTheme.cpp:742-748` | T5 (semantic tokens) |
| F-19 | ViewerPE's menu bar is themed only at creation (UNVERIFIED at runtime) | `ViewerPE.cpp:743-803` | T5 |
| F-20 | No DPI or theme re-push on a DPI change | `FolderWindow.Layout.cpp:1458-1497` | T5 (DPI channel) |
| F-21 | RedXe services always get a black background | `PluginHost.cpp:2409` | T7 |
| F-22 | AV Control picks dark from Windows but paints the dashboard background | `AVControlView.cpp:212-217` | T7 |
| F-23 | DxUi Selector Button recolors a cached brush (UNVERIFIED at runtime) | DxUi `Controls.cpp:2749` | DxUi owner |
| F-24 | RedPrism High Contrast leaves alert, tooltip and smoke roles at defaults (UNVERIFIED visually) | `ThemePalette.h:32-85` | T6 |
| F-25 | System High Contrast forces `dark=false` even on dark schemes | `AppTheme.cpp:794-797` | T3/T5 |
| F-26 | Every palette build calls `SystemParametersInfoW` | DxUi `Theme.cpp:12-21`, `956-960` | T5 |
| F-27 | ViewerSqlite's theme DPI is a dead store; `TerminalTheme::ansiArgb` is never filled or read | `ViewerSqlite.cpp:1049`; `TerminalHostSupport.cpp:104-119` | T9 |

### 2.8 DxUi extraction lessons (still apply)

- Ban product includes from day one, and apply seams before the import (RTA-8).
- Freeze the moved symbols and delete the legacy copy promptly.
- Libraries hold no global state (registrations are module-owned). Test an EXE together with a DLL.
- Six configurations, a seeded ASan probe, and annotation passthrough.
- Reuse each product's restore code by library name.
- Path budget.
- Build every configuration of every consumer before calling a pin bump safe.
- Condition every consumer import.
- Re-add include paths, because `Common.vcxproj:113` overwrites its list.

## 3. Architecture

### 3.1 Data versus code

| Concern | Where it lives | Copies per process |
|---|---|---|
| Resolved theme data | One immutable `ThemeSnapshot` per content hash, owned by the host publisher | **1** (shared by reference) |
| Engine (documents, expressions, selection, compiler) | DxTheme, linked by the host binary (`Common.dll`, `RedXe.exe`, `RedPrism.exe`) | 1 |
| Dynamic evaluation | Compiled by the host into snapshot tables; `Evaluate` for continuous programs runs in the host | 1 |
| Reader | Header-only and index-only (bounds-checked loads); in each consuming module | Trivial |
| DxUi palette per `ControlHost` | Copied from the snapshot's role section; skipped when the equality key matches | One small copy per host, by DxUi's design |

### 3.2 `ThemeSnapshot` format v1

**Header**
- Fields: magic `DXTH`, format major and minor, total bytes, a 64-bit content hash, and the generation.
- Flags: dark, darkBase, highContrast, rainbow, transparency, reducedMotion.
- Also: the base theme id, density, overlay material and the time epoch.
- **Content hash.** It covers every byte a consumer may read: every section, and every header field except `generation` and the hash itself.
  - Raw probe inputs are not stored. They stay in the publisher, and only their resolved effects reach the snapshot.
  - So a probe change that alters no resolved value does not republish (for example, an accent change under a theme that ignores the accent).
  - A change to reduced motion, density, overlay material or the epoch always republishes.
- Layout: 8-byte aligned, position-independent, little-endian.

**Directory.** Entries of {tag, version, offset, count, stride}, so each section can grow (RTB-10).

**Sections**

| Section | Content |
|---|---|
| Tokens | {keyHash32, keyOffset, value as uint32 ARGB (the engine's own representation), programIndex or none}, sorted by hash. Key strings are compared to resolve hash collisions. |
| Roles | DxUi role id (append-only, the S7 list) → float32 RGBA (an exact copy of today's `D2D1_COLOR_F` value), plus an optional program index, computed once by the host. |
| Programs and tables | Compiled seeded-hue tables (per base, 360 entries), choice candidates, state variants, on-color outputs; each program has a fallback. |
| Overlays | Named role overlays for surface variants (RedSalamander's folder-content, navigation-edit, credential-prompt and alert builders). |
| Strings | UTF-8 key strings. |

**Precision.** Each value keeps its legacy representation, so parity is byte-exact (RTB-9, PR #126 review):
- **Tokens** are uint32 ARGB, as the engine resolves them.
- **DxUi roles** are float32, exact copies of today's `D2D1_COLOR_F` values: DxUi accent .47/.84, RedPrism background .118/.133, alphas .14 and .30. A 16-bit channel cannot hold those exactly.
- **Conversions.** There is one documented conversion each way:
  - ARGB → float divides each channel by 255 (DxUi `ColorFromArgb`);
  - float → ARGB is `lround(channel × 255)`, as the viewer records use (`FolderWindow.Viewers.cpp:309-323`).
- **Exceptions.** Fields that legacy code truncates instead (`AppTheme.cpp:215-221`) keep their legacy conversion, and the plan lists them by name. No tolerance is declared anywhere else.

**Versioning.** Each build uses an exact format version (RedXe's one-contract rule). Readers ignore unknown flags; AV's check is rewritten in T7 (RTB-12). Opcode and table meanings are append-only and never redefined (RTA-11). No snapshot crosses products, so there is no cross-product binary compatibility promise.

### 3.3 Interfaces, threading and lifetime

**`IThemeSnapshot : IUnknown`**
- `STDMETHODCALLTYPE`, `noexcept`.
- `Data()` returns the immutable header.
- `Evaluate(...)` is reserved.
- Defined in DxTheme's ABI headers. Each product's PlugInterfaces includes it, with `sizeof`/`offsetof` asserts.

**Who constructs and frees snapshots**
- Only the host image constructs snapshots. A static library in a plugin must never create one, because its vtable would vanish at `FreeLibrary` (RTB-4).
- The final `Release` only frees heap memory: no back-pointer, no logging, no pool.
- The publisher is leaked at process exit, like `HostServices`.
- Product quiet points release theme references and unsubscribe before `SetCallback(nullptr, nullptr)`.

**Publisher**
- Its current pointer is guarded by an SRWLOCK.
- Readers copy and `AddRef` under the shared lock; publishing swaps under the exclusive lock.
- `GetSnapshot` never marshals to the UI thread.
- Thread affinity is documented on every method. `GetTheme` is forbidden inside RedXe's `Render`/`OnDataSnapshot` (RTB-3).

**Lifetime rules for consumers**
- A view or token reference never outlives its COM pointer.
- Caches swap the pointer and its derived values together, keyed by content hash.
- Debug builds poison memory on the final `Release`.
- An ASan test re-enters `SetTheme` from a nested message loop (RTB-5, RTB-8).

**New interfaces** (the old ones are kept until T9; RTB-1):

| Product | Interface | Notes |
|---|---|---|
| RedSalamander viewers and terminal | `IThemeSnapshotSink` (new IID, found through `QueryInterface`) | Old `SetTheme` slots stay until T9. Null is invalid, as today. Plugins get pushes, so there is no Subscribe (RTB-6). |
| RedSalamander, other consumers | `IHostTheme` on the `HostServices` QI chain | Callable from any thread. |
| RedXe widgets | Snapshot in the Prepare record, plus a generation check that forces re-prepare (RTB-13) | The record keeps its 24-byte prefix and swaps the 32-byte appearance for an 8-byte pointer (§7.3). |
| RedXe services and window widgets | `IRedXeHost::GetTheme`, plus `IRedXeThemeObserver` | The observer runs on the UI thread, has a synchronous `Unsubscribe` barrier, publishes copy-on-iterate, and hands off to lanes (RTB-6). |

**Per-instance data stays out of the snapshot (RTB-2, RTB-14).**
- **DPI** goes through each instance's open/attach context and `WM_DPICHANGED`.
- **RedXe per-widget background and color overrides** go through a per-instance override record next to Prepare.

### 3.4 Reader and host-compiled tables

**Reader.** `ThemeView` provides:
- `Find(key)`: once per content hash;
- `Argb(token)`;
- `Role(roleId)`;
- `Sample(token, seed, stateBits)`, which indexes the compiled tables.

**Validation.** On each new content hash, the reader checks the magic, major version, sizes, ranges, NUL termination and indices once. On failure it falls back to built-in defaults.

**Seeds and state**
- Seed kinds are {hash, ordinal}, plus a declared legacy byte-wise kind for RedSalamander items until T5 (RTC-13).
- State bits are extensible: hover, pressed, selected, focused, keyboardFocus, checked, current, alternateRow, paneInactive, windowInactive, disabled. Unknown bits are ignored (RTC-3).

**`Evaluate`** (host-implemented) returns {argb, validUntil}. It is reserved for continuous programs (§3.7). It is allocation-free and may be called on the render thread, but only on the snapshot already pinned during Prepare (D-16).

### 3.5 Engine, publisher and change filter (host side)

**Engine import.** The imported RedSalamander engine (T2) does selection: base, High Contrast precedence, accent rebuild, static overrides and dynamic programs. A builder then produces the snapshot.

**Token registry.** It covers DxUi roles, suite keys, product keys and plugin keys.
- **Bundled plugins:** specs come from a build-time manifest in the host's bundled catalog (RedXe `BundledPlugins.h`; RedSalamander's bundled plugin list) and register before the first resolve.
- **Plugins discovered later:** they get a per-plugin overlay snapshot on first bind, resolved against the current process snapshot. Their specs join the process snapshot at the next real theme change (D-18).

Each spec has:
- a default per base;
- an optional derived default;
- a High Contrast mapping;
- a semantic flag;
- an opt-out.

Lazy plugin binding then never forces a republish (RTB-15).

**One documented order (RTC-15).** RedSalamander's current quirk is recorded in T2 parity: document expressions see the pre-accent base, and hard-coded derivations run after the overrides.

**Change filter**
1. Probe on appearance messages.
2. Compare with the previous probe.
3. Re-resolve.
4. Publish only if the content hash (§3.2) changed. A failed rebuild keeps the current snapshot (D-11).

Schedules such as day/night or a wallpaper accent use a publisher one-shot timer that publishes a new generation (RTC-5).

**Probe: the exact union of today's reads (PR #126 review).** The probe records the raw value of every system read
that any product's theme path makes today. T0 enumerates it from code.

| Reader | Reads |
|---|---|
| Engine, `MakeSystemThemeResolutionContext` (`ThemeExpression.cpp:708-731`) | `DwmGetColorizationColor` with its success flag; `COLOR_HIGHLIGHT` (the fallback); `COLOR_WINDOW`, `COLOR_WINDOWTEXT`, `COLOR_HIGHLIGHTTEXT` |
| RedSalamander AppTheme (`AppTheme.cpp:174-237`, `:797`) | `SPI_GETHIGHCONTRAST`, `AppsUseLightTheme`, the DWM accent (fallback (0, 0.478, 1)); under system High Contrast the accent is `COLOR_HIGHLIGHT` |
| RedXe `RefreshAppearance` (`Application.cpp:5239-5266`) | `AppsUseLightTheme`, `SPI_GETHIGHCONTRAST`, and seven `GetSysColor` values: window, windowText, highlight, highlightText, button, buttonText, disabledText |
| RedPrism (`ThemePalette.h:24-104`) | `SPI_GETHIGHCONTRAST`, `AppsUseLightTheme`, and the `GetSysColor` indices of its High Contrast mapping (`:32-85`, enumerated in T0) |
| DxUi | `SPI_GETCLIENTAREAANIMATION` (`Theme.cpp:12-21`); `COLOR_GRAYTEXT` (`Menu.cpp:430`) |

The engine's role mapping is imported verbatim:
- Accent is the DWM colorization, falling back to `COLOR_HIGHLIGHT`.
- AccentLight and AccentDark are 25% blends toward white and black.
- Window, WindowText and HighlightText are the matching `COLOR_*` values.
- **Highlight is the accent, not `COLOR_HIGHLIGHT`.**

The T2 oracle builds the legacy and the new resolution context from the same probe record.

**Publishers.** There is one per process that shows themed UI (D-14). `Common.dll` provides the classes and never holds a process-wide instance. RedConfigure previews resolve private snapshots that are never published.

**JSON5.** The JSON5 adapter is header-only and uses the host's yyjson. DxTheme's own manifest is test-only (RTA-8).

### 3.6 Vocabulary, failure isolation, High Contrast and semantic colors

**Layers.** DxUi roles ← suite keys ← product keys ← plugin keys. References point only down the list (RTC-1).
- Suite keys are `app.accent`, `window.background`, `window.text`, `selection.background`, `selection.text` and the three alert pairs.
- `menu.*` stays RedSalamander's.

**Failure isolation.** An unknown reference, program or contrast failure in a suite, plugin or foreign token yields that token's fallback plus a diagnostic. RedSalamander's own product keys keep whole-theme failure in R3, until the owner decides otherwise.

**High Contrast.** Every token declares one of two mappings:
- a system role, limited to the values in the probe union (§3.5), such as GrayText and ButtonFace/ButtonText;
- "semantic: keep hue, ensure contrast ≥4.5 against Window".

High Contrast outranks per-widget user settings (RTC-7). F-25 is fixed when the builder computes `dark` from the High Contrast background.

**Semantic tokens** (diff added/removed, weather alert levels, ok/warn/critical, Logicon invalid, RedPrism marching ants and curve channels):
- never inherit Rainbow or accent programs;
- allow only lightness and contrast repair, validated so levels stay distinct and readable;
- keep a non-color cue (RTC-8).

**User colors versus theme.** An absent setting means "follow the theme". An explicit setting wins, except under High Contrast. User files are never rewritten. Instead, RedXe's default theme reproduces today's literals, and shipped templates stop authoring colors that equal it (D-19, RTC-9).

### 3.7 Rainbow and dynamic themes

**Rainbow**
- Rainbow becomes a theme whose seeded-tint tokens are compiled tables. Static overrides suppress it per token, as `Core_SettingsStore.md:760` already requires (D-5).
- Per-base program parameters let S/V differ between dark and light.
- The kinds (b)–(e) in §2.3 stay product logic that reads theme parameters and categorical tokens.

**Composition (bounded, no allocation)**
1. Seeded or choice program.
2. StateBlend: a static reference, per-state amounts, and alpha taken from a token.
3. OnColor: one WCAG rule.

Depth is limited.

**Composition is internal to the host compiler**, for built-in recipes only. Theme authors keep the v2 grammar:
- expressions stay non-nested;
- references to paint-time sources stay validation errors (`Core_SettingsStore.md:752`; the Done plan's "terminal" rule at `Theme_ExpressivePaletteAndReferencesPlan_2026-07-13.md:234` stands).

Exposing composition to authors would need `formatVersion` 3 and a spec change.

**Scoped accent family.** A scope seed produces accent, hover, pressed, header tint and on-accent together, for per-file viewer accents (RTC-14).

**Time-based programs** (animated accents, progress shimmer). They stay disabled until the gate from `Theme_ExpressivePaletteAndReferencesPlan_2026-07-13.md:240, :615` is met, extended per RTC-5 and RTA-9:
- time is the snapshot epoch plus the system monotonic clock;
- `Evaluate` returns `validUntil`;
- each surface keeps a "painted an animated token" bit, aggregated like RedXe's scheduled widgets;
- repaint coalescing and background-window throttling;
- reduced motion freezes time;
- a deterministic test clock;
- wake-ups are counted per module, with power evidence.

Logicon LCD keys and other surfaces without frames use the fallback.

### 3.8 DxUi changes (one DxUi pull request in T3)

- **Role ids and metadata.** S7's typed role ids and metadata become the roles section's id space.
- **Equality key.** `ControlHost::SetTheme` skips work when the key (content hash, overlay id, local overrides) matches. The legacy overload clears the key (RTB-7).
- **Scratch-brush API.** Dynamic colors are drawn with a host-owned scratch brush, never by recoloring a cached one. This fixes F-23. Grid, Tree, Menu and ThroughputGraph move to it.
- **Allocation gate.** A Rainbow scenario joins DxUi's allocation gate (RTB-16, RTC-11).
- **Seeded tint tables.** A DxUi-owned POD (a pointer, a count and a seed kind) feeds Grid, Tree, Menu and ComboBox tints. Grid also gets alternation and underlay inputs. ThroughputGraph keeps a capability flag for its geometry; its samples carry slots or seeds (RTC-12).
- **Surface policy.** Settings such as reduced motion, density and overlay material are applied by each surface after the role values, so AV Control's forced reduced motion stays possible (RTC-6).
- **Charter text.** These are the only charter changes. DxUi still owns no theme files, documents, persistence or watchers.

### 3.9 GPU and JavaScript consumers

- **GPU.** Each plugin declares its own token list, mapped to float4 slots that are resolved once per content hash. The values are copied into the plugin's existing per-draw constant buffer, so there is no cross-plugin HLSL register ABI (RTB-13, RTC-10). HLSL only reads values computed on the CPU.
- **RedPrism.** Theme fields are appended to `BlitConstants` with defaults equal to today's values (0.18/0.22). A shared C++/HLSL header is added with `sizeof`/`offsetof` asserts (RTB-17).
- **ViewerWeb JavaScript** receives token values as data and stops duplicating the color math.

### 3.10 Packaging: the answer to "a DLL for theme?"

**Data is shared by reference, with or without a DLL.**
- **RedSalamander:** the theme's binary home is `Common.dll`. It is already a shared DLL loaded by the exe and linked by all 8 UI plugins, so no loader or packaging change is needed.
- **RedXe and RedPrism:** the exe, because a DLL adds staging cost and no sharing.

**Plugins carry only the index-only reader.**

**Packaging can change later without changing plugin contracts.** The snapshot ABI is POD plus COM. A future DxUi DLL (§0.3), a C/COM theme DLL or a shared-memory section for out-of-process renderers would leave the contracts as they are.

### 3.11 What stays product-owned

- Settings and persistence: selected theme, inline themes, theme directories, watching.
- Localization of names and errors.
- Editors and workbenches; the theme-cycle overlay.
- Product and plugin token vocabularies and defaults.
- Plugin ABI interface definitions, which embed the DxTheme snapshot ABI.
- Per-instance DPI and overrides.

## 4. Repository and consumption (DxTheme)

The r2 contract still holds, plus the ABI headers.

**Repository (lean governance)**
- `AGENTS.md`, README, MIT `LICENSE.txt`, origin notices.
- `Directory.Build.*` with six configurations, an ASan forced include and a seeded probe; `VcpkgEnabled=false` for library projects.
- `build.ps1`, `test.ps1`, `test-consumer.ps1`, `format.ps1`.
- Test-only yyjson legs resolved through the consumers' exact vcpkg baselines:
  - `9e593bb18ea69cc5095e012465dcd675a822ed0d` gives yyjson 0.12.0 (DxUi, RedXe and RedPrism);
  - `d53f25d56e7806475054f624439e93ee38e1e81d` gives yyjson 0.13.0 (RedSalamander).

  These hashes are vcpkg registry baselines, not yyjson commits. Upstream: `https://github.com/ibireme/yyjson`.
- Specs: format, resolution, snapshot ABI, build/consumption, testing.
- `capabilities.json`, with `apiRevision` and a separate `snapshotFormat` version.
- One skill; CI if hosted runners are available, otherwise local receipts.

**Validator.** It checks:
- the path budget;
- no product or DxUi includes;
- one static library;
- `<Windows.h>` only in the Win32 and JSON5 headers, and never in the ABI headers;
- no mutable statics at any scope;
- ASCII-only public headers;
- the build matrix.

**Consumption**
- `Dependencies/DxTheme.lock.json`, restored by generalizing each product's existing DxUi restore and provenance module by library name. RedSalamander does this only after I31 closes, or through I31's owner (D-13). Restores clone to a temporary sibling, verify, rename atomically under a per-commit lock, and never use a live `--reference`.
- `DxTheme.Consumer.props/.targets` work in two modes:
  - **link mode** for host binaries: a ProjectReference that passes through the configuration, the annotation policy, `PreferredToolArchitecture` and WholeProgramOptimization, re-adds the include path, and checks the pin through a stamp file;
  - **include-only mode** for plugin and test projects that need only the ABI headers.
- RedSalamander, RedXe and RedPrism bind the DxTheme lock and module into their build evidence and receipts. In RedSalamander that is `BuildEvidence.psm1:675, 756, 1000` and `ValidationFingerprint.psm1:604-606`, plus an impact rule that makes a lock change trigger Full.
- The warning profiles from r2 §4.3 apply: RedSalamander `/Wall` with `/external`; RedXe and RedPrism `/W4 /WX /sdl`; RedPrism analysis; Windows.h permutations.

## 5. Extraction procedure

1. **Freeze** by fingerprinting the extracted symbols, for about two weeks at most; otherwise take a new snapshot. A backport ledger records changes made meanwhile.
2. **Apply the seams in RedSalamander before importing** (RTA-8):
   - no `Helpers.h` or `SettingsStore.h` includes;
   - injected context instead of live probes;
   - `ResolveStats` instead of `Debug::Perf`;
   - a profile for the paint key, built-in ids and the `user/` prefix.
3. **Import in one commit:**
   - the working copy goes under `src/` and `include/`;
   - the unmodified legacy snapshot goes under `Tests/Oracle/Legacy/` with minimal shims, compiled only into tests and deleted after T4;
   - `source-import.json` records the origin URL, commit, per-file SHA-256 and the clean flag.
4. **Duplicate engine-level tests** with their origin identity. RedSalamander's suites stay unchanged.
5. **Delete the legacy engine sources** from `Common` atomically in T4. The legacy plugin interfaces go in T9.

## 6. Slices

| Slice | Work | Exit gate |
|---|---|---|
| **P-1** RedXe refresh | Handle the D-15 allow-list (`WM_SETTINGCHANGE` with `"ImmersiveColorSet"` or `SPI_SETHIGHCONTRAST`; `WM_THEMECHANGED`; `WM_SYSCOLORCHANGE`; `WM_DWMCOLORIZATIONCOLORCHANGED`) → `RefreshAppearance` → compare → one coalesced frame. Ignore every other broadcast. No library. | One refresh per appearance broadcast, none for an unrelated one. `test.ps1 -Full`. The three specs in §2.4 marked implemented. |
| **T0** RedSalamander characterization | <ul><li>Snapshot-injection seam in `ResolveAppTheme` and the Monitor resolver, with a same-machine A/B.</li><li>Engine corpus: a new RedConfigureTests scope file.</li><li>Product dump: ProductUiTests plus a Monitor harness.</li><li>Divergence baselines: host (Path B) and viewer (Path A) role tables, where the host table is the T5 oracle (D-9); Rainbow hue per seed per site; all byte-wise hash call sites.</li><li>Probe record: each product's raw system reads (§3.5).</li><li>Delivery log: trigger, instance, bytes, DPI source.</li><li>Performance baselines, including `App.Startup.ApplyAppTheme`.</li></ul> | Identical bytes across reruns. Raw output under `X:\RedSalamander.Perf`. Evidence ≤2 MiB per artifact and ≤5 MiB per run. |
| **T1** DxTheme bootstrap | §4 repository, plus the ABI headers (format v1, `IThemeSnapshot`, `ThemeView`) and conformance tests. | Six configurations. The validator passes. ABI headers compile standalone under every profile. `sizeof`/`offsetof` asserts hold. |
| **T2** Engine import | §5. | Zero oracle differences on x64 (FMA3 off) and native ARM64. ASan clean. |
| **T3** Runtime | <ul><li>Builder, table compiler, publisher (§3.3) and change filter.</li><li>The DxUi pull request (§3.8).</li><li>Consumer fixture: an EXE and a DLL share one snapshot; every program × seed corpus row is identical in both; refcount returns to 0; Release after publisher teardown and re-entrant `SetTheme` are both ASan-clean.</li><li>`Samples/HelloTheme`.</li></ul> | Fixture green in six configurations. U2 and U3 met. DxUi pull request merged with its gates. |
| **T4** RedSalamander parity switch | <ul><li>Prerequisite: the restore generalization per D-13 (after I31 closes, or by I31's owner).</li><li>`Common.dll` links DxTheme and hosts the publisher behind `IHostTheme`; Monitor and RedConfigure get their own publishers (D-14); a snapshot is published before the first plugin open (D-11).</li><li>Legacy engine sources deleted.</li><li>`ViewerTheme`/`TerminalTheme` synthesized from the snapshot.</li><li>The same push cadence, call sites and DPI source; the ~20 resolve sites and the alert System theme unchanged (RTB-18).</li></ul> | Synthesized records byte-identical to the legacy builders over T0 × injected appearances, per architecture. Delivery log equal. Existing tests unchanged and green (ProductUi goldens, PluginContract sized records, ViewerSqlite theme cycle, ViewerPE test theme, self-test `BuildViewerTheme` copies). Startup within noise. Receipts bound to the pin. |
| **T5** RedSalamander unification (owner-approved) | <ul><li>`IThemeSnapshotSink` in every viewer and the terminal.</li><li>Roles from the snapshot using the host mapping (D-9); host builders become overlays.</li><li>Rainbow tables (D-5) and one on-color rule.</li><li>Semantic diff tokens (F-18).</li><li>Per-paint palette builds removed.</li><li>Refresh filtered.</li><li>Alerts on the configured theme.</li><li>DPI channel (F-20).</li><li>ViewerWeb reads values.</li><li>Status-bar hue state kept as an ordinal seed.</li></ul> | Viewer and terminal role tables equal the host's Path B table recorded in T0 (deltas recorded before, all equal after). Rainbow hue per seed equal across host, viewers and DxUi controls. These counters are zero: per-paint builds, clean-frame allocations (including Rainbow scrolling), and pushes/ViewerSpace device discards/Terminal re-grids on an unchanged probe. DPI equals `GetDpiForWindow` after a monitor move. High Contrast roles equal the system colors. |
| **T6a** RedPrism DxUi pin | Move RedPrism to a DxUi revision that carries T3's pull request, through its normal pin loop (D-12). | RedPrism `test.ps1` in Debug, Release and ASan Debug on the new pin. If the pin cannot move, record why; T6 then uses the D-12 fallback. |
| **T6** RedPrism | <ul><li>Prerequisite: agree timing with both active branch owners.</li><li>Chrome roles (through the new API, or the D-12 fallback), overlay accent, clear-color token.</li><li>`BlitConstants` theme fields.</li><li>Rail icons masked then tinted.</li><li>`reducedMotion` override removed.</li><li>High Contrast alert roles (F-24).</li><li>Semantic canvas colors declared and kept fixed: curve channels, levels handles, the marching-ant pair, and the document transparency checker and pixel grid (`CompositorBlit.hlsl:68-75`).</li><li>Filtered refresh (D-15 messages); dark caption.</li><li>Theme files under `%LocalAppData%\RedPrism\Themes` (D-17); Rainbow.</li></ul> | `test.ps1` Debug/Release/ASan Debug. CompositorTests 46/56 unchanged. Roles byte-equal in their float32 representation. A numeric dump shows the semantic canvas values unchanged under every shipped theme, Rainbow and High Contrast, except declared contrast repair. SVG documents created after a change ≤ visible icons, and 0 per settled frame. 0 frames on an unchanged probe. |
| **T7** RedXe | <ul><li>Census every color literal in every bundled plugin before changing code, as T0 does for RedSalamander.</li><li>Snapshot in Prepare plus a generation check; `GetTheme`; observer.</li><li>Per-instance override record. The provider-key change and its 10 readers land in one lockstep step; services follow the background.</li><li>Token sets per module (§7.3 table; `zoom.action` and `SystemData` have no painted surface); opt-outs for 5H4D3R5 art and the camera.</li><li>Default theme equal to today's literals; templates and settings spec updated with the §7.3 replacement text (D-19).</li><li>Theme files (D-17); AV flag check rewritten.</li></ul> | `test.ps1 -Full`, six configurations. With the default theme, a numeric dump of every declared token per plugin equals the census values, and AV Control's DxUi roles equal today's palette. AV Prepare returns S_OK for every flag combination. A background edit does not rebuild the runtime. An omitted background resolves to `#000000`. P-1 witnesses hold. |
| **T8** Dynamic | State blends and categorical tokens everywhere; then the first time-based program (candidate: ViewerSpace progress) once the §3.7 gate is met. | Gate evidence: wake-ups per module, power, reduced motion, deterministic clock. |
| **T9** Closeout | <ul><li>Remove `SetTheme(const ViewerTheme*)`, `SetTheme(const TerminalTheme*)`, `RedXeAppearance` and the compatibility headers.</li><li>Fix F-6 and F-27.</li><li>Update the §13 specs and the S7 note.</li><li>Done admission.</li></ul> | `Get-SpecInventory -FailOnFindings` and every repository's validators green. |

**Order:**
- P-1 is independent.
- T0 → T1 → T2 → T3 → T4 → T5.
- T6a precedes T6. T6 and T7 follow T3, in either order.
- T8 follows T5.
- T9 comes last.

## 7. Product adoption details

### 7.1 RedSalamander

**Build (T4)**
- Lock; generalized restore, only after I31 closes or through I31's owner (D-13), with tool inventory, help, focused test and spec; `Build/RedSalamander.DxTheme.props/.targets`.
- Link mode for `Common`. Include-only mode for every project whose sources include the theme or plugin-ABI headers; that set is derived from includes by a source contract. The include path is added once in `Directory.Build.targets`.
- Evidence binding.
- `AGENTS.md:217` ("Common → no dependencies") is updated.

**Compatibility (T4)**
- `Common/ThemeExpression.h` becomes the compatibility header. It holds:
  - the aliases (the forward declaration is removed);
  - `static_assert`s that freeze member names, order, defaults and enumerator values;
  - using-declarations for unchanged functions;
  - distinctly named DxTheme functions for the wrapped ones (resolve with metrics, profile-dependent checks, system context).
- `ParseThemeDefinitionFromValue` and the writer stay verbatim in `Common` (F-14 unifies them later).

**Gates for each T4 commit** (all non-interactive)
- Commands `settings` theme cases, `theme-overlay`, and `preferences` `cmd_preferences_dialog_themes_*`.
- RedConfigureTests, ProductUiTests (including `ResolveAppThemeSelection`), SettingsSchemaTests, PluginContractTests.
- The `DocumentationDriftContracts` function-spelling case, extended to the pinned DxTheme.
- `ThemeDistributionContracts.Tests.ps1`: lines `:64-78` stay unchanged. A new case asserts there is no engine source under `Common/` and that only the compatibility forms remain.

**Unification (T5).** In order, each a separately revertible commit:
1. The DPI channel.
2. Snapshot sinks in plugins, with ViewerTheme fallback until T9.
3. Roles and overlays.
4. Rainbow tables.
5. Semantic tokens.
6. Refresh filter.
7. Alert theme.

### 7.2 RedPrism (T6)

**Build**
- Lock, generalized `Build/DxUiRestore.psm1`, `Build/RedPrism.DxTheme.props/.targets` (conditioned, `/external:I`).
- Imports at `RedPrism.vcxproj:68/:410`.
- `build.ps1:74-75/:150-152`.
- BuildProcessTests.
- `AGENTS.md:86-89`.
- A new `Specs/Core/Core_DxThemeIntegration.md`.

**Theme sources.**
- The exe hosts the publisher, and publishes before the chrome is created (D-11).
- ChromeHost and the renderer read the snapshot.
- Without the T6a pin, ChromeHost fills today's `ThemePalette` fields from the snapshot's role values (D-12).

**Semantic canvas colors stay fixed.** Curve channels and levels handles (`AdjustControls.cpp:62-75, 365-388`), the marching-ant contrast pair (`CanvasOverlay.cpp:334-335`), and the document transparency checker and pixel grid (`CompositorBlit.hlsl:68-75`) are declared semantic tokens. They never take Rainbow or accent programs; only declared High Contrast repair may change them.

**Settings.** The `theme` setting accepts `system`/`light`/`dark` plus `builtin/*` and `user/*` ids. Theme files are DxTheme JSON5 documents (D-17). Errors show one STRINGTABLE message.

**Specs changed:** `Core_Settings.md:35, 43-44`; `UI_ApplicationChrome.md:73-77, 153-155, 203-210, 344, 718-725`; `Rendering_Compositor.md:290-293`; `Core_DxUiIntegration.md:89`. The CompositorTests pins stay.

### 7.3 RedXe (T7)

**Build**
- Lock, generalized `Build/DxUiRestore.psm1` and `DxUiProvenance.psm1`, `Build/RedXe.DxTheme.props/.targets` (conditioned).
- Link mode for `RedXe.exe`; include-only mode for every plugin and test host.
- `build.ps1:88-93/:152-154`, `ci.yml`, `Tests/test-scopes.json`.

**ABI changes in one lockstep commit**
- **Records:**
  - The Prepare record (`RedXeGpuPreparationContext`, `Widget.h:103-113`) keeps its 24-byte prefix (size, DPI, tile and raised extents). It replaces the 32-byte `RedXeAppearance` with an 8-byte snapshot pointer at offset 24, for 32 bytes in total, with `sizeof == 32` and `offsetof == 24` asserts.
  - The frame record carries no theme; Prepare plus the generation check is enough.
  - A per-instance override record is added.
- **`IRedXeHost`:** gains `GetTheme`.
- **New interface:** `IRedXeThemeObserver`.
- **Removed:** `backgroundColor` leaves the factory options and the provider key, together with its 10 readers.
- **Tests:** test hosts use a snapshot factory, or treat null as defaults.

**Plugins.** Each module declares its tokens. The T7 census supplies the default values (D-19).

| Module | Token namespaces | Notes |
|---|---|---|
| `RedXe.exe` (canvas, host chrome, failure placeholder) | `host.*` | `host.canvas` defaults to `#000000` |
| `ProcessViewer.dll` (the System Data meters: CPU, memory, network, storage, GPU, power, thermal and the rest) | `systemData.*`, `pageIndicator.*` | ok/warn/critical and the heat ramps are semantic |
| `Weather.dll` | `weather.*`, `pageIndicator.*` | Alert levels are semantic |
| `Launcher.dll` | `launcher.*` | Its page dots live in `LauncherPixel.hlsl` |
| `MatrixRain.dll` | `matrixRain.*` | Explicit head and trail settings win |
| `StudioClock.dll` | `studioClock.*` | Default `#FF1616` |
| `DeskClock.dll` | `deskClock.*` | Defaults: card `#FF3B43`, digit `#FFFFFF`, date `#D8D8D8` |
| `GdiOrbit.dll` | `gdiOrbit.*` | GDI pens and brushes rebuilt outside paint |
| `RotatingTriangle.dll` | `triangle.*` | — |
| `5H4D3R5.dll` | `shaders.{fade,tint}` | The third-party art (CC BY-NC-SA) opts out |
| `AVControl.dll` | DxUi roles (Path A) | Keeps its High Contrast mapping |
| `Logicon.dll` (LCD key faces; Debug monitor) | `logicon.*` | Invalid color is semantic; faces recompose on the device lane after an observer handoff |
| `SystemData.dll`, `zoom.action.dll` | — | No painted surface (not applicable) |
| `AVControlCamera.dll` (loaded by Frame Server) | — | Camera frames; opts out |

**Settings spec replacement text** (`Core_Settings.md`, D-19):
- **`:85` becomes:** "`backgroundColor` \| No \| Exact `#RRGGBB`. Omitted follows the theme's `host.canvas` value; the default theme's value is `#000000`. …"
- **`:164-165` become:** "A change to the document `backgroundColor`, or a theme change, updates the canvas through the theme snapshot without rebuilding the page. Shipped templates do not author colors that equal the default theme. An explicit value wins over the theme, except under High Contrast."
- **The settings test** that requires the templates to author `backgroundColor` changes: it now asserts that an omitted value resolves to `#000000` under the default theme.
- **Clock defaults:** the plugin settings defaults for the clock and Matrix colors (`Core_Settings.md:286-290`) change the same way.

**Specs changed:**
- `Plugins_API.md:15-19, 229-240, 657-661, 685-687, 946-949, 1179-1183, 1189, 1225-1228, 1283-1296, 1363-1373`;
- `Core_PerformanceAndResources.md:63-65`;
- `Core_Settings.md:65-66, 85, 156-165, 286-290`;
- `Plugins_AVControl.md:203-207`;
- `Plugins_Weather.md:95, 160-164`;
- `Plugins_Logicon.md:67-68`;
- `UI_Dashboard.md:288-296`;
- `RFC_Plugins_AVControl.md:247-252`;
- `Core_DxUiIntegration.md`;
- docs pages.

## 8. Validation and evidence

### 8.1 Commands (reference, not run for this plan)

```powershell
# DxTheme
.\validate.ps1
.\format.ps1 -Check
foreach ($c in 'Debug', 'Release', 'ASan Debug') { .\test.ps1 -Configuration $c -Platform x64; .\test-consumer.ps1 -Configuration $c -Platform x64 }
.\test-consumer.ps1 -Configuration 'ASan Debug' -Platform x64 -DisableStlAnnotations
# Repeat with -Platform ARM64 on a native ARM64 host; cross-builds only prove compilation.

# RedSalamander
.\Test-Changes.ps1 -Explain -Configuration Debug -Platform x64
.\Test-Changes.ps1 -Suite Commands -CommandsFamily settings -CaseFilter theme_ -Configuration Release
.\Test-Changes.ps1 -Suite Commands -CommandsFamily theme-overlay -Configuration Debug
.\Test-Changes.ps1 -Suite Commands -CommandsFamily preferences -CaseFilter cmd_preferences_dialog_themes_ -Configuration Debug
.\Test-Changes.ps1 -Mode PrePush -Configuration Debug -Platform x64
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -Configuration Debug -Platform x64
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
.\Tools\Get-ToolInventory.ps1 -Validate

# RedXe (also ARM64 natively)
foreach ($c in 'Debug', 'Release', 'ASan Debug') { .\test.ps1 -Full -Configuration $c -Platform x64 }

# RedPrism
foreach ($c in 'Debug', 'Release', 'ASan Debug') { .\test.ps1 -Configuration $c -Platform x64 }
```

`-Explain` must list RedConfigureTests, ProductUiTests and SettingsSchemaTests; add any it omits. Interactive cases
need the owner's agreement and `-Interactive` with the existing warning and lease.

### 8.2 Parity witnesses

**Per architecture, same machine, in emission order.** These are compared as emitted, not sorted:
- the T2 oracle;
- T0/T4 dumps built from one dumper source over injected snapshots;
- the synthesized 96/124-byte records;
- the delivery log;
- RedSalamander and Monitor dumps;
- the `affected` order and the error that wins;
- the writer's bytes.

**Corpus.** The r2 inputs, plus divergence baselines and the byte-wise hash call-site list. Cross-architecture differences are recorded, not STOPs; FMA3 state is recorded too.

**Products.** RedPrism and RedXe use numeric dumps (DxUi roles and tokens) at an injected appearance, never screenshots. Values are compared in their legacy representation: uint32 ARGB tokens and records, float32 roles. They must be byte-equal. The only exceptions are the named legacy-truncation fields (§3.2).

### 8.3 Performance

An interleaved A/B/A in test-enabled Release, legacy versus candidate, covering:
- `theme_v2_runtime_resolution_and_dynamic_perf`;
- the RedConfigure preview fan-out;
- a perceptual-function microbenchmark;
- settings load;
- `theme.cycle.apply_us`;
- `App.Startup.ApplyAppTheme`;
- snapshot build time;
- per-module wake-ups.

T5 must show zero per-paint palette builds and zero clean-frame allocations, including during Rainbow scrolling. Never rebaseline silently.

## 9. Alternatives (ranked by the red team)

| Rank | Option | Verdict |
|---:|---|---|
| 1 | Snapshot by reference + DxTheme home + `Common.dll` in RedSalamander + host-compiled tables | **Recommended**: meets R1–R6 with no new DLL, no change to DxUi's charter beyond roles and brushes, and no pin-cadence coupling. |
| 2 | Same runtime with the engine in DxUi | One repository fewer, but blocked by RedXe's header rule, the products' pin differences (RedPrism 232 commits behind), DxUi's validators and its gates on every parser fix. |
| 3 | C/COM theme DLL | Workable, but it still needs the same delivery ABI, plus a second copy in RedXe's `Plugins\` and packaging entries, for no data benefit. |
| 4 | DxUi DLL | Real code savings in RedSalamander; a separate packaging decision (§0.3); theming does not need it. |
| 5 | C++ theme DLL | Violates RedXe's ban on C++/STL across its plugin ABI. |
| 6 | Shared-memory section | Needed only for out-of-process renderers; none exist. The position-independent format keeps it open. |
| 7 | Plugins evaluating programs themselves | Duplicates evaluator code in about 20 modules plus HLSL and JavaScript, and risks drift. |
| 8 | Roles-only snapshot | Fails R3 and R4: plugin and canvas tokens are missing. |
| 9 | Per-module engines and pushed copies (r2) | Fails R2. |

## 10. Red-team attack matrix

| Challenge | Answer / witness |
|---|---|
| "A stale plugin misreads the new pointer." | New IIDs. Old slots kept until T9. |
| "A reader races the publisher's swap." | SRWLOCK copy and AddRef. Never marshalled. ASan swap test. |
| "The last Release runs after unload." | The host image implements the snapshot. Release only frees memory. The publisher is leaked. Quiet-point rule. |
| "A view dangles across a nested message loop." | Views tied to their pointer. Debug poisoning. Re-entrancy test. |
| "DPI is lost when the theme stops carrying it." | Per-instance DPI channel, before T5. T4 keeps today's source. |
| "One bad token reverts the whole theme." | Per-token isolation for suite, plugin and foreign tokens. |
| "Rainbow still differs between modules." | Host-compiled tables. Hue-per-seed equality witness. |
| "Dynamic colors allocate or wake the CPU." | Scratch brushes. Allocation gate. Time gate with per-module wake-up counts. |
| "High Contrast or semantic colors become unreadable." | A High Contrast mapping per token. Hue-locked semantic tokens. |
| "Colors shift by a rounding step." | Each value keeps its legacy representation (uint32 tokens, float32 roles), with one documented conversion. Parity is byte-equal. |
| "A reduced-motion or density change is missed, or an accent-only change republishes everything." | The content hash covers every consumer-visible byte; raw probe inputs stay in the publisher (§3.2). |
| "A plugin starts before any theme exists, or a broken theme file blanks the UI." | Publish before the first open; a failed rebuild keeps the previous snapshot (D-11). |
| "A plugin that loads later forces a process-wide republish." | Bundled manifest before the first resolve; per-plugin overlay for late plugins (D-18). |
| "Theme authors can now reference paint-time sources." | No. Composition is internal to the host compiler; the v2 grammar is unchanged (§3.7). |
| "RedXe defaults never follow the theme." | Absent versus explicit settings, with a migration. |
| "Theme changes rebuild RedXe pages." | Background leaves the provider key. A no-rebuild witness. |
| "A theme DLL would have been simpler." | §3.10 and §9: it shares code, not data. `Common.dll` already plays that role in RedSalamander. |

## 11. Claims considered and rejected

- **"A theme DLL is needed for the exe and plugins to share data."** Rejected: data is shared by reference. RedSalamander's existing engine DLL (`Common.dll`) did not prevent divergence.
- **"Put the theme runtime in DxUi" (r3 draft).** Rejected for the engine and the ABI (RTA-2, RTA-4, RTB-11). Accepted only for role metadata, brushes and tint tables.
- **"One evaluator compiled into every module."** Replaced by host-compiled tables (RTA-5).
- **"A tolerant, suite-wide binary format."** Replaced by an exact format version per build. Theme documents get the versioning instead (RTA-11, RTB-12).
- **"Swap `SetTheme`'s parameter in place."** Rejected (RTB-1).
- **"T4 is parity while also filtering refresh and retiring resolve sites."** Split out (RTB-18).
- **"Background moves out of the provider key easily."** Corrected: it becomes a per-instance record, with every reader changed in one step (RTB-14).
- **"All 19 Rainbow recipes become theme tokens."** Narrowed to seeded tints (RTC-2).
- **Retracted in r3.1 (PR #126 review):**
  - "RGBA16 keeps float roles exact." Replaced by legacy representations with one documented conversion.
  - "Hash only the resolved sections." Replaced by a hash over every consumer-visible byte.
  - "An 8-name probe list." Replaced by the exact union of today's reads, with the engine mapping kept verbatim.
  - "The Prepare record becomes 32 bytes plus a pointer." Corrected to a 24-byte prefix plus an 8-byte pointer, 32 bytes in total.
  - "Flag-driven Rainbow is an option." It would delete `Core_SettingsStore.md:760`.
  - "Composition supersedes the paint-source rule." It is compiler-internal only.
  - "A settings migration rewrites explicit defaults." Replaced by D-19.
  - "Path A as the shared role target." It drops authored `menu.*`/`folderView.*`/`fileOps.*` values; D-9 picks the host mapping instead.
- **Review claims adjusted:**
  - "No prefix in §7.3 names ProcessViewer." `systemData.*` was ProcessViewer's set; the module table now says so.
  - The review called the drift-path and yyjson-pin notes resolved. The web edits had broken both; r3.1 restores the real paths and the full vcpkg baselines.
- **From r2:**
  - Product adapters per module and no runtime sharing (fails R2).
  - RedXe dark-only and widgets deferred (fails R3).
  - RedPrism partial theming (fails R4).

## 12. STOP conditions

Stop the affected slice and keep the previous pin or legacy path when any of these happens:
- any parity witness differs on the same architecture;
- performance or the allocation gate regresses without accepted evidence;
- ASan reports a finding;
- a consumer exceeds U1a (fix the kit);
- an ABI header fails a profile or an assert;
- RedSalamander's or RedXe's DxUi pin cannot take T3's pull request (RedPrism takes the D-12 fallback instead of stopping);
- a branch owner objects to the timing;
- the DxUi owner rejects the role and brush pull request.

Never rebaseline silently. Never reclassify a failing required test as optional.

## 13. Authoritative updates

- **DxTheme:** format, resolution, snapshot ABI, build/consumption and testing specs; `capabilities.json`; skill; README; sample.
- **DxUi:**
  - `AGENTS.md:11-12`, `Specs/DesignSystem/Theming.md:7, :44` and `Specs/UI/UI_ThemeAndTypography.md:33-34` describe role ids, tint tables and the scratch brush;
  - `Core_Architecture.md:40-42, :80-83, :87`;
  - the S7 plan's forbid list is rescoped.
- **RedSalamander:**
  - `Specs/Core/Core_SettingsStore.md` keeps the acceptance requirements: grammar, limits, High Contrast precedence, performance contract, keys.
  - Also updated: `Specs/Plugins/Plugins_ViewerPlugins.md`, `Specs/Plugins/Plugins_PluginAPI.md`, `Specs/UI/UI_VisualStyle.md`, `Specs/Build/Build_Toolchain.md` (pinned DxTheme), `Specs/Core/Core_SharedHelpers.md`, `Specs/NormativeConsistency.json`, the theming skill, `docs/Themes.md`, `docs/DeveloperGuide.md`, a new `docs/DxTheme.md`, `AGENTS.md:217`, the tool inventory and the tooling-governance spec.
- **RedXe and RedPrism:** the specs listed in §7.2 and §7.3; new `Core_DxThemeIntegration.md` files; AGENTS consumption rules.

## 14. Done criteria and routed follow-ups

**Done when:**
- every checklist item passes with retained evidence;
- each process resolves once and shares one snapshot;
- no legacy engine or legacy theme interface remains;
- durable rules live in the §13 specs.

**Follow-ups:**
- **F-7:** consolidate `Common::Colors`.
- **F-8:** remove the compatibility header.
- **F-10:** deterministic resolver ordering.
- **F-14:** RedSalamander adopts DxTheme's parser through a lossless typed draft.
- **F-15:** RedSalamander adopts DxTheme's probe and window chrome, verified by a recording DWM shim.
- **F-23:** DxUi's Selector Button recolors a cached brush. Owner: DxUi.
- **Time-based programs** beyond the first one (T8).
- **A suite theme location** for sharing theme files across products.
- **DxUi DLL** (§0.3). The owner scoped "no duplication" to data for now, so this is not pursued here.
