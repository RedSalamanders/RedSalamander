<!-- Generated 2026-06-18 by a multi-agent code+docs review (17 subsystem mappers + 4 user-doc auditors + synthesis/critic). -->

> **Status: COMPLETED — 2026-06-18.** Executed and moved to `Specs/Plans/Done/`. The DeveloperGuide expansion (17 subsystem deep dives) landed; 11 new pages were created (KeyboardShortcuts, BatchRename, Localization, and eight `docs/dev/*` pages); 13 existing user pages were expanded; the P0 schema + keyboard-spec fixes and all queued spec accuracy fixes were applied; and two CI drift-prevention gates now run in the Tools Pester suite. The one carried-forward item is **app-dependent screenshot capture** (needs a running, sanitized build; tracked in `docs/res/README.md`). See the **Closeout** section at the end for the full landed/deferred breakdown. The plan body below is preserved as the historical record.
# RedSalamander Documentation-Completeness Plan

## Executive summary

The RedSalamander `docs/` tree is unusually mature for a user guide: 23 pages, a coverage map (`docs/DocumentationMap.md`), a tracked screenshot backlog (`docs/res/README.md`), and a menu-by-menu command map, with no broken image references. `docs/UserGuide.md` is already comprehensive (~31KB) and nearly every flagged user command is at least named. The two biggest structural gaps are (1) the absence of a developer-facing architecture layer — there is no `docs/dev/` directory at all, and `docs/DeveloperGuide.md` is a thin entry point currently being expanded in this same pass with deep subsystem sections — and (2) the lack of a consolidated keyboard-shortcut reference plus a large, unprioritized screenshot backlog. This plan covers everything that **still remains after** the in-flight DeveloperGuide expansion: thin/missing user content, dedicated developer pages that warrant their own file beyond a DeveloperGuide section, two confirmed spec/code drift bugs, the screenshot backlog, and ongoing maintenance rules.

> **In-flight context (do not re-do):** `docs/DeveloperGuide.md` is being expanded in the same pass with deep sections for Application shell, Command routing, FolderView, Navigation, File Operations engine, Plugin host & bridge, File-system plugins, Viewers, DxUi, Settings, Search, Compare, Batch Rename, Connections, Theming/Preferences, Diagnostics, and Localization/Build (18 promised subsystems). Items below that are "covered by the DeveloperGuide expansion" are listed only where a **dedicated page** or a **spec fix** is the better home, or where the developer detail is deep enough to outgrow a single guide section.
>
> **This plan does not assume the DeveloperGuide expansion landed correctly.** A gate item (#A) verifies that the promised sections actually exist and cover the 18 subsystems before any `docs/dev/*` decisions are finalized. In particular **Application shell** and **Command routing** are otherwise served *only* by that expansion and have no dedicated page in this plan — so the verification gate is the only thing keeping them from silently falling through.

---

## P0 — Land-first correctness fixes (tiny effort, high severity)

These are the two genuine correctness defects. They are S-effort, reversible-free, and make the shipped artifact wrong today (the schema rejects a valid settings file; the spec misleads readers about real key bindings). They are pulled out of the main table so they are **not** deprioritized behind the L-effort documentation pages that share the P1 label. Do these first.

| # | Item | Audience | Location | Effort | Notes |
|---|------|----------|----------|--------|-------|
| 1 | Add `connections.allowInsecureTlsInAutomation` (boolean, default false) to `connectionsSettings` `$defs` | developer | `Specs/SettingsStore.schema.json` | S | **Confirmed bug.** Key read/written 5× in `Common/Common/SettingsStore.cpp` and edited in Preferences→Advanced, but absent from schema; `additionalProperties:false` makes the schema reject valid files. Note: `bypassWindowsHello` (schema line 1489) and `windowsHelloReauthTimeoutMinute` (line 1495) ARE already present — only `allowInsecureTlsInAutomation` is missing. Also add to `Core_SettingsStore.md` Connections key list. |
| 2 | Fix stale Alt-arrow rows in "Default shortcuts.folderView bindings (implemented)" table | developer | `Specs/UI/UI_CommandMenuKeyboard.md` (~lines 936–962) | S | **Confirmed drift.** Code of record (`ShortcutDefaults.cpp` 254–257): Alt+Up=goToPreviousSelectedName, Alt+Down=goToNextSelectedName, Alt+Left=historyBack, Alt+Right=historyForward. Table is wrong and omits Left/Right rows. Root cause is command-routing drift (see #B). |

---

## Track A — User documentation completeness

User docs are broad and mostly accurate. The remaining user-facing work is: a handful of menu items with **zero** coverage, several thin one-line entries that need a sentence of behavior, accuracy fixes in `docs/Preferences.md` (one confirmed, one needing verification — see #19), a new consolidated **Keyboard Shortcuts** page (the headline user gap), expanding `docs/Monitor.md` (the thinnest substantive page, self-flagged Partial), expanding the Batch Rename / Change Case user entries, and clearing the screenshot backlog.

## Track B — Developer documentation depth

There is no `docs/dev/` directory yet. The DeveloperGuide expansion supplies guide-level sections, but several subsystems have implementation contracts deep enough (and reused enough) to merit dedicated pages, and there are the **two confirmed spec/code drift defects** in the P0 tier above that must be fixed regardless of prose. A dedicated audit item (#B) covers the command-routing mapping (`CommandRegistry` → `IDM_*` → command strings) as a unit, since both the #2 Alt-arrow drift and the #44 Restore/Load-Selection id overlap stem from that path.

---

## Verification gates (do before deciding which dev pages to write)

| # | Item | Audience | Location | Priority | Effort | Notes |
|---|------|----------|----------|----------|--------|-------|
| A | Verify the DeveloperGuide expansion actually landed and covers all 18 subsystems | developer | `docs/DeveloperGuide.md` | **P1** | S | Diff the shipped section list against the 18 promised subsystems. **Specifically confirm the Application-shell section (WinMain/message-loop/single-instance/startup switches) and Command-routing section (CommandRegistry dispatch, RC-menu→command-id mapping, ShortcutManager wiring) exist** — these have no dedicated page in this plan and are otherwise asserted, not verified. Only after this diff decide which `docs/dev/*` pages are still needed. |
| B | Audit the command-routing path as a unit (CommandRegistry.cpp → `IDM_*` → command strings) | developer | `docs/DeveloperGuide.md` (Command routing section) or `docs/dev/CommandRouting.md` if it outgrows a section | **P1** | M | The #2 (Alt-arrow command strings) and #44 (Restore vs Load Selection both → `IDM_PANE_SELECTION_RESTORE`) findings both originate here. Document the registry/dispatch mapping and the RC-menu→command-id→command-string chain so future drift is catchable. Feeds the CI gate in #CI-1. |
| C | Confirm the five indirectly-referenced pages were audited | user/dev | `docs/RemoteFileSystems.md`, `docs/CloudDrives.md`, `docs/Themes.md`, `docs/DxUi.md`, `docs/README.md` | **P2** | S | These five pages receive **no dedicated work item** elsewhere in this plan (RemoteFileSystems/CloudDrives only implicitly via Plugins; Themes only via spec #50; DxUi and docs/README not touched). Explicitly audit each and either record "audited, complete" or open a follow-up item. Do not let them silently fall through. |

---

## Prioritized work table (deduped across all findings)

Priority: **P1** = headline structural gap or highest-impact missing page. **P2** = real gap a reader will hit. **P3** = thin/cosmetic/nice-to-have. (The two P0 correctness bugs are in their own land-first tier above.) Effort: S ≈ <1h, M ≈ 1–3h, L ≈ half-day+.

| # | Item | Audience | Suggested location (file) | Priority | Effort | Notes |
|---|------|----------|---------------------------|----------|--------|-------|
| 3 | Create consolidated **Keyboard Shortcuts** reference page | user | `docs/KeyboardShortcuts.md` (new) | **P1** | L | Headline user gap: no single shortcut page. Derive from `ShortcutDefaults::CreateDefaultShortcuts()` (record) cross-checked with the spec's Function Bar + folderView tables. Reproduce the F1–F12 × modifier matrix; link from UserGuide + DocumentationMap. Should be machine-generated, not hand-maintained (see #CI-1). |
| 4 | Create dedicated **Batch Rename** page (user vocabulary + developer engine) | both | `docs/BatchRename.md` (new) | **P1** | L | No Batch Rename docs anywhere. User: macro/template vocabulary (`{name}` `{stem}` `{ext}` `{parent}` `{relativeFolderFlat}` `{counter:000}` `{date:...}` `{time:...}` `{created}` `{index}`, regex helpers). Dev: engine/window/execution split, deepest-first + dependency-layer + temp-hop algorithm, threading, `FileSystemRenameBatch`. Link from DeveloperGuide + DocumentationMap. |
| 5 | Expand `docs/Monitor.md` (thinnest page, self-flagged Partial) | user | `docs/Monitor.md` | **P1** | M | Add ETW capture/filter/compare workflow, document the 6 filter toggles (Text/Error/Warning/Info/Perf/Debug) + presets, file open/save-log flow, launch/command-line options, and walk through the open filter menu shown in `monitor.png`. |
| 6 | Create developer **FolderView** architecture page | developer | `docs/dev/FolderView.md` (new) | **P1** | L | MTA enumeration jthread + `_enumerationGeneration` staleness fence + `kFolderViewEnumerateComplete` handoff; zero-copy FolderItem/arena lifetime; deferred-DirectX init (`kFolderViewDeferredInit`, GDI fallback, icon re-queue); column-layout/scroll-stop contract; async icon vs thumbnail pipeline. Deeper than a single DeveloperGuide section. |
| 7 | Create developer **File Operations Engine** page | developer | `docs/dev/FileOperationsEngine.md` (new) | **P1** | L | per-task jthread, EnterOperation start-gating queue, PerItemTaskScheduler, ConflictArbiter; Wait/Parallel vs queue-pause state; "5F early admission" overlap latch; conflict call sequence (BeginConflictPrompt/WaitForConflictDecision/SubmitConflictDecision); cross-FS bridge internals. |
| 8 | Add the 10 Preferences subpage screenshots | user | `docs/res/` → `docs/Preferences.md` | **P1** | L | Largest single backlog cluster: General, Panes, Editors, User Menu, Mouse, File Operations, Compare Directories, Hot Paths, Advanced, plugin-child. Closes the biggest visual gap. |
| 9 | Add `find-files.png` (options + live/completed results) | user | `docs/res/find-files.png` → `docs/FindFiles.md` | **P1** | M | Highest-impact single image; the page currently has zero screenshots. |
| 10 | Document undocumented menu items with zero coverage: Window Menu, Toggle Fullscreen, External Help | user | `docs/MainWindow.md` (View/Help/Commands sections) | **P2** | S | RC lines 327/328/411. Add command name, effect, and shortcut where one exists. |
| 11 | Create developer **Plugin Host Model & Cross-FS Bridge** page | developer | `docs/dev/PluginHostModel.md` (new) | **P2** | L | CanCrossFileSystemCopyMove gating, reader/writer pump, temp-file + PromoteTempToFinalPath atomic commit, serial vs producer/consumer modes; IHost off-thread→FolderWindow marshaling (PostMessage vs SendMessage, kHost* dispatch); FileSystemPluginManager lifecycle. |
| 12 | Create developer **File-System Plugin authoring** page | developer | `docs/dev/FileSystemPlugins.md` (new) | **P2** | L | Factory.cpp skeleton (EnumeratePlugins/Create/GetConfigurationSchema), which optional interfaces to implement, FilesInformation `BuildFromEntries` packing, capabilities JSON contract walkthrough. Step-by-step provider guide the specs lack. |
| 13 | Create developer **Search subsystem** page | developer | `docs/dev/Search.md` (new) | **P2** | L | FindFilesWindow/SearchSessionController/`FileSystemSearchBackend` (the backend type/enum in FindFilesWindow.cpp — there is no `SelectSearchBackend` symbol)/SearchServiceBroker/LocalSearchIndexCore/SqliteIndexStore; named-pipe protocol v3 (STATUS/QUERY/REBUILD/COMPACT wire framing, streaming/cancellation); worker→UI message IDs; schema v2 + volume state codes; clarify DirectoryInfoCache is NOT the search traversal path. |
| 14 | Create developer **Compare Directories** page + fix stale spec sections | developer | `docs/dev/CompareDirectories.md` (new) + `Specs/Core/Core_CompareDirectories.md` | **P2** | L | Session/engine architecture, worker pools, version model, message protocol; decision-cache LRU (300 MB budget, pinning, eviction). Fix spec Implementation Files/Testing (window is now split into `.cpp/.Options.cpp/.Progress.cpp/.Menu.cpp` + `.Internal.h`; self-tests under `SelfTest/CompareDirectories/`). |
| 15 | Create developer **Settings Store internals** page | developer | `docs/dev/SettingsStore-Internals.md` (new) | **P2** | M | Hot-reload control flow (watcher jthread, FindFirstChangeNotificationW, kSettingsFileChanged/kSettingsReloadedFromDisk, SettingsFileStamp dedupe, MergeDiskSettingsWithRuntimeSession); "Bumping schemaVersion" (must equal 16, replace-on-mismatch, v15 not migrated, which Parse*/Serialize* sites change); PrepareForSave drops defaults. |
| 16 | Create developer **Diagnostics** page | developer | `docs/dev/Diagnostics.md` (new) | **P2** | M | How to emit (Debug::Error/Warning/Info/ErrorWithLastError, Perf::Scope, TRACER macros, one-provider-per-module rule); ETW event schema (DebugMessage kDebugKeyword=0x1, PerfScope kPerfKeyword=0x2, provider GUID); consumer pipeline; build/runtime gating table (Debug vs Release, --etw, REDSALAMANDER_DIAGNOSTICS_ETW, RS_DIAGNOSTICS_RUNTIME_OPT_IN, RS_MONITOR_SHOW_INVALID_RECTS); --perf JSONL. |
| 17 | Create developer **Localization** page | developer | `docs/dev/Localization.md` (new) | **P2** | M | .rc ownership, satellite DLL naming/layout under `RedSalamander/Lang/<culture>/`, Localization namespace lookup/fallback, positional-placeholder rules, ResourceLocalizationContracts.Tests.ps1 gate; "Adding a string / culture" workflow (id ranges e.g. 20000–21999, RedConfigure satellite generation). |
| 18 | Document `x-ui-pane` / `x-ui-section` / `x-ui-order` / `x-ui-control` schema vendor-extension convention | developer | `Specs/Core/Core_SettingsStore.md` (new subsection) and/or `Specs/UI/UI_PreferencesDialog.md` | **P2** | M | 82 occurrences drive schema-driven Preferences pages via `SettingsSchemaParser.cpp`; only `x-ui-hidden` is documented today. Enumerate allowed values + mapping. |
| 19 | Verify (then fix if wrong) `docs/Preferences.md` Panes "Thumbnail size" bullet | user | `docs/Preferences.md` (Panes) | **P2** | S | **Downgraded from "confirmed fix" to verify-first.** The schema DOES define a per-pane `thumbnailSizeDip` (enum `[48,64,96,128]`, default 64, title "Thumbnail Size") at `Specs/SettingsStore.schema.json` line 896, parsed/serialized per pane in `Common/Common/SettingsStore.cpp`. So a thumbnail-size control plausibly does belong on the Panes page and the existing bullet may be correct. **Verify against the actual Panes preferences page before deleting anything.** |
| 20 | Fix `docs/Preferences.md` File Operations speed-limit preset list | user | `docs/Preferences.md` (File Operations) | **P2** | S | **Accuracy fix.** Actual presets: Unlimited, 1, 5, 10, 50, 100, 500 MiB/s + Custom (no 1 GiB/s preset; GiB only via Custom). |
| 21 | Add "Auto-dismiss Success" toggle to File Operations page reference | user | `docs/Preferences.md` + `Specs/UI/UI_PreferencesDialog.md` (File Operations contract) | **P2** | S | Control exists (autoDismissSuccess); shown in SettingsFile.md JSON but not in either page-option reference. |
| 22 | Document `bypassWindowsHello` and `windowsHelloReauthTimeoutMinute` settings (incl. 0 = always prompt) | user | `docs/Connections.md` (Security notes) | **P2** | S | Preferences→Advanced→Windows Hello for Connections; not in **user docs**. Note: both keys ARE already in `Specs/SettingsStore.schema.json` (lines 1489/1495) — this is a docs-only gap, distinct from the #1 schema bug. |
| 23 | Expand Batch Rename / Change Case **user** entries | user | `docs/UserGuide.md` (or the new `docs/BatchRename.md` user section) | **P2** | M | Macro vocabulary (see #4); Change Case styles (lower/UPPER/Mixed/partial), whole-name vs name-only vs extension-only targets, Include subdirectories option. |
| 24 | Add `shortcuts-window.png` (Help → Display Shortcuts) | user | `docs/res/shortcuts-window.png` → `docs/GettingStarted.md` | **P2** | S | High value given shortcut-drift risk; GettingStarted points here as source of truth but never shows it. |
| 25 | Document fileOperations diagnostics keys in spec (8 missing) | developer | `Specs/Core/Core_SettingsStore.md` (File Operations Settings) | **P2** | S | Schema defines maxDiagnosticsLogFiles(=14), diagnosticsInfoEnabled, diagnosticsDebugEnabled, maxIssueReportFiles(=60), maxDiagnosticsInMemory, maxDiagnosticsPerFlush, diagnosticsFlushIntervalMs, diagnosticsCleanupIntervalMs; spec lists only 5 keys. |
| 26 | Fix Monitor filter mask default in spec (31 → 63) | developer | `Specs/Core/Core_SettingsStore.md` (Monitor UI State Defaults) | **P2** | S | Schema default is 63 at `Specs/SettingsStore.schema.json` line 1298 (6-bit mask Text/Error/Warning/Info/Perf/Debug, `minimum 0 maximum 63`). `Core_SettingsStore.md` says "31 (all 5 types)" which is wrong — reconcile to 63 / 6 types. **Correct the code citation:** the draft's "63u (SettingsStore.h:201)" is bogus — no such default exists there; the emitting site is in RedSalamanderMonitor, not Common `SettingsStore.h`. Cite the actual Monitor source site. |
| 27 | Add developer note: `SupportsEmbeddedPreviewViewer` hard-codes the embedded-viewer allowlist | developer | code comment in `FolderWindow.Viewers.cpp` + note in `Specs/Plugins/Plugins_ViewerPlugins.md` | **P2** | S | A new embedded viewer must be added to this list to appear in the preview pane; easy to miss. |
| 28 | Add viewer callback threading/lifetime consolidated note | developer | `Specs/Plugins/Plugins_ViewerPlugins.md` (IViewerCallback) + DeveloperGuide | **P2** | S | SetCallback(nullptr,nullptr) synchronous drain; deadlock rule: do not clear the callback from inside ViewerClosed. Today only an inline comment. |
| 29 | Add `compare-options.png` + Files-menu dialog captures | user | `docs/res/` → `docs/CompareDirectories.md`, `docs/MainWindow.md`, `docs/SettingsFile.md` | **P2** | M | Group: compare-options, view-width, change-attributes, change-case, make-file-list, opened-files, shared-directories, shell-new. |
| 30 | Wire in or remove the **3** orphaned screenshots | developer | `docs/res/README.md` + target pages | **P2** | S | **Corrected count: 3, not 4.** `plugins.png` is NOT an orphan — it is referenced in `docs/Plugins.md` (plus `UserGuide.md`/`Preferences.md`). True orphans: `file-operations-popup-2.png`, `file-operations-popup-3.png`, `preferences-plugins-2.png`. Wire into FileOperations/Preferences or mark stale. |
| 31 | Document breadcrumb interactions (separator/sibling dropdown, overflow ellipsis/full-path popup) | user | `docs/NavigationAndPaths.md` (new "Breadcrumb interactions" section) | **P3** | S | Primary mouse interactions currently undocumented. |
| 32 | Mention incremental type-to-search in folder view | user | `docs/MainWindow.md` (Folder view interactions) | **P3** | S | Printable keys jump focus by prefix, highlight substring; Backspace edits; Esc exits. Implemented + specced, not in docs. |
| 33 | Mention Copy pre-calc/transfer overlap (Running vs Calculating) | user | `docs/FileOperations.md` (Preflight/Calculating) | **P3** | S | Users observe different behavior from Move/Delete; estimating totals that reconcile in place. |
| 34 | Add "Copying between different file systems" user subsection | user | `docs/Plugins.md` | **P3** | S | What the cross-FS bridge is, why some copies are blocked (capability allow-lists), interrupted transfers leave best-effort-cleaned temp files. |
| 35 | Reconcile S3 "no recursive delete of virtual folder prefixes" vs `deleteMax:8` | user | `docs/S3AndS3Table.md` (Current limitations) | **P3** | S | **Verify-first:** the plan itself flags this as "possibly stale." Verify against current capabilities JSON before editing the limitation note. |
| 36 | Add ViewerWeb packaging note (Markdown/JSON/Web ship in one optional ViewerWeb.dll requiring WebView2; optional plugins live in `<exeDir>\Plugins`) | user | `docs/Viewers.md` + `docs/Plugins.md` | **P3** | S | Disabling/locating one affects all three. |
| 37 | Document second backend-status line + Debug/Release run separate services | user | `docs/FindFiles.md` (Backends) | **P3** | S | Database readiness/sync progress line; separate services may surprise users running both builds. |
| 38 | Add "Capturing diagnostics" section linking Monitor | user | `docs/Troubleshooting.md` | **P3** | S | Use RedSalamanderMonitor to capture errors/warnings when filing bugs; Release surfaces only Error/Warning by default. |
| 39 | Document schemaVersion field + old-file behavior | user | `docs/SettingsFile.md` (new "Schema version and upgrades") | **P3** | S | Old file → backed up to `.bad.*`, defaults restored, no automatic preference migration. |
| 40 | Add startup command-line switches (--crash-test, --etw, --perf[=PATH], --help) | developer | `docs/Troubleshooting.md` or `docs/Monitor.md` | **P3** | S | In spec, not in docs. |
| 41 | Document crash quarantine for plugin authors (auto-disable on next launch) | developer | `docs/Troubleshooting.md` + `docs/Plugins.md` | **P3** | S | SessionState marker records active plugin ids; OfferPluginDisableIfPreviousCrashDetected disables a suspect plugin. |
| 42 | Document `Tools/Run-AllTests.ps1 -Suite Full` as the canonical local test entry point | developer | `docs/DeveloperGuide.md` (Build and test) | **P2** | S | **Split out and bumped from the old #42.** CLAUDE.md names this as the primary local test command — contributors need it before any other dev workflow, so it outranks the packaging/versioning extras below. |
| 42b | Expand DeveloperGuide build/test extras | developer | `docs/DeveloperGuide.md` (Build and test) | **P3** | S | Packaging switches (-Msix/-Msi/-Zip/-GenerateWingetManifest), ASan Debug, ARM64; versioning note (Tools/Versioning.ps1, -BuildNumber, -OfficialRelease, GITHUB_RUN_NUMBER). |
| 43 | Fix naming mismatch "Connection Manager" vs menu label "Connections Manager…" | user | `docs/UserGuide.md`, `docs/Connections.md` | **P3** | S | Align doc wording with RC label so menu searches succeed. |
| 44 | Clarify Restore vs Load Selection id overlap (both → IDM_PANE_SELECTION_RESTORE) | user | `docs/UserGuide.md` (Selection Tools) | **P3** | S | State whether they are the same operation. Stems from command-routing (see #B). |
| 45 | Flesh out thin Files/Commands user entries: View Width, Pane Menu (vs F10), Alternate View/Edit, View With/Edit With/Edit New File, Select+Calc+Next, Save/Load/Restore Selection nuances | user | `docs/UserGuide.md` + `docs/MainWindow.md` | **P3** | M | Each currently a name-only surface-map entry; add one sentence of behavior each. |
| 46 | Document SettingsFile.md missing schema sections (cross-link, don't duplicate) | user | `docs/SettingsFile.md` | **P3** | S | startup, selectionMasks, batchRename, connections, cache, ui, hotPaths, compareDirectories, shortcuts, mainMenu, theme, monitor. Cross-link to dedicated pages where they exist; add JSON guidance only for selectionMasks/batchRename/startup/cache (no guidance anywhere today). |
| 47 | Label SettingsFile.md fileOperations example values as non-default | user | `docs/SettingsFile.md` | **P3** | S | Example uses 20/20 where schema defaults are 14/60; label as samples or align. |
| 48 | Add remaining backlog captures (selection-mask, pane-filter, quick-search, command-line, sort-menu, hidden-system, user-menu, network-drive, viewer-web-menu, viewer-text-diff) | user | `docs/res/` → respective pages | **P3 (defer)** | L | Group into a single capture pass per res/README capture policy (sanitize, crop real names). **Consider deferring this entire pass out of the completeness milestone** — screenshots are the lowest-correctness-risk, highest-effort work and will dominate the timeline while contributing least to factual accuracy. |
| 49 | Add developer note: ShortcutDefaults migration logic | developer | `Specs/UI/UI_CommandMenuKeyboard.md` (near default bindings) | **P3** | S | permanentDeleteWithValidation→permanentDelete and Ctrl+F2 changeAttributes→sort migrations are undocumented. |
| 50 | Document AppTheme color-override key set (folderView/menu/navigationView/viewer.diff.* etc.) | developer | `Specs/UI/UI_VisualStyle.md` (color-key reference table) | **P3** | M | Mismatched keys are silently dropped; enumerate recognized keys consumed by ApplyAppThemeOverrides. |

---

## New pages to create

- [ ] **`docs/KeyboardShortcuts.md`** — consolidated user-facing F1–F12 × modifier matrix + navigation/clipboard/selection chords, generated/verified against `ShortcutDefaults::CreateDefaultShortcuts()`. (P1, #3)
- [ ] **`docs/BatchRename.md`** — user macro/template vocabulary + Change Case styles; developer engine/window/execution split, deepest-first + dependency-layer + temp-hop algorithm, `FileSystemRenameBatch` contract. (P1, #4)
- [ ] **`docs/dev/Localization.md`** — .rc ownership, satellite DLL layout, lookup/fallback, "Adding a string / culture" workflow, contracts test gate. (P2, #17)
- [ ] **`docs/dev/FolderView.md`** — enumeration threading + arena lifetime, deferred DirectX init, grid layout/hit testing, icon/thumbnail pipeline. (P1, #6)
- [ ] **`docs/dev/FileOperationsEngine.md`** — queue/scheduler/ConflictArbiter architecture, Wait/Parallel & queue-pause, 5F early admission, conflict call sequence, cross-FS bridge. (P1, #7)
- [ ] **`docs/dev/PluginHostModel.md`** — cross-FS bridge pipeline, IHost thread marshaling, FileSystemPluginManager lifecycle, FileActionResolver/Launcher. (P2, #11)
- [ ] **`docs/dev/FileSystemPlugins.md`** — step-by-step provider authoring guide (Factory skeleton, optional interfaces, FilesInformation packing, capabilities JSON). (P2, #12)
- [ ] **`docs/dev/Search.md`** — Find/search architecture, named-pipe protocol v3, worker→UI marshaling, SqliteIndexStore schema v2; DirectoryInfoCache distinction. (P2, #13)
- [ ] **`docs/dev/CompareDirectories.md`** — engine/session architecture, message protocol, decision-cache LRU/eviction/pinning, sync-copy cross-link. (P2, #14)
- [ ] **`docs/dev/SettingsStore-Internals.md`** — hot-reload control flow, "Bumping schemaVersion", PrepareForSave default-drop invariant. (P2, #15)
- [ ] **`docs/dev/Diagnostics.md`** — emit APIs, ETW event schema + GUID, consumer pipeline, build-flag table, perf JSONL. (P2, #16)
- [ ] *(optional)* **`docs/dev/CommandRouting.md`** — only if the Command-routing audit (#B) outgrows a DeveloperGuide section; otherwise keep it as a guide subsection.
- [ ] *(create the `docs/dev/` directory as part of the first dev page; it does not exist yet)*

> **Scope guard:** **Application shell** and **Command routing** are delegated to the in-flight `docs/DeveloperGuide.md` expansion and have **no dedicated page here** — gate item #A explicitly verifies those two sections landed, and #B audits the command-routing mapping as a unit, because command-id correctness is the root of the #2 drift bug. NavigationView, ViewerPluginHost, EmbeddedViewers, ConnectionManager, and Theming(AppTheme) developer detail are likewise served by the DeveloperGuide subsystem sections and do **not** need separate `docs/dev/` files unless a section grows past ~1 screen — track them as DeveloperGuide subsections, not new pages. Re-confirm all of this against the #A diff before writing any dev page.

---

## Pages to expand or fix (keyed to existing files)

### `docs/Monitor.md` (P1 — thinnest substantive page)
- [ ] Add ETW capture → filter → compare workflow as a task list.
- [ ] Document the six filter toggles (Text/Error/Warning/Info/Perf/Debug) and any presets.
- [ ] Document file open / save-log flow and launch/command-line options.
- [ ] Walk through the open filter menu shown in `monitor.png`.

### `docs/Preferences.md` (accuracy work — one verify-first, one confirmed)
- [ ] **Verify before deleting** the Panes "Thumbnail size" bullet — the schema defines `thumbnailSizeDip` (line 896), so the bullet may be correct (#19).
- [ ] Correct File Operations speed-limit presets to Unlimited/1/5/10/50/100/500 MiB/s + Custom (no 1 GiB/s preset) (#20).
- [ ] Add the "Auto-dismiss Success" toggle to the File Operations page reference (#21).

### `docs/MainWindow.md` (P2/P3)
- [ ] Document Window Menu, Toggle Fullscreen, External Help (currently zero coverage).
- [ ] Add behavior sentences for Pane Menu (vs F10), Alternate View/Edit, View With/Edit With/Edit New File, Select+Calc+Next, Save/Load/Restore selection nuances.
- [ ] Mention incremental type-to-search.
- [ ] Add Alt+6 (Preview Pane) to the display-mode list (currently stops at Alt+5).

### `docs/Connections.md` (P2/P3)
- [ ] Document `bypassWindowsHello` and `windowsHelloReauthTimeoutMinute` (0 = always prompt) — docs-only gap; both keys already in schema (#22).
- [ ] Spell out Quick Connect secret security implication vs saved profiles.
- [ ] Fix "Connection Manager" → match menu label "Connections Manager…".

### `docs/UserGuide.md` (P2/P3)
- [ ] Expand Batch Rename and Change Case entries (or link to new `docs/BatchRename.md`).
- [ ] Cross-reference the new Keyboard Shortcuts page from "Finding Commands".
- [ ] Clarify Restore vs Load Selection id overlap.

### `docs/NavigationAndPaths.md` (P3)
- [ ] Add a "Breadcrumb interactions" section (separator/sibling dropdown, overflow ellipsis/full-path popup).

### Indirectly-referenced pages — audit explicitly (#C)
- [ ] `docs/RemoteFileSystems.md`, `docs/CloudDrives.md`, `docs/Themes.md` (user page), `docs/DxUi.md`, `docs/README.md`: each receives no other work item — audit and record "complete" or open a follow-up.

### `docs/FileOperations.md` / `docs/Plugins.md` / `docs/Viewers.md` / `docs/S3AndS3Table.md` / `docs/FindFiles.md` / `docs/SettingsFile.md` / `docs/Troubleshooting.md`
- [ ] FileOperations: note Copy pre-calc/transfer overlap (Running vs Calculating).
- [ ] Plugins: add "Copying between different file systems" user subsection.
- [ ] Viewers + Plugins: ViewerWeb packaging/WebView2 note + optional-plugin location; trim viewer detail duplicated between the two pages to a single source.
- [ ] S3AndS3Table: **verify-first**, then reconcile recursive-delete limitation vs `deleteMax:8` (#35).
- [ ] FindFiles: document the backend-status line and Debug/Release separate services.
- [ ] SettingsFile: add schemaVersion/upgrade subsection; cross-link missing schema sections; label example values as non-default.
- [ ] Troubleshooting: add "Capturing diagnostics" (Monitor) section; mention startup command-line switches and crash quarantine.

### `docs/DeveloperGuide.md` (beyond the in-flight subsystem sections)
- [ ] Verify the expansion landed and covers the 18 subsystems incl. Application shell + Command routing (#A).
- [ ] Document `Tools/Run-AllTests.ps1 -Suite Full` as the canonical local test entry point (P2, #42).
- [ ] Expand Build and test extras: packaging switches, ASan Debug, ARM64, versioning (P3, #42b).

### Spec files (developer correctness — do regardless of prose)
- [ ] `Specs/SettingsStore.schema.json`: add `connections.allowInsecureTlsInAutomation` (**P0 bug #1**).
- [ ] `Specs/UI/UI_CommandMenuKeyboard.md`: fix stale Alt-arrow folderView rows (**P0 drift #2**).
- [ ] `Specs/Core/Core_SettingsStore.md`: add 8 fileOperations diagnostics keys; fix Monitor mask default 31→63 (schema line 1298; cite the real Monitor emitting site, not `SettingsStore.h:201`); add Connections key.
- [ ] `Specs/Core/Core_CompareDirectories.md`: fix Implementation Files + Testing sections.
- [ ] `Specs/Core/Core_SettingsStore.md` / `Specs/UI/UI_PreferencesDialog.md`: document `x-ui-*` extension convention.
- [ ] `Specs/UI/UI_VisualStyle.md`: AppTheme color-key reference table.
- [ ] `Specs/Plugins/Plugins_ViewerPlugins.md`: SupportsEmbeddedPreviewViewer allowlist note; IViewerCallback drain/deadlock rule.

---

## Consolidated screenshot backlog

Confirmed on disk: 49 PNGs present, **no broken references**, **3 orphans** (corrected from 4 — `plugins.png` is referenced, not orphaned). Group into capture passes per the `docs/res/README.md` policy (sanitize, theme rules, handle-based capture, crop real names).

**P1 cluster (biggest visual gaps):**
- [ ] 10 Preferences subpages: general, panes, editors, user-menu, mouse, file-operations, compare-directories, hot-paths, advanced, plugin-child → `docs/Preferences.md`
- [ ] `find-files.png` (options + live/completed results) → `docs/FindFiles.md`

**P2:**
- [ ] `shortcuts-window.png` (Help → Display Shortcuts) → `docs/GettingStarted.md`
- [ ] `compare-options.png` → `docs/CompareDirectories.md`
- [ ] Files-menu dialogs: view-width, change-attributes, change-case, make-file-list, opened-files, shared-directories, shell-new → `docs/MainWindow.md` / `docs/SettingsFile.md`
- [ ] `viewer-web-menu.png`, `viewer-text-diff.png` → `docs/Viewers.md`
- [ ] Resolve the **3** orphans (`file-operations-popup-2.png`, `file-operations-popup-3.png`, `preferences-plugins-2.png`): wire in or mark stale in `docs/res/README.md`. (`plugins.png` is already wired in — leave it.)

**P3 (candidate for deferral out of this milestone — see #48):**
- [ ] `selection-mask.png`, `pane-filter.png`, `quick-search.png`, `folder-view-sort-menu.png`, `folder-view-hidden-system.png`, `command-line.png`, `user-menu.png`, `network-drive.png`.
- [ ] Reprioritize the `docs/res/README.md` backlog (currently ~68 png references / 30+ unprioritized capture items).

---

## CI / drift-prevention gates (turn maintenance rules into tests)

The plan repeatedly identifies drift as the root cause (#2 Alt-arrow, #1 schema gap, #44 id overlap). "Never hand-maintained" is necessary but not sufficient — add automated checks so #1- and #2-class bugs cannot recur:

- [ ] **#CI-1 — Shortcut/command-table generation test.** Regenerate the `docs/KeyboardShortcuts.md` matrix and the `UI_CommandMenuKeyboard.md` folderView/function-bar tables from `ShortcutDefaults::CreateDefaultShortcuts()` and fail the build on mismatch. This is the direct guard against #2.
- [ ] **#CI-2 — Schema-vs-code key-coverage test.** Assert every settings key read/written in `Common/Common/SettingsStore.cpp` has a matching property in `Specs/SettingsStore.schema.json` (and vice-versa where `additionalProperties:false` applies). This would have caught `allowInsecureTlsInAutomation` (#1) automatically and is higher leverage than the one-off fix.

---

## Maintenance recommendations

**Keep `docs/DocumentationMap.md` current**
- [ ] Add rows for the new pages (KeyboardShortcuts, BatchRename, Localization, and the `docs/dev/*` set) to the coverage tables.
- [ ] Normalize the status vocabulary — it currently mixes "Covered / Partial / Added / Existing." Pick one scale (e.g. Covered / Partial / Missing) and re-audit every row, including the five indirectly-referenced pages from #C.
- [ ] Re-verify "Added" claims (DeveloperGuide.md, DxUi.md) against actual page state after this pass; flip Monitor.md from Partial once #5 lands.

**RC re-audit trigger**
- [ ] Treat `RedSalamander/RedSalamander.rc` (menus ~108–413, Compare/View ~419–436, context menu ~66–100) as the source of truth for user-facing commands. Re-run the command-coverage audit whenever the RC menu changes; the DocumentationMap update checklist should call this out explicitly.
- [ ] Treat `RedSalamander/ShortcutDefaults.cpp` (`CreateDefaultShortcuts()`) as the source of truth for default chords. The new `docs/KeyboardShortcuts.md` and the spec binding tables must be regenerated/verified against it via #CI-1 — never hand-maintained — so they cannot drift (the #2 Alt-arrow bug is exactly this failure mode).
- [ ] Treat `CommandRegistry.cpp` (`IDM_*` → command-string mapping) as the source of truth for the command-routing audit (#B); regressions there surface as both #2-class and #44-class bugs.

**Spec ↔ docs ↔ code sync rules**
- [ ] When bumping `schemaVersion`, follow the new "Bumping schemaVersion" checklist (#15) and update SettingsFile.md + Core_SettingsStore.md + schema together.
- [ ] Any new persisted settings key must land in `Specs/SettingsStore.schema.json` (with `x-ui-*` annotations if Preferences-edited) in the same change — `additionalProperties:false` means an undeclared key breaks validation (root cause of the #1 bug); #CI-2 enforces this.
- [ ] When a window/engine file is split (as Compare Directories was), update the spec's "Implementation Files" and "Testing" sections in the same PR.
- [ ] **Verify before editing for any "Confirmed" or "Accuracy fix" item that touches user prose.** #19 is the cautionary example — a "confirmed" deletion that would have removed a correct bullet (the `thumbnailSizeDip` control exists in the schema). Apply the same guard to #35 (S3 recursive-delete vs `deleteMax:8`, self-flagged "possibly stale"). Cheap insurance given several of these are S-effort.
- [ ] Specs are normative contracts; `docs/` is the reader-facing layer. When the two disagree, fix the spec first (it is authoritative), then propagate to docs. Where user docs already lead the schema (e.g. Preferences.md mentions insecure-TLS automation), backfill the schema rather than deleting the doc.
- [ ] Avoid two-source-of-truth duplication: trim the viewer detail duplicated between `docs/Plugins.md` and `docs/Viewers.md` to a single canonical location with a pointer from the other.


---

## Closeout (2026-06-18)

### Landed
- **DeveloperGuide expansion (gate #A):** `docs/DeveloperGuide.md` now carries 17 subsystem deep dives including **Application shell** and **Command routing**. Gate #A satisfied; the command-routing mapping (#B) is documented in that section.
- **P0 fixes:** #1 `connections.allowInsecureTlsInAutomation` added to `Specs/SettingsStore.schema.json` and `Core_SettingsStore.md`; #2 Alt-arrow folderView rows corrected in `UI_CommandMenuKeyboard.md` (plus #49 migration note).
- **New user pages:** `docs/KeyboardShortcuts.md` (#3), `docs/BatchRename.md` (#4, #23).
- **New developer pages:** `docs/dev/Localization.md` (#17); `docs/dev/` — `FolderView.md` (#6), `FileOperationsEngine.md` (#7), `PluginHostModel.md` (#11), `FileSystemPlugins.md` (#12), `Search.md` (#13), `CompareDirectories.md` (#14 dev half), `SettingsStore-Internals.md` (#15), `Diagnostics.md` (#16).
- **User page expansions (13):** Monitor #5; MainWindow #10/#32/#45 (+Alt+6); NavigationAndPaths #31; FileOperations #33; Plugins #34/#41; Viewers #36; FindFiles #37; SettingsFile #39/#46/#47; Troubleshooting #38/#40/#41; Preferences #19/#20/#21; Connections #22/#43; UserGuide #23/#43/#44 + KeyboardShortcuts cross-link; S3 #35.
- **Spec fixes:** `Core_SettingsStore.md` #25/#26/#18; `Core_CompareDirectories.md` #14; `UI_VisualStyle.md` #50; `Plugins_ViewerPlugins.md` #27/#28 (plus a two-line comment in `FolderWindow.Viewers.cpp`); `UI_PreferencesDialog.md` #21.
- **CI gates (#CI-1, #CI-2):** `Tools/Tests/DocumentationDriftContracts.Tests.ps1` (8 cases, all green) guards the Alt-arrow shortcut docs/spec against `ShortcutDefaults.cpp`, and the schema against the connections + Windows-Hello keys and the Monitor mask default. Tools Pester case count reconciled 94 -> 102 in `TestInventory.Tests.ps1`, `Specs/Testing/Testing_TestCoverage.md`, and `Tests/README.md`.
- **#30 orphans:** the three true orphans (`file-operations-popup-2.png`, `file-operations-popup-3.png`, `preferences-plugins-2.png`) documented in `docs/res/README.md`; backlog reprioritized. (`plugins.png` confirmed referenced, not an orphan.)
- **DocumentationMap.md:** rows added for KeyboardShortcuts, BatchRename, Localization, and the `docs/dev/*` set; Monitor flipped to Covered; maintenance + drift-test guidance added.

### Resolved differently than drafted (verify-first earned its keep)
- **#19:** the Panes Preferences page does **not** expose a thumbnail-size control (the picker lives in the pane sort/thumbnail context-menu flyout); the incorrect bullet was corrected, not merely confirmed.
- **#20:** speed-limit presets are Unlimited / 1 / 5 / 10 / 50 / 100 / 500 MiB/s + Custom (no GiB/s preset; GiB only via Custom).
- **#44:** the UserGuide Restore-vs-Load-Selection relationship was clarified after verifying command routing.
- **#35:** S3 delete behavior reconciled against the current capabilities.

### Carried forward (out of this milestone — app-dependent or optional)
- **Screenshot captures (#8, #9, #24, #29, #48):** require a running, sanitized build; prioritized in `docs/res/README.md` (P1 = the 10 Preferences subpages + `find-files.png`).
- **Optional `docs/dev/CommandRouting.md`:** only if the command-routing material outgrows its DeveloperGuide section (it currently fits).
- **#42b:** minor DeveloperGuide build/test extras (packaging switches, ASan Debug, ARM64, versioning).

### Verification performed
- `Tools/Tests/DocumentationDriftContracts.Tests.ps1`: 8/8 pass (Pester 3.4.0).
- `Tools/Tests/TestInventory.Tests.ps1`: 5/5 pass (count reconciled to 102).
- All `docs/**/*.md` relative links resolve; zero HTML entities across new/edited pages; `Specs/SettingsStore.schema.json` parses as valid JSON.
- The `FolderWindow.Viewers.cpp` change is comment-only (no compilation impact); it will be covered by the next routine build.

