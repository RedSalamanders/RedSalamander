# Documentation Coverage Map

This page is the maintenance map for the public `docs/` tree. Use it when a
feature lands and the durable behavior needs to be reflected in product docs,
developer docs, or both.

> The documentation-completeness pass that produced the developer deep dives, the
> new user/developer pages, the spec/P0 fixes, and the drift-guard tests is
> recorded in
> [`Specs/Plans/Done/Operation_Codex_DocumentationCompleteness_2026-06-18.md`](../Specs/Plans/Done/Operation_Codex_DocumentationCompleteness_2026-06-18.md).
> The only carried-forward items are app-dependent screenshot captures (tracked in
> [res/README.md](res/README.md)). Update both alongside this map.

## Product documentation coverage

| Area | Primary page | Status | Notes |
| --- | --- | --- | --- |
| Getting started and first run | [GettingStarted.md](GettingStarted.md) | Covered | Keep screenshots and default shortcuts aligned with resource changes. |
| Keyboard shortcuts | [KeyboardShortcuts.md](KeyboardShortcuts.md) | Covered | Consolidated F1-F12 x modifier matrix plus navigation/selection/clipboard/view chords. Source of truth is `ShortcutDefaults.cpp`; the `DocumentationDriftContracts` Pester test guards the Alt-arrow rows against drift. |
| Main window, panes, navigation, preview | [MainWindow.md](MainWindow.md), [NavigationAndPaths.md](NavigationAndPaths.md) | Covered | Includes menu-by-menu command map plus current screenshots for Brief, Detailed, Extra Detailed, Thumbnails, and Preview Pane. Update when pane layout, path syntax, history, filters, preview behavior, or main menu labels change. |
| File operations | [FileOperations.md](FileOperations.md) | Covered | Update for copy/move/delete queue behavior, conflict policy, ShellNew, ZIP pack/unpack, and operation diagnostics. |
| Batch rename and change case | [BatchRename.md](BatchRename.md) | Covered | Macro/template vocabulary, search/replace (literal + regex), case changes, manual mode, preview grid, and execution/undo reports. Update when MaskSyntax tokens or the rename engine change. |
| Find files | [FindFiles.md](FindFiles.md) | Covered | Add backend-specific limits when new search providers land. |
| Compare directories | [CompareDirectories.md](CompareDirectories.md) | Covered | Keep option names, sync behavior, and performance notes aligned with selftests/specs. |
| Connections and remote file systems | [Connections.md](Connections.md), [RemoteFileSystems.md](RemoteFileSystems.md), [CloudDrives.md](CloudDrives.md), [S3AndS3Table.md](S3AndS3Table.md) | Covered | Split service-specific caveats into the service page, not the overview. |
| Preferences, themes, settings | [Preferences.md](Preferences.md), [Themes.md](Themes.md), [SettingsFile.md](SettingsFile.md) | Covered | Includes page-by-page option reference. Settings schema changes should update both user workflows and advanced JSON examples. |
| Plugins and viewers | [Plugins.md](Plugins.md), [Viewers.md](Viewers.md) | Covered | Includes viewer menu map. Add new plugin IDs, capabilities, association behavior, and viewer-specific troubleshooting. |
| Monitor and diagnostics | [Monitor.md](Monitor.md), [Troubleshooting.md](Troubleshooting.md) | Covered | Monitor now documents the capture/filter/compare workflow, the six message-type filters and presets, log open/save, and command-line options; Troubleshooting adds diagnostics capture, startup switches, and crash quarantine. Developer internals in [dev/Diagnostics.md](dev/Diagnostics.md). |
| Installer and Winget publishing | [WingetIntegration.md](WingetIntegration.md) | Covered | Keep package metadata and submission steps aligned with `Installer/winget/`. |

## Developer documentation coverage

| Area | Primary page | Status | Notes |
| --- | --- | --- | --- |
| Developer entry point | [DeveloperGuide.md](DeveloperGuide.md) | Covered | Now includes 17 subsystem deep dives (app shell, command routing, FolderView, navigation, file-operations engine, plugin host & bridge, file-system plugins, viewers, DxUi, settings, search, compare, batch rename, connections, theming/preferences, diagnostics, localization/build) plus build, specs, tests, and source layout. |
| Shared DirectX UI layer | [DxUi.md](DxUi.md) | Covered | Covers background, host lifecycle, theme/background usage, examples, and test expectations; the DeveloperGuide DxUi deep dive cross-links it. |
| Localization and resources | [dev/Localization.md](dev/Localization.md) | Covered | .rc ownership, satellite DLL layout, lookup/fallback chain, positional-placeholder rules, add-a-string/add-a-culture workflows, and the `ResourceLocalizationContracts` test gate. |
| Subsystem deep-dive pages | `docs/dev/*` (FolderView, FileOperationsEngine, PluginHostModel, FileSystemPlugins, Search, CompareDirectories, SettingsStore-Internals, Diagnostics) | Covered | Deeper implementation contracts that outgrew a DeveloperGuide section. Keep each aligned with its matching `Specs/` contract. |
| Architecture specs | `Specs/` | Existing | Specs remain the authoritative contract for subsystem behavior. Link from docs when public or onboarding value exists. |
| Agent/project rules | [../AGENTS.md](../AGENTS.md) | Existing | Keep repo-wide engineering rules here, not duplicated deeply in user docs. |

## Missing or thin areas

- **Remaining work is app-dependent screenshot capture only.** The prose, specs,
  and developer pages from the completeness pass have landed (developer deep
  dives, the new user/developer pages, the spec/P0 fixes). The named capture
  backlog — the 10 Preferences subpages, `find-files.png`, the Files-menu dialog
  captures, and the viewer-menu captures — is prioritized in
  [res/README.md](res/README.md) and the Done plan; capture requires a running,
  sanitized build.
- Three orphaned screenshots are flagged in [res/README.md](res/README.md):
  `file-operations-popup-2.png`, `file-operations-popup-3.png`, and
  `preferences-plugins-2.png` — wire each into a page or remove it.
- Drift guards are now enforced by `Tools/Tests/DocumentationDriftContracts.Tests.ps1`,
  which fails the Tools Pester suite if the Alt-arrow shortcut docs/spec drift
  from `ShortcutDefaults.cpp`, or if `connections.allowInsecureTlsInAutomation`,
  the Windows Hello keys, or the Monitor filter-mask default fall out of
  `Specs/SettingsStore.schema.json`.
- User docs should be re-audited against `RedSalamander.rc` whenever menu labels
  or Preferences resource strings change (the audit source of truth).

## Update checklist

- Add or update the user-facing topic page when behavior is visible in the app.
- Add or update [DeveloperGuide.md](DeveloperGuide.md), [DxUi.md](DxUi.md), or a
  new developer page when the change affects how contributors should build,
  test, theme, extend, or debug the product.
- Keep links relative so the same Markdown works in GitHub, GitHub Pages, and a
  local checkout.
- For perf-sensitive behavior, link the user/developer explanation to the
  authoritative spec or archived test run under `Specs/TestRuns/` when useful.
- When adding a settings key or changing a default shortcut, update the matching
  doc/spec in the same change and keep
  `Tools/Tests/DocumentationDriftContracts.Tests.ps1` green; add a new
  `docs/dev/*` page for any implementation contract that outgrows a
  DeveloperGuide section.

