# Documentation Coverage Map

This page is the maintenance map for the public `docs/` tree. Use it when a
feature lands and the durable behavior needs to be reflected in product docs,
developer docs, or both.

## Product documentation coverage

| Area | Primary page | Status | Notes |
| --- | --- | --- | --- |
| Getting started and first run | [GettingStarted.md](GettingStarted.md) | Covered | Keep screenshots and default shortcuts aligned with resource changes. |
| Main window, panes, navigation, preview | [MainWindow.md](MainWindow.md), [NavigationAndPaths.md](NavigationAndPaths.md) | Covered | Includes menu-by-menu command map plus current screenshots for Brief, Detailed, Extra Detailed, Thumbnails, and Preview Pane. Update when pane layout, path syntax, history, filters, preview behavior, or main menu labels change. |
| File operations | [FileOperations.md](FileOperations.md) | Covered | Update for copy/move/delete queue behavior, conflict policy, ShellNew, ZIP pack/unpack, and operation diagnostics. |
| Find files | [FindFiles.md](FindFiles.md) | Covered | Add backend-specific limits when new search providers land. |
| Compare directories | [CompareDirectories.md](CompareDirectories.md) | Covered | Keep option names, sync behavior, and performance notes aligned with selftests/specs. |
| Connections and remote file systems | [Connections.md](Connections.md), [RemoteFileSystems.md](RemoteFileSystems.md), [CloudDrives.md](CloudDrives.md), [S3AndS3Table.md](S3AndS3Table.md) | Covered | Split service-specific caveats into the service page, not the overview. |
| Preferences, themes, settings | [Preferences.md](Preferences.md), [Themes.md](Themes.md), [SettingsFile.md](SettingsFile.md) | Covered | Includes page-by-page option reference. Settings schema changes should update both user workflows and advanced JSON examples. |
| Plugins and viewers | [Plugins.md](Plugins.md), [Viewers.md](Viewers.md) | Covered | Includes viewer menu map. Add new plugin IDs, capabilities, association behavior, and viewer-specific troubleshooting. |
| Monitor and diagnostics | [Monitor.md](Monitor.md), [Troubleshooting.md](Troubleshooting.md) | Partial | Needs more ETW workflow examples when monitor scenarios stabilize. |
| Installer and Winget publishing | [WingetIntegration.md](WingetIntegration.md) | Covered | Keep package metadata and submission steps aligned with `Installer/winget/`. |

## Developer documentation coverage

| Area | Primary page | Status | Notes |
| --- | --- | --- | --- |
| Developer entry point | [DeveloperGuide.md](DeveloperGuide.md) | Added | Points to build, specs, tests, and source layout. |
| Shared DirectX UI layer | [DxUi.md](DxUi.md) | Added | Covers background, host lifecycle, theme/background usage, examples, and test expectations. |
| Architecture specs | `Specs/` | Existing | Specs remain the authoritative contract for subsystem behavior. Link from docs when public or onboarding value exists. |
| Agent/project rules | [../AGENTS.md](../AGENTS.md) | Existing | Keep repo-wide engineering rules here, not duplicated deeply in user docs. |

## Missing or thin areas

- Monitor docs need a task-oriented flow for capturing, filtering, and comparing
  ETW traces once the command-line workflow is stable.
- Developer docs now cover DxUi, but plugin authoring and viewer authoring still
  deserve dedicated guides if those interfaces become public extension points.
- Screenshot upkeep is manual. [res/README.md](res/README.md) tracks current
  assets, the latest sanitized capture pass, and the named backlog of missing
  captures.
- User-facing command coverage is broad but some implemented workflows are still
  visually thin: selection masks, pane filter and Quick Search, command-line
  insertion, View Width, Change Attributes, Change Case, Make File List, List of
  Opened Files, Shared Directories, ShellNew, User Menu, and network-drive flows.
  Keep these as screenshot/documentation follow-ups unless the command behavior
  changes first.
- User docs should be re-audited against `RedSalamander.rc` whenever menu labels
  or Preferences resource strings change.

## Update checklist

- Add or update the user-facing topic page when behavior is visible in the app.
- Add or update [DeveloperGuide.md](DeveloperGuide.md), [DxUi.md](DxUi.md), or a
  new developer page when the change affects how contributors should build,
  test, theme, extend, or debug the product.
- Keep links relative so the same Markdown works in GitHub, GitHub Pages, and a
  local checkout.
- For perf-sensitive behavior, link the user/developer explanation to the
  authoritative spec or archived test run under `Specs/TestRuns/` when useful.
