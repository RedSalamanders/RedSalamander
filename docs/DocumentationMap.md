# Documentation Coverage Map

Use this map to find the guide and source contract for each product area. It was
reviewed against the command registry, shipped plugin families, Preferences pages
and top-level settings schema on 2026-10-19. Coverage here describes documentation;
it is not runtime, hardware or release qualification.

## Product workflows

| Area | User documentation | Source of behavior |
| --- | --- | --- |
| Installation, first run and shortcuts | [Getting Started](GettingStarted.md), [Keyboard Shortcuts](KeyboardShortcuts.md) | Startup, packaging and `ShortcutDefaults.cpp` |
| Main window, panes, splitter, preview and menus | [Main Window](MainWindow.md), [User Guide](UserGuide.md) | FolderWindow, FolderView and `RedSalamander.rc` |
| Command discovery and configured actions | [Command Palette](CommandPalette.md), [Preferences](Preferences.md) | CommandRegistry, CommandPaletteWindow and file-action settings |
| Paths, history, drives, WSL and bookmarks | [Navigation and Paths](NavigationAndPaths.md) | NavigationView and connection/path resolution |
| Selection, filtering, natural sorting, path text and reports | [Selections and File Lists](SelectionsAndLists.md), [Main Window](MainWindow.md) | FolderView selection/sort/filter, Make File List and clipboard commands |
| Copy, Move, Delete, verification, conflicts and Issues | [File Operations](FileOperations.md), [Troubleshooting](Troubleshooting.md) | [File Operations spec](../Specs/FileSystem/FileSystem_FileOperations.md), [popup spec](../Specs/UI/UI_FileOperationsPopup.md) |
| Batch rename and case changes | [Batch Rename](BatchRename.md), [Selections and File Lists](SelectionsAndLists.md) | BatchRename and MaskSyntax |
| Find and indexed/native/fallback search | [Find Files](FindFiles.md) | [Search spec](../Specs/Core/Core_Search.md), SearchService and search providers |
| Compare, content differences and synchronization | [Compare Directories](CompareDirectories.md) | [Compare spec](../Specs/Core/Core_CompareDirectories.md) |
| Local files, archives and cross-provider transfers | [Plugins](Plugins.md), [File Operations](FileOperations.md) | FileSystem, FileSystem7z and host bridge |
| FTP, SFTP, SCP and IMAP | [Remote File Systems](RemoteFileSystems.md), [Connections](Connections.md) | FileSystemCurl and Connection Manager |
| Google Drive, OneDrive and SharePoint | [Cloud Drives](CloudDrives.md), [Connections](Connections.md) | Google/Microsoft drive plugins and per-provider capability limits |
| S3 objects and read-only S3 Table metadata | [S3 and S3 Table](S3AndS3Table.md) | FileSystemS3 |
| Phones, cameras and portable storage | [Portable Devices](PortableDevices.md), [Connections](Connections.md) | [MTP/PTP spec](../Specs/FileSystem/FileSystem_Mtp.md), FileSystemMtp and device/storage pickers |
| Embedded/floating Terminal, tabs, WSL, history and paste | [Terminal](Terminal.md), [Keyboard Shortcuts](KeyboardShortcuts.md) | [Terminal spec](../Specs/Terminal/Terminal_EmbeddedPlugin.md), Terminal plugin and host integration |
| Text, hex, diff, images and RAW | [Viewers](Viewers.md) | ViewerText and ViewerImgRaw |
| Space map, Explorer/Analyst, table, inspector and refresh | [Viewers](Viewers.md), [User Guide](UserGuide.md) | ViewerSpace and retained scan model |
| PE, SQLite, VLC, Web, JSON and Markdown | [Viewers](Viewers.md), [Plugins](Plugins.md) | ViewerPE, ViewerSqlite, ViewerVLC and ViewerWeb |
| Plugin enable/configure/test/quarantine | [Plugins](Plugins.md), [Troubleshooting](Troubleshooting.md) | Plugin managers, factory metadata and crash recovery |
| Preferences, pointer focus, shortcuts, themes and settings | [Preferences](Preferences.md), [Themes](Themes.md), [Settings File](SettingsFile.md) | Preferences page owners and [settings schema](../Specs/SettingsStore.schema.json) |
| Monitor, ETW, perf, operation logs and crash reports | [Monitor](Monitor.md), [Troubleshooting](Troubleshooting.md) | Independent RedMonitor, diagnostic sinks and File Operations logs |
| Installer and Winget workflows | [Getting Started](GettingStarted.md), [Winget Integration](WingetIntegration.md) | Installer/ZIP/Winget manifests and release tooling |

The [plugin catalog](Plugins.md) lists every shipped file-system protocol and
viewer family, including MTP and Terminal. The [settings section index](SettingsFile.md#settings-sections-at-a-glance)
accounts for every top-level schema key. Parameterized viewer/editor/User Menu
commands are explained by their configuration workflow; configured IDs are not
a fixed list of product commands.

## Developer workflows

| Area | Entry point |
| --- | --- |
| Source layout, startup, command routing and subsystem ownership | [Developer Guide](DeveloperGuide.md) |
| File Operations admission, traversal, workers, conflicts and data safety | [File Operations Engine](dev/FileOperationsEngine.md) |
| Plugin host/bridge and file-system provider contracts | [Plugin Host Model](dev/PluginHostModel.md), [File System Plugins](dev/FileSystemPlugins.md) |
| Folder rendering, navigation and input | [FolderView](dev/FolderView.md), [DxUi](DxUi.md) |
| Search and Compare internals | [Search](dev/Search.md), [Compare Directories](dev/CompareDirectories.md) |
| Settings and localization | [Settings Internals](dev/SettingsStore-Internals.md), [Localization](dev/Localization.md) |
| ETW, diagnostics and performance evidence | [Diagnostics](dev/Diagnostics.md), [Tests](../Tests/README.md) |
| Scoped tests, sandbox cleanup, receipts and screenshot publication | [Tests, Evidence and Screenshots](dev/TestingAndScreenshots.md), [Screenshot Inventory](res/README.md) |
| Durable rules and authoritative contracts | [AGENTS.md](../AGENTS.md), [Specs](../Specs/README.md) |

## Refresh scope and remaining visual work

The [Visual Tour](VisualTour.md) now spans 45 native product frames beyond the ten
File Operations examples: Search, thirteen Space lifecycle/workspace states, actual viewers, plugin/device
setup, all Preferences pages, command discovery and synthetic workspace themes.
Each frame has adjacent provenance and full-frame privacy review. These examples
illustrate presentation; they do not claim live provider/device or playback coverage.

The October review adds missing portable-device, Terminal, command-discovery and
selection/report workflows. It corrects stale pre-calculation architecture,
File Operations/Mouse Preferences, diagnostic logging and test-output guidance.
The earlier June completeness pass is historical context, not proof that later
features are documented.

Current File Operations examples use native harness frames with invented data,
including open counting, a fully wrapped terminal note and the actual Issues
pane. A whole-frame/OCR privacy review of `docs/res` retires illustrations with
workspace/device-run paths, actual filenames or obsolete UI. Retired images are
not replaced with redaction labels or placeholder screenshots. The backlog in
[res/README](res/README.md) separates written coverage from missing visuals.

Historical qualification galleries and completed-plan evidence retain their own
source/build and review scope. They are not automatically republished as current
user-guide illustrations.

## Maintain this map

- Check commands and menu labels against `CommandRegistry.cpp`, `ShortcutDefaults.cpp`
  and `RedSalamander.rc`; account for new public commands in a workflow.
- Check plugin factories/protocols and Preferences pages against their guides,
  including read-only and unsupported-operation limits.
- Add top-level settings keys to the section index and explain their owning UI.
- Keep lasting requirements in the owning spec. Prose, screenshots and focused
  tests do not establish Full qualification.
- Use harness captures, invented identifying data, complete-frame review and
  provenance. Follow the [screenshot contract](../Specs/Testing/Testing_SelfTests.md#documentation-screenshot-scenarios).
- Run documentation drift, link/schema coverage and affected native checks.
  Foreground input needs an agreed time; these static captures do not need it.

The supplemental gallery closes twenty-one further frames: eleven utility/shortcut
scenes, three native Terminal layouts, five fake-provider browse scenes and two
real generated-media playback states. Loopback provider work and synthetic VT
transcripts retain explicit qualification limits; no personal data enters captures.
