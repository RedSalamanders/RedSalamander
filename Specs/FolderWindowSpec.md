# FolderWindow Specification (Dual Pane)

## Overview

`FolderWindow` is the host container for file browsing in `RedSalamander`. In v1 it implements a **dual-pane** layout:

- **Left pane**: `NavigationView` + `FolderView` + per-pane **Status Bar**
- **Right pane**: `NavigationView` + `FolderView` + per-pane **Status Bar**
- A vertical splitter between panes for horizontal resizing

Each pane has its own navigation state (current path) and its own FolderView state (display mode + sort + status bar visibility). Folder history is a **global MRU list** shared across panes.

## Implementation Files

FolderWindow implementation is intentionally split into focused translation units for easier navigation:

- `RedSalamander/FolderWindow.cpp`: window class registration, creation/destruction, and top-level `WndProc` routing.
- `RedSalamander/FolderWindow.Layout.cpp`: layout calculation, child window positioning, painting, split ratio, status bar visibility, and DPI handling.
- `RedSalamander/FolderWindow.Interaction.cpp`: focus routing and splitter interaction (mouse/cursor/capture).
- `RedSalamander/FolderWindow.StatusBar.cpp`: per-pane status bar custom paint + hit testing + `UpdatePaneStatusBar`.
- `RedSalamander/FolderWindow.SelectionSize.cpp`: async selection folder-bytes traversal + worker thread + UI notifications.
- `RedSalamander/FolderWindow.FileOperations.cpp`: FolderWindow bridge for file-ops commands + completion handler.
- `RedSalamander/FolderWindow.FileOperations.State.cpp`: file operation task scheduling/queueing + worker execution.
- `RedSalamander/FolderWindow.FileOperations.Dialog.cpp`: file operation progress dialog + UI updates.
- `RedSalamander/FolderWindow.FileOperationsInternal.h`: private file-ops state (`FileOperationState` + `Task`) shared by the file-ops TUs.
- `RedSalamander/FolderWindow.FileSystem.cpp`: plugin selection, path parsing/formatting, history/view/sort state, and NavigationView↔FolderView path synchronization.
- `RedSalamander/FolderWindowInternal.h`: private shared declarations/constants for the above `FolderWindow.*.cpp` files.

## Extension Points (Scoped Windows)

Some features embed a `FolderWindow` but require additional host integration (e.g., Compare Directories). FolderWindow exposes optional callbacks and APIs for these scoped windows:

- **Pane callbacks**: path changed, enumeration completed, and per-item details text provider.
- **Selection helpers**: apply selection based on display-name predicates.
- **Informational tasks**: create/update/dismiss read-only task cards in the File Operations popup for background work that is not a file operation (see `Specs/FileOperationsSpec.md`).

## File System Plugin (v2)

FolderWindow uses `IFileSystem` plugins selected per-pane:
- Each pane can have its own active `IFileSystem` (no cross-pane interference).
- The plugin is selected by the **path prefix** (the plugin short ID / scheme): `<shortId>:<pluginPath>`.
  - Example: `fk:/` selects plugin `fk` and passes `/` to the plugin.
- Plugins that need a per-instance “mount context” (archives, FTP, S3, etc.) use: `<shortId>:<instanceContext>|<pluginPath>`.
  - Example: `7z:C:\\Downloads\\archive.zip|/` mounts `archive.zip` as the root and passes `/` to the plugin.
  - `|` is chosen as the delimiter because it is not valid in Windows paths.
  - When the mount context changes and the plugin supports `IFileSystemInitialize`, FolderWindow calls `Initialize()` and clears cached enumerations for that instance.
- If no prefix is present:
  - Windows absolute paths (`C:\...`, `\\server\share\...`, `\\server\`, `\\?\...`) are routed to `file`.
  - Otherwise the pane keeps its current plugin and interprets the path as that plugin’s path.
- If the prefix is unknown or the plugin is unavailable, FolderWindow falls back to the default plugin (`plugins.currentFileSystemPluginId`, typically `builtin/file-system`).
- When a pane switches plugins, FolderWindow cancels enumeration for that pane and (if the previous plugin is no longer used by any pane) purges `DirectoryInfoCache` entries for the old plugin to allow safe DLL unload.
- When switching to the `file` plugin, non-absolute paths are replaced with the default Windows drive root.
- FolderWindow maintains both:
  - A **canonical location** string for persistence/history (`file`: Windows path; non-`file`: `<shortId>:<pluginPath>` or `<shortId>:<instanceContext>|<pluginPath>`).
  - A **pluginPath** passed to `FolderView`/viewers/`IFileSystem` (never includes `<shortId>:` and never includes the mount delimiter `|`).

### Host-reserved navigation prefixes

FolderWindow also recognizes host-reserved, non-plugin prefixes for Connection Manager and URI input:

- Connection Manager routing:
  - `nav:<connectionName>` / `nav://<connectionName>`
  - `@conn:<connectionName>` (alias; includes `@conn:@quick` for Quick Connect)
  - These resolve to: `<pluginShortId>:/@conn:<connectionName><path>` where `<path>` is the profile’s `initialPath` (or an override path when provided).
  - If `<connectionName>` is empty, FolderWindow opens the Connection Manager dialog instead of failing navigation.
- File URIs:
  - `file:` URIs (ex: `file:///C:/Windows/`, `file://server/share/`) are percent-decoded (UTF-8) and routed to the `file` plugin.

Navigation notes:
- `/@conn:<connectionName>/` is treated as a terminal root; navigate-up does not climb above it.

## Error UI

FolderWindow is responsible for routing user-action errors to the most relevant UI surface:

- **Pane-scoped alerts**: non-fatal errors that are contextual to a specific pane are shown as an in-pane alert overlay in that pane’s `FolderView` (shared `RedSalamander::Ui::AlertOverlay` renderer).
  - Examples: “both panes must use the same file system”, “create directory unsupported/failed”, “settings save failed”.
- **Window/app modal dialogs**: confirmations and startup/fatal errors remain blocking dialogs (scoped to the appropriate owner window).
- **Inline validation in modal dialogs**: input validation errors are displayed within the active dialog/control (no secondary message box).

### Drive connect/disconnect

- FolderWindow refreshes affected panes on `WM_DEVICECHANGE` volume events (USB removal, mapped-drive removal) to force a background re-enumeration and surface the FolderView “Disconnected” in-pane overlay when the location is no longer available.
- FolderWindow also subscribes to network interface changes (`NotifyIpInterfaceChange`) and refreshes panes browsing UNC paths or `DRIVE_REMOTE` drives (debounced) so network disconnect/reconnect is detected without polling or user-triggered revalidation.

## Layout

- The window is split horizontally into two panes with a vertical splitter between them.
- Each pane contains:
  - `NavigationView` at the top (fixed height in DIPs)
  - `FolderView` filling the remaining height (above the status bar)
  - **Status Bar** at the bottom (optional per pane)

### Split Ratio

- `folders.layout.splitRatio` controls the divider position as a fraction of the available width (excluding the splitter width).
  - `0.0..1.0` (clamped)
  - Default: `0.5` (equal split)
- There is no minimum pane width: the splitter can be moved all the way to either edge (a pane may be `0px` wide).

### User Interaction

- Dragging the splitter updates the split ratio continuously.
- Double-clicking the splitter resets the ratio to `0.5` (equal split).
- `cmd/pane/zoomPanel` toggles the focused pane between maximized (splitter at the edge) and restored (using `folders.layout.zoomRestoreSplitRatio`).
  - If the user drags the splitter while maximized, the restore state is cleared; the next toggle maximizes again.
- The splitter renders a small centered grip handle to indicate it is draggable (see `Specs/VisualStyleSpec.md`).

## Active/Focused Pane

- The "focused pane" is the pane that contains the current keyboard focus (either its `NavigationView` or its `FolderView`).
- Keyboard accelerators that target "the active pane" apply to the focused pane.

## Keyboard Management

The canonical shortcut and routing spec is `Specs/CommandMenuKeyboardSpec.md`.

FolderWindow responsibilities:
- Define “focused pane” vs “active pane” routing and apply pane-targeted commands accordingly.
- Support Commander-style pane switching (`Tab`) and function-key operations (F2/F3/F5/F6/F7/F8).
- Keep existing sort/view accelerators (`Ctrl+F3..F6`, `Alt+2/3`) targeting the focused pane.

## Status Bar (per pane)

Each pane optionally shows a status bar at the bottom. It is a distinct control per pane (not shared across panes).

### Contents

- **Left part**: selection summary for that pane:
  - Selected **files** count
  - Selected **folders** count
  - Total selected **bytes**:
    - File sizes are summed directly.
    - Folder sizes are computed by traversing the folder subtree (all descendants) and asynchronously when explicitly requested (via the Insert selection workflow). The traversal MUST be implemented iteratively (explicit stack/queue; no call-stack recursion); while pending, the status bar shows a localized “calculating” indicator and the **current bytes computed so far**.
    - Until folder subtree bytes are available (never requested, or canceled), folder bytes show a localized **unknown** placeholder.
  - When exactly **one** item is selected, show item details instead:
    - Folder: `DIR • <size?> • YYYY-MM-DD HH:MM • <attrs>` (size may be unknown or pending)
    - File: `<size> • YYYY-MM-DD HH:MM • <attrs>`
- **Right part**: current sort indicator for that pane. This region is clickable and opens the pane’s **Sort by** menu.
  - The right part is **always visible** (fixed width) because it is a persistent clickable target.
  - When the pane is **unsorted**, the indicator shows a small placeholder glyph (localized via resources) rather than disappearing.
    - Resource: `IDS_STATUS_SORT_INDICATOR`
  - When the window is too narrow, the left part truncates with an ellipsis but the right indicator remains visible.

Example (wide):

| Left part (selection summary) | Right part (sort) |
|---|---|
| `No selection` | `↑◰` |
| `3 files: 7.42 MB selected` | `↑◰` |

Example (narrow):

| Left part (selection summary) | Right part (sort) |
|---|---|
| `No selection` | `↑◰` |
| `3 files: 7.42 M...` | `↑◰` |

### Visual style

- Status bar uses the application `MenuTheme` (background/text), with a subtle separator line between `FolderView` and the status bar.
- The status bar renders a `2 DIP` focus indicator line at its top edge:
  - Focused pane: uses the theme accent color; in Rainbow mode the hue advances each time pane focus changes.
  - Unfocused pane: uses the normal separator color.
- The pane splitter applies to the status bar region as well: each pane owns its own status bar control and the splitter visually and logically separates them (no overlap).
- Implementation note: status bars are themed via custom paint (subclass `WM_PAINT`) rather than `SBT_OWNERDRAW`/`WM_DRAWITEM`, to avoid message-routing issues and to match menu theming.

### Interaction

- Clicking the sort indicator opens the same sort menu as the pane menu bar (`Left → Sort by` or `Right → Sort by`), anchored to the sort indicator region (right-aligned to its edge).
- Hovering the sort indicator shows a tooltip indicating it is clickable.
- Status bar visibility is persisted per pane.

## Menus

FolderWindow exposes two top-level pane menus in the application menu bar:

- **Left** (left side)
- **Right** (right-justified)

**Localization requirement**: the menu structure and static labels are declared in resources (`RedSalamander/RedSalamander.rc`). Runtime code must only rebuild the dynamic **History** submenu entries (see `Specs/LocalizationSpec.md`).

Each menu controls its corresponding pane.

FolderWindow also integrates with the top-level **Plugins** menu:
- The dynamic plugin list applies to the focused pane’s active `IFileSystem` plugin.
- `Manage Plugins...` opens the plugin manager dialog.

### Sort By

Submenu items:

- None (`Ctrl+F2`)
- Name (`Ctrl+F3`)
- Extension (`Ctrl+F4`)
- Time (`Ctrl+F5`)
- Size (`Ctrl+F6`)
- Attributes (no default shortcut)

Sort behavior:

- Directories are listed before files.
- Sort supports **direction** and an **unsorted** state.
- Selecting **None** sets the pane to the **unsorted** state and restores the initial order for the current directory snapshot.
- Default directions:
  - Name / Extension / Attributes: ascending
  - Time / Size: descending (newest/largest first)
- Reselecting the same sort key toggles direction: **default direction ↔ opposite direction**.
- Selecting a different sort key selects that key’s default direction.
- The pane sort menu displays a key glyph for each sort option; the active sort also shows an arrow indicator for the current direction. The status bar mirrors the same arrow + key glyph.
- Arrow semantics:
  - `↑` = ascending
  - `↓` = descending
- Status bar sort indicator examples (right-side glyph):
  - Name: `↑≣` (A→Z), `↓≣` (Z→A)
  - Extension: `↑ⓔ` (A→Z), `↓ⓔ` (Z→A)
  - Size: `↓◲` (big→small), `↑◰` (small→big)
  - Time: `↓⏱` (newest→oldest), `↑⏱` (oldest→newest)
  - Attributes: `↑Ⓐ` (ascending), `↓Ⓐ` (descending)
  - Unsorted: placeholder glyph (still clickable; does not disappear)

### Status Bar toggles

Status bar visibility must be controllable from:

- **View → Pane**:
  - Status Bar (Left) (checkable)
  - Status Bar (Right) (checkable)
- **Pane menus**:
  - `Left` menu contains a checkable **Status Bar** item controlling the left pane.
  - `Right` menu contains a checkable **Status Bar** item controlling the right pane.

### Display As

Submenu items:

- Brief (`Alt+2`)
- Detailed (`Alt+3`)

Detailed mode (v1) is multi-line per item: name + a single details line.

### History

- A dynamic submenu showing the global folder history (most recent first, bounded by `folders.historyMax`, default `20`, clamped `1..50`).
- The current folder is shown with a checkmark.
- Selecting an entry navigates that pane to the chosen path.

## Persistence

FolderWindow state is persisted in settings (see `Specs/SettingsStoreSpec.md`):

- `folders.active`
- `folders.layout.splitRatio`
- `folders.layout.zoomedPane`
- `folders.layout.zoomRestoreSplitRatio`
- `folders.historyMax`
- `folders.history`
- `folders.items[]` per-pane:
  - `current`
  - `view.display`
  - `view.sortBy`
  - `view.sortDirection`
  - `view.statusBarVisible`
