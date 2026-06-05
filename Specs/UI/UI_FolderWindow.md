# FolderWindow Specification (Dual Pane)

## Overview

`FolderWindow` is the host container for file browsing in `RedSalamander`. In v1 it implements a **dual-pane** layout:

- **Left pane**: optional `NavigationView` + `FolderView` + optional per-pane **Filter Bar** + optional per-pane **Status Bar**, or a Folder/Preview tabbed host when previewing the right pane
- **Right pane**: optional `NavigationView` + `FolderView` + optional per-pane **Filter Bar** + optional per-pane **Status Bar**, or a Folder/Preview tabbed host when previewing the left pane
- A vertical splitter between panes for horizontal resizing
- A bottom application **Function Bar** when function-key chrome is enabled

Each pane has its own navigation state (current path) and its own FolderView state (display mode + sort + view-option chrome visibility). Folder history is a **global MRU list** shared across panes.

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
- **Informational tasks**: create/update/dismiss read-only task cards in the File Operations popup for background work that is not a file operation (see `Specs/FileSystem/FileSystem_FileOperations.md`).

## File System Plugin (v2)

FolderWindow uses `IFileSystem` plugins selected per-pane:
- Each pane can have its own active `IFileSystem` (no cross-pane interference).
- The plugin is selected by the **path prefix** (the plugin short ID / scheme): `<shortId>:<pluginPath>`.
  - Example: `fk:/` selects plugin `fk` and passes `/` to the plugin.
- Plugins that need a per-instance “mount context” (archives, FTP, S3, etc.) use: `<shortId>:<instanceContext>|<pluginPath>`.
  - Example: `7z:C:\\Downloads\\archive.zip|/` mounts `archive.zip` as the root and passes `/` to the plugin.
  - `|` is chosen as the delimiter because it is not valid in Windows paths.
  - When the mount context changes and the plugin supports `IFileSystemInitialize`, FolderWindow calls `Initialize()` and re-registers the provider in `DirectoryInfoCache`.
  - Cache and watch state are keyed by logical context (`pluginId + normalized instanceContext`), so panes sharing the same mount context also share cache/watch state even if they use different live COM instances.
- If no prefix is present:
  - Windows absolute paths (`C:\...`, `\\server\share\...`, `\\server\`, `\\?\...`) are routed to `file`.
  - Otherwise the pane keeps its current plugin and interprets the path as that plugin’s path.
- If the prefix is unknown or the plugin is unavailable, FolderWindow falls back to the default plugin (`plugins.currentFileSystemPluginId`, typically `builtin/file-system`).
- When a pane switches plugins, FolderWindow cancels enumeration for that pane and unregisters the previous provider from `DirectoryInfoCache`. Logical-context state drains only when the last provider for that context goes away.
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

## Cross-Pane Mutation Propagation

Visible pane correctness is mandatory across all built-in file system plugins.

- Watchers are an invalidation source, not the source of truth.
- Successful host-side mutations and watch callbacks are routed through `DirectoryInfoCache`, which posts a directory-impact payload to visible panes.
- If a visible folder's direct child list changes, `FolderView` refreshes automatically with local debounce.
- If the current folder or one of its ancestors is deleted or moved away, `FolderWindow` relocates the pane to the nearest surviving ancestor in the same logical context.
- If no ancestor survives, the pane falls back to the plugin root.
- When a mounted context resolves to a Windows backing path, local file-system rename or move retargets the mount and keeps the internal plugin path when possible.
- When that backing item is deleted, the pane exits the mount and navigates to the nearest surviving local ancestor or the default local root.
- Focus memory is preserved where possible so cross-pane updates feel stable instead of jumping to an unrelated item.
- Off-screen folders may be marked dirty without eager re-enumeration.

## Error UI

FolderWindow is responsible for routing user-action errors to the most relevant UI surface:

- **Pane-scoped alerts**: non-fatal errors that are contextual to a specific pane are shown as an in-pane alert overlay in that pane’s `FolderView` (shared `RedSalamander::Ui::AlertOverlay` renderer).
  - Examples: “both panes must use the same file system”, “create directory unsupported/failed”, “settings save failed”.
- **Window/app modal dialogs**: confirmations and startup/fatal errors remain blocking dialogs (scoped to the appropriate owner window).
- **Inline validation in modal dialogs**: input validation errors are displayed within the active dialog/control (no secondary message box).

## Item Properties Dialog

- `Alt+Enter` / Properties opens a themed DxUi dialog for the focused pane item when the active filesystem supports `IFileSystemIO::GetItemProperties`.
- Property content is presented as card sections with the section name above, not inside, the card. Each regular section uses two-column rows (`key` / `value`) with a larger section title, so values are scannable without relying on a plain multiline text dump.
- Compact sections such as Timestamps and Attributes are laid out side by side when the dialog is wide enough, and stack vertically on narrow widths.
- Long values, especially item names and paths, wrap within the value column and grow the containing row/card instead of clipping or drawing past the card edge.
- The built-in local filesystem General section avoids duplicate rows for parent path, root, and extension; it shows name, full path, type, and file size when applicable.
- Built-in local filesystem file sizes use the same compact units as the folder view and include the exact byte count in parentheses for sizes of 1 KB and larger.
- The dialog footer uses a compact Close button; Enter and Escape route to Close.
- When named streams are present, the dialog includes a Streams card. It lists every named stream surfaced by the filesystem with stream name and size. The unnamed default data stream is not shown.
- If no named streams are present, the dialog omits the Streams section rather than showing an empty card.
- If the active filesystem implements `IFileSystemItemStreams` and a stream row has `canRemove=true`, the row exposes an enabled Remove button. Remove deletes that stream via the filesystem interface, refreshes properties in place, and removes the row from the card.
- If stream deletion fails, the dialog keeps the row visible and shows an error message with the failing HRESULT.

## Shared Directories Dialog

- `cmd/pane/shares` opens `RedSalamander.SharedDirectoriesWindow`, a modeless FolderWindow-owned DxUi window with a shared-directories `DxUi::Grid`, Open Path, Manage, and Close commands.
- The dialog is part of FolderWindow's themed modeless surface area: it inherits the active app theme, avoids visible native dialog-template controls, exposes the grid rows through UI Automation, and closes through deferred FolderWindow message handling so its WndProc never destroys its own state while processing a command.

### Drive connect/disconnect

- FolderWindow refreshes affected panes on `WM_DEVICECHANGE` volume events (USB removal, mapped-drive removal) to force a background re-enumeration and surface the FolderView “Disconnected” in-pane overlay when the location is no longer available.
- FolderWindow also subscribes to network interface changes (`NotifyIpInterfaceChange`) and refreshes panes browsing UNC paths or `DRIVE_REMOTE` drives (debounced) so network disconnect/reconnect is detected without polling or user-triggered revalidation.

## Layout

- The window is split horizontally into two panes with a vertical splitter between them.
- Each pane contains:
  - `NavigationView` at the top when navigation bar visibility is enabled
  - `FolderView` filling the remaining height above the optional filter/status bars
  - **Filter Bar** below the folder view when enabled for that pane; it is a themed DxUi host, not a native `STATIC` label
  - **Status Bar** at the bottom when enabled for that pane
- When preview mode is open, the opposite pane becomes a themed DxUi tabbed host. The tab strip is placed at the top of that pane, the `Preview` tab hides the host pane's normal navigation/filter/folder/status chrome, and the preview content fills the host pane down to the function bar or the bottom edge when the function bar is hidden.
- When enabled, the application **Function Bar** is a visible child HWND laid out below the pane area; toggling chrome must keep its HWND visibility and layout in sync.

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
- The splitter renders a small centered grip handle to indicate it is draggable (see `Specs/UI/UI_VisualStyle.md`).
- The splitter also exposes two click targets sized to the navigation bar height:
  - Restored split: the top target maximizes the left pane and moves focus into it; the bottom target maximizes the right pane and moves focus into it.
  - Arrows point in the direction the splitter will move to create that maximized state: maximizing left shows a right chevron, maximizing right shows a left chevron.
  - Maximized split: both targets maximize/focus the hidden pane instead of restoring the current pane, and both arrows point in the direction the splitter will move to switch sides.
  - Arrows use compact chevrons colored exactly like the centered grip dots.
  - These arrow targets are not drag handles; dragging and double-click reset remain on the non-arrow splitter region.

## Active/Focused Pane

- The "focused pane" is the pane that contains the current keyboard focus (either its `NavigationView` or its `FolderView`).
- Keyboard accelerators that target "the active pane" apply to the focused pane.
- When keyboard focus is outside both pane `FolderView` controls but still belongs to main-window chrome or a transient menu, `FolderWindow`'s active pane is the fallback pane for focus restoration.
- Modeless external actions that call into `FolderWindow` to navigate/open in the active pane must commit the same-folder preparation or new target path before activating the main window or moving keyboard focus into the `FolderView`, so foreground/focus changes cannot visibly outrun the navigation work.
- Any user navigation action inside a pane's `NavigationView` activates that pane before changing navigation state. This includes the drive/menu button, history and disk-info dropdowns, breadcrumb segment clicks, sibling menus, and menu-selected path changes. If the action originates from an unfocused pane, keyboard focus returns to that pane's `FolderView` rather than staying in the previously focused pane.
- A first `Escape` from main-window chrome or transient menu focus MUST restore keyboard focus to the active pane's `FolderView` without changing that pane's selection. A subsequent `Escape` delivered while the `FolderView` already has focus uses normal `FolderView` behavior.

## Keyboard Management

The canonical shortcut and routing spec is `Specs/UI/UI_CommandMenuKeyboard.md`.

FolderWindow responsibilities:
- Define “focused pane” vs “active pane” routing and apply pane-targeted commands accordingly.
- Support Commander-style pane switching (`Tab`) and function-key operations (F2/F3/F5/F6/F7/F8).
- Keep existing sort/view accelerators (`Ctrl+F3..F6`, `Alt+2/3`) targeting the focused pane.

## Function Bar

- The function bar renders command short display names, not full menu display names, so labels remain useful when space is limited.
- When function-key chrome is enabled, the function bar must have a visible HWND and non-empty layout rectangle at the bottom of `FolderWindow`.
- A visible function bar must also paint its themed background, separators/key glyphs, and command labels. Its Direct2D HWND render target is DPI-aware, so Win32 pixel rectangles MUST be converted to DIPs before drawing; scaled-DPI layouts must not produce a blank strip.
- Function Bar key glyph, modifier, hit-test, and label measurement MUST use DirectWrite text formats / `DxUi.Typography` specs. The Function Bar must not create or select `HFONT` for app-owned visible text measurement.
- The Function Bar visible paint path MUST NOT bind an `ID2D1DCRenderTarget` to a paint HDC for app-owned text/chrome. It uses the same Direct2D HWND-target model as the pane status bars so toggle/show operations cannot leave an unpainted child strip.
- Every command must expose a localized short display resource via the `IDS_CMD_SHORT_BASE + IDS_CMD_*` convention. Examples: `Make Directory` -> `MakeDir`, `User Menu` -> `UsrMenu`, `Sort by Time` -> `ByTime`.
- Resource ids `20000..21999` are reserved for command short labels.
- Short labels should be concise enough for the function bar target width; the command registry self-test guards that every command has a non-empty short label and that representative labels match the expected compact text.
- Full command display names remain the source for menus, Preferences, and shortcut lists. The function bar may fall back to the full display name only for unknown/custom command IDs.

## Find Files and Directories

`cmd/pane/find` is implemented as an independent modeless `Find Files and Directories` window.

FolderWindow responsibilities:
- resolve the target pane for `cmd/pane/find`,
- expose `Alt+F7` and `Ctrl+F` as default bindings for the same command,
- provide the current plugin, instance context, and root path as the default search scope,
- allow multiple independent modeless Find windows,
- route result actions back into the pane navigation/open flow,
- route queued pane-level viewer/editor commands posted through a `FolderView` back to the main command dispatcher after the folder view has focus,
- keep search execution off the UI thread.

The Find window is part of the existing themed host UI stack and must follow runtime theme changes and normal keyboard/focus behavior. See `Specs/Core/Core_Search.md`.

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
- When no items are selected, show the current focused item details using the same file/folder format. Use the localized `No selection` text only when the pane has no current focused item.
- When there is no selection, keyboard or mouse focus changes inside the `FolderView` MUST refresh the status bar immediately so the displayed item details follow the current focused item.
- **Right part**: current sort indicator for that pane. This region is clickable and opens the pane’s **Sort by** menu.
  - The right part is **always visible** (fixed width) because it is a persistent clickable target.
  - When the pane is **unsorted**, the indicator shows a small placeholder glyph (localized via resources) rather than disappearing.
    - Resource: `IDS_STATUS_SORT_INDICATOR`
  - When the window is too narrow, the left part truncates with an ellipsis but the right indicator remains visible.

Example (wide):

| Left part (selection summary) | Right part (sort) |
|---|---|
| `1.24 MB • 2026-04-25 12:19 • A` | `↑◰` |
| `No selection` | `↑◰` |
| `3 files: 7.42 MB selected` | `↑◰` |

Example (narrow):

| Left part (selection summary) | Right part (sort) |
|---|---|
| `1.24 MB • 2026-04-25...` | `↑◰` |
| `3 files: 7.42 M...` | `↑◰` |

### Visual style

- Status bar uses the application `MenuTheme` (background/text), with a subtle separator line between `FolderView` and the status bar.
- The status bar renders a `2 DIP` focus indicator line at its top edge:
  - Focused pane: uses the theme accent color; in Rainbow mode the hue advances each time pane focus changes.
  - Unfocused pane: uses the normal separator color.
- Status bar text uses `12 DIP` UI text typography. The inactive pane's status text is dimmed while remaining readable; high-contrast mode preserves system contrast instead of applying dimming.
- Status bar Direct2D paint uses DPI-aware render targets. All Win32 client/part rectangles are physical pixels and MUST be converted to DIPs before drawing text or focus lines, so `12 DIP` status text remains vertically visible and unclipped at scaled DPI.
- The pane splitter applies to the status bar region as well: each pane owns its own status bar control and the splitter visually and logically separates them (no overlap).
- Implementation note: status bars are themed via custom paint (subclass `WM_PAINT`) rather than `SBT_OWNERDRAW`/`WM_DRAWITEM`, to avoid message-routing issues and to match menu theming.
- Status bar text paint and part sizing MUST stay on the shared `DxUi.Typography` / DirectWrite path; do not reintroduce ambient `DEFAULT_GUI_FONT`, `HFONT`-selected GDI paint, or visible GDI text measurement for this surface.
- The status-bar HWND is a transitional host only: it must not own an `HFONT`, receive `WM_SETFONT`, or use the native font message path. Typography ownership belongs to the DirectWrite render resources used by `FolderWindow.StatusBar.cpp`.

### Interaction

- Clicking the sort indicator opens the same sort menu as the pane menu bar (`Left → Sort by` or `Right → Sort by`), anchored to the sort indicator region (right-aligned to its edge) and placed above the status bar so the menu grows into the pane area. The popup MUST render through the shared DxUI popup menu path rather than a native `TrackPopupMenu` popup.
- The pane sort popup also exposes a **Thumbnail size** discrete slider with four stops: Small `48 DIP`, Medium `64 DIP`, Large `96 DIP`, and Extra Large `128 DIP`. The slider targets the pane whose status-bar sort region opened the popup, updates only that pane, and keeps left/right pane sizes independent.
- Hovering the sort indicator shows a tooltip indicating it is clickable.
- Status bar visibility is persisted per pane.

## Pane View Options

Pane view option commands are defined by `Specs/UI/UI_CommandMenuKeyboard.md` and implemented by `FolderWindow`/`FolderView` state.

- File-extension visibility is per pane and display-only. Hidden extensions affect rendered names, text measurement, and quick-search highlighting, but item operations, command-line insertion, clipboard actions, and plugin calls continue to use real names and full paths.
- Thumbnails is a per-pane display mode, exclusive with Brief/Detailed/Extra Detailed. When selected, the owning `FolderView` uses the pane's configured thumbnail size, queues thumbnail extraction for visible local items only, applies completed thumbnails on the UI thread, preserves image aspect ratio, and keeps icon fallback available for unsupported providers, extraction failures, bad files, or pending work.
- Preview Pane is an independent per-source toggle. The preview host resolves the configured viewer plugin for the focused item and uses it when the plugin supports embedded preview hosting; when saved viewer associations are missing or only resolve the default text viewer, the host consults the built-in embedded viewer defaults and uses only a specific embedded-capable match, not the catch-all default text rule. If the newly focused item resolves to the same embedded viewer plugin as the current preview, the host reuses that viewer instance and refreshes it with the new open context; it replaces the preview window only when resolution chooses a different plugin or the refresh fails. Focus-driven preview refresh is coalesced with a short debounce so fast navigation across image/media/PDF/raw/text items opens only the latest focused candidate instead of churning every intermediate viewer, while explicit preview open/tab actions remain immediate. It must not force images, media, web/document formats, PE files, SQLite files, folders, or text files through `builtin/viewer-text` when a better configured or built-in embedded viewer is available. Known embedded-capable built-ins are `builtin/viewer-text`, `builtin/viewer-space`, `builtin/viewer-imgraw`, `builtin/viewer-vlc`, `builtin/viewer-web`, `builtin/viewer-json`, `builtin/viewer-markdown`, `builtin/viewer-pe`, and `builtin/viewer-sqlite`. If embedded preview is unavailable or fails to open, the host falls back to a focused item Properties preview from `IFileSystemIO::GetItemProperties`; if properties are unavailable too, it may show localized empty/unsupported/binary fallback text. Embedded preview viewers MUST NOT take keyboard focus from the source pane, preview close/replacement must persist viewer plugin configuration changes such as media volume/mute, and rapid same-plugin or cross-plugin preview navigation must not block the UI on slow media/player teardown.
- Preview pane visibility is tied to the source pane. The preview is hosted in the opposite pane with compact themed DxUi `Folder` and `Preview` tabs; the Preview tab tracks the focused item from the source pane, and the Folder tab keeps the host pane usable without closing preview mode. The tabs are real pointer targets, are painted as attached pane tabs rather than pill buttons, and do not steal keyboard focus from the source pane. Inactive tabs have no border, selected tabs blend into the pane below with square lower corners, the Folder tab tooltip shows the host pane path after the standard hover delay, and the Preview tab close glyph is visible when selected or hovered. The close glyph closes preview mode. Closing preview removes the tabs and restores normal host pane layout. File preview uses an embedded viewer-capable plugin window for supported files. Embedded media preview, including audio-only files and any visualizer-capable player path, must keep all playback and media output parented inside the Preview host and must not create unowned/top-level player or visualization windows; standalone viewers may keep their normal visualization behavior. Embedded viewers that expose standalone menu options must expose the Preview-appropriate subset from a themed right-click context menu, without creating a hidden menubar gap at the top of the embedded preview window. Preview context menus must remove standalone-only actions such as Exit, Open, and internal other-file navigation, remove empty submenus/separator runs left by that pruning, and omit viewer shortcut labels because keyboard focus remains in the source pane. Focused files and folders without a specific embedded preview show the same normalized Properties contract used by the Properties dialog as compact DxUi cards inside a vertical `ScrollPanel`; long values wrap, the scrollbar appears only when content exceeds the preview area, wheel scrolling must increase the preview properties scroll offset without stealing source-pane focus, and Rainbow theme applies restrained rainbow section-header accents while high-contrast mode keeps system-safe colors.
- Navigation bar visibility is per pane. Explicit left/right menu entries affect their named pane; shortcut invocation affects the active pane. Commands that focus the address bar must first show the owning navigation bar, then focus the address edit.
- Filter bar visibility is per pane. The DxUi filter bar is a compact alternate surface for the same workflow as `cmd/pane/filter`: it starts with an editable history combo whose placeholder/accessibility name is Filter, is bound to `selectionMasks.filterHistory`, and uses a right-side Use Filter toggle. It must not add a redundant static Filter label. Typing a non-empty mask applies it live with the same wildcard parsing and pane refresh behavior as the dialog without automatically opening the history dropdown; Enter or choosing a history entry stores the mask in history; turning Use Filter off keeps the text but disables filtering. The combo history must still open as a visible flyout when explicitly requested even though the bar itself is compact. The bar updates when filters are applied, cleared, restored from history, or panes are swapped, and it remains separate from transient Quick Search.
- Generic status bar shortcut routing applies to the active pane and shares the existing left/right status-bar implementation.
- Pane view option changes must update child HWND visibility, recompute layout, preserve folder-view focus when a bar is hidden, refresh menu check states, and persist through `folders.items[].view.*`.

## Menus

FolderWindow exposes two top-level pane menus in the application menu bar:

- **Left** (left side)
- **Right** (right-justified)

**Localization requirement**: the menu structure and static labels are declared in resources (`RedSalamander/RedSalamander.rc`). Runtime code must only rebuild the dynamic **History** submenu entries (see `Specs/Core/Core_Localization.md`).

Each menu controls its corresponding pane.

FolderWindow also integrates with the top-level **Plugins** menu:
- The dynamic plugin list applies to the focused pane’s active `IFileSystem` plugin.
- `Manage Plugins...` opens the plugin manager dialog.

## Pane Prompts

The pane-scoped file-system prompts for change-attributes, change-case, and selection-mask actions use owned DxUi prompt windows, not legacy Win32 dialog-template routes.

- `Change Attributes...` MUST route through the owned prompt-window path in `RedSalamander/FolderWindow.FileSystem.Commands.Part.cpp`. The prompt MUST expose tri-state attribute toggles, date/time rows, alternate-stream removal, and an Include subdirectories option that is enabled only when a folder is in scope. Recursive Change Attributes runs asynchronously as a File Operations informational task with enumeration/apply status.
- `Change Case...` MUST route through the owned prompt-window path in `RedSalamander/FolderWindow.FileSystem.cpp`.
- The selection-mask commands (`Select by Mask...` / `Unselect by Mask...`) MUST route through the owned prompt-window path in `RedSalamander/FolderWindow.FileSystem.cpp`.
- These prompts MUST expose the shared DxUi host contract and MUST NOT reintroduce legacy `DialogBoxParamW`, owner-draw button/toggle, or extra visible combo/input-frame fallback surfaces for the active product path.
- Validation anchors for this contract are `cmd_pane_changeAttributes_options_dialog_uses_dxui_not_win32_template`, `cmd_pane_changeAttributes_recurse_applies_datetime_with_progress`, `cmd_pane_changeCase_prompt_*`, `cmd_pane_changeCase_dialog`, and `cmd_pane_selection_mask_dialogs`.

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

- **Pane menus**:
  - `Left` → `Show` contains a checkable **Status Bar** item controlling the left pane.
  - `Right` → `Show` contains a checkable **Status Bar** item controlling the right pane.

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

FolderWindow state is persisted in settings (see `Specs/Core/Core_SettingsStore.md`):

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
  - `view.thumbnailSizeDip`
  - `view.fileExtensionsVisible`
  - `view.navigationBarVisible`
  - `view.filterBarVisible`
  - `view.statusBarVisible`
