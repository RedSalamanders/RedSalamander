# RedSalamander User Guide

This guide is the complete user-facing tour of RedSalamander: the dual-pane file manager, file operations, search, compare, connections, themes, monitor, and all built-in plugin families.

Screenshots in this guide were captured from the running application and cropped to the UI area being explained. Theme examples intentionally use Light, Dark, Rainbow, and High Contrast App.

## Product Overview

RedSalamander is a Windows dual-pane file manager built for fast keyboard-driven work, remote file-system access, rich file viewers, and themeable DirectX-rendered UI.

![RedSalamander main window](res/main-window.png)

The application is organized around two independent folder panes. Most commands use the focused pane as the source and the other pane as the destination.

Core capabilities:

- Browse local folders, drives, known folders, UNC paths, archives, remote systems, cloud drives, and object stores.
- Copy, move, rename, delete, and inspect files across compatible file systems.
- Search files and directories with native, indexed, and fallback search backends.
- Compare two directory trees and synchronize differences.
- Open files with viewer plugins for text, hex, images, RAW files, SQLite databases, PE files, web content, media, and folder space usage.
- Store reusable remote/cloud profiles in the Connections Manager.
- Customize themes, keyboard shortcuts, viewers, plugins, file operations, hot paths, and advanced settings.
- Watch application diagnostics with RedSalamanderMonitor.

## Quick Start

1. Launch `RedSalamander.exe`.
2. Use the left and right panes to navigate to source and destination folders.
3. Press `Tab` or use **View -> Switch Pane Focus** to change the active pane.
4. Press `F5` to copy selected items to the other pane.
5. Press `F6` to move selected items to the other pane.
6. Press `F3` to view the selected file.
7. Press `Alt+F7` or use **Commands -> Find Files and Directories...** to search from the focused pane.
8. Press `Ctrl+F10` or use **Commands -> Compare Directories...** to compare the two sides.

The function bar at the bottom shows the main keyboard actions:

- `F2` Rename
- `F3` View
- `F4` Edit
- `F5` Copy
- `F6` Move
- `F7` MakeDir
- `F8` Delete
- `F9` User Menu
- `F10` Menu
- `F11` Connect
- `F12` Disconnect

The function bar also updates as you hold `Ctrl`, `Alt`, or `Shift` to show the alternate command for each key. For the complete default-binding reference, including the modifier rows, see [Keyboard Shortcuts](KeyboardShortcuts.md).

## Main Window

The main window contains:

- **Menu bar**: pane, file, edit, command, plugin, view, and help commands.
- **Navigation bars**: one per pane, with drive/menu access, breadcrumbs, path entry, history, and disk information.
- **Folder views**: the two file panels where you select, sort, and operate on files.
- **Filter bars**: optional per-pane filter summaries for persistent folder filters.
- **Status bars**: selected item, directory, and operation context for each pane.
- **Function bar**: keyboard command strip.

### Pane Focus

Only one pane is focused at a time. Commands such as View, Copy, Move, Delete, Search, and Properties operate on the focused pane or its selection.

Useful pane commands:

- **View -> Switch Pane Focus** switches focus.
- **Left/Right -> Maximize/Restore Pane** gives one side more room.
- **Left/Right -> Swap Panes** exchanges the two pane locations.
- **Left/Right -> Refresh** reloads the focused pane.
- **Left/Right -> Brief / Detailed / Extra Detailed** changes the folder view layout.
- **Left/Right -> Sort by Name / Extension / Time / Size / Attributes / None** changes ordering.

### Folder View

Folder views support common file-manager interactions:

- Single and multi-selection.
- Keyboard navigation.
- Enter to open or execute.
- `Backspace` to go to the parent folder.
- Context menu for file-system-specific actions.
- Folder refresh.
- Display modes for compact or detailed work.
- Optional hidden/system item visibility.

### Disk Space and Drive Actions

The right side of each navigation bar can show disk information for local `file:` locations: capacity text and a usage bar.

![Navigation bar with disk information](res/navigation-view.png)

Use the drive/menu dropdown to reach:

- Known folders such as Desktop, Documents, Downloads, Pictures, Music, Videos, and OneDrive.
- WSL distributions when available.
- Saved Connections Manager profiles.
- Local drive roots with current capacity information.
- Disk actions such as **Disk Properties** and **Disk Cleanup** when the active item is a local drive and Windows exposes those actions.

### Display, Sort, and Pane Layout

Folder views can be switched between **Brief**, **Detailed**, **Extra Detailed**, and **Thumbnails** layouts. Sorting is available by name, extension, time, size, attributes, or no sort.

Use the **Left** and **Right** menus to toggle per-pane display and chrome options. **Show -> File Extensions** changes only displayed labels in the named pane; file operations still use real names and paths. **Thumbnails** uses larger item visuals and background thumbnail loading with shell/WIC thumbnail fallback to icons. The sort indicator at the bottom right of each pane opens a sort popup that also includes a per-pane thumbnail-size slider. **Preview Pane** opens a Folder/Preview tabbed host in the opposite pane, follows the named pane's focused item, and shows scrollable Properties cards when no specific embedded preview is available. **Filter Bar** shows the current persistent folder filter below the named pane. **Navigation Bar** targets the named pane, and address-bar focus commands show the bar first when needed. Persistent options are saved per pane.

![Brief and Detailed folder views](res/folder-view-brief-detailed.png)

![Extra Detailed and Thumbnails folder views](res/folder-view-extra-thumbnails.png)

![Preview Pane with text content](res/preview-pane-text.png)

Default shortcuts:

- `Alt+2`: Brief
- `Alt+3`: Detailed
- `Alt+4`: Extra Detailed
- `Alt+5`: Thumbnails
- `Alt+6`: Preview Pane
- `Ctrl+F2`: Sort None
- `Ctrl+F3`: Sort Name
- `Ctrl+F4`: Sort Extension
- `Ctrl+F5`: Sort Time
- `Ctrl+F6`: Sort Size
- `Ctrl+U`: Swap panes
- `Ctrl+F11`: Maximize/restore the focused pane

### Status and Function Bars

The status bars summarize the current folder and selection. They show details such as selected files, folders, bytes, timestamps, attributes, and sort state where available.

The function bar shows the current `F1` through `F12` bindings. It updates when modifier keys are active and can be hidden from **View -> Function Bar**.

### Finding Commands

When a workflow is not visible in the first screen, check these command surfaces:

- The **Left** and **Right** menus for pane-specific view, sort, history, filter, and layout commands.
- The **Files** menu for open/view/edit, copy/move/delete, properties, security, archive, attribute, case, and ShellNew commands.
- The **Edit** menu for selection masks, saved selections, hide/show-name filters, clipboard commands, and path-copy helpers.
- The **Commands** menu for search, compare, occupied space, make file list, opened-files list, shared directories, connections, shell/command-line helpers, and File Explorer shortcuts.
- The **Plugins** menu for plugin management and plugin-provided commands.
- The **View** menu for theme, fullscreen, File Operations Failed Items, menu/function bar toggles, and Preferences.
- `F1` / **Help -> Display Shortcuts** for the shortcut list that matches the current command set.

See [Main Window & Panes](MainWindow.md) for the full menu-by-menu map, and [Keyboard Shortcuts](KeyboardShortcuts.md) for the consolidated default-binding reference.

## Navigation

The navigation bar accepts local paths, UNC paths, file URIs, plugin-prefixed paths, mounted archive paths, and saved connection aliases.

![Drive/menu dropdown](res/drive-menu.png)

### Common Navigation Commands

- `Ctrl+L` or `Alt+D`: focus the address bar.
- `Alt+F1`: open the left pane drive/menu dropdown.
- `Alt+F2`: open the right pane drive/menu dropdown.
- `Alt+Left` / `Alt+Right`: history back / forward.
- `Alt+F12`: show folder history.
- `Backspace`: parent folder.
- `Shift+Backspace`: root folder.
- `Ctrl+.`: set current pane path from the other pane.
- `Ctrl+F12`: filter the current folder.
- `Shift+Space`: start Quick Search in the current pane. Type to highlight every matching name, move to the best prefix match first, use arrows to move through matches, `Esc` to clear, or `Enter` to accept the focused item.
- `Ctrl+1` through `Ctrl+0`: go to Hot Path slots 1 through 10.
- `Ctrl+Shift+1` through `Ctrl+Shift+0`: assign the current folder to a Hot Path slot.

### Path Forms

Local Windows paths:

```text
C:\Windows\
\\server\share\folder\
```

`file:` URIs:

```text
file:///C:/Windows/
file://server/share/
```

Plugin-prefixed paths:

```text
ftp:/
s3:/
s3table:/
7z:C:\Downloads\archive.zip|/
```

Connections Manager aliases:

```text
nav:MySftpServer
nav://MySftpServer
@conn:MySftpServer
s3://@conn/MyAwsS3/logs/
sharepoint:/@conn:TeamSite/Shared Documents/
```

If you type `nav:`, `nav://`, or `@conn:` without a name, RedSalamander opens the Connections Manager.

## File Management

Use **Files** and **Edit** menu commands, the context menu, or the function keys to work with the current selection.

### Opening, Viewing, and Inspecting

- **Open/Execute**: runs the focused file or opens a folder.
- **View** (`F3`): opens the viewer action associated with the focused file.
- **View Width**: opens the viewer using width-oriented behavior where supported.
- **Properties**: shows item properties exposed by the active file system.
- **Context Menu**: opens the item shell/context menu when available.
- **Context Menu (Current Directory)**: opens the Windows shell menu for the active pane's current local folder.
- **Security**: opens the Windows Security property page for the focused local item.
- **Change Attributes** (`Ctrl+F8`): changes read-only, hidden, system, and archive attributes for the selection and can remove alternate data streams, then reports what changed.
- **Go to Shortcut or Link Target**: follows a focused `.lnk`, local `file:` `.url`, junction, mount point, or directory symlink. File targets open their parent folder and focus the target file; folder targets open directly. Web URLs stay in place and report that they are unsupported for pane navigation.
- **List of Opened Files** (`Alt+F11`): shows files currently open through RedSalamander viewers, external viewer/editor launches, and the Preview pane. The list shows the source, opener, and file path; double-click or **Focus Item** returns the owning pane to that item when it is still reachable.
- **Pack** (`Alt+F5`) and **Unpack** (`Alt+F6`): create or extract local ZIP archives from the active pane. Pack preserves selected empty folders and sorted entries; Unpack validates ZIP paths before writing and can preserve existing files when overwrite is disabled.

![File properties dialog](res/file-properties.png)

Properties can be opened with **Files -> Properties** or `Alt+Enter`. The dialog is theme-aware and can show plugin-provided metadata, including general item identity, path, type, size, timestamps, attributes, shortcut/link/reparse targets, and named streams when the provider exposes them. `Esc` closes the dialog and `Ctrl+C` copies all property text.

### Creating and Renaming

- **Rename** (`F2`): rename the focused item. If more than one item is selected, RedSalamander opens Batch Rename instead. The rename prompt also offers **Batch...**: for a folder it opens Batch Rename rooted at that folder, for a file it opens Batch Rename with that file in the preview.
- **Batch Rename**: preview selected files before bulk renaming them. The Rules area combines a new-name template, search/replace options, and independent file-name/extension case changes, or you can switch to Manual mode and type one target name per row. The preview grid validates every proposed name before any rename can run. The window is modeless, so you can keep working in the panes while it is open. For the full set of macros, search/replace options, case styles, validation rules, and reporting, see [Batch Rename](BatchRename.md).
- **New Folder / MakeDir** (`F7`): create a directory in the focused pane.
- **New -> Shell templates**: lists Windows ShellNew templates for local folders. Choosing one prompts for a name, creates the file, refreshes the pane, and focuses the new item when it is visible.
- **Change Case** (`Ctrl+F7`): change name casing for selected items (lower, upper, mixed) without opening the full preview window. See [Batch Rename](BatchRename.md#change-case) for the available case styles.
- **Change Attributes**: uses mixed tri-state boxes for multi-selection so unchanged attributes stay untouched.

### Copy, Move, and Delete

![File operations popup](res/file-operations-popup.png)

- **Copy** (`F5`): copy selected items to the other pane.
- **Move/Rename** (`F6`): move selected items to the other pane or rename depending on context.
- **Delete** (`Del` or `F8`): delete selected items, using the Recycle Bin when supported.
- **Permanent Delete** (`Shift+Del` or `Shift+F8`): asks for confirmation, then bypasses the Recycle Bin.

Long operations run in the background. The File Operations popup shows:

- Current operation, source, destination, progress, and remaining time.
- Pause, resume, cancel, and dismiss controls.
- Wait vs Parallel execution mode.
- Per-task speed limit for copy/move tasks.
- Conflict prompts for overwrite, retry, skip, and cancel decisions.
- Issue export and diagnostic log access.

### Clipboard and Drag-and-Drop

- `Ctrl+C`: copy selected Windows-path items to the clipboard.
- `Ctrl+X`: mark selected local Windows-path items for move using the standard Explorer clipboard format.
- `Ctrl+V`: paste into the current folder.
- **Paste Shortcut** creates `.lnk` files in the current local folder for file paths already on the clipboard.
- Drag items between panes to queue file operations.
- Drag items to external applications where standard Windows drop formats are available.

### Selection Tools

The **Edit** menu provides selection commands for batch work:

- Select / unselect by mask.
- Select all and unselect all.
- Invert selection.
- Restore the previous selection. The **Edit -> Advanced -> Load Selection...** command restores the same saved selection (both **Restore Selection** and **Load Selection...** invoke the same restore action), and **Edit -> Advanced -> Save Selection** captures the current selection so it can be restored later.
- Select next matching item.
- Select same extension.
- Calculate directory size and continue to next item.
- Hide selected or unselected names.

### Path, Shell, and Clipboard Helpers

The file and command menus include several everyday workflow helpers:

- **Copy Path and File Name**, **Copy Path**, **Copy Name**, and **Copy UNC Path and Name** copy text forms of the current selection.
- **Cut** and **Paste Shortcut** use standard Windows clipboard formats for local file-system selections.
- **Open File Explorer -> Current Folder** opens the active folder in Windows File Explorer when it has a local backing path.
- **Command Shell** opens a shell in the current location when possible.
- **Bring Current Directory to Command Line** (`Ctrl+Space`) opens the command-line input and appends the active local folder path with command-line quoting.
- **Bring Filename to Command Line** (`Ctrl+Enter`) opens the command-line input and appends the focused item name, or selected full local paths when files are selected. Press `Enter` in the input to run the command in that pane's current folder, or `Esc` to close it.
- **User Menu** (`F9`) opens configured external commands for the focused pane. Entries are managed in Preferences, can be filtered by file extension and computer name, and use the same path/selection macros as viewer and editor actions.
- **Reread Associations** reloads the settings-backed viewer, editor, User Menu, plugin, and extension association data without moving the current pane folders. It rebuilds the dynamic action menus, clears icon association caches, refreshes both panes, and keeps the previous valid settings when the settings file cannot be reloaded.
- **Make File List** opens an options dialog for the current selection, focused item, or current folder. It can recurse into folders, write JSON/CSV/text, choose which fields are included, use text macros such as `{fullPath}` and `{size}`, and save the result to the clipboard or a UTF-8 file. The last selected options are remembered.
- **Shared Directories** (`Ctrl+Shift+F9`) opens a modeless list of local Windows disk shares. The dialog shows share name, local path, type, and remark, can open reachable local share paths in the focused pane, and offers the Windows Shared Folders management console.
- **Connect Network Drive** (`F11`) and **Disconnect** (`F12`) use Windows network-drive integration from local `file:` panes.
- **File Operations Failed Items** (`Ctrl+J`) opens the issues pane for failed copy/move/delete work.

### Folder Size and Space Usage

Use **Calculate Occupied Space** (`Alt+F10`) to open ViewerSpace, a treemap-style folder usage viewer.

Pressing `Space` on a folder selects it, calculates its size, and moves to the next item.

## Find Files and Directories

Use **Commands -> Find Files and Directories...** or `Alt+F7` to open the modeless search window for the focused pane.

Search features include:

- Search by name and path.
- Search by file content where supported by the backend.
- Scope from the current pane, with multiple independent Find windows supported.
- Native, indexed, and fallback search backend handling.
- Backend status and degraded-mode diagnostics.
- Reusable recent values stored in settings.
- Result-list actions for shell copy/cut, configured viewer/editor, copy/move to the other pane, and delete/permanent delete.
- A `?` help button that summarizes result-list actions and shows the current copy/move destination folder.

Because the search window is modeless, you can keep browsing while a search is open.

More detail: [Find Files and Directories](FindFiles.md)

## Compare Directories

Compare Directories opens a dedicated two-pane comparison window for the current left and right roots.

![Compare Directories window](res/compare-directories.png)

Open it with:

- **Commands -> Compare Directories...**
- `Ctrl+F10`

Key features:

- Match items by relative folder and name.
- Compare by size, timestamp, attributes, and content.
- Include or exclude subdirectories.
- Apply file and directory ignore patterns.
- Show or hide identical items.
- Rescan after path or option changes.
- Restore, invert, and manage difference selection.
- Use copy/move operations as synchronization actions when compare mode is active.

If a compare scan runs long enough, it can also appear as an informational task in the File Operations popup.

More detail: [Compare Directories](CompareDirectories.md)

## Connections

The Connections Manager (opened from **Commands -> Connections Manager...**) stores reusable profiles for remote, cloud, and object-storage file systems.

![Connections Manager with sanitized example profiles](res/connections-manager.png)

Open it from:

- **Commands -> Connections Manager...**
- `nav:`, `nav://`, or `@conn:` in the address bar.
- A protocol prefix with no target, such as `sftp:`, `gdrive:`, `s3:`, or `s3table:`.

The Connections Manager supports:

- FTP, SFTP, SCP, and IMAP profiles.
- Google Drive, OneDrive Personal, OneDrive Business, and SharePoint profiles.
- S3 and S3 Table profiles.
- Quick Connect for one-time profiles.
- Secure secret storage through Windows Credential Manager.
- Optional Windows Hello gating before saved secrets are released.
- Per-connection operation overrides where supported.

More detail: [Connections](Connections.md)

## Preferences

Open Preferences with **View -> Preferences...**.

![Preferences dialog](res/preferences.png)

Preferences pages cover:

- **General**: startup and common application behavior.
- **Panes**: folder pane display and interaction settings.
- **Viewers**: viewer Actions and Associations for View, Alternate View, and View With.
- **Editors**: editor Actions and Associations for Edit, Alternate Edit, Edit New, and Edit With.
- **User Menu**: external command actions shown by `F9` and Commands -> User Menu.
- **Keyboard**: shortcut bindings.
- **Themes**: built-in, file-based, and user themes.
- **Plugins**: plugin enablement, diagnostics, and plugin configuration.
- **File Operations**: pre-calculation, speed, bridge buffer, and task defaults.
- **Compare Directories**: default comparison options.
- **Hot Paths**: ten bookmark slots.
- **Advanced**: lower-level diagnostics and storage settings.

## Themes

Themes can be switched quickly from **View -> Theme** or managed in **Preferences -> Themes**.

Built-in choices:

- System
- Light
- Dark
- Rainbow
- High Contrast App

![Light theme](res/theme-light.png)

![Dark theme](res/theme-dark.png)

![Rainbow theme](res/theme-rainbow.png)

![High Contrast App theme](res/theme-high-contrast-app.png)

Theme features:

- Panes, navigation bars, status bars, file operations UI, dialogs, and viewers follow the active theme.
- Version 2 themes can be loaded from `Themes\*.theme.json5`; shipped choices include Dracula and Catppuccin Latte, Frappé, and Mocha.
- User themes can define reusable palette entries, semantic references, derived colors, system colors, and stable runtime Rainbow/choice sources.
- Rainbow follows the Windows light/dark base and gives multiple selected items stable, repeatable color variation; Windows Contrast Themes still take precedence.
- User themes can be created, duplicated, edited, previewed, and saved without flattening their authored sources.
- Temporary theme previews are restored if Preferences is canceled.

More detail: [Themes](Themes.md)

## Plugin System

Plugins provide virtual file systems and file viewers. Open **Plugins -> Plugin Manager...** or **View -> Preferences... -> Plugins** to inspect, enable, disable, test, and configure plugins.

![Plugin configuration](res/preferences-plugins.png)

### Built-In File-System Plugins

| Prefix | Plugin | Main use | User-visible capabilities |
|--------|--------|----------|---------------------------|
| `file:` | Windows File System | Local drives, known folders, UNC paths, WSL shortcuts | Browse, open, copy, move, rename, delete, recycle-bin delete, properties, clipboard paths, drag-and-drop where Windows supports it. |
| `7z:` | 7-Zip Archive File System | Archive browsing | Open supported archives as virtual folders, browse archive contents, and read/copy out items where the archive backend supports it. |
| `ftp:` | FTP | Remote FTP servers | Browse, read, upload, overwrite, delete, and use Connections Manager profiles; server capability may limit operations. |
| `sftp:` | SFTP | SSH file transfer | Browse, read, write, delete, and use SSH credentials or keys through settings or the Connections Manager. |
| `scp:` | SCP | SSH copy access | Browse and transfer through the shared curl-based plugin; directory operations depend on server SFTP support. |
| `imap:` | IMAP | Mailbox browsing | Browse mailboxes and messages exposed as `.eml` files; deleting an `.eml` deletes the message when allowed. |
| `gdrive:` | Google Drive | Google Drive metadata browsing | Browse folders and inspect drive metadata; the current user-facing milestone is read-only metadata and requires an existing OAuth refresh token. |
| `onedrivep:` | OneDrive Personal | Microsoft Graph personal drive | Browser sign-in, browse, read, write, rename, move, delete, and calculate directory size. |
| `onedriveb:` | OneDrive Business | Microsoft Graph work drive | Browser sign-in, browse, read, write, rename, move, delete, and calculate directory size. |
| `sharepoint:` | SharePoint | Microsoft Graph SharePoint sites/libraries | Browser sign-in, browse document libraries, read, write, rename, move, delete, and calculate directory size. |
| `s3:` | S3 object storage | Buckets, prefixes, and objects | Browse buckets/prefixes, open/download objects, upload, overwrite, delete objects, and bridge-copy with other file systems. |
| `s3table:` | S3 Table | Table buckets, namespaces, and tables | Browse table buckets/namespaces and open generated `*.table.json` table documents; read-only. |
| `fk:` | Dummy/Test File System | Development and test scenarios | Exposes a deterministic fake file system for testing. |

### File-System Plugin Notes

- The Connections Manager is preferred for credentials because it stores secrets through Windows facilities.
- Plugin defaults in Preferences can be useful for local testing but may store non-secret and some plugin-specific values in the settings file.
- Cross-file-system copy/move uses the host bridge: the source plugin reads, the destination plugin writes.
- Read-only plugins reject writes and destructive operations.
- Some remote operations depend on server permissions and protocol support.
- File-system plugins use the common pane shortcuts: `Enter` to open, `F3` to view, `F5` to copy, `F6` to move, `F7` to create a directory, `F8`/`Del` to delete, `Alt+Enter` for Properties, and `Backspace` to go up.

For full per-plugin behavior, limitations, configuration, and shortcut notes, see [Plugins](Plugins.md).

### Built-In Viewer Plugins

Viewer actions and associations are configured in **Preferences -> Viewers**. Press `F3` to open the focused file with the associated primary viewer action, use `Alt+F3` for the associated alternate viewer action, or choose **View With** to pick from applicable configured viewer actions. If no specific rule matches, the default `*` association normally opens the Text viewer.

| Viewer | Plugin id | Typical files | User-visible capabilities |
|--------|-----------|---------------|---------------------------|
| Text / Hex Viewer | `builtin/viewer-text` | Text, logs, config, CSV, XML, unknown files | Text and hex modes, find, next/previous match, encoding selection, line numbers, wrap, byte color options, diff viewing modes, and other-files navigation. |
| Image / RAW Viewer | `builtin/viewer-imgraw` | Common image formats and camera RAW files | Fit/actual size, zoom, pan, rotate, flip, reset orientation, export, brightness/contrast/gamma/grayscale/negative adjustments, thumbnail vs RAW source, Exif overlay, and other-files navigation. |
| Space Viewer | `builtin/viewer-space` | Folders | Treemap-style occupied-space visualization, scan current or selected folder, navigate up, refresh, and exit. |
| PE Viewer | `builtin/viewer-pe` | `.exe`, `.dll`, `.sys`, and other PE files | Parsed Portable Executable metadata, top/bottom navigation, refresh, other-files navigation, and export to text or Markdown. |
| SQLite Viewer | `builtin/viewer-sqlite` | SQLite databases | Table selector, paged table preview, read-only custom SQL queries, row caps, sortable preview grids, and status strip. |
| VLC Viewer | `builtin/viewer-vlc` | Audio and video files | Media playback through VLC/libVLC, configurable VLC path, plugin path, caching, hardware acceleration, output, visualizer, and extra arguments. |
| Web Viewer | `builtin/viewer-web` | HTML, PDF, and browser-rendered content | WebView2-based rendering for web/PDF content when the WebView2 Runtime is installed. |
| JSON Viewer | `builtin/viewer-json` | `.json`, `.json5` | WebView2-based structured rendering with Text viewer fallback if unavailable. |
| Markdown Viewer | `builtin/viewer-markdown` | `.md` | WebView2-based Markdown rendering with Text viewer fallback if unavailable. |

More detail: [Plugins](Plugins.md) and [Viewers](Viewers.md)

## Viewer Workflows

The viewer screenshots in this section are real plugin windows opened on generated sample files. They are intended to show the plugin in action, not the settings page.

### Text and Diff Work

The Text viewer is the fallback for unmapped file types and is also useful for direct text, hex, and diff-style inspection.

![Text viewer](res/viewer-text.png)

Common actions:

- Find text with `Ctrl+F`.
- Move to next/previous match with `F3` / `Shift+F3`.
- Switch Text and Hex modes.
- Toggle wrapping and line numbers.
- Choose text encoding.
- Navigate to peer files in the same folder.

### Images and RAW Files

![Image/RAW viewer](res/viewer-imgraw.png)

Use the Image/RAW viewer for inspection, zooming, orientation corrections, and lightweight visual adjustments.

### Folder Space Usage

![Space viewer](res/viewer-space.png)

Use **Commands -> Calculate Occupied Space** or the folder context menu to open a treemap of folder size distribution.

### PE, Web, Media, and SQLite Files

![PE viewer](res/viewer-pe.png)

The PE viewer parses Windows executables and libraries, then lets you export the report as text or Markdown.

![SQLite viewer](res/viewer-sqlite.png)

The SQLite viewer opens databases read-only, previews tables by page, and can run read-only custom SQL queries.

![Web viewer](res/viewer-web.png)

The Web viewer renders HTML and PDF-style content through Microsoft Edge WebView2.

![JSON viewer](res/viewer-json.png)

The JSON viewer formats structured JSON, supports tree-style inspection, and can hand oversized files to the Text viewer.

![Markdown viewer](res/viewer-markdown.png)

The Markdown viewer renders headings, lists, emphasis, and code blocks, with a source toggle for inspection.

![VLC viewer](res/viewer-vlc.png)

VLC playback requires VLC to be installed or configured in plugin settings. Web, JSON, Markdown, and PDF rendering require Microsoft Edge WebView2 Runtime.

## RedSalamanderMonitor

RedSalamanderMonitor is a companion ETW/log viewer for RedSalamander diagnostics.

![RedSalamanderMonitor](res/monitor.png)

It supports:

- Live ETW streaming.
- Opening captured logs.
- Append-only high-throughput display.
- Auto-scroll tail mode.
- Filtering by message type.
- Find with next/previous navigation.
- Copying selected log lines.
- Always-on-top behavior.
- The same theme system as the main app.

More detail: [RedSalamanderMonitor](Monitor.md)

## Settings and Advanced Configuration

RedSalamander stores durable user choices in its settings file, including:

- File-action associations for viewer/editor behavior.
- Theme choices and user theme data.
- Hot Paths.
- File operation defaults.
- Compare Directory defaults.
- Search recent values.
- Connections Manager non-secret profile fields.
- Plugin configuration.

Secrets are not meant to be stored in the settings JSON by the Connections Manager. Saved connection secrets are stored with Windows Credential Manager, optionally protected by Windows Hello.

More detail: [Settings File & Advanced Configuration](SettingsFile.md)

## Troubleshooting

Use these first checks:

- If a remote connection fails, open the Connections Manager and verify host, protocol, authentication, and saved secret state.
- If a cloud sign-in fails, verify the plugin client ID and whether first-time OAuth is supported for that plugin.
- If media playback fails, configure the VLC installation path in the VLC viewer plugin settings.
- If copy/move behaves unexpectedly, open the File Operations popup and check conflicts, issues, and exported diagnostics.
- If a theme preview looks wrong, cancel Preferences to restore the prior theme or switch through **View -> Theme**.

More detail: [Troubleshooting / Reset](Troubleshooting.md)
