# Plugins (File Systems & Viewers)

RedSalamander uses plugins for:

- **File systems**: browsing + operations (local, archives, remote protocols, cloud)
- **Viewers**: opening files in dedicated viewer windows

Plugins are loaded from:

- the `Plugins\` folder next to `RedSalamander.exe`, and
- any additional custom plugin paths you add in Preferences.

![Plugins UI](res/preferences-plugins.png)

## Open the plugin manager UI

- **Plugins → Plugin Manager…** (opens the Plugins page in Preferences)
- **View → Preferences… → Plugins**

From there you can:

- Enable/disable plugins
- Configure a plugin (when it exposes a configuration schema)
- Run **Test** / **Test All** to validate plugin availability

## Built-in file system plugins

File systems are selected by a **short prefix** in the address bar:

| Prefix | Plugin id | Main use | User-visible capabilities |
|--------|-----------|----------|---------------------------|
| `file:` | `builtin/file-system` | Windows file system | Drives, known folders, UNC paths, WSL shortcuts, browse/open/copy/move/rename/delete, Recycle Bin delete where supported, properties, clipboard, and drag-and-drop. |
| `7z:` | `builtin/file-system-7z` | Archives | Open supported archives as virtual folders, usually by pressing Enter on an archive from `file:`. |
| `ftp:` | `builtin/file-system-ftp` | FTP | Browse/read/write/delete through the curl-based remote file-system plugin; server capability and credentials may limit operations. |
| `sftp:` | `builtin/file-system-sftp` | SFTP | SSH file transfer with password or key-based authentication via settings or Connection Manager. |
| `scp:` | `builtin/file-system-scp` | SCP | SSH copy access; directory operations depend on server SFTP support. |
| `imap:` | `builtin/file-system-imap` | IMAP mailboxes | Browse mailbox hierarchy and expose messages as `.eml` files; deleting an `.eml` deletes the message when allowed. |
| `gdrive:` | `builtin/file-system-gdrive` | Google Drive | Browse folders and inspect drive metadata; current user-facing milestone is read-only metadata and requires an existing OAuth refresh token. |
| `onedrivep:` | `builtin/file-system-onedrive-personal` | OneDrive Personal | Browser sign-in, browse/read/write/rename/move/delete, and directory size. |
| `onedriveb:` | `builtin/file-system-onedrive-business` | OneDrive Business | Browser sign-in, browse/read/write/rename/move/delete, and directory size. |
| `sharepoint:` | `builtin/file-system-sharepoint` | SharePoint | Browser sign-in, document library browsing, read/write/rename/move/delete, and directory size. |
| `s3:` | `builtin/file-system-s3` | S3 object storage | Browse buckets/prefixes, open/download objects, upload, overwrite, delete objects, and bridge-copy with compatible file systems. |
| `s3table:` | `builtin/file-system-s3table` | S3 Table | Browse table buckets, namespaces, and generated `*.table.json` documents; read-only. |
| `fk:` | `builtin/file-system-dummy` | Test file system | Deterministic fake file system for development and tests. |

See also:

- [Navigation & Path Syntax](NavigationAndPaths.md)
- [Remote File Systems](RemoteFileSystems.md)
- [Cloud Drives](CloudDrives.md)
- [S3 / S3 Table](S3AndS3Table.md)
- [Connections](Connections.md)

## Common file-system shortcuts

File-system plugins share the same pane commands unless a plugin is read-only or the server rejects the operation.

| Shortcut | Command |
|----------|---------|
| `Enter` | Open folder, execute/open file, or mount an archive when an association exists. |
| `Backspace` | Parent folder. |
| `F3` | View selected file with the associated viewer action. |
| `F5` | Copy selection to the other pane. |
| `F6` | Move selection to the other pane when supported. |
| `F7` | Create directory when supported. |
| `F8` / `Del` | Delete, using Recycle Bin only when the active file system supports it. |
| `Shift+F8` / `Shift+Del` | Permanent delete after confirmation. |
| `Alt+Enter` | Properties through the active file-system plugin. |
| `Ctrl+L` / `Alt+D` | Type a plugin path or saved connection alias. |
| `Alt+F1` / `Alt+F2` | Open the left/right plugin drive menu. |

## File-system plugin details

### Windows File System (`file:`)

Behavior:

- Browses local drives, removable media, UNC shares, known folders, OneDrive shell locations, and WSL entries shown by the drive menu.
- Supports the broadest operation set: open, read, write, copy, move, rename, create directory, delete, Recycle Bin delete, properties, shell context menus, clipboard `CF_HDROP`, and external drag-and-drop.
- Shows local disk capacity and usage in the navigation bar.
- Exposes drive actions such as **Disk Properties** and **Disk Cleanup** when Windows provides them for the current drive.

Shortcuts:

- Uses all common file-system shortcuts above.
- `Shift+A` through `Shift+Z` jumps to a drive root when that drive exists.
- `F11` / `F12` connect or disconnect Windows network drives from local panes.

### 7-Zip Archive File System (`7z:`)

Behavior:

- Opens supported archive/container files as virtual folders.
- Usually starts by pressing `Enter` on a supported archive from a `file:` pane.
- Supports browsing and reading/copying items out of archives.
- Treats archives as mounted contexts, for example `7z:C:\Downloads\archive.zip|/`.
- Current archive browsing is read-oriented; destructive operations inside mounted archives should be treated as unsupported unless the UI explicitly enables them. Local ZIP creation/extraction is available separately through **Files -> Pack** and **Files -> Unpack**.

Shortcuts:

- `Enter` opens folders inside the archive.
- `Backspace` navigates toward the archive root/parent.
- `F3` views an archived file.
- `F5` copies items out through the host file-operation bridge.

### FTP (`ftp:`)

Behavior:

- Connects to FTP servers through the curl-based remote file-system plugin.
- Supports browsing, reading, uploading, overwriting, deleting, and creating directories when the server and account allow it.
- Can use URI-style paths such as `ftp://user@example.com/pub/`, plugin defaults, or Connection Manager profiles.

Shortcuts:

- Uses common file-system shortcuts for browse/copy/move/delete/properties where the server supports the operation.
- Typing `ftp:` with no target opens Connection Manager filtered to FTP.

### SFTP (`sftp:`)

Behavior:

- Uses SSH file transfer for remote folders.
- Supports password or key-based authentication through plugin settings or Connection Manager.
- Supports browse, read, write, delete, rename, create directory, and host-bridged copy/move when the server permits it.

Shortcuts:

- Uses common file-system shortcuts.
- Typing `sftp:` with no target opens Connection Manager filtered to SFTP.

### SCP (`scp:`)

Behavior:

- Provides SSH copy-style access through the same curl plugin family.
- Directory behavior depends on server support for SFTP over SSH.
- Best used through Connection Manager when credentials, private keys, or known-host files are required.

Shortcuts:

- Uses common file-system shortcuts where supported.
- Typing `scp:` with no target opens Connection Manager filtered to SCP.

### IMAP (`imap:`)

Behavior:

- Browses IMAP mailbox hierarchy as folders.
- Exposes messages as `.eml` files using names based on subject, sender, and UID.
- Opening or viewing a message uses the normal viewer flow.
- Deleting an `.eml` deletes the server-side message when the account/server permits it.
- TLS behavior follows the configured port and server capabilities.

Shortcuts:

- `Enter` opens mailbox folders.
- `F3` views `.eml` messages.
- `F8` / `Del` deletes messages when allowed.
- Typing `imap:` with no target opens Connection Manager filtered to IMAP.

### Google Drive (`gdrive:`)

Behavior:

- Browses Google Drive folders and metadata.
- The current user-facing milestone is read-only metadata browsing.
- First-time browser sign-in is not launched by the current plugin UI; a usable profile requires an existing refresh token from host OAuth plumbing or test tooling.
- Google Drive shortcut entries are surfaced, but shortcut dereference/open semantics are not implemented yet.

Shortcuts:

- `Enter` opens folders.
- `F3` views readable generated/metadata files when available.
- Write, rename, move, and delete shortcuts are expected to be disabled or rejected in this read-only milestone.
- Typing `gdrive:` with no target opens Connection Manager filtered to Google Drive.

### OneDrive Personal (`onedrivep:`)

Behavior:

- Uses Microsoft Graph with browser sign-in.
- Supports browse, read, write, rename, move, delete, and directory size calculation.
- Connection Manager stores the non-secret profile data and Windows stores refresh/sign-in secrets.

Shortcuts:

- Uses common file-system shortcuts.
- Typing `onedrivep:` with no target opens Connection Manager filtered to OneDrive Personal.

### OneDrive Business (`onedriveb:`)

Behavior:

- Uses Microsoft Graph with work/school browser sign-in.
- Supports browse, read, write, rename, move, delete, and directory size calculation.
- Supports saved profiles through Connection Manager.

Shortcuts:

- Uses common file-system shortcuts.
- Typing `onedriveb:` with no target opens Connection Manager filtered to OneDrive Business.

### SharePoint (`sharepoint:`)

Behavior:

- Uses Microsoft Graph to browse SharePoint sites and document libraries.
- A profile host can be a tenant host or a tenant plus site path.
- Supports browse, read, write, rename, move, delete, and directory size calculation when permissions allow it.

Shortcuts:

- Uses common file-system shortcuts.
- Typing `sharepoint:` with no target opens Connection Manager filtered to SharePoint.

### S3 (`s3:`)

Behavior:

- Browses buckets, prefixes, and objects.
- Supports opening/downloading objects, uploading new objects, overwriting existing objects, deleting objects, and host-bridged copy/move with other file systems.
- Folders are virtual prefixes, not normal directory objects.
- Current limitations include no server-side rename/move and no recursive delete of virtual folder prefixes.

Shortcuts:

- `Enter` opens prefixes or objects.
- `F3` views an object through the associated viewer action.
- `F5` uploads/copies to the other pane when the destination supports it.
- `F8` / `Del` deletes selected objects.
- Typing `s3:` with no target opens Connection Manager filtered to S3.

### S3 Table (`s3table:`)

Behavior:

- Browses S3 Table buckets, namespaces, and tables.
- Tables are exposed as generated `*.table.json` documents.
- Current behavior is browse/read-only: no upload, write, delete, rename, or move.

Shortcuts:

- `Enter` opens buckets, namespaces, and generated table documents.
- `F3` views generated table JSON.
- Write/destructive shortcuts are expected to be disabled or rejected.
- Typing `s3table:` with no target opens Connection Manager filtered to S3 Table.

### Dummy/Test File System (`fk:`)

Behavior:

- Exposes a deterministic fake file system used by development and self-tests.
- Useful for exercising plugin operations without relying on a live remote service.

Shortcuts:

- Uses common file-system shortcuts according to the scenario exposed by the dummy provider.

### Archive auto-mount (advanced)

When browsing `file:`, RedSalamander can automatically open certain extensions as virtual file systems instead of executing them.

This is controlled by the settings map:

- `extensions.openWithFileSystemByExtension` *(currently not exposed in Preferences UI)*

Defaults include many archive/container formats such as `.zip`, `.7z`, `.rar`, `.iso`, `.vhd`, `.vhdx`, `.qcow2`, …

To disable auto-mount behavior, set `extensions.openWithFileSystemByExtension` to `{}` in the settings file. See: [Settings File & Advanced Configuration](SettingsFile.md)

## Viewer plugins

Viewers are chosen by configured viewer actions and associations:

- Preferences -> **Viewers** has **Associations** for `F3` and `Alt+F3`, plus **Actions** for internal viewer plugins and external viewer programs
- `F3` opens the viewer action associated with the focused file's extension/pattern/computer, falling back to the default `*` association
- `Alt+F3` opens the associated alternate viewer action when configured
- **View With** lists applicable configured viewer actions for the focused item

Built-in/available viewers include:

- `builtin/viewer-text` — Text/Hex viewer
- `builtin/viewer-imgraw` — Images + camera RAW
- `builtin/viewer-space` — “occupied space” treemap for folders
- `builtin/viewer-pe` — Portable Executable (EXE/DLL/SYS) viewer
- `builtin/viewer-sqlite` — SQLite table/query viewer
- `builtin/viewer-vlc` — media playback using VLC (requires VLC installation)
- `builtin/viewer-web` / `builtin/viewer-json` / `builtin/viewer-markdown` — WebView2-based viewers (requires WebView2 runtime)

The screenshots in the viewer sections are real plugin windows opened on generated sample files. They are not Preferences screenshots and do not use personal data.

## Viewer plugin details

### Text / Hex / Diff Viewer (`builtin/viewer-text`)

![Text viewer opened on a generated log file](res/viewer-text.png)

Behavior:

- Fallback viewer for files without a more specific association.
- Text mode with encoding detection and selectable code pages.
- Hex mode with byte/text correspondence.
- Parsed diff modes for `.diff`, `.patch`, and `.rej`: side-by-side, inline, unchanged text toggle, and hunk navigation.
- Save As with optional encoding conversion.

Shortcuts:

- `Ctrl+O`: open file.
- `Ctrl+S`: save as.
- `F5`: refresh.
- `Ctrl+F`, `F3`, `Shift+F3`: find, next, previous.
- `Home` / `End`: top/bottom.
- `Ctrl+G`: go to offset.
- `F7` / `Shift+F7`: next/previous diff hunk.
- `F8` / `Shift+F8`: next/previous encoding.
- `Space` / `Backspace` / `Ctrl+Down` / `Ctrl+Up`: other-files navigation.
- `Esc`: close.

### Image / RAW Viewer (`builtin/viewer-imgraw`)

![Image/RAW viewer opened on a generated PNG](res/viewer-imgraw.png)

Behavior:

- Displays common image formats and many camera RAW formats.
- Supports fit/actual-size modes, smooth zoom, pan, rotate, flip, reset orientation, export, brightness/contrast/gamma adjustments, grayscale/negative toggles, RAW vs thumbnail source, Exif overlay, and neighbor prefetch.

Shortcuts:

- `F5`: refresh.
- `Ctrl+S`: export.
- `Right` / `PgDn` / `Space`: next file.
- `Left` / `PgUp` / `Backspace`: previous file.
- `Home` / `End`: first/last file.
- `Ctrl+F` or double click: fit to window.
- `F`: toggle fit/100%.
- `1`: actual size.
- `+` / `-` / `0`: zoom in/out/reset.
- `R`, `Ctrl+R`/`Shift+R`: rotate clockwise/counterclockwise.
- `H` / `V`: flip horizontal/vertical.
- `O`: reset orientation.
- `G` / `N`: grayscale/negative.
- `I`: Exif overlay.
- `Ctrl+Alt+Arrow` and `Ctrl+Alt+PgUp/PgDn`: adjustment shortcuts.
- Mouse wheel zooms; dragging pans when zoomed.

### Space Viewer (`builtin/viewer-space`)

![Space viewer treemap for a generated folder](res/viewer-space.png)

Behavior:

- Visualizes folder disk usage as a treemap.
- Opens from **Commands -> Calculate Occupied Space** (`Alt+F10`) or folder context menu **View Space**.
- If exactly one folder is selected, it scans that folder; otherwise it scans the current folder.

Shortcuts:

- `F5`: refresh/rescan.
- `Backspace`: up.
- `Esc`: close.

### PE Viewer (`builtin/viewer-pe`)

![PE viewer opened on a sample executable](res/viewer-pe.png)

Behavior:

- Parses Portable Executable files such as `.exe`, `.dll`, and `.sys`.
- Shows structured PE metadata and supports export.

Shortcuts:

- `Ctrl+S`: export as text.
- `Ctrl+Shift+S`: export as Markdown.
- `F5`: refresh.
- `Space` / `Ctrl+Down`: next peer file.
- `Backspace` / `Ctrl+Up`: previous peer file.
- `Ctrl+Home` / `Ctrl+End`: first/last peer file.
- `Home` / `End`: top/bottom.
- `Esc`: close.

### SQLite Viewer (`builtin/viewer-sqlite`)

![SQLite viewer opened on a generated sample database](res/viewer-sqlite.png)

Behavior:

- Opens SQLite databases read-only.
- Provides a table selector, paged table preview, previous/next page buttons, read-only custom SQL query field, Run Query command, Table Preview command, results grid, and status strip.
- Plugin settings control preview page size, custom query row cap, and direct-open behavior for local files.

Shortcuts:

- `F3` opens the viewer from the host.
- Inside the viewer, use normal dialog navigation (`Tab`, arrow keys, Enter/Space on focused buttons) for Reload, Previous, Next, Run Query, and Table Preview.

### VLC Viewer (`builtin/viewer-vlc`)

![VLC viewer playing a generated audio sample](res/viewer-vlc.png)

Behavior:

- Plays audio/video files through VLC/libVLC.
- Provides play, pause, stop, timeline, and snapshot controls.
- Requires VLC media player or a configured VLC installation folder.

Shortcuts:

- `F3` opens the viewer from the host.
- Current user-facing controls are the visible playback buttons; no separate viewer menu accelerator table is defined in the current resources.

### Web / JSON / Markdown Viewers (`builtin/viewer-web`, `builtin/viewer-json`, `builtin/viewer-markdown`)

![Web viewer rendering a generated HTML file](res/viewer-web.png)

![JSON viewer rendering a generated JSON file](res/viewer-json.png)

![Markdown viewer rendering a generated Markdown file](res/viewer-markdown.png)

Behavior:

- Use Microsoft Edge WebView2.
- Web viewer opens HTML and PDF content.
- JSON viewer supports pretty, tree, and JSONL/card-style modes with syntax highlighting.
- Markdown viewer renders Markdown and can toggle source view.
- Oversized JSON/Markdown documents can be handed off to Text Viewer.

Shortcuts:

- `Ctrl+S`: save as.
- `F5`: refresh.
- `Ctrl+F`, `F3`, `Shift+F3`: find, next, previous.
- `Ctrl++`, `Ctrl+-`, `Ctrl+0`: zoom in/out/reset.
- `F12`: DevTools when enabled.
- `Ctrl+L`: copy URL.
- `Ctrl+Enter`: open in browser.
- `Ctrl+backtick`: toggle Markdown source.
- `Space` / `Backspace` / `Ctrl+Down` / `Ctrl+Up`: other-files navigation.
- `Esc`: close.

See: [Viewers](Viewers.md)
