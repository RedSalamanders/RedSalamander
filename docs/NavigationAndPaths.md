# Navigation & Path Syntax

## Quick navigation

- Focus the address bar: `Ctrl+L` or `Alt+D`
- Change directory (command): `Shift+F7`
- Open the drive/menu dropdown:
  - click the left section of the Navigation bar, or
  - `Alt+F1` (Left) / `Alt+F2` (Right)
- Open folder history: `Alt+F12` (or **Commands → Show Folders History**)
- History back / forward: `Alt+Left` / `Alt+Right`
- Hot Paths (bookmarked folders):
  - Go to Hot Path: `Ctrl+1` .. `Ctrl+0`
  - Set Hot Path to current folder: `Ctrl+Shift+1` .. `Ctrl+Shift+0`
  - Manage Hot Paths: `Shift+F9` (opens Preferences → Hot Paths)
- Go up one directory: `Backspace`
- Go to previous/next selected name: `Alt+Up` / `Alt+Down`
- Go to root directory: `Shift+Backspace`
- Jump to a drive root: `Shift` + a drive letter (example: `Shift+C` → `C:\`)
- Set path from other pane: `Ctrl+.`
- Filter current folder: `Ctrl+F12`

![Drive/menu dropdown](res/drive-menu.png)

### Hot Paths details

Hot Paths are 10 bookmark slots for quick folder navigation:

- Slots map to digits: `Ctrl+1` = slot 1, …, `Ctrl+9` = slot 9, `Ctrl+0` = slot 10.
- Each slot can optionally have a **Label** (display name). If the label is empty, the path is shown.
- You can choose whether a slot appears in the **drive/menu dropdown** via **Preferences → Hot Paths → Show in Change Drive menu**.

## Breadcrumb interactions

The path is shown as a breadcrumb: each folder in the current path is a clickable segment, with a chevron separator (`›`) between segments. The whole bar reacts to the mouse so you can move around without typing a path.

### Clicking a segment

- Click any ancestor segment (for example `Users` or `Documents` in `C:\Users\Documents\Projects`) to navigate straight to that folder.
- Clicking the **last** segment (the folder you are already in) does not re-navigate. It just moves keyboard focus back to the file list, so a click on the current crumb never triggers a redundant refresh.
- The clickable area covers the whole segment, and hovering highlights it with a subtle rounded background.

### Separator / sibling dropdown

- The chevron separator between two folders is itself clickable. Clicking it opens a dropdown of the **sibling folders** at that level — the other folders inside the same parent.
- For example, clicking the `›` after `Documents` lists the siblings of `Documents` (such as `Pictures` and `Downloads`). The folder you are currently in is marked with an accent bar; hover and keyboard selection move independently of that marker.
- Pick an entry to navigate into that sibling. The dropdown only commits on click or `Enter`; clicking outside or pressing `Esc` cancels without navigating.

### Overflow ellipsis and the full-path popup

When the path is too long to fit, the breadcrumb collapses from the middle outward and inserts a clickable `...` segment, keeping the first and current folders visible as long as possible.

- Click the `...` segment (or a chevron next to it) to open a **full-path popup** that shows the complete breadcrumb.
- The popup uses the same interaction model as the bar: click a segment to navigate (this closes the popup), or click a separator to open that segment's sibling dropdown.
- If the full path is wider than the screen, the popup wraps onto multiple lines; if it is also too tall, it clamps to the screen and scrolls with the mouse wheel.
- The popup can switch to edit mode to type or paste a path — double-click empty space inside it, or press `F4`, `Ctrl+L`, or `Alt+D`. Pressing `Enter` navigates if the path is valid and changed; an unchanged path just leaves edit mode.
- Click outside the popup, press `Esc`, or move focus elsewhere to close it.

To edit the path directly in the bar instead, double-click empty space in the path area (or use `Ctrl+L` / `Alt+D`) to enter address-bar edit mode with the current path selected.

## Paths you can type

The address bar accepts multiple forms. RedSalamander chooses the correct file system plugin based on prefixes.

### 1) Windows paths (local/UNC)

Examples:

- `C:\Windows\`
- `\\server\share\folder\`

### 2) `file:` URIs

Examples:

- `file:///C:/Windows/`
- `file://server/share/`

### 3) Plugin-prefixed paths

General form:

- `<shortId>:<pluginPath>`

Examples:

- `ftp:/`
- `s3:/`
- `s3table:/`

Mounted file systems (archives, etc) use a mount context:

- `<shortId>:<instanceContext>|<pluginPath>`

Example:

- `7z:C:\Downloads\archive.zip|/`

### 4) Connection Manager routing (`nav:` / `nav://` / `@conn:`)

To navigate using a saved Connection Manager profile:

- `nav:<connectionName>`
- `nav://<connectionName>`
- `@conn:<connectionName>`

Examples:

- `nav:MySftpServer`
- `nav://MySftpServer`
- `nav:MyAwsS3`
- `@conn:WorkImap/INBOX/`

If the name is empty (`nav:`, `nav://`, or `@conn:`), RedSalamander opens the Connection Manager dialog.

Advanced protocol-local forms also work:

- `<shortId>:/@conn:<connectionName>/...`
- `<scheme>://@conn/<connectionName>/...`

See: [Connections](Connections.md)

## Special shorthand

### 7z mount shorthand

When the `7z` file system is available, typing:

- `7z:<archivePath>`

is treated as “mount this archive and open its root”.

## History

RedSalamander maintains a shared “most recently used” folder history list and uses it for the History dropdown.

![History dropdown](res/history-dropdown.png)

### Filter state in history

- When a pane filter is active, the filter state is persisted in folder history.
- Navigating to a history entry restores the saved filter state for that path.
- Entries that would restore an active filter display a small filter icon in the history list.

## Disk info actions (Win32 file system)

When browsing a local drive (the `file` file system), the disk info section can expose:

- **Disk Properties**
- **Disk Cleanup**

The disk info text and usage bar are also visible in the navigation bar:

![Navigation bar disk info](res/navigation-view.png)

## Preview Pane

**Left/Right -> Preview Pane** opens the named pane's preview in the opposite pane. The host pane shows Folder and Preview tabs, follows the source pane's focused item, and extends to the function bar or to the bottom of the window when the function bar is hidden.

See: [Main window](MainWindow.md)
