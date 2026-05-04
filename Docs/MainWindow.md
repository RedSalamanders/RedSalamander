# Main Window & Panes

RedSalamander is a **dual-pane** file manager:

- **Left pane** and **Right pane** each have an optional Navigation bar, Folder view, optional Filter bar, and optional Status bar.
- A splitter between panes controls the width ratio.
- Most commands apply to the **focused pane** (the one that currently has keyboard focus).

![Main window overview](res/main-window.png)

## Panes, focus, and “other pane”

Some actions explicitly use both panes:

- **Copy to other pane** (`F5`): copies the focused pane selection into the other pane’s current folder.
- **Move to other pane** (`F6`): moves the focused pane selection into the other pane’s current folder.

Pane management:

- **Swap panes** (`Ctrl+U`): swaps the left/right locations.
- **Maximize/Restore pane** (`Ctrl+F11`): temporarily zooms the focused pane.

## Navigation bar (per pane)

The Navigation bar is the strip above the Folder view:

- **Drive/Menu**: click to open the plugin-provided menu (drives, known folders, WSL, …).
- **Path/Breadcrumb**: shows the current location.
- **History**: opens the folder history dropdown.
- **Disk info** (when available): shows total/used/free and a usage bar; may provide disk actions.

![Navigation bar](res/navigation-view.png)

Disk actions are exposed by the active file-system plugin. For local `file:` drives, the disk area can provide Windows drive actions such as **Disk Properties** and **Disk Cleanup**.

## Folder view (per pane)

Folder view shows items in a grid with icons.

Common interactions:

- `Enter`: open a folder / execute a file (see [Navigation & Path Syntax](NavigationAndPaths.md))
- `F3`: open the focused file in a viewer (see [Viewers](Viewers.md))
- `F2`: rename
- `Alt+Enter`: properties
- `Shift+F10` or the **Apps/Menu** key: open the item context menu
- `Ctrl+F12`: filter the current folder (pane filter dialog)
- Mouse: multi-select, drag & drop

Display and sort (defaults):

- `Alt+2`: Brief mode
- `Alt+3`: Detailed mode
- `Ctrl+F2`: Sort = None
- `Ctrl+F3..F6`: Sort by Name/Ext/Time/Size

![Folder view](res/folder-view.png)

### File properties

Open item properties with `Alt+Enter` or **Files -> Properties**.

![File properties](res/file-properties.png)

The properties dialog can show local or plugin-provided fields: general identity, path, type, size, timestamps, attributes, and named streams when exposed by the active file system. `Ctrl+C` copies all visible property text; `Esc` closes the dialog.

### Selection helpers

- `Ctrl+A`: select all items in the focused pane.
- `Esc`: clear the current selection.
- `Ctrl` + the key left of `Backspace`: open the select-by-mask dialog.
- `Ctrl` + the key right of `0`: open the unselect-by-mask dialog.
- `Ctrl+Shift` + those same keys: select/unselect same extension.
- `Ctrl+Shift+F5`: save the current selection.
- `Ctrl+Shift+F6`: restore the saved selection.
- `Alt+Up` / `Alt+Down`: jump to the previous/next selected item.
- The **Edit** menu also exposes **Invert Selection**, **Hide Selected Names**, **Hide Unselected Names**, and **Show Hidden Names**.

### View options

- **Left/Right → Show → Hidden Files**: toggle display of hidden items (hidden items use a dim icon when shown).
- **Left/Right → Show → System Files**: toggle display of system items.
- **Left/Right → Show → File Extensions**: toggle displayed file extensions in the named pane without changing real file names or paths.
- **Left/Right → Thumbnails**: toggle thumbnail-sized visuals in the named pane. Thumbnails load in the background and fall back to normal file/folder icons when unavailable.
- **Left/Right → Preview Pane**: preview the named pane's focused item in the opposite pane. The opposite pane shows Folder and Preview tabs; selecting Folder keeps that pane browsable, while selecting Preview shows the current preview. Closing preview removes the tabs and restores the normal pane.
- **Left/Right → Filter Bar**: show or hide the named pane's persistent filter summary.
- **Left/Right → Navigation Bar**: show or hide the named pane's navigation bar. Address-bar commands show it automatically before focusing the address edit.

![Preview pane showing an embedded Image/RAW viewer](res/preview-pane-image.png)

## Status bar (per pane)

The optional status bar shows:

- Selection summary (files/folders/bytes)
- A sort indicator you can click to open the Sort menu

Toggle:

- **Left/Right → Status Bar**

## Function bar

The Function bar shows the current `F1..F12` bindings (including modifiers).

Toggle:

- **View → Function Bar**

## Menu bar

Toggle:

- **View → Menu Bar**

## Useful commands

- **Commands → Connect Network Drive…** (`F11`) / **Disconnect…** (`F12`) *(Win32 `file:` pane only)*
- **Commands → Find Files and Directories…** (`Alt+F7`) opens the modeless search window for the focused pane. See: [Find Files and Directories](FindFiles.md)
- **Commands → List of Opened Files** (`Alt+F11`) shows viewer, editor, and Preview pane entries and can focus the source item from the list.
- **Commands → Shared Directories…** (`Ctrl+Shift+F9`) lists local Windows disk shares, opens reachable share paths in the focused pane, and links to Windows Shared Folders management.
- **Commands → Command Shell**: opens a shell in the current location when possible
- **Commands → Bring Current Directory/Filename to Command Line**: opens the pane command-line input and inserts quoted local paths for execution
- **Commands → Open File Explorer → Current Folder** (`Shift+F3`)
