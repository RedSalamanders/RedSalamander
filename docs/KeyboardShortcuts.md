# Keyboard Shortcuts

This page is the consolidated reference for RedSalamander's default keyboard shortcuts, organized by category. It complements the task-oriented [User Guide](UserGuide.md).

RedSalamander is built for keyboard-driven work. Most commands act on the **focused pane** as the source and the **other pane** as the destination. Chords listed here are the factory defaults defined in `RedSalamander/ShortcutDefaults.cpp`.

## Source of truth and customization

- The live, complete binding list is shown by **Help -> Display Shortcuts...** (or press `F1`). That list always reflects your current configuration, including any changes you made.
- The factory defaults are defined in code in `RedSalamander/ShortcutDefaults.cpp` (`ShortcutDefaults::CreateDefaultShortcuts`). If a chord on this page ever disagrees with what the application does, the code and the in-app list win.
- Every binding is user-configurable in **View -> Preferences... -> Keyboard**. You can rebind, add, or clear chords for both the function bar (F1-F12) and the folder-view chords.
- Unmodified letter and digit keys are intentionally left unbound so they can be used for incremental / quick-search typing inside a pane.

The tables below use `Ctrl`, `Alt`, and `Shift` for the modifier keys. A blank cell means no default binding for that key/modifier combination.

## Function bar (F1-F12)

The function bar along the bottom of the window shows the current `F1`-`F12` actions and updates as you hold modifier keys. It can be hidden from **View -> Function Bar**.

| Key | (none) | Ctrl | Alt | Shift | Ctrl+Shift | Alt+Shift |
|-----|--------|------|-----|-------|------------|-----------|
| F1  | Display Shortcuts | | Open Left Drive Menu | | | |
| F2  | Rename | Sort: None | Open Right Drive Menu | | | |
| F3  | View | Sort: Name | Alternate View | Open Current Folder (Explorer) | View Width | |
| F4  | Edit | Sort: Extension | Exit | Edit New | Alternate Edit | |
| F5  | Copy to Other Pane | Sort: Time | Pack | | Save Selection | |
| F6  | Move to Other Pane | Sort: Size | Unpack | | Restore Selection | |
| F7  | Create Directory | Change Case | Find | Change Directory | | |
| F8  | Delete | Change Attributes | | Permanent Delete | | |
| F9  | User Menu | Refresh | Unpack | Hot Paths | Shared Directories | |
| F10 | Menu | Compare Directories | View Space | Context Menu | | Context Menu (Current Directory) |
| F11 | Connect Network Drive | Zoom Panel | List of Opened Files | Previous Theme | Full Screen | |
| F12 | Disconnect Network Drive | Filter | Show Folders History | Next Theme | | |

## Navigation and paths

Folder-view chords apply while focus is inside a folder pane. See [Navigation and Paths](NavigationAndPaths.md) for the address bar and breadcrumb details.

| Chord | Action |
|-------|--------|
| `Backspace` | Up one directory |
| `Shift+Backspace` | Go to root directory |
| `Tab` / `Shift+Tab` | Switch pane focus |
| `Alt+Left` | History back |
| `Alt+Right` | History forward |
| `Alt+Up` | Go to previous selected name |
| `Alt+Down` | Go to next selected name |
| `Alt+F12` | Show folders history |
| `Ctrl+L` or `Alt+D` | Focus the address bar |
| `Ctrl+.` | Set current pane path from the other pane |
| `Shift+A` .. `Shift+Z` | Go to drive root (`<drive>:\`) |
| `Enter` | Execute / open the focused item |
| `Alt+Enter` | Open Properties |

## Selection

| Chord | Action |
|-------|--------|
| `Ctrl+A` | Select all |
| `Esc` | Unselect all |
| `Insert` | Select current item and move to next |
| `Space` | Select, calculate directory size, and move to next |
| `Shift+Space` | Quick Search in the current pane |
| `Ctrl+=` | Select... (select-by-pattern dialog) |
| `Ctrl+-` | Unselect... (unselect-by-pattern dialog) |
| `Ctrl+Shift+=` | Select items with the same extension |
| `Ctrl+Shift+-` | Unselect items with the same extension |
| `Ctrl+Shift+F5` | Save selection |
| `Ctrl+Shift+F6` | Restore selection |

## Clipboard, copy paths and drag

| Chord | Action |
|-------|--------|
| `Ctrl+C` or `Ctrl+Insert` | Clipboard copy |
| `Ctrl+X` | Clipboard cut |
| `Ctrl+V` or `Shift+Insert` | Clipboard paste |
| `Alt+Insert` | Copy path + name as text |
| `Ctrl+Alt+Insert` | Copy path as text |
| `Alt+Shift+Insert` | Copy name as text |
| `Ctrl+Shift+Insert` | Copy UNC path + name as text |

Drag-and-drop between panes and with Explorer is supported where the underlying file system allows it; see [File Operations](FileOperations.md).

## View, sort and pane layout

| Chord | Action |
|-------|--------|
| `Alt+2` | Display as Brief |
| `Alt+3` | Display as Detailed |
| `Alt+4` | Display as Extra Detailed |
| `Alt+5` | Toggle thumbnails |
| `Alt+6` | Toggle preview pane |
| `Ctrl+F2` | Sort: None |
| `Ctrl+F3` | Sort: Name |
| `Ctrl+F4` | Sort: Extension |
| `Ctrl+F5` | Sort: Time |
| `Ctrl+F6` | Sort: Size |
| `Ctrl+U` | Swap panes |
| `Ctrl+F11` | Zoom panel (maximize / restore the focused pane) |
| `Ctrl+Shift+F3` | View width |
| `Ctrl+Shift+F11` | Full screen |
| `Shift+F11` | Previous theme |
| `Shift+F12` | Next theme |

## Hot Paths

Hot Paths are ten quick-jump bookmark slots. See [Navigation and Paths](NavigationAndPaths.md) for managing them.

| Chord | Action |
|-------|--------|
| `Ctrl+1` .. `Ctrl+9`, `Ctrl+0` | Go to Hot Path slot 1-10 |
| `Ctrl+Shift+1` .. `Ctrl+Shift+9`, `Ctrl+Shift+0` | Set Hot Path slot 1-10 to the current folder |
| `Shift+F9` | Open the Hot Paths list |

`Ctrl+0` / `Ctrl+Shift+0` correspond to slot 10.

## File operations

See [File Operations](FileOperations.md) for behavior, confirmation, and failed-item handling.

| Chord | Action |
|-------|--------|
| `F5` | Copy to other pane |
| `F6` | Move to other pane |
| `F7` | Create directory |
| `F8` | Delete (Recycle Bin where supported) |
| `Del` | Move to Recycle Bin |
| `Shift+F8` or `Shift+Del` | Permanent delete (with confirmation) |
| `Ctrl+Shift+Del` | Permanent delete (with confirmation) |
| `F2` | Rename (Batch Rename when multiple items are selected) |
| `Ctrl+F7` | Change case |
| `Ctrl+F8` | Change attributes |
| `Alt+F5` | Pack (create ZIP) |
| `Alt+F6` | Unpack (extract ZIP) |
| `Ctrl+J` | Toggle File Operations Failed Items |

## Search and compare

See [Find Files and Directories](FindFiles.md) and [Compare Directories](CompareDirectories.md).

| Chord | Action |
|-------|--------|
| `Ctrl+F` or `Alt+F7` | Find files and directories |
| `Ctrl+F10` | Compare directories (the two panes) |
| `Ctrl+F12` | Filter the current folder |
| `Shift+Space` | Quick Search in the current pane |

## Command line, shell and miscellaneous

| Chord | Action |
|-------|--------|
| `Ctrl+Enter` | Bring filename to the command line |
| `Ctrl+Space` | Bring current directory to the command line |
| `Ctrl+Shift+Space` | Bring current directory to the command line |
| `Ctrl+Shift+Enter` | Bring filename to the command line |
| `Alt+Space` | Window menu |
| `Ctrl+Alt+T` | Open command shell |
| `F9` | User Menu |
| `Ctrl+Shift+F9` | Shared directories |
| `Alt+F11` | List of opened files |
| `Alt+F10` | View Space (occupied space) |
| `Alt+/` or `Alt+Shift+/` | About |
| `Shift+F3` | Open current folder (File Explorer) |
| `Ctrl+F9` | Refresh |

## See also

- [User Guide](UserGuide.md) - full feature tour, including the function bar and menus.
- [Navigation and Paths](NavigationAndPaths.md) - address bar, breadcrumb, history, and Hot Paths.
- [File Operations](FileOperations.md) - copy, move, delete, and failed-item handling.
- [Find Files and Directories](FindFiles.md) and [Compare Directories](CompareDirectories.md).
- [Preferences](Preferences.md) - the **Keyboard** page for customizing every binding.
