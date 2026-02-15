# Compare Directories Specification (`cmd/app/compare`)

## Overview

`cmd/app/compare` opens a dedicated **Compare Directories** window that behaves like a full dual‑pane `FolderWindow`, but with a **comparison scope**:

- The user chooses two folder roots (**Left** and **Right**).
- The window shows **differences** between the two hierarchies.
- Default view shows **only different items** (a “differences filter”).
- Standard file operations are available to **sync** the two folders (copy/move/delete/rename/view/etc.) using the existing FolderWindow command set and shortcuts.

The feature is directory‑oriented but compares **files with the same name** under the same relative folder.

## Implementation Files

- `RedSalamander/CompareDirectoriesWindow.h/.cpp` (top‑level window + options panel UI)
- `RedSalamander/CompareDirectoriesEngine.h/.cpp` (compare engine + compare-scoped filesystem wrappers)
- `Common/SettingsStore.h` + `Common/Common/SettingsStore.cpp` (persisted settings: `compareDirectories`)
- `Common/WindowMessages.h` (custom message IDs; no `WM_APP`/`WM_USER` definitions outside this file)

## Window Behavior (FolderWindow‑like)

The Compare Directories window must behave like the main dual‑pane folder window:

- Two panes with `NavigationView` + `FolderView` + (optional) status bars.
- Pane focus indicator, `Tab` focus switching, function bar, and the same shortcut handling as the main window.
- A themed vertical splitter between panes for resizing (same behavior as `FolderWindow`).

### Placement

- Window placement is persisted using `WindowPlacementPersistence` with the window id:
  - `CompareDirectoriesWindow`
- If there is no saved placement yet, the compare window opens with the **same normal position and size** as the invoking main window (not always-on-top).

### Banner

Above the panes, a banner is displayed:

- Title: **Compare Folder**
- Buttons:
  - **Options…** (shows the scan options panel)
  - **Rescan** (invalidates cached decisions and refreshes panes)

## Navigation & Scope Sync

Comparison is defined by **two roots**:

- `_leftRoot` (Left root)
- `_rightRoot` (Right root)

When the user navigates within one pane:

1. Compute the relative path from the changed pane root.
2. Resolve the same relative path under the other pane root.
3. Navigate the other pane to the resolved folder.

### Root changes (navigation outside scope)

If the user navigates to a folder that is **not under the current root** for that pane:

- Treat the new folder as the new root for that pane.
- Keep the other root unchanged.
- Invalidate the compare session (scan decisions) and resync both panes to the new roots.

## Scan Options (Settings Panel)

Scan options are configured via an in‑window panel shown before the first scan and accessible later via:

- Banner **Options…**
- Menu **Compare → Options…**

The panel uses themed preference‑style cards (title + description + toggle) and supports vertical scrolling when the window is small.

### Options

**Compare files with same name by**

- `compareSize` — Bigger files are selected.
- `compareDateTime` — Newer files are selected.
- `compareAttributes` — Files with different attributes are selected.
- `compareContent` — Files with different content are selected.

**Subdirectories options**

- `compareSubdirectories` — Compare subdirectories; directories with different content are selected; directories existing only in one panel are selected.

**Additional options**

- `compareSubdirectoryAttributes` — Directories with different attributes are selected.
- `selectSubdirsOnlyInOnePane` — Select subdirectories contained only in one panel.

**More options**

- `ignoreFiles` + `ignoreFilesPatterns` — Wildcards separated by `;`.
- `ignoreDirectories` + `ignoreDirectoriesPatterns` — Wildcards separated by `;`.

### Persistence

All options persist in settings:

- `Common::Settings::Settings::compareDirectories` (`Common::Settings::CompareDirectoriesSettings`)
- Settings are re-applied automatically when the compare command is used again.

## Scanning & Progress UI

The engine performs on-demand folder decisions for a given relative folder.

During scanning, the window shows a progress row with:

- A text status including the relative path being scanned and counts.
- A marquee progress bar (indeterminate).

Progress notifications use a posted payload message declared in `Common/WindowMessages.h`:

- `WndMsg::kCompareDirectoriesScanProgress`

## Differences Model

For each folder (relative to the roots), the engine produces an item decision per name:

- Existence: only in left, only in right, both
- Type mismatch: file vs directory
- Criteria differences (optional): size, time, attributes, content

### Default filter

By default, the panes show **only different items**:

- “Only in Left/Right” items
- Type mismatches
- Items different by any enabled compare criterion
- Directories that differ (when `compareSubdirectories` is enabled)

The menu item “Show Identical Items” (`showIdenticalItems`) switches to a full list view (no differences filter).

## Selection Rules

Initial selection after a scan:

- Items that exist **only in one pane** are selected in that pane.
- When a compare criterion is enabled and a file differs:
  - Size: select the bigger file’s side.
  - Date/time: select the newer file’s side.
  - Attributes: select both.
  - Content: select both.

After file operations, when the view is filtered to differences only, items that become identical naturally disappear from the list.

### Selection helpers (post-scan)

The window provides:

- **Restore Differences Selection**: re-select what the current decision model marks as “different”.
- **Invert Selection**: invert selection in the active pane (FolderView selection semantics).

## Per-item “Why is this different?” UI

Each item shows a second line of details under the name (FolderView detailed text) describing the difference, for example:

- “Only in left folder”
- “Only in right folder”
- “Different size (left bigger)”
- “Different date/time (right newer)”
- “Different attributes”
- “Different content”
- “Type mismatch”

Items should also have a clear visual difference (color/icon) consistent with the current theme.

## File Operations Semantics (Sync)

All standard operations should work with the same shortcuts as the main FolderWindow:

- View / View With
- Rename
- Copy / Move to other pane
- Delete / Permanent delete
- Create directory

### Directory copy/move in compare mode

When the user requests to copy/move a **folder** from one pane to the other in compare mode:

- Only items that are currently considered **different** under that folder’s hierarchy are copied/moved.
- The operation must preserve folder structure under the destination.

(This differs from normal FolderWindow behavior, which copies the entire directory tree.)

### Cache invalidation after operations

After a successful operation, the compare session invalidates cached decisions for affected paths so the UI updates quickly without forcing a full rescan.

## Content Compare I/O Contract

Content comparison must use the active filesystem plugin I/O:

- Use `IFileSystemIO::CreateFileReader` to open files.
- Read with `IFileReader::Read` (streaming) until a difference is found or EOF is reached.

This allows content compare to work for non-Win32 filesystems (archives, FTP/S3, etc.) and avoids direct Win32 file APIs (`CreateFileW`/`ReadFile`) in compare logic.

## Theming Requirements

All surfaces must be themed and respond to runtime theme changes:

- Compare window background, menu, banner buttons, and scan progress controls.
- Options panel cards, toggles, edits, and scrollbars.
- FolderViews inside the embedded FolderWindow.

No custom `WM_APP`/`WM_USER` message IDs should be introduced outside `Common/WindowMessages.h`.

## Testing

- Engine self-tests: `RedSalamander/CompareDirectoriesEngine.SelfTest.cpp` via `--compare-selftest`.
- Manual UI checks:
  - Shortcuts (function keys, selection, copy/move/delete) behave like the main window.
  - Navigation sync and root change invalidation.
  - Options panel theming and scroll behavior at small window sizes.
  - Progress UI visible during long scans.

