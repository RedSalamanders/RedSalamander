TODO

  1. Consider splitting CompareDirectoriesWindow.cpp into multiple files (following the FolderView pattern)
    - CompareDirectoriesWindow.Options.cpp — options panel dialog
    - CompareDirectoriesWindow.Progress.cpp — progress UI and ETA
    - CompareDirectoriesWindow.Menu.cpp — themed menu rendering

  2. Extract magic pixel values into named constants
  3. Harden wildcard matching against pathological patterns
  4. Add tests for empty directories and confirm progress counter consistency during invalidation

# Compare Directories Specification (`cmd/app/compare`)

## Overview

`cmd/app/compare` opens a dedicated **Compare Directories** window. It behaves like a dual‑pane `FolderWindow`, but with an explicit **comparison scope** defined by two folder roots:

- The user navigates each pane to the desired folders, then presses **Options → OK** or **Rescan** to establish (or re-establish) the compare roots and start a scan.
- The panes show differences between the two directory hierarchies.
- When **Show Identical Items** is off, the panes show only items that differ (a "differences filter").
- Standard file operations (copy/move/delete/rename/view, etc.) are available via the normal `FolderWindow` command set and shortcuts.

The comparison is directory-oriented and matches items by **name under the same relative folder**.

## Implementation Files

- `RedSalamander/CompareDirectoriesWindow.h/.cpp` (window, banner, options panel, sync logic, progress UI)
- `RedSalamander/CompareDirectoriesEngine.h/.cpp` (compare session/engine + compare-scoped filesystem wrappers)
- `Common/SettingsStore.h` + `Common/SettingsStore.cpp` (persisted settings: `compareDirectories`)
- `Common/WindowMessages.h` (custom message IDs; no `WM_APP`/`WM_USER` definitions outside this file)
  - Posted payload helpers: `Common/Helpers.h` (`PostMessagePayload` / `TakeMessagePayload` / `DrainPostedPayloadsForWindow`)

## Window Behavior (FolderWindow-like)

The Compare Directories window uses an embedded `FolderWindow`:

- Two panes with `NavigationView` + `FolderView` + per-pane status bars.
- A themed vertical splitter between panes (same interaction model as `FolderWindow`).
- Both panes use a compare-specific display mode (see [Display Modes](#display-modes-verbosity-levels)).
- Status bars are forced visible for both panes.
- On the first start of the compare view, the split ratio is set to **50%**.

### Placement

- Window placement is persisted using `WindowPlacementPersistence` with window id `CompareDirectoriesWindow`.
- If there is no saved placement yet, the compare window opens at the same **normal position and size** as the invoking main window.

### Banner (Title + Actions + Progress)

Above the panes, a banner region is displayed. The banner is themed consistently with the rest of the app and supports smooth progress updates (a Direct2D implementation is recommended but not required by the spec).

Contents:

- Title: **Compare Folder**
- Buttons:
  - **Options…** — shows the in-window options panel.
  - **Rescan** — establishes compare roots from the panes' current folders and starts a fresh scan. While a scan is active, this button changes to **Cancel**.
    - **Cancel** cancels the current scan and pending content compares and returns the window to an idle state.
    - Cancel must be responsive (no long UI-thread stalls while stopping/clearing large compare state).
    - After Cancel, background compare work remains paused until the next explicit **Rescan** (or **Options → OK**).

When scanning is active and/or background content compare is pending, a **progress row** appears below the banner title/buttons:

- A status text describing what is being processed (scan path + counts; and/or content compare path + bytes).
- Scan elapsed time and (when possible) a content-compare ETA.
- **Indeterminate progress**: a themed spinner/indeterminate animation while directory scanning is active or when total content bytes are unknown (no Win32 default marquee visuals).
- **Determinate progress**: a themed progress bar (following the current theme colors and rainbow mode, consistent with the file operations dialog — see `FolderWindow.FileOperations.Popup` rendering) when content compare reports a known total byte count.

The progress row hides automatically when there are no active scans and no pending content compares.

### File Operations popup integration (scan/content progress)

Compare scanning and content-compare progress must also be visible in the **File Operations** popup (`FolderWindow.FileOperations.Popup.*`) as a task card.

Requirements:

- When a compare run starts (Options → OK / Rescan), create or **re-use** a single **Compare Folders** task card per compare window. A Rescan replaces the previous task card's state (no accumulation of old cards).
- The File Operations popup should be shown (or brought to the foreground) when this task starts, consistent with how file operation tasks surface progress (shown without stealing focus).
  - If the popup window exists but is off-screen (monitor/layout changes), it is repositioned into the current monitor work area.
- The task card shows:
  - The two roots being compared (Left / Right).
  - A **Scan / Preflight** phase:
    - Current relative folder / entry being scanned.
    - Monotonic counters (folders scanned, entries scanned).
    - Elapsed scan time.
    - When `compareContent` is enabled: total candidate file count and total bytes to compare (accumulated as sizes become known).
  - A **Content compare** phase (when `compareContent` is enabled and pending compares exist):
    - Parallel in-flight compares: show up to `kMaxContentInFlightFiles` entries (one per content worker) as individual lines with mini progress bars (like copy/move in-flight lines).
    - Each line shows the current file (relative path) and per-file progress (completed bytes / total bytes when known).
    - Overall progress across all pending content compares (total completed bytes / total bytes when sizes are known).
    - A pending/completed count for content compares.
    - An ETA (when total bytes are known and a stable rate estimate can be computed).
- When both scan and content compare are complete, the task card transitions to a compact **Done** state showing summary counts (files compared, differences found). The Done state auto-dismisses after **5 seconds** or when the compare window is closed, whichever comes first.
  - If `fileOperations.autoDismissSuccess` is enabled, the File Operations layer may auto-dismiss the compare task immediately on success/cancel.
- This compare task is **informational only**: it must not participate in file-ops Wait/Parallel queuing and must not block normal copy/move/delete tasks. Implementation: use a distinct task category (e.g., `FileOperationTaskCategory::Informational`) that the queue logic skips.

## Compare Mode vs Browsing Mode

The compare window has two user-visible modes:

- **Compare mode (active)**:
  - The panes are synchronized by relative path under the current roots.
  - Enumeration is driven by the compare engine (differences filter, per-item status text, background content compare).
  - Progress UI is shown while scanning/content compare is running.
  - The banner **Rescan** button changes to **Cancel**.
- **Browsing mode (inactive)**:
  - The panes behave like a normal `FolderWindow` (independent navigation).
  - Enumeration is delegated to the underlying filesystem (no compare decisions/progress).
  - Any selection from a previous scan is cleared.

Compare mode is entered (or re-entered) only by **Options → OK** or **Rescan**. Navigation alone never implicitly changes the compare roots.

## Navigation & Scope Sync

Comparison is defined by two roots:

- Left root (`_leftRoot`)
- Right root (`_rightRoot`)

While compare mode is active, when the user navigates within one pane:

1. Compute the relative path from that pane's root.
2. Resolve the same relative path under the other pane's root.
3. Navigate the other pane to that resolved folder.

If the resolved folder does not exist on one side, compare mode remains active:

- The existing side navigates normally.
- The missing side enumerates as **empty** (missing-path errors are treated as "no entries", not as an error overlay).
- The missing side shows a localized empty-state message centered in the pane: **"This folder doesn't exist in this hierarchy."** (resource string `IDS_COMPARE_FOLDER_NOT_FOUND`).

### Leaving the compare scope (navigation outside roots)

If the user navigates to a folder that is not under the current root for that pane:

- If a scan or content-compare is active, prompt before canceling compare mode:
  - Message: "Leaving the compare scope will cancel the current scan and pending content compares."
  - Buttons:
    - **Cancel**: keep the current folder (do not navigate).
    - **OK**: cancel the scan and proceed with navigation.
- If no scan is active (or the user confirmed OK), compare mode is canceled:
  - Synchronized navigation stops.
  - Progress UI is hidden.
  - Enumeration delegates to the underlying filesystem so both panes can browse normally.
- The roots are **not** implicitly changed.
- A subsequent **Options → OK** or **Rescan** establishes a new scope using the panes' current folders as the new roots.

## Options Panel (Scan Settings)

Scan options are configured via an in-window panel shown on first open and accessible later via:

- Banner **Options…**
- Menu **Compare → Options…**

The panel uses themed preference-style cards (title + description + toggle) and supports vertical scrolling when the window is small.

### Panel layout

Options are organized in four sections:

| Section | Options |
|---------|---------|
| **Compare files with same name by** (`IDS_COMPARE_OPTIONS_SECTION_COMPARE`) | `compareSize`, `compareDateTime`, `compareAttributes`, `compareContent` |
| **Subdirectories** (`IDS_COMPARE_OPTIONS_SECTION_SUBDIRS`) | `compareSubdirectories` |
| **Additional** (`IDS_COMPARE_OPTIONS_SECTION_ADVANCED`) | `compareSubdirectoryAttributes`, `selectSubdirsOnlyInOnePane` |
| **Ignore patterns** (`IDS_COMPARE_OPTIONS_SECTION_IGNORE`) | `ignoreFiles` + `ignoreFilesPatterns`, `ignoreDirectories` + `ignoreDirectoriesPatterns` |

### Panel actions

- **OK**:
  - Saves settings to `SettingsStore`.
  - Cancels any in-progress scan.
  - Enters (or re-enters) compare mode and starts a fresh scan using the panes' current folders as roots.
- **Cancel**:
  - If compare mode has not started yet (first open), closes the compare window.
  - Otherwise, hides the panel and returns to the panes without changing compare mode or settings.

### Options

**Compare files with same name by**

- `compareSize` — Different sizes mark the item different; the bigger side is selected.
- `compareDateTime` — Different timestamps mark the item different; the newer side is selected.
- `compareAttributes` — Different attributes mark the item different; both sides are selected.
- `compareContent` — Different content marks the item different; both sides are selected. Content is evaluated asynchronously (see [Content Compare Architecture](#content-compare-architecture)).

**Subdirectories**

- `compareSubdirectories` — Compare subdirectories. A directory is marked "different" when any descendant differs under the same relative path.

**Additional**

- `compareSubdirectoryAttributes` — Directories with different attributes are selected in both panes.
- `selectSubdirsOnlyInOnePane` — When enabled, directories that exist only on one side are selected on that side. (Files that exist only on one side are always selected regardless of this setting.)

**Ignore patterns**

- `ignoreFiles` + `ignoreFilesPatterns` — File name wildcards separated by `;` (e.g., `*.obj;*.pdb`). Matched files are excluded from comparison.
- `ignoreDirectories` + `ignoreDirectoriesPatterns` — Directory name wildcards separated by `;` (e.g., `.git;node_modules`). Matched directories (and their subtrees) are excluded from comparison.

### Persistence

All options persist under:

- `Common::Settings::Settings::compareDirectories` (`Common::Settings::CompareDirectoriesSettings`)

Settings are loaded when the compare window opens and saved on Options → OK.

## Differences Model

For each relative folder, the engine computes an item decision per (normalized) name:

- **Existence**: only in left, only in right, both
- **Type mismatch**: file vs directory with the same name
- **Criteria differences** (when both exist and are the same type): size, time, attributes, content

### Difference bits (`CompareDirectoriesDiffBit`)

| Bit | Meaning |
|-----|---------|
| `OnlyInLeft` | Item exists only on the left side |
| `OnlyInRight` | Item exists only on the right side |
| `TypeMismatch` | Same name but one side is a file, the other a directory |
| `Size` | File sizes differ |
| `DateTime` | Last-write timestamps differ |
| `Attributes` | File attributes differ |
| `Content` | File content differs (confirmed by content compare) |
| `ContentPending` | Content compare is queued or in progress (not yet resolved) |
| `SubdirAttributes` | Directory attributes differ (when `compareSubdirectoryAttributes` enabled) |
| `SubdirContent` | A descendant differs (when `compareSubdirectories` enabled) |
| `SubdirPending` | A descendant comparison is pending (e.g., content compares still running) |

### Name matching

- Names are compared using Windows ordinal, case-insensitive semantics (`CompareStringOrdinal(..., TRUE)`).
- The decision map uses an ordered `std::map` with a `CompareStringOrdinal`-based comparator (`WStringViewNoCaseLess`). This avoids the hash/equality contract violation that would arise from using `std::unordered_map` with a different hash function and ordinal equality.
- Trailing spaces/dots are ignored for comparison to reduce false mismatches across enumeration backends and Win32 path semantics.

### Default filter (Show Identical Items)

By default, the panes show only different items:

- "Only in left/right" items
- Type mismatches
- Items different by any enabled compare criterion
- Directories that differ (when `compareSubdirectories` is enabled)
- Items that are still in-flight (`ContentPending` / `SubdirPending`) so the user can see progress even when `showIdenticalItems` is off.

The menu item **Compare → Show Identical Items** (`showIdenticalItems`) toggles the differences filter globally for all folders in the compare session. When enabled, identical items are shown alongside different items. The setting is persisted in `CompareDirectoriesSettings`.

Toggling the filter triggers an immediate pane refresh (re-enumeration through the compare filesystem wrapper with the updated filter).

### Subdirectory traversal

When `compareSubdirectories` is enabled:

- Subtree comparison is performed **iteratively** using an explicit stack/worklist (no recursive call stack growth, safe for deeply nested trees like `node_modules`).
- Directory reparse points (symlinks/junctions) are **not** followed.

### Error handling for enumeration failures

If one side fails to enumerate a folder (e.g., access denied):

- The decision for that folder is marked with a failed `HRESULT`.
- The failed folder is treated as "different" (it appears in the differences list).
- The failed decision is **not cached** — the engine retries enumeration on the next access (e.g., after the user fixes permissions and navigates back).
- No error overlay is shown in the pane; the folder simply appears as a different item. The per-item details line shows the failure reason (e.g., "Access denied").

## Selection Rules

Selection is applied after enumeration completes (per-pane) while compare mode is active.

Initial selection behavior:

- Files that exist only in one pane are selected in that pane.
- Directories that exist only in one pane are selected only when `selectSubdirsOnlyInOnePane` is enabled.
- Directories with differing attributes (when `compareSubdirectoryAttributes` is enabled) are selected in both panes.
- Directories with differing subtree content (when `compareSubdirectories` is enabled) are selected in both panes.
- Items with `ContentPending` / `SubdirPending` are **not** selected by default (pending comparisons are not treated as final differences).
- When a compare criterion is enabled and a file differs:
  - Size: select the bigger file's side.
  - Date/time: select the newer file's side.
  - Attributes: select both.
  - Content: select both.

During file operations in compare mode, items that become identical (after a copy/move) naturally disappear from the filtered list (when `showIdenticalItems` is off). Remaining different items stay selected so the user can see sync progress visually.

### Selection helpers

The window provides:

- **Compare → Restore Differences Selection**: re-apply the decision model's default selection (as described above) to both panes. Useful after the user manually changes selection.
- **Compare → Invert Differences Selection**: for each item in both panes, select items that the decision model would **not** select by default, and deselect items that it **would** select. This is a decision-model inversion (not a simple toggle of each item's current selected state).

## Display Modes

The compare window supports the standard `FolderView` display modes and applies the chosen mode to **both panes**. A top-level **View** menu provides the same commands (labeled with the current shortcuts), and the same commands are available via the normal shortcut system.

This requires adding a third value to `FolderView::DisplayMode`:

```cpp
enum class DisplayMode : uint8_t
{
    Brief,          // 1 line:  name only
    Detailed,       // 2 lines: name + details
    ExtraDetailed,  // 3 lines: name + details + metadata
};
```

The compare window intercepts the normal Brief/Detailed shortcut commands and cycles through these three levels:

| Shortcut | Level | Lines | Content |
|----------|-------|-------|---------|
| **Alt+2** | Brief | 1 | File name only. |
| **Alt+3** | Detailed | 2 | File name + details line (difference status when applicable; otherwise metadata). |
| **Alt+4** | Extra Detailed | 3 | File name + difference details line + metadata line (size, date/time, attributes). |

The default level on entering compare mode is **Detailed** (Alt+3).

### Per-item details line examples

The details line describes why the item is different (or pending). When an item has no difference/pending status, **Detailed** mode shows metadata on the details line.

- "Only in left folder"
- "Only in right folder"
- "Bigger • Newer • Content differs" (multiple criteria combined with separator)
- "Smaller • Older"
- "Attributes differ"
- "Content differs"
- "Comparing..." (content compare pending — `ContentPending` bit is set)
- "Computing..." (descendant comparisons pending — `SubdirPending` bit is set on a directory)
- "Type mismatch"
- "Access denied" (enumeration failure on one side)

All detail strings are loaded from `.rc` string resources (`IDS_COMPARE_DETAILS_*`).

### Per-item metadata line

The metadata line (third line in Extra Detailed mode) shows:

- File size (formatted consistently with `FolderView` size display).
- Last-write date/time.
- File attributes (archive, read-only, hidden, system, etc.).

Items should have a clear visual distinction (foreground color or icon) based on their difference status, consistent with the current theme.

### Transition behavior for content-pending items

When a content compare completes for an item:

- The `ContentPending` bit is cleared.
- If content differs, the final `Content` bit is set and the item becomes different/selected.
- If content is equal, the item remains identical and unselected (no content bits remain set).
- The pane refreshes via `WndMsg::kCompareDirectoriesDecisionUpdated`.
- If the item is now identical and `showIdenticalItems` is off, the item is removed from the list on the next enumeration refresh.
- No explicit animation or fade-out is applied — the item simply disappears from the filtered list or updates its details text.
  - Ancestor directories update `SubdirPending` / `SubdirContent` so directory status transitions correctly (e.g., "Computing..." → final).

## Content Compare Architecture

Content comparison uses a **two-phase** model to keep the UI responsive:

### Phase 1: Metadata scan (synchronous, per-folder)

When `GetOrComputeDecision` is called for a folder:

1. Enumerate both left and right sides.
2. Match entries by name (case-insensitive ordinal).
3. Compare metadata (existence, type, size, date/time, attributes) based on enabled settings.
4. For files that need content comparison (same size or size-unknown, and `compareContent` is enabled): set the `ContentPending` diff bit and enqueue a content-compare job. `ContentPending` does not imply a final difference and must not select the item.
5. Return the decision immediately — the UI shows "Comparing..." for content-pending items.

### Phase 2: Background content compare (asynchronous, worker pool)

- A pool of `std::jthread` workers (sized to `std::thread::hardware_concurrency() / 2`, minimum 1) processes the content-compare queue (add a setting for the level of paraellelism 0 or no setting use default value).
- Workers are created lazily on first content-compare enqueue (`EnsureContentCompareWorkersLocked`).
- Each worker dequeues a `ContentCompareJob`, checks the version (bail out if stale), opens both files via `IFileSystemIO::CreateFileReader`, and streams/compares until a difference or EOF.
- Results are stored in `_pendingContentCompareUpdates`, keyed by folder and entry name.
- Pending updates are applied by `FlushPendingContentCompareUpdates()` (called by the UI when `WndMsg::kCompareDirectoriesDecisionUpdated` is received) and also on-demand the next time `GetOrComputeDecision` runs for a specific folder.
- Applying a pending update must also update ancestor directories' subtree state (`SubdirPending` / `SubdirContent`) so directory status transitions correctly without requiring navigation.
- On completion, the engine posts `WndMsg::kCompareDirectoriesDecisionUpdated` so the panes refresh.

### Cancellation

- Workers check `std::stop_token` and the session `_version` inside the read loop.
- When the user clicks Rescan, changes settings, or navigates outside the roots: `_version` increments, the content-compare queue is cleared, and workers bail out of in-progress jobs on the next cancellation check.

### Notification throttling

Progress callbacks (`ScanProgressCallback`, `ContentProgressCallback`, `DecisionUpdatedCallback`) are throttled by a minimum interval (tracked via `_*LastNotifyTickMs` atomics) to avoid flooding the UI thread with `PostMessage` calls. A final forced notification is always sent when a scan or the last content-compare completes, ensuring the UI reaches the final state.

## Cache Coherence & Version Model

The engine uses an atomic `uint64_t _version` counter as the primary cache coherence mechanism:

- **Version increments** on: `SetRoots`, `SetSettings`, `Invalidate`.
- **Cached decisions** (`_cache`) are tagged with the version at computation time. A cached entry is valid only if its version matches the current `_version`; stale entries are recomputed on next access.
- **In-flight content-compare jobs** carry the version at enqueue time. Workers discard results if the version changed during comparison.
- **Content-compare caches** (`_contentCompareCache`, `_contentCompareInFlight`) are cleared on version change (`ClearContentCompareStateLocked`).

### UI version (`_uiVersion`)

A separate `_uiVersion` tracks the last version the UI has observed. This prevents redundant pane refreshes: the window only triggers a re-enumeration when `_uiVersion` differs from `_version`.

### Cache invalidation after file operations

`InvalidateForAbsolutePath(path, includeSubtree)`:

1. Converts the absolute path to a relative path under the left or right root.
2. Removes the matching cache entry (and all subtree entries if `includeSubtree` is true) under lock.
3. Bumps `_version` so stale decisions are not served.

This is currently O(n) over cache entries for subtree invalidation. For large trees, consider using the ordered `std::map` with prefix-range lookup (`lower_bound`) for O(log n + k) invalidation.

## IFileSystem Wrapper (`CompareDirectoriesFileSystem`)

The compare engine exposes a per-pane filesystem wrapper created via:

```cpp
wil::com_ptr<IFileSystem> CreateCompareDirectoriesFileSystem(ComparePane pane, std::shared_ptr<CompareDirectoriesSession> session);
```

The wrapper implements `IFileSystem` and `IInformations`. Behavior:

### `ReadDirectoryInfo` (enumeration)

- If compare mode is **active** and the requested folder is **within** the pane's root:
  - Call `session->GetOrComputeDecision(relativeFolder)` to get the folder decision.
  - For each item in the decision:
    - Skip items that don't exist on this pane's side (`existsLeft`/`existsRight`).
    - Skip identical items if `showIdenticalItems` is off and `isDifferent` is false.
    - Populate the directory entry from the decision's metadata (size, date/time, attributes).
    - Set the per-item details text from the decision's difference mask.
  - Apply the decision model's selection rules to the returned entries.
- If compare mode is **inactive**, or the folder is **outside** the roots:
  - Delegate to the base filesystem (`_baseFileSystem->ReadDirectoryInfo`).

### Missing-side folders

When one side doesn't exist for a relative path:

- `ReadDirectoryInfo` returns an empty entry list (success, zero entries).
- The `FolderView` renders the localized empty-state message (`IDS_COMPARE_FOLDER_NOT_FOUND`).

### File operations (copy/move/delete/rename)

All mutation operations delegate to the base filesystem. The compare wrapper does not intercept or filter these — filtering for sync-copy is handled at a higher level (see [Directory sync-copy](#directory-copymove-behavior-sync)).

## Scanning & Progress Notifications

Enumeration in compare mode is driven by the compare-scoped filesystem wrapper (see above).

Progress notifications use posted payload messages declared in `Common/WindowMessages.h`:

- `WndMsg::kCompareDirectoriesScanProgress` — scan phase updates (folder/entry counts, current path).
- `WndMsg::kCompareDirectoriesContentProgress` — content compare updates (per-file bytes, pending count).
- `WndMsg::kCompareDirectoriesDecisionUpdated` — causes the panes to refresh so "Comparing..." transitions to final results.

Notes:

- Posted payload receivers must use `TakeMessagePayload<T>(lParam)` and windows must drain pending payloads in `WM_NCDESTROY`.
- The compare window tags progress updates with a **run id** (the `_version` at scan start) and ignores stale updates whose run id does not match the current version.

## File Operations Semantics (Sync)

All standard operations work with the same shortcuts and command set as the main `FolderWindow`:

- View / View With
- Rename
- Copy / Move to other pane
- Delete / Permanent delete
- Create directory

### Directory copy/move behavior (sync)

When compare mode is active, Copy/Move to other pane for **directories** acts as a **sync** operation:

- It operates on the selected paths.
- The compare engine provides a **diff manifest** for the selected directory's subtree: the window iteratively walks the compare decisions for all descendant folders and collects only the items marked as different.
- The collected different items are passed to the normal file-operation pipeline as the source list. The pipeline creates destination directories as needed, but **only copies/moves files that are in the diff manifest**.
- Identical descendants are **not** copied/moved.
- The expansion respects the current compare settings (ignore patterns, enabled criteria, `compareSubdirectories`).
- When a directory exists **only on one side**, its entire subtree is considered different and is eligible for copy/move (no per-file filtering needed — the whole tree is new).
- Implementation approach: the compare window builds a flat list of `(sourceAbsolutePath, destinationAbsolutePath)` pairs from the diff manifest and submits this list to the existing file-operation copy/move handler, bypassing the normal recursive directory enumeration that the copy handler would otherwise perform.

### Cache invalidation after operations

When compare mode is active, after file operations complete the compare session calls `InvalidateForAbsolutePath` for affected source/destination paths so the UI updates quickly without forcing a full rescan.

## Content Compare I/O Contract

Content comparison uses the active filesystem plugin I/O:

- Use `IFileSystemIO::CreateFileReader` to open files on both sides.
- Read with `IFileReader::Read` (streaming, 256 KB buffer) until a difference is found or EOF is reached.
- Short reads are permitted by the `IFileReader` contract; content compare must tolerate them (continue reading until EOF or a mismatch).
- If `IFileReader::GetSize` succeeds for both sides and sizes differ, return "different" immediately without reading content.
- If `IFileReader::GetSize` succeeds for both sides and both sizes are 0, return "equal" immediately.
- If sizes are unknown (plugin doesn't report sizes), stream until EOF or a mismatch.

This keeps content compare filesystem-agnostic (archives, FTP/S3, etc.) and avoids direct Win32 file APIs (`CreateFileW`/`ReadFile`) in compare logic.

## Theming Requirements

All surfaces must be themed and respond to runtime theme changes:

- Compare window background, menu, banner (D2D-rendered), and progress controls.
- Options panel cards, toggles, edits, and scrollbars.
- FolderViews inside the embedded `FolderWindow`.
- Per-item difference colors/icons must follow the current theme palette.

No custom `WM_APP`/`WM_USER` message IDs should be introduced outside `Common/WindowMessages.h`.

## Settings Requirements

The global **Preferences** dialog includes a dedicated **Compare Directories** page that edits the persisted defaults used by `cmd/app/compare`.

This section maps directly to `Common::Settings::CompareDirectoriesSettings`:

- Compare files: `compareSize`, `compareDateTime`, `compareAttributes`, `compareContent`
- Subdirectories: `compareSubdirectories`, `compareSubdirectoryAttributes`, `selectSubdirsOnlyInOnePane`
- Ignore patterns: `ignoreFiles` + `ignoreFilesPatterns`, `ignoreDirectories` + `ignoreDirectoriesPatterns`
- Display: `showIdenticalItems`

The Compare Directories Options panel reads/writes the same settings; Preferences provides a centralized way to configure these defaults outside the compare window.

## Testing

- Engine self-tests: `RedSalamander/CompareDirectoriesEngine.SelfTest.cpp` via `--compare-selftest`.
- Manual UI checks:
  - Shortcuts (function keys, selection, copy/move/delete) behave like the main window.
  - Compare scope behavior: OK/Rescan establishes roots; navigation outside roots cancels compare mode (with a prompt if scan/content compare is active) until the next OK/Rescan.
  - Options panel theming and scroll behavior at small window sizes.
  - Progress appears in the compare banner and in the File Operations popup task card; it disappears when idle.
  - All three display modes (Alt+2, Alt+3, Alt+4) render correctly.
  - Directory sync-copy only copies different items (verify identical files are skipped).
  - Content-pending items show "Comparing..." and transition to final state without getting stuck.
  - Missing-side folders show the empty-state message.
  - Enumeration failures (access denied) show the failure reason in the details line.
