# Compare Directories Specification (`cmd/app/compare`)

## Overview

`cmd/app/compare` opens a dedicated **Compare Directories** window. It behaves like a dual‑pane `FolderWindow`, but with an explicit **comparison scope** defined by two folder roots:

- The user navigates each pane to the desired folders, then presses **Options → OK** or **Rescan** to establish (or re-establish) the compare roots and start a scan.
- The panes show differences between the two directory hierarchies.
- By default, the panes show only items that differ (a "differences filter"). If `keepIdenticalItems` is enabled, the user can toggle **Compare → Show Identical Items** to include identical items in the view.
- Standard file operations (copy/move/delete/rename/view, etc.) are available via the normal `FolderWindow` command set and shortcuts.

The comparison is directory-oriented and matches items by **name under the same relative folder**.

Related normative documents:

- `Specs/Testing/Testing_SelfTests.md`
- `Specs/Testing/Testing_PerformanceValidation.md`

## Performance Validation Contract

Compare Directories is a performance-sensitive subsystem across both engine and window paths. Any new feature or optimization that can affect scan throughput, content-compare queueing, visible progress latency, pane refresh work, cache behavior, or memory retention MUST:

- identify the user-visible scenario being protected,
- add or reuse measurable instrumentation,
- add deterministic coverage in `--compare-selftest`, `--commands-selftest`, or another deterministic harness,
- archive validation runs under `Specs/TestRuns/`,
- ground any claimed improvement in archived before/after evidence.

This requirement applies from the beginning of feature work, including baseline-establishing landings.

When the behavior is window-facing, the preferred metric family is `compare.ui.*`. When the behavior is engine/storage-facing, the change SHOULD keep or extend deterministic compare selftest coverage alongside the instrumentation. Compare sync manifest planning and resolved-item submission use `compare.sync.manifest.*` counters for build count, elapsed time, ready item count, blocker count/reason, and submitted resolved-item count.

## Implementation Files

- `RedSalamander/CompareDirectoriesWindow.h` + `.Internal.h` (window declarations + internal helpers shared across the window translation units)
- `RedSalamander/CompareDirectoriesWindow.cpp` (window, banner, sync logic)
- `RedSalamander/CompareDirectoriesWindow.Options.cpp` (in-window options panel)
- `RedSalamander/CompareDirectoriesWindow.Progress.cpp` (progress UI and ETA)
- `RedSalamander/CompareDirectoriesWindow.Menu.cpp` (themed menu rendering)
- `RedSalamander/CompareDirectoriesEngine.h/.cpp` (compare session/engine + compare-scoped filesystem wrappers)
- `Common/SettingsStore.h` + `Common/Common/SettingsStore.cpp` (persisted settings: `compareDirectories`)
- `Common/WindowMessages.h` (custom message IDs; no `WM_APP`/`WM_USER` definitions outside this file)
  - Posted payload helpers: `Common/Helpers.h` (`PostMessagePayload` / `TakeMessagePayload` / `DrainPostedPayloadsForWindow`)

### Window Controller State

`CompareDirectoriesWindow` remains the Win32 message owner, but its major UI
state is grouped by owned controller buckets in `.Internal.h`:

- `ChromeController` owns menu-bar and banner HWND/DxUi state plus rescan/cancel
  chrome state.
- `OptionsPanelController` owns the Options dialog host/body state, DxUi
  controls, draft settings, scrolling/layout flags, and Options-specific theme
  brushes/colors.
- `ProgressController` owns banner progress data, scan/content ETA state,
  spinner/watermark timing, progress host state, and File Operations task-card
  state.

New compare-window state should be added to the narrowest controller bucket that
owns its lifecycle. Keep cross-cutting run/session fields on
`CompareDirectoriesWindow` only when they are not owned by chrome, options, or
progress.

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

### Lifetime / shutdown

- The Compare Directories window is an independent top-level window (not a child/owned window).
- Closing the Compare Directories window MUST restore focus to the invoking main-window folder view when that view still exists. The surrounding main navigation shell MUST return without lingering address-bar edit, history, suggestion, or full-path popup state, and the current path, focused item, selection count, item count, and refresh count for the invoking pane MUST remain stable across open/close.
- When the application shuts down (main window closes), the host closes any remaining unowned top-level RedSalamander windows during shutdown teardown so Direct2D resources can be released before process exit.

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
- A themed indeterminate spinner is shown **to the left of the Rescan/Cancel button** while any scan/content compare work is pending.
  - The buttons + spinner are vertically centered within the banner region; when the progress row appears/disappears, they re-center.

The progress row hides automatically when there are no active scans and no pending content compares.

### In‑pane background watermark (run state)

While a compare run is active, both panes MUST render a centered watermark in the **pane background** (drawn behind items; not a separate overlay window) to make it obvious the comparison is still in progress:

- **In progress** (scan/content work running): show an **animated** watermark (e.g., pulsing text opacity).
  - The animation pulse MUST be driven by compare run state and pane invalidation, not by the existence or visibility of the banner spinner control. If the spinner window is unavailable, the in-pane watermark must continue to animate while scan/content work is active.
- **Cancelled** (user pressed **Cancel**): show a **static** watermark indicating the run was cancelled, and keep it visible until the next explicit **Rescan** / **Options → OK**.
- **Done** (run finished successfully): show **no** watermark.

Because the watermark is rendered as part of the pane background, it is inherently click‑through (must not block normal pane interaction).

While a compare watermark is visible, panes MUST suppress the standard **Empty folder** empty‑state UI (it is misleading for differences‑filtered compare views).

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
    - An ETA (when total bytes are known and a stable, finite rate estimate can be computed). The UI must suppress ETA display for near-stalled rates or estimates above the configured sane display threshold instead of rendering very large durations.
- When both scan and content compare are complete, the task card transitions to a compact **Done** state showing summary counts (files compared, differences found). The Done state auto-dismisses after **5 seconds** or when the compare window is closed, whichever comes first.
  - If `fileOperations.autoDismissSuccess` is enabled, the File Operations layer may auto-dismiss the compare task immediately on success/cancel.
- Task-card progress updates may be throttled for UI responsiveness, but a throttled intermediate update MUST schedule a one-shot trailing flush after the throttle window. The trailing flush MUST be cancelled when the task reaches Done, is dismissed, or the compare window is torn down.
- This compare task is **informational only**: it must not participate in file-ops Wait/Parallel queuing and must not block normal copy/move/delete tasks. Implementation: use a distinct task category (e.g., `FileOperationTaskCategory::Informational`) that the queue logic skips.

## Compare Mode vs Browsing Mode

The compare window has two user-visible modes:

- **Compare mode (active)**:
  - The panes are synchronized by relative path under the current roots.
  - Enumeration is driven by the compare engine (differences filter, per-item status text, background content compare).
    - Enumeration MUST be cache-only: the compare filesystem wrapper MUST NOT compute decisions synchronously during `ReadDirectoryInfo()`.
      - When background work is enabled and a folder decision is not yet cached, the host MUST queue a background scan for that folder and MAY return an empty enumeration until the decision becomes available.
  - Progress UI is shown while scanning/content compare is running.
  - The banner **Rescan** button changes to **Cancel**.
- **Browsing mode (inactive)**:
  - The panes behave like a normal `FolderWindow` (independent navigation).
  - Enumeration is delegated to the underlying filesystem (no compare decisions/progress).
  - Any selection from a previous scan is cleared.

Compare mode is entered (or re-entered) only by **Options → OK** or **Rescan**. Navigation alone never implicitly changes the compare roots.

## Navigation & Scope Sync

Comparison scope is defined by the combination of **filesystem context** and **roots**:

- Left filesystem context: `(pluginId, instanceContext)`
- Right filesystem context: `(pluginId, instanceContext)`
- Left root (`_leftRoot`) — a pane-local filesystem path (Win32 absolute for `file`, plugin path like `/...` for non-`file` plugins)
- Right root (`_rightRoot`) — same rules as left

While compare mode is active, the filesystem contexts are **fixed**:

- If either pane’s `(pluginId, instanceContext)` changes, the host MUST cancel compare mode before doing any further compare-root/path calculations.

### Root/path semantics (Win32 vs plugin paths)

- If the pane uses the `file` filesystem (Win32 paths), roots and navigation are treated as **Windows absolute paths** (case-insensitive; extended paths like `\\?\` are accepted).
- Otherwise, roots and navigation are treated as **plugin paths**:
  - forward slashes (`/`)
  - a leading `/` is required
  - host normalization uses `NavigationLocation::NormalizePluginPathText(...)` rules (and MUST NOT apply Win32 `\\?\`-style normalization).

### Synchronized navigation (relative path under roots)

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
  - While the prompt is pending, the pane MUST be reverted to an in-scope folder before the navigation callback returns. If there is no previous in-scope folder, the session root for that pane is the required fallback. The prompt is then posted to the compare window; it MUST NOT be shown inline from the navigation callback.
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

While the panel is visible, the **Options…** command MUST be disabled/greyed out (to prevent re-entrant panel activation).

The panel uses themed preference-style cards (title + description + toggle) and supports vertical scrolling when the window is small. When the body is not scrollable, wheel delta remainders must be reset so a later non-scrollable-to-scrollable layout transition cannot jump unexpectedly.

### Panel layout

Options are organized in four sections:

| Section | Options |
|---------|---------|
| **Compare files with same name by** (`IDS_COMPARE_OPTIONS_SECTION_COMPARE`) | `compareSize`, `compareDateTime`, `compareAttributes`, `compareContent` |
| **Subdirectories** (`IDS_COMPARE_OPTIONS_SECTION_SUBDIRS`) | `compareSubdirectories` |
| **Additional** (`IDS_COMPARE_OPTIONS_SECTION_ADVANCED`) | `compareSubdirectoryAttributes`, `selectSubdirsOnlyInOnePane`, `keepIdenticalItems` |
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
  - If either side does not support `IFileSystemIO`, content comparison is unavailable for that compare scope. The host MUST disable and uncheck the toggle in the options UI for that compare run, persist `compareContent=false` when Options is accepted, and treat `compareContent` as **off** (no `Content`/`ContentPending` differences are computed).

**Subdirectories**

- `compareSubdirectories` — Compare subdirectories. A directory is marked "different" when any descendant differs under the same relative path.

**Additional**

- `compareSubdirectoryAttributes` — Directories with different attributes are selected in both panes.
- `selectSubdirsOnlyInOnePane` — When enabled, directories that exist only on one side are selected on that side. (Files that exist only on one side are always selected regardless of this setting.)
- `keepIdenticalItems` — Retain identical entries so **Compare → Show Identical Items** can be toggled without rescanning. Uses more memory on very large trees.

**Ignore patterns**

- `ignoreFiles` + `ignoreFilesPatterns` — File name wildcards separated by `;` (e.g., `*.obj;*.pdb`). Matched files are excluded from comparison.
- `ignoreDirectories` + `ignoreDirectoriesPatterns` — Directory name wildcards separated by `;` (e.g., `.git;node_modules`). Matched directories (and their subtrees) are excluded from comparison. This exclusion applies both when the directory is encountered as a child entry and when the user navigates directly into that directory or any descendant under the compare roots.
- Ignore pattern parsing and wildcard matching MUST stay bounded: the engine caps the number and length of parsed patterns, and wildcard matching uses a per-match work budget. Inputs that exceed those bounds are ignored or treated as non-matches rather than stalling scan work.

### Persistence

All options persist under:

- `Common::Settings::Settings::compareDirectories` (`Common::Settings::CompareDirectoriesSettings`)

Settings are loaded when the compare window opens and saved on Options → OK.

### Panel state model

While the Options panel is visible, `Common::Settings::CompareDirectoriesSettings`
draft state is the single source of truth for the panel body. Live DxUi toggles
and edits write directly to that draft. DxUi state is synchronized from the draft
after load, reload, interlock changes, and programmatic debug/test updates.

The Options body MUST NOT create hidden native statics, checkboxes, or edit
controls as a parallel state model. The native host HWND is only the scroll/body
container that lets DxUi render and receive input. Footer command plumbing may
use normal dialog commands, but option values must never be persisted by reading
hidden native body controls.

`Options → OK` persists the current draft to `compareDirectories` immediately
before restarting the compare run. `Cancel` discards the draft and leaves both
compare mode and persisted settings unchanged.

### External settings reload while Options is open

The Compare Directories **Options** panel participates in the shared settings-editor registry used by live settings reload.

Behavior:
- Clean options panel: the draft reloads from the newly applied `compareDirectories` defaults immediately, then visible DxUi controls synchronize from the draft.
- Dirty options panel: prompt `Reload from disk` / `Keep editing`.
- `Keep editing` marks the panel stale.
- A stale panel must prompt on the next `OK` with `Overwrite disk` / `Reload from disk` / `Cancel`.
- `Options → OK` persists the updated `compareDirectories` defaults to the main settings file immediately before restarting the compare run.

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
| `SubdirPending` | A descendant comparison is pending (descendant folder scans and/or content compares still running) |

### Name matching

- Names are compared using Windows ordinal, case-insensitive semantics (`CompareStringOrdinal(..., TRUE)`).
- The decision map uses an ordered `std::map` with a `CompareStringOrdinal`-based comparator (`WStringViewNoCaseLess`). This avoids the hash/equality contract violation that would arise from using `std::unordered_map` with a different hash function and ordinal equality.
- Trailing spaces/dots are ignored for comparison to reduce false mismatches across enumeration backends and Win32 path semantics.
- If multiple entries on the same side normalize to the same trailing-space/dot-trimmed name, the engine MUST NOT drop either entry. Singleton normalized entries may still use the normalized key for cross-side matching, but same-side normalized-collision groups MUST remain represented by their original enumerated names so the UI and sync manifest can address the real item.

### Default filter (Show Identical Items)

By default, the panes show only different items:

- "Only in left/right" items
- Type mismatches
- Items different by any enabled compare criterion
- Directories that differ (when `compareSubdirectories` is enabled)
- Items that are still in-flight:
  - Directories that are still computing (`SubdirPending`), so subtree progress is visible.
  - File-level `ContentPending` placeholders MAY be elided when `keepIdenticalItems` is off to keep memory bounded on very large folders. In that mode, the host MUST still surface progress via folder-level UI (scan/content progress indicators) and MUST set `anyPending` until work completes.

The menu item **Compare → Show Identical Items** (`showIdenticalItems`) is a view filter toggle for all folders in the compare session. When enabled, identical items are shown alongside different items. The setting is persisted in `CompareDirectoriesSettings`.

Availability and scan impact:

- When `keepIdenticalItems` is off, the host MAY prune identical entries from cached decisions to keep memory bounded. In this mode, the **Compare → Show Identical Items** menu item MUST be disabled/greyed out.
- When `keepIdenticalItems` is on, cached decisions retain identical entries so `showIdenticalItems` can be toggled without invalidating cached decisions or restarting the compare scan.
- Changing `keepIdenticalItems` changes the cached decision shape; the host MUST invalidate cached decisions and restart the compare scan.

When `showIdenticalItems` is off and a folder has no differences after a successful compare run completes, panes MUST show a localized **No differences** empty‑state message instead of the standard **Empty folder** UI.

### Subdirectory traversal

When `compareSubdirectories` is enabled:

- Subtree comparison is performed by **background scan workers** and is **progressive**:
  - A directory may initially show `SubdirPending` ("Computing...") until descendant folders are scanned and any in-flight content compares complete.
  - The UI MUST tolerate intermediate states (pending bits) and refresh via coalesced, budgeted updates (see [Transition behavior for content-pending items](#transition-behavior-for-content-pending-items)).
- Subtree traversal uses an iterative worklist (no recursive call-stack growth, safe for deeply nested trees like `node_modules`).
- Directory reparse points (symlinks/junctions) are **not** followed.
- The UI thread MUST NOT perform subtree traversal or synchronous folder enumeration as part of rendering (details/selection/empty-state). It may only read cached decisions and apply budgeted pending updates.

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
- Invert Differences Selection is a one-shot command, not a persistent mode. After the command applies to the current panes, future pane enumeration, navigation, and refresh callbacks must use the default decision-model selection again unless the user invokes Invert again.
- Compare-window pane refresh commands must restore the default decision-model selection before forcing the pane refresh so the `FolderView` refresh-preservation path cannot carry an inverted selection forward.

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

Compare uses the same `CompactDetails` normalized-metadata profile defined by
`Specs/UI/UI_FolderView.md`: Windows-local fixed-minute timestamps and exact `RHSACETOP` compact
attributes. Compare retains its own omission policy: empty timestamp tokens are omitted, directories
omit the size token, and remaining tokens are joined without leading or duplicate separators.

Items should have a clear visual distinction (foreground color or icon) based on their difference status, consistent with the current theme.

### Transition behavior for content-pending items

When a content compare completes for an item:

- If a per-item `ContentPending` placeholder exists, the `ContentPending` bit is cleared.
- If content differs, the final `Content` bit is set and the item becomes different/selected.
- If content is equal, the item remains identical and unselected (no content bits remain set).
- If the item was not surfaced as a per-item placeholder (differences-only mode elision), the item appears only if it becomes different; otherwise it remains absent from the filtered list.
- The pane refreshes via `WndMsg::kCompareDirectoriesDecisionUpdated`. The UI MUST coalesce refresh work (apply pending updates + refresh panes) to avoid UI-thread stalls during large content-compare bursts.
  - When there is a large backlog, the UI SHOULD apply pending updates in bounded batches (e.g., N folders per tick) and schedule additional refresh ticks until drained.
  - If the coalescing timer cannot be armed, the UI MUST log the Win32 failure and post a one-shot same-window fallback refresh message. A pending decision refresh must never be left waiting solely on a failed timer.
  - The UI MUST treat decision computation as background-only work: no UI-thread code path should trigger subtree traversal or synchronous `ReadDirectoryInfo()` I/O for compare decisions.
- If the item is now identical and `showIdenticalItems` is off, the item is removed from the list on the next enumeration refresh.
- No explicit animation or fade-out is applied — the item simply disappears from the filtered list or updates its details text.
  - Ancestor directories update `SubdirPending` / `SubdirContent` so directory status transitions correctly (e.g., "Computing..." → final).

## Content Compare Architecture

Content comparison uses a **two-phase** model to keep the UI responsive:

### Phase 1: Metadata scan (synchronous, per-folder)

When `GetOrComputeDecision` is called for a folder:

1. Enumerate both left and right sides by calling `IFileSystem::ReadDirectoryInfo` on the base filesystems.
   - Compare background scans MUST NOT enumerate via `DirectoryInfoCache::BorrowDirectoryInfo(..., AllowEnumerate)`: `DirectoryInfoCache` is a global cache sized for interactive browsing and can grow to GiB scale when a subtree scan touches many folders.
2. Match entries by name (case-insensitive ordinal).
3. Compare metadata (existence, type, size, date/time, attributes) based on enabled settings.
4. For files that need content comparison (same size or size-unknown, and `compareContent` is enabled):
   - Enqueue a content-compare job.
   - When the item is surfaced in the list, set the `ContentPending` diff bit. `ContentPending` does not imply a final difference and must not select the item.
   - When `keepIdenticalItems` is off, the host MAY elide per-file `ContentPending` placeholders to keep memory bounded, but it MUST keep folder-level pending state (`anyPending`) accurate until all content jobs complete.
5. Return the decision immediately — the UI shows progress via per-item "Comparing..." (when placeholders are surfaced) and/or folder-level progress indicators.

### Phase 2: Background content compare (asynchronous, worker pool)

- A pool of `std::jthread` workers processes the content-compare queue:
  - If `contentCompareWorkerCount == 0` (Auto): use `std::thread::hardware_concurrency()` (fallback `2` when it returns `0`), clamped to `1..4`.
  - Otherwise: use `contentCompareWorkerCount` clamped to `1..4`.
- Workers are created lazily on first content-compare enqueue (`EnsureContentCompareWorkersLocked`).
- Each worker dequeues a `ContentCompareJob`, checks the version (bail out if stale), opens both files via `IFileSystemIO::CreateFileReader`, and streams/compares until a difference or EOF.
- Results are applied to cached decisions and must update ancestor directories' subtree state (`SubdirPending` / `SubdirContent`) so directory status transitions correctly without requiring navigation.
- Implementations may apply completed results directly from workers, or queue them and apply them on the UI thread via `FlushPendingContentCompareUpdatesBudgeted()`.
  - The UI SHOULD batch/coalesce decision-update refreshes (e.g., timer-based) instead of doing a full refresh on every posted notification.
  - Budgeted application may require multiple UI ticks when many updates are outstanding.
- On completion, the engine posts `WndMsg::kCompareDirectoriesDecisionUpdated` so the panes refresh.

### Cancellation

- Workers check `std::stop_token` and the session `_version` inside the read loop.
- When the user clicks Rescan, changes settings, or navigates outside the roots: `_version` increments, the content-compare queue is cleared, and workers bail out of in-progress jobs on the next cancellation check.

### Notification throttling

Progress callbacks (`ScanProgressCallback`, `ContentProgressCallback`, `DecisionUpdatedCallback`) are throttled by a minimum interval (tracked via `_*LastNotifyTickMs` atomics) to avoid flooding the UI thread with `PostMessage` calls. A final forced notification is always sent when a scan or the last content-compare completes, ensuring the UI reaches the final state.

### Performance and memory bounds (normative)

- The engine MUST keep background compare memory bounded (bounded work queues and bounded caches).
- Under sustained backlog (“backpressure”), the engine MUST prioritize work for **currently visible folders** over low-priority background subtree scanning.
- UI thread work MUST remain bounded and coalesced (see [Transition behavior for content-pending items](#transition-behavior-for-content-pending-items)).
- The decision cache is budgeted. Eviction MUST preserve the root decision and pinned visible folders/ancestors, but MUST NOT treat `anyPending` as an absolute eviction veto.
- Pending child directory decisions may be evicted only after the engine has captured the child aggregate needed for ancestor repair (`version`, `HRESULT`, `anyPending`, `anyDifferent`). Budgeted pending-subdir flushes MUST be able to update ancestor `SubdirPending` / `SubdirContent` from that aggregate even if the full child decision was evicted.
- Active scan worker keys are temporarily protected from decision-cache eviction until their pending-subdir aggregate is recorded. Once recorded, non-pinned pending decisions are recomputable and may be evicted to keep both live and high-water decision-cache bytes within a bounded multiple of the configured budget.
- Pending content-update and pending subdir-aggregate repair retries MUST be bounded per folder/key. If the decision remains unavailable after the retry cap, the pending repair entry and retry counter are discarded so an inaccessible or non-cacheable folder cannot self-sustain rescan/timer work indefinitely.
- Eviction work itself MUST remain bounded; a budget pass should scan the LRU tail once, collect a bounded number of evictable candidates, and erase them after candidate collection instead of restarting from the tail for each eviction.

## Cache Coherence & Version Model

The engine uses an atomic `uint64_t _version` counter as the primary cache coherence mechanism:

- **Version increments** on: `SetRoots`, `SetSettings`, and full `Invalidate`.
- **Cached decisions** (`_cache`) are tagged with the version at computation time. A cached entry is valid only if its version matches the current `_version`; stale entries are recomputed on next access.
- **In-flight scan jobs** carry `(version, cancelToken)` stamps. A scan may write back decisions, enqueue child work, or queue pending-subdir propagation only while its live in-flight stamp still exists.
- **In-flight content-compare jobs** carry the version at enqueue time. Workers discard results if the version changed during comparison.
- **In-flight content-compare state** (`_contentCompareInFlight`, content queues, pending content updates) is cleared on version-changing resets.
- **Completed content-compare cache entries** (`_contentCompareCache`) are metadata-keyed by relative path, size, and last-write time. `SetSettings` invalidates cached folder decisions but preserves completed content results so a settings-only recompute of unchanged files does not re-open readers. `SetRoots`, full `Invalidate`, background-work cancellation, and targeted `InvalidateForAbsolutePath(...)` still clear or evict affected content-cache entries.

### UI version (`_uiVersion`)

A separate `_uiVersion` tracks the last version the UI has observed. This prevents redundant pane refreshes: the window only triggers a re-enumeration when `_uiVersion` differs from `_version`.

### Cache invalidation after file operations

`InvalidateForAbsolutePath(path, includeSubtree)`:

1. Converts the absolute path to a relative path under the left or right root.
2. Removes matching decision-cache entries and pending-content/subdir updates under lock. With `includeSubtree=false`, only the leaf decision key is removed; ancestors stay cached so unrelated in-flight ancestor scans are not cancelled by a single-file notification.
3. Prunes affected scheduled, duplicate, and in-flight scan/content work using the same exact-or-subtree key semantics.
4. Leaves `_version` stable, bumps `_uiVersion`, and relies on the live in-flight scan stamp check above to prevent stale pre-invalidation workers from writing back.

This is currently O(n) over cache entries for subtree invalidation. For large trees, consider using the ordered `std::map` with prefix-range lookup (`lower_bound`) for O(log n + k) invalidation.

## IFileSystem Wrapper (`CompareDirectoriesFileSystem`)

The compare engine exposes a per-pane filesystem wrapper created via:

```cpp
wil::com_ptr<IFileSystem> CreateCompareDirectoriesFileSystem(ComparePane pane, std::shared_ptr<CompareDirectoriesSession> session);
```

The wrapper implements `IFileSystem` and `IInformations`. Behavior:

### `ReadDirectoryInfo` (enumeration)

- If compare mode is **active** and the requested folder is **within** the pane's root:
  - Call `session->TryGetCachedDecision(relativeFolder)` to get the cached folder decision (cache-only; no synchronous I/O).
  - If no cached decision exists yet, the wrapper MUST call `session->RequestScanForFolder(relativeFolder)` and return an empty enumeration.
  - For each item in the decision:
    - Skip items that don't exist on this pane's side (`existsLeft`/`existsRight`).
    - Skip identical items if `showIdenticalItems` is off and `isDifferent` is false.
    - Populate the directory entry from the decision's metadata (size, date/time, attributes).
    - Set the per-item details text from the decision's difference mask.
  - Apply the decision model's selection rules to the returned entries.
- If compare mode is **inactive**, or the folder is **outside** the roots:
  - Delegate to the wrapper's base filesystem for that pane.

### Missing-side folders

When one side doesn't exist for a relative path:

- `ReadDirectoryInfo` returns an empty entry list (success, zero entries).
- The `FolderView` renders the localized empty-state message (`IDS_COMPARE_FOLDER_NOT_FOUND`).

### File operations (copy/move/delete/rename)

All mutation operations delegate to the wrapper's base filesystem for that pane. The compare wrapper does not intercept or filter these — filtering for sync-copy is handled at a higher level (see [Directory sync-copy](#directory-copymove-behavior-sync)).

## Scanning & Progress Notifications

Enumeration in compare mode is driven by the compare-scoped filesystem wrapper (see above).

Progress notifications use posted payload messages declared in `Common/WindowMessages.h`:

- `WndMsg::kCompareDirectoriesScanProgress` — scan phase updates (folder/entry counts, current path).
- `WndMsg::kCompareDirectoriesContentProgress` — content compare updates (per-file bytes, pending count).
- `WndMsg::kCompareDirectoriesDecisionUpdated` — causes the panes to refresh so "Comparing..." transitions to final results.

Notes:

- Posted payload receivers must use `TakeMessagePayload<T>(lParam)` and windows must drain pending payloads in `WM_NCDESTROY`.
- The compare window tags progress updates with a **run id** (the `_version` at scan start) and ignores stale updates whose run id does not match the current version.
- Scan and content progress posts also carry that run id in `wParam`. The UI may reduce latest-wins progress
  payloads only while the unfiltered UI-thread queue head has the same window, message, and run id. A completion,
  input message, or payload for another run is an ordering boundary and remains queued for the normal message
  pump. Coalescing uses the shared `TakeAndCoalesceContiguousPostedPayloads` helper so every removed payload is
  unregistered through `TakeMessagePayload` and teardown draining remains authoritative.

## File Operations Semantics (Sync)

All standard operations work with the same shortcuts and command set as the main `FolderWindow`:

- View / View With
- Rename
- Copy / Move to other pane
- Delete / Permanent delete
- Create directory

### Directory copy/move behavior (sync)

When compare mode is active, Copy/Move to other pane acts as a **sync** operation over the selected/focused compare items:

- The compare window gathers selected plugin paths, converts them to compare-relative paths, and asks the session for a cache-only `CompareSyncManifest`.
- The session-owned manifest planner is the only component that interprets compare decisions. It respects current compare settings, ignore patterns, side selection, `compareSubdirectories`, and the current session version.
- The planner fails closed: missing/stale decisions, failed decisions, `ContentPending`, `SubdirPending`, and incomplete pending-content counts return a blocker instead of a partial manifest. The UI requests high-priority scan work for the blocker path/parent, shows a localized "comparison still in progress" status, and does not fall back to generic recursive copy/move.
- Mixed directories that exist on both sides are expanded into a `DirectoryShell` item plus explicit changed/source-selected children. Identical descendants are not represented, copied, moved, or deleted.
- A source-only directory is the only recursive exception. It is represented as one `WholeSubtree` item with `FILESYSTEM_FLAG_RECURSIVE` because the whole tree is new on the destination side.
- Manifest items carry exact `sourceAbsolutePath`, `destinationAbsolutePath`, relative path, kind, and per-item flags. The file-operation layer executes resolved items against those exact destinations instead of recomputing `destinationFolder / sourceLeaf`.
- Destructive Move uses sync-specific confirmation text with the manifest item count and one-sided whole-subtree count. Move deletes only manifest items, and only after their destination operation succeeds.
- After completion, exact resolved destination paths are invalidated alongside source paths so the compare view converges without a full rescan.

### Cache invalidation after operations

When compare mode is active, after file operations complete the compare session calls `InvalidateForAbsolutePath` for affected source/destination paths so the UI updates quickly without forcing a full rescan.

## Content Compare I/O Contract

Content comparison uses per-side filesystem plugin I/O:

- If either side does not support `IFileSystemIO`, content comparison MUST be treated as unavailable for that compare scope (see the `compareContent` option notes above).
- Use `leftIo->CreateFileReader(leftPath)` to open the left file.
- Use `rightIo->CreateFileReader(rightPath)` to open the right file.
- Read with `IFileReader::Read` (streaming, 256 KB buffer) until a difference is found or EOF is reached.
- Short reads are permitted by the `IFileReader` contract; content compare must tolerate them (continue reading until EOF or a mismatch).
- If `IFileReader::GetSize` succeeds for both sides and sizes differ, return "different" immediately without reading content.
- If `IFileReader::GetSize` succeeds for both sides and both sizes are 0, return "equal" immediately.
- If sizes are unknown (plugin doesn't report sizes), stream until EOF or a mismatch.
- Reader creation, `GetSize` (when the consumer requires a size), seek, or `Read` failure is an I/O failure, not evidence that contents differ. The content decision MUST carry the HRESULT/failure reason distinctly from `Different`; callers may show the per-item error and retry/invalidate, but MUST NOT convert an unreadable file into a content mismatch.

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
- Content compare: `contentCompareWorkerCount`
- Subdirectories: `compareSubdirectories`, `compareSubdirectoryAttributes`, `selectSubdirsOnlyInOnePane`
- Ignore patterns: `ignoreFiles` + `ignoreFilesPatterns`, `ignoreDirectories` + `ignoreDirectoriesPatterns`
- Display: `keepIdenticalItems`, `showIdenticalItems`

The Compare Directories Options panel reads/writes the same settings; Preferences provides a centralized way to configure these defaults outside the compare window.

## Testing

- Engine self-tests live under `RedSalamander/SelfTest/CompareDirectories/` and run via `--compare-selftest` (or `--selftest`):
  - `CompareDirectoriesEngine.SelfTest.cpp` + `CompareDirectoriesEngine.SelfTest.h` (harness/registration)
  - `CompareDirectoriesEngine.SelfTest.Cases.CoreDiffs.cpp` (core difference cases)
  - `CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp` (search/index cases)
  - `CompareDirectoriesEngine.SelfTest.Cases.RuntimeAndRemote.cpp` (runtime and remote/plugin filesystem cases)
- Required engine coverage includes empty-directory roots/folders, bounded wildcard matching for pathological ignore patterns, targeted invalidation that preserves unaffected sibling caches, and stale in-flight scan writeback suppression after scoped invalidation.
- Manual UI checks:
  - Shortcuts (function keys, selection, copy/move/delete) behave like the main window.
  - Compare scope behavior: OK/Rescan establishes roots; navigation outside roots cancels compare mode (with a prompt if scan/content compare is active) until the next OK/Rescan.
  - Options panel theming, scroll behavior, and forward/reverse keyboard traversal at small window sizes.
  - Options panel tests that exercise in-process DxUi UIA ValuePattern reads or writes must pump the UI thread while the UIA worker performs provider calls, because retained DxUi controls are now the live value providers.
  - Progress appears in the compare banner and in the File Operations popup task card; it disappears when idle.
  - All three display modes (Alt+2, Alt+3, Alt+4) render correctly.
  - Directory sync Copy/Move only operates on manifest items: different files transfer, identical descendants are skipped, source-only directories are whole-subtree items, and Move preserves unlisted identical source files.
  - Content-pending items show "Comparing..." and transition to final state without getting stuck.
  - Missing-side folders show the empty-state message.
  - Enumeration failures (access denied) show the failure reason in the details line.
