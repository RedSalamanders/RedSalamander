# File Operations Specification

Last updated: 2026-02-19

Normative sections use RFC-2119 keywords (MUST/SHOULD/MAY). Appendices are informative.

## Overview

This document defines the **execution model**, **progress UI contract**, and **pre-calculation phase** for long-running file operations initiated from RedSalamander panes:

- Copy to other pane (`F5`)
- Move to other pane (`F6`)
- Delete (`Del` / `F8`, context menu)

It also defines the **speed limit** behavior (host-provided bandwidth cap) and the `FileSystemDummy` **virtual speed limit** interaction.

## Goals

- File operations MUST not block the UI thread.
- Users MUST see operation progress (items + throughput) and be able to **Pause** and **Cancel**.
- Users MUST be able to run operations in **Wait** (sequential) or **Parallel** mode.
- Copy/Move operations MUST support a per-task **Speed Limit**.
- The popup MUST follow the active `AppTheme` (light/dark/high contrast) and support `menu.rainbowMode`.

## Terminology

- **Operation**: Copy / Move / Delete (executed by an `IFileSystem` plugin).
- **Task**: One user-requested operation with a stable ID and mutable options (pause/cancel/speed limit).
- **Informational Task**: A host-created, read-only progress card shown in the File Operations popup for long-running background work that is not a file operation (e.g., Compare Directories scan/content-compare).
- **Pre-calculation (pre-calc)**: A scan phase that computes totals (bytes + file/folder counts) before the operation begins.
- **Wait mode**: Sequential mode. Only one task may execute at a time. UI label: `IDS_FILEOPS_BTN_MODE_QUEUE` (`"Wait"`).
- **Parallel mode**: Concurrent mode. Multiple tasks may execute at once. UI label: `IDS_FILEOPS_BTN_MODE_PARALLEL` (`"Parallel"`).
- **Queue pause**: A host-driven pause applied to tasks that already started, used when switching to Wait mode while multiple tasks are active.

## Architecture

### Key types and files

- Core state + worker thread:
  - `RedSalamander/FolderWindow.FileOperationsInternal.h` (`FolderWindow::FileOperationState`, `Task`)
  - `RedSalamander/FolderWindow.FileOperations.State.cpp` (queueing + pre-calc + operation execution)
- Progress popup (Direct2D/DirectWrite):
  - `RedSalamander/FolderWindow.FileOperations.Popup.h`
  - `RedSalamander/FolderWindow.FileOperations.Popup.cpp`
- Plugin contracts:
  - `Common/PlugInterfaces/FileSystem.h` (`IFileSystem`, `IFileSystemCallback`, `IFileSystemDirectoryOperations`, `IFileSystemIO`, `IFileReader`, `IFileWriter`)
- Dummy plugin:
  - `Plugins/FileSystemDummy/FileSystemDummy.cpp`

### Related specs

- Plugin operation and callback contracts: `Specs/PluginsVirtualFileSystem.md`
- Theme key list (including file ops keys): `Specs/SettingsStoreSpec.md`

## Execution Model (Normative)

### Threading

- The host MUST execute each Task (including pre-calc and `IFileSystem::*`) on a background worker thread (one per task).
- When a Task uses per-item execution with per-item concurrency (`maxConcurrency > 1`), the host SHOULD schedule per-item work using a **shared worker pool** across all active Tasks (especially in Parallel mode) so the total worker thread count stays bounded and workers can be reassigned between Tasks after each item.
- When a file-system plugin uses internal parallelism (e.g. plugin max concurrency knobs for Copy/Move/Delete), the plugin SHOULD schedule that work using a **shared bounded worker pool** across all in-flight operations (rather than spawning per-operation thread pools) so worker threads can be reused/reassigned between operations as items finish.
- The host MUST drive progress via `IFileSystemCallback` and forward updates to the UI thread.
- The host MAY surface background work as Informational Tasks. Informational Tasks:
  - MUST NOT participate in Wait/Parallel queueing rules.
  - MUST NOT block or pause file-operation Tasks.
  - MUST be read-only in the UI (no conflict prompts; no overwrite/replace-readonly decisions).

### Preconditions (cross-pane Copy/Move)

- The host MUST reject cross-pane **Copy**/**Move** when both panes point to the **same effective destination folder** (same normalized path text), because this is almost always user error (accidental self-copy / self-move).
  - The host MUST show a localized error (see `IDS_MSG_PANE_OP_REQUIRES_DIFFERENT_FOLDER`).

- The host MUST reject **Copy**/**Move** when any source item would be copied/moved onto itself or into its own subtree (destination folder overlaps the source item), because this can recurse indefinitely or produce confusing no-op operations.
  - The host MUST show a localized error (see `IDS_FMT_FILEOPS_INVALID_DESTINATION_OVERLAP`).

- **Same-context copy/move**:
  - If both panes are operating on the same effective file system context (same filesystem plugin id + same per-instance mount context when the plugin uses `IFileSystemInitialize`), the host SHOULD execute Copy/Move using that plugin instance directly.

- **Cross-context (cross-filesystem) copy/move**:
  - If the effective contexts differ, the host MUST NOT silently fall back to passing “foreign” paths to an arbitrary `IFileSystem` instance.
  - Instead, the host MUST either:
    1) Execute the operation via a host-driven **cross-filesystem bridge** (see “Cross-filesystem bridge”), or
    2) Reject the operation and show a localized error (see `IDS_MSG_PANE_OP_REQUIRES_COMPATIBLE_FS`).

### Pause / Cancel

- **Cancel**:
  - Host sets a cancel flag; `IFileSystemCallback::FileSystemShouldCancel` returns `TRUE`.
  - Plugins MUST abort promptly and return `HRESULT_FROM_WIN32(ERROR_CANCELLED)` (or `E_ABORT`, normalized by the host).
- **Pause**:
  - Host pauses by blocking inside progress callbacks (`IFileSystemCallback::FileSystemProgress` and `IFileSystemDirectorySizeCallback::DirectorySizeProgress`) at progress checkpoints.
  - Host MAY also pause “non-primary” tasks via **queue pause** when Wait mode is enabled while multiple tasks are already active.

### Conflict Handling (Normative)

Conflict handling covers per-item failures that require a user decision (overwrite, replace read-only, permissions, etc.).

- The host MUST provide conflict handling for all in-app entry points that can trigger Copy/Move/Delete (keyboard shortcuts, menus/context menus, pane → pane drag/drop).
- The host MUST NOT silently auto-resolve conflicts by default (no implicit overwrite, replace-readonly, or continue-on-error without user intent).

#### Defaults

- Copy/Move MUST start without allowing overwrite and without allowing replace-readonly (conflicts must surface).
- Delete SHOULD start by using Recycle Bin when supported.
- Continue-on-error MUST be user-driven (via per-conflict Skip/Skip All decisions), not a default behavior.

#### Conflict detection + bucketing

- When an item operation fails, the host SHOULD map the failure (`HRESULT`, usually `HRESULT_FROM_WIN32(...)`) into a stable *bucket* so the UI can show bucket-specific messaging and support apply-to-all caching.
- The bucket set SHOULD include:
  - already exists
  - read-only
  - access denied
  - sharing violation
  - disk full
  - path too long
  - recycle bin failed (delete)
  - network / offline
  - unknown

#### Prompting + task concurrency

- Conflicts MUST be resolved by an inline prompt in the file operations popup (no modal dialogs).
- When a task encounters a conflict, the host MUST block that task’s forward progress until a decision is made.
- If a task is executing multiple items concurrently (plugin or host internal parallelism), the host MUST serialize prompts (at most one active prompt per task) and MUST ensure all in-flight workers for that task converge to a stopped/paused state at progress checkpoints while waiting for the decision.
- For recursive directory operations, conflicts SHOULD be raised at the most-specific failing path (file/subdir), not by aborting the entire top-level directory item.
  - Plugins SHOULD invoke `IFileSystemCallback::FileSystemIssue(...)` and continue traversal based on the returned action.
  - If any sub-items are skipped (or partially fail) and the operation continues, the plugin SHOULD return `HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY)` for the top-level item.
- The host MUST treat `HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY)` returned by a plugin item call as a terminal "partial success" for that item (no additional conflict prompting for that item), while still surfacing the partial status in the task result/diagnostics.

#### Actions

The prompt MUST offer a subset of these actions as appropriate for the operation + bucket:

- **Overwrite** (Copy/Move): retry the item allowing overwrite.
- **Replace read-only** (Copy/Move): retry the item allowing replace-readonly.
- **Permanent delete** (Delete): when a Recycle Bin delete fails, retry the item without the Recycle Bin.
- **Retry**: retry the item (primarily for transient buckets like sharing/network/unknown).
- **Skip**: skip the current item and continue.
- **Skip All**: skip all future items that fail with the same bucket in this task.
- **Cancel**: cancel the entire task.

Retry and apply-to-all rules:

- Retry MUST be capped to at most one retry per (item, bucket). If the retry fails with the same bucket again, Retry MUST NOT be offered again for that (item, bucket) and the UI MUST indicate that the retry failed.
- “Apply to all similar” MUST only apply to non-retry actions and MUST be cached per-task per-bucket. Retry MUST NOT be cached.

### Wait / Parallel mode

The host supports two mechanisms that together implement Wait/Parallel behavior:

1. **Start gating** (per-task wait-for-others):
   - If a task is configured to wait, it MUST block at task start (before pre-calc and before invoking `IFileSystem::*`) until it becomes the active queued task.
   - The host maintains a FIFO queue of waiting tasks.
   - Pre-calc runs while holding the queue slot (so queued tasks do not interleave pre-calc + operation).

2. **Queue pause** (for already-active tasks):
   - When global Wait mode is enabled while multiple tasks are already active, only the oldest active task SHOULD continue; other active tasks SHOULD pause at the same progress checkpoints until the previous task completes or mode is switched back to Parallel.

**Default behavior**: when starting a new task while another task is active, the new task SHOULD start in Wait mode (queued execution).

### Mode switching

- **Wait → Parallel**:
  - Tasks blocked at start MUST be unblocked and proceed.
  - Tasks paused only due to Wait mode SHOULD resume.
- **Parallel → Wait**:
  - The oldest active task SHOULD continue running.
  - All other active tasks (including tasks still in pre-calc) SHOULD become queue-paused at the next progress checkpoint (serialized execution).
  - Tasks not yet started (and newly created tasks) MUST queue at start and wait until no active tasks remain.

### Cross-filesystem bridge (Normative)

When Copy/Move must execute across different filesystem contexts, the host MAY perform the operation itself by bridging between plugins:

- The host MUST treat the source and destination file system instances as separate:
  - Source: directory enumeration + attribute queries + file reading (`ReadDirectoryInfo`, `IFileSystemIO::GetAttributes`, `IFileSystemIO::CreateFileReader`).
  - Destination: directory creation + file writing (`IFileSystemDirectoryOperations::CreateDirectory`, `IFileSystemIO::CreateFileWriter`).
- The host SHOULD prefer in-memory streaming using `IFileReader` → `IFileWriter` (no temp files).
- If streaming is not possible for a given plugin pair, the host MAY fall back to a temp-folder materialization strategy (implementation-defined), but MUST preserve cancel/pause responsiveness and MUST not block the UI thread.

Move semantics under the bridge:

- Cross-filesystem **Move** SHOULD be implemented as Copy + Delete (delete after successful copy).
- On partial failure, the host SHOULD follow the existing conflict/continue-on-error rules for the task.

### Drag & Drop (Normative)

- **In-app drag/drop (pane → pane)** MUST follow the same execution model as `F5`/`F6`:
  - It MUST enqueue a Task and show progress in the file operations popup.
  - It MUST respect same-folder rejection and cross-filesystem bridge rules above.

- **External drag/drop (to Explorer or other apps)** SHOULD continue to use standard shell formats (e.g. `CF_HDROP`) so that drops “just work” for filesystem-backed items.

## Pre-Calculation Phase

### Purpose

Pre-calculation scans the operation’s source paths to compute:

- Total bytes (Copy/Move/Delete)
- Total file count
- Total directory count

This enables accurate progress totals and improves ETA accuracy.

### When it runs

- Pre-calc runs on the same task worker thread, before invoking `IFileSystem::*`.
- In Wait mode, pre-calc MUST execute while holding the queue slot (sequential with respect to other queued tasks).

### Interface contract

The host queries `IFileSystemDirectoryOperations` from the active `IFileSystem` instance:

- If `QueryInterface(__uuidof(IFileSystemDirectoryOperations))` fails or returns null, pre-calc MUST be skipped and the task proceeds without totals.
- Otherwise, the host calls `IFileSystemDirectoryOperations::GetDirectorySize(...)` for each source path (typically with `FILESYSTEM_FLAG_RECURSIVE`).

Progress uses `IFileSystemDirectorySizeCallback`:

- This callback is **NOT COM** (no `IUnknown`); its lifetime is host-owned.
- `cookie` (if used) MUST be passed back verbatim by the plugin.
- Skip/cancel:
  - `DirectorySizeShouldCancel` MUST return `TRUE` when the task is cancelled or pre-calc is skipped, so plugins can abort promptly.

### Skip button behavior (UI → model)

- Pressing **Skip** sets a pre-calc skipped flag.
- Pre-calc stops and the operation begins without updating totals from pre-calc.
  - If the task is paused (including **queue pause**), the host MUST stop pre-calc promptly but MUST NOT start the `IFileSystem::*` operation until the pause is lifted.
- In Wait mode, the task retains its queue slot and continues in order (it is NOT re-queued).

### Cancel button behavior

- Pressing **Cancel** cancels the task.
- If cancellation occurs during pre-calc, the task MUST release the queue slot and complete as cancelled.

### Delete operations

Delete operations also run pre-calc:

- Size totals are useful for displaying “completed / total bytes” (when available).
- Throughput is displayed as items/sec (bytes/sec may remain 0 depending on plugin).

## Progress Popup UI (Normative)

### Window contract

- A single modeless top-level window associated with `FolderWindow` (one per FolderWindow instance).
- Standard Win32 chrome (icon + caption + minimize/close), resizable, appears in the taskbar.
- Closing the popup window (system close button / `WM_CLOSE`) MUST behave like **Cancel all**:
  - Show the same localized confirmation dialog.
  - If confirmed: cancel all tasks and close the popup.
  - If declined: keep the popup open and do not cancel tasks.
- The window MUST be independent and minimize/restore independently of the main application window; restoring it while the main window is minimized MUST NOT change the main window.
- The client area is rendered using Direct2D/DirectWrite.
- The UI updates on a timer (~100ms) by reading the latest progress snapshot collected via callbacks.
- The operations list is vertically scrollable using the standard window scrollbar (`WS_VSCROLL`) and MUST be themed consistently with `FolderView`.
- **Auto-sizing**:
  - The window height MUST automatically adjust to fit the content when tasks are added or removed.
  - The window height MUST automatically adjust when task cards are collapsed or expanded.
  - The maximum height is limited to the current monitor's work area (excluding taskbar).
  - When content exceeds screen height, the scrollbar becomes visible.
  - Auto-resize is disabled during user resize operations.
- Positioning:
  - Reuse the last window rectangle if it is still fully visible on a monitor work area.
  - Otherwise, center the popup over the current main application window (clamped to the monitor work area).

### Layout

- The popup has three regions:
  - **Title bar** (always visible)
  - **Operations list** (scrollable): task cards stacked vertically (oldest at top)
  - **Global footer** (always visible)

Global footer MUST provide:

- **Cancel all** (requires confirmation)
- **Wait / Parallel** toggle

### Operation cards

Each task card MUST support collapse/expand and MUST adapt to task state:

**Header**
- During pre-calc: `IDS_FILEOPS_CALCULATING` / `IDS_FMT_FILEOPS_CALCULATING_TIME`
- During pre-calc while waiting/queue-paused: `IDS_FILEOPS_GRAPH_WAITING`
- Otherwise: `IDS_FMT_FILEOPS_OP_COUNTS` (e.g. `Copy: 3/12`)
  - For Copy/Move, when pre-calc totals (files + folders) are available, the host SHOULD use those totals for the `X/Y` counts; if per-entry completion counts are not available, `X` MAY be estimated from byte progress for display.
- A per-task collapse/expand chevron

**Body (expanded)**
- During pre-calc: display the currently accumulated **item totals** (files + folders) and total bytes so far (as they are discovered).
- When pre-calc is skipped (totals unknown), Copy/Move cards SHOULD display a best-effort breakdown of completed **top-level** items by type (files vs folders) when that classification is available.
  - The counts MUST be monotonic.
  - When classification is complete for all top-level items, the host SHOULD keep the breakdown consistent with `completedItems` (`completedFiles + completedFolders == completedItems`).
  - Throughput:
  - Copy/Move: bytes/sec
  - Delete: items/sec
- ETA (Copy/Move):
  - Shown only when total bytes are known and current speed is > 0.
- Paths:
  - Copy/Move: `From:` and `To:` lines
    - Before a Copy/Move task starts, the UI MAY offer a destination selector menu (other panel + history) on the `To:` line.
  - Delete: `Deleting:` line
  - During pre-calc, the host SHOULD display the current directory being scanned (from `IFileSystemDirectorySizeCallback::DirectorySizeProgress.currentPath`) when available; otherwise it SHOULD display the first planned source path.
  - Parallel Copy/Move (multi-file in-flight):
    - When a single Task is executing Copy/Move with multiple files actively copying at once (plugin internal parallelism), the UI MUST display multiple `From:` file lines with per-file progress (instead of only one “current item”).
    - The UI SHOULD display one stable line per active progress stream (up to max concurrency lines).
      - A stream is identified by `(cookie, progressStreamId)` from `IFileSystemCallback::FileSystemProgress(...)`.
      - A line SHOULD be updated in-place as the stream advances to new items (so completed items are replaced by the next item rather than lingering at `100%`).
      - When a stream completes an item and becomes idle, its line SHOULD disappear promptly (so “done” lines don’t linger at `100%` for long-running tasks).
    - Entries MAY still show `100%` briefly (e.g., end-of-file), optionally with a short grace window (~300ms), before disappearing or advancing to the next item.
- Progress bars:
  - Pre-calc: indeterminate marquee bar
  - Copy/Move: current-item bar (primary) + overall bar (secondary)
  - Delete: overall item bar; MAY additionally show size progress when pre-calc data is available
- Bandwidth graph (Copy/Move):
  - Shows recent throughput history.
  - When speed limit is active, shows a horizontal line at the effective limit.
  - Y-axis MUST auto-scale with headroom so the graph remains readable.
  - Overlay text MUST have a drop shadow for visibility against colored graph backgrounds.
  - Overlays:
    - Pre-calc: `IDS_FILEOPS_GRAPH_CALCULATING` + animation
    - Paused: `IDS_FILEOPS_GRAPH_PAUSED` (graph frozen)
    - Waiting/queued: `IDS_FILEOPS_GRAPH_WAITING` (graph frozen)

**Controls**
- During pre-calc: **Skip** + **Cancel**
- During operation:
  - Copy/Move: **Pause/Resume**, **Speed Limit**, **Cancel**
  - Delete: **Pause/Resume**, **Cancel**

**Conflict prompts (inline)**
- When a task is blocked on a conflict decision, the popup MUST display an inline prompt associated with that task (not a separate modal dialog).
- The prompt MUST include:
  - bucket-specific message text (localized)
  - the relevant item path(s) (`From` and `To` for Copy/Move; `Deleting` for Delete)
  - an optional “Apply to all similar” toggle (only for non-retry actions), placed adjacent to the action buttons and clearly visible
  - buttons for the available actions (Overwrite / Replace read-only / Permanent delete / Retry / Skip / Skip All / Cancel)
- While a task is blocked on a prompt, its progress UI SHOULD appear paused/waiting (frozen counters/graph overlays) until the decision is applied.

### Path truncation

When paths do not fit, the UI MUST truncate using a **middle ellipsis** (`…`) so the most important portion remains visible:

- Source line: preserve the file/folder name at the end.
- Destination line: destination is a folder path; do not show a filename.
- When showing a per-file mini progress indicator on the right (parallel in-flight display), the UI MUST reserve space and clip text so the filename/path never renders underneath the progress indicator.

## Speed Limit (Normative)

### Semantics

- The per-task speed limit applies to Copy/Move via `FileSystemOptions::bandwidthLimitBytesPerSecond`.
- `0` means unlimited.
- Plugins MAY clamp the host-provided limit and report an effective applied limit by writing back to `FileSystemOptions::bandwidthLimitBytesPerSecond` before progress callbacks.
- Presets (bytes/sec):
  - 1 MiB/s, 5 MiB/s, 10 MiB/s, 50 MiB/s, 100 MiB/s, 1 GiB/s
- `Custom...` opens an input dialog.

### Throughput text parsing (host UI)

The host parses user-entered speed limits using a whitespace-tolerant, case-insensitive grammar:

- Number:
  - Accepts integers or decimals (e.g. `1.5`)
- Optional unit (defaults to KiB when absent):
  - `B`
  - `K`, `KB`, `KiB`
  - `M`, `MB`, `MiB`
  - `G`, `GB`, `GiB`
  - `T`, `TB`, `TiB`
  - `P`, `PB`, `PiB`
- Optional `"/s"` suffix is accepted.
- Rounding: the result is rounded to the nearest integer bytes/sec.
- Clamping: values larger than `std::numeric_limits<uint64_t>::max()` are clamped.
- Empty input or `0` is treated as unlimited.

### FileSystemDummy virtual speed

`FileSystemDummy` exposes a configuration setting `virtualSpeedLimit` (text). The dummy plugin’s **effective speed** is:

- `min(hostLimit, virtualSpeedLimit)` when both are non-zero
- otherwise whichever is non-zero

**Delete simulation (Dummy):**
- Delete uses a fixed **virtual bytes per item** budget with deterministic jitter to simulate non-instant deletes.
- The per-item delay is computed from the **effective speed** (`virtualBytesPerItem / effectiveBytesPerSecond`).
- If the effective speed is `0` (unlimited / `virtualSpeedLimit == 0` and no host limit), delete does **not** add any wait.

## Theming (Normative)

- All UI text MUST be in `.rc` resources.
- All progress + graph colors MUST be derived from the active theme (no hard-coded colors).
- Theme updates MUST re-skin the existing popup window on theme changes.

### Theme keys

The file operations popup supports these optional theme override keys:

- Progress:
  - `fileOps.progressBackground` (track)
  - `fileOps.progressTotal` (overall)
  - `fileOps.progressItem` (current-item; ignored in Rainbow mode)
- Bandwidth graph:
  - `fileOps.graphBackground`
  - `fileOps.graphGrid`
  - `fileOps.graphLimit`
  - `fileOps.graphLine`
- Scrollbar (if custom-colored):
  - `fileOps.scrollbarTrack`
  - `fileOps.scrollbarThumb`

### Rainbow mode

When `menu.rainbowMode` is enabled:

- The current-item progress color MUST change per file (stable per-item color derived from the current source path or leaf name).
- The bandwidth graph line and fill MUST use per-sample colors matching the progress bar:
  - Each data point in the graph history stores the hue derived from the source file path at the time of capture.
  - As the graph scrolls, older samples retain their original colors, creating a multicolored rainbow trail.
  - The graph does NOT use time-based color cycling (no blinking effect).
- The pre-calc animation MAY use time-based HSV color cycling (visual-only; no semantic meaning).

## String Resources

All user-facing strings referenced by the file operations UI MUST be localizable resources. Current IDs include:

- Operations + labels: `IDS_FILEOP_OPERATION_COPY`, `IDS_FILEOP_OPERATION_MOVE`, `IDS_FILEOP_OPERATION_DELETE`, `IDS_FILEOPS_LABEL_FROM`, `IDS_FILEOPS_LABEL_TO`, `IDS_FILEOPS_LABEL_DELETING`
- Buttons + menus: `IDS_FILEOP_BTN_PAUSE`, `IDS_FILEOP_BTN_RESUME`, `IDS_FILEOP_BTN_CANCEL`, `IDS_FILEOP_BTN_SPEED_LIMIT`, `IDS_FILEOPS_BTN_SKIP`, `IDS_FILEOPS_BTN_CANCEL_ALL`, `IDS_FILEOPS_BTN_MODE_QUEUE`, `IDS_FILEOPS_BTN_MODE_PARALLEL`, `IDS_FILEOP_SPEED_LIMIT_MENU_UNLIMITED`, `IDS_FILEOP_SPEED_LIMIT_MENU_CUSTOM`
- Format strings: `IDS_FMT_FILEOPS_OP_COUNTS`, `IDS_FMT_FILEOPS_ETA`, `IDS_FMT_FILEOPS_CALCULATING_TIME`, `IDS_FMT_FILEOPS_FILES_FOLDERS`, `IDS_FMT_FILEOPS_SIZE_PROGRESS`, `IDS_FMT_FILEOP_SPEED_LIMIT_BUTTON_BYTES`, `IDS_FMT_FILEOP_SPEED_LIMIT_MENU_BYTES`
- Overlay text: `IDS_FILEOPS_GRAPH_PAUSED`, `IDS_FILEOPS_GRAPH_WAITING`, `IDS_FILEOPS_GRAPH_CALCULATING`
- Confirmations:
  - `IDS_CAPTION_FILEOPS_CANCEL_ALL` / `IDS_MSG_FILEOPS_CANCEL_ALL_POPUP` — shown when clicking Cancel All button in popup
  - `IDS_CAPTION_FILEOPS_EXIT` / `IDS_MSG_FILEOPS_CANCEL_ALL_EXIT` — shown when exiting application with active operations
- Errors: `IDS_MSG_FILEOP_SPEED_LIMIT_INVALID`

## App Exit Behavior (Normative)

- When the main application window is closing while file operations are active (running or queued), the host MUST prompt with a context-specific confirmation dialog (title: "Exit Application", message explaining operations will be cancelled).
- The confirmation dialog MUST be centered on the main window.
- If the user cancels the dialog, the application close MUST be aborted.
- If the user confirms, the host MUST cancel all operations and continue closing.

## Performance Considerations

- Callback implementations SHOULD be fast; heavy work MUST remain off the UI thread.
- The popup SHOULD read progress using thread-safe snapshots (atomics / protected state) and render at a fixed cadence (~100ms).
- Plugins SHOULD throttle progress callbacks for large directory trees (time- and/or entry-based) to avoid excessive overhead.

## Future Enhancements

- Parallelize directory pre-calc across multiple source roots within a single task.
- Cache directory size results for recently-scanned directories.
- Expand cross-filesystem support and properties UX (see Appendix A, Phase 8).

## Appendix A — End-to-End Plan (Informative)

_Merged from the former `FileOperationsEndToEndPlan.md` (removed after merge; last updated: 2026-02-04)._

### Goal

Deliver reliable, scalable file operations (Copy/Move/Delete) with a responsive progress popup, by aligning:

- Filesystem plugin behavior (`Plugins/FileSystem/*`)
- Host execution model (`RedSalamander/FolderWindow.FileOperations*.cpp`)
- Progress popup UI + interactions (`RedSalamander/FolderWindow.FileOperations.Popup.*`)

This plan is explicitly tied to the existing specs and current codebase state (as of 2026-02-03).

### References (read first)

- Filesystem plugin improvements: `Specs/FileSystemPluginImprovementPlan.md`
- Plugin interface definitions: `Common/PlugInterfaces/FileSystem.h`
- VFS contract notes: `Specs/PluginsVirtualFileSystem.md`
- Theme key list: `Specs/SettingsStoreSpec.md`

### Current Implementation Snapshot (2026-02-03)

#### Host (already present)

- Background tasks with pre-calc + queueing:
  - `RedSalamander/FolderWindow.FileOperationsInternal.h`
  - `RedSalamander/FolderWindow.FileOperations.State.cpp`
- Pause/cancel implemented by blocking inside callback checkpoints:
  - `FolderWindow::FileOperationState::Task::WaitWhilePaused()` blocks inside:
    - `Task::FileSystemProgress(...)`
    - `Task::DirectorySizeProgress(...)`
- Wait/Parallel mode implemented via:
  - Start gating: `_waitForOthers` + `EnterOperation(...)` queue
  - Queue pause: `UpdateQueuePausedTasks()` sets `_queuePaused`
- Speed limit plumbing:
  - Popup sets desired limit → host stores `_desiredSpeedLimitBytesPerSecond`
  - Host writes `options->bandwidthLimitBytesPerSecond` from inside `FileSystemProgress/ItemCompleted`
  - Host passes `FileSystemOptions` into `CopyItems/MoveItems`

#### Popup (already present)

- Modeless top-level D2D popup with auto-sizing, scrollbar, theme + rainbow support:
  - `RedSalamander/FolderWindow.FileOperations.Popup.h`
  - `RedSalamander/FolderWindow.FileOperations.Popup.cpp`
- Speed limit parsing and presets are implemented:
  - `TryParseThroughputText(...)`, `ShowSpeedLimitMenu(...)`
- Long path text is clipped so it never renders underneath the right-side mini progress bar.
- Cancel All confirmation is implemented:
  - `ConfirmCancelAll(...)`
- Close button prompts Cancel All and destroys the window (`OnClose(...)`).

#### Local Filesystem plugin (needs improvements)

- Directory watch:
  - `Plugins/FileSystem/FileSystem.Watch.cpp` uses a small buffer pool and re-arms `ReadDirectoryChangesW` **before** invoking the host callback (I/O completion is decoupled from parsing/callback delivery).
  - `overflow=TRUE` is already emitted on OS overflow and internal caps.
- Directory enumeration:
  - `Plugins/FileSystem/FileSystem.DirectoryOps.cpp` uses progressive growth up to a soft cap and resumes enumeration on growth (no restart); it can fall back to a higher hard cap before returning `ERROR_INSUFFICIENT_BUFFER`.
    - defaults: 512 MiB soft / 2048 MiB hard
    - configurable via plugin configuration (`enumerationSoftMaxBufferMiB` / `enumerationHardMaxBufferMiB`)
- Batch file ops:
  - `Plugins/FileSystem/FileSystem.FileOps.cpp` runs `CopyItems/MoveItems` with bounded internal parallelism across top-level items:
    - default max concurrency: 4
    - configurable via `copyMoveMaxConcurrency` (max 8)
  - `DeleteItems` parallelizes delete with ordering safety across overlapping inputs (children before parents) and supports Recycle Bin deletes with bounded concurrency:
    - default max concurrency: 8
    - default max concurrency (Recycle Bin): 2
    - configurable via `deleteMaxConcurrency` / `deleteRecycleBinMaxConcurrency`
  - `CopyProgressRoutine` enforces a task-global bandwidth cap across workers (best-effort) and serializes callback delivery (host never sees concurrent callback calls).
  - Delete populates `completedBytes` best-effort using deleted file sizes (to support size progress when pre-calc totals exist).

### Cross-Layer Contract Decisions (resolve before parallel ops)

These are the “product rules” that determine the safest implementation strategy.

1) **Pause model vs plugin callback threading**
   - Host pause works by blocking inside `IFileSystemCallback::FileSystemProgress`.
   - Therefore, the plugin must only advance work at progress checkpoints that eventually call into the host callback (so that blocking the callback blocks work).
   - Decision: keep callbacks serialized and invoked from worker progress checkpoints (not a detached dispatcher that allows work to run ahead while UI is paused).

2) **Speed limit semantics under parallel Copy/Move**
   - `FileSystemOptions::bandwidthLimitBytesPerSecond` is per-task.
   - Decision: enforce a *task-global* limiter across all concurrent workers (not “each worker can use the full limit”).

3) **Delete bytes semantics**
   - Delete progress MAY report bytes when available (best-effort) so that the popup can show size progress when pre-calc totals exist.
   - The host SHOULD treat delete `completedBytes` as advisory; item counts remain the primary completion signal.

4) **Directory watcher overflow semantics**
   - `overflow==TRUE` MUST be treated as “resync required” (not just “some events may be coalesced”).

5) **Popup close behavior**
   - The popup close button continues to mean “Cancel All” (with confirmation).

6) **Multi-file in-flight display**
   - Under parallel Copy/Move, the popup shows multiple MRU in-flight file lines with per-file progress (up to max concurrency).

### Implementation Plan (phased)

#### Phase 1 — Contract & Spec Updates (documentation-first)

Files:

- `Common/PlugInterfaces/FileSystem.h`
- `Specs/FileOperationsSpec.md`
- `Specs/FileSystemPluginImprovementPlan.md`

Tasks:

- [x] Clarify in `Common/PlugInterfaces/FileSystem.h` that `overflow==TRUE` means “resync required” (not just “dropped/coalesced”).
- [x] Document callback expectations that affect pause/responsiveness:
  - progress callbacks may be blocked by the host
  - plugins must avoid holding locks that would deadlock if callbacks block
- [x] Align file-op progress semantics across layers:
  - out-of-order `itemIndex` completion is allowed
  - `totalBytes` may be unknown/0 (host pre-calc provides totals when available)
  - delete bytes policy (see Decision #3)
- [x] Update `Specs/FileOperationsSpec.md` to explicitly define popup close behavior (Decision #5).

Acceptance:

- The desired behavior is unambiguous and implementable without hidden coupling between host pause and plugin threading.

#### Phase 2 — Directory Watch Resiliency (plugin) + Overflow Resync (host)

Files:

- `Plugins/FileSystem/FileSystem.Watch.cpp`
- `RedSalamander/FolderWatcher.h`
- `RedSalamander/FolderWatcher.cpp`
- (likely) `RedSalamander/DirectoryInfoCache.cpp`

Tasks (plugin):

- [x] Implement buffer pool (2–4 buffers) so we can re-arm reads immediately on completion.
- [x] Re-issue `ReadDirectoryChangesW` *before* invoking callbacks (decouple I/O from processing).
- [x] Add a bounded “completed buffer” processing queue; on queue overflow set `overflow=TRUE` and coalesce.
- [x] Keep the “no callbacks after `UnwatchDirectory` returns” guarantee using RAII + `CancelIoEx` + `WaitForThreadpoolIoCallbacks`.

Tasks (host):

- [x] Update `FolderWatcher::PluginCallback::FileSystemDirectoryChanged(...)` to inspect `notification->overflow`.
- [x] On `overflow==TRUE`, schedule a full refresh (coalesced/rate-limited per folder) instead of relying on best-effort incremental updates.
  - Note: current `DirectoryInfoCache` behavior is “refresh on any event”; overflow is logged as telemetry so loss is visible.

Acceptance:

- Under heavy churn + slow host processing, the watch keeps re-arming and the app remains stable.
- Overflow results in a reliable resync path without UI stalls.

#### Phase 3 — Directory Enumeration Scalability (plugin)

Files:

- `Plugins/FileSystem/FileSystem.DirectoryOps.cpp`

Tasks:

- [x] Keep `struct FileInfo` layout unchanged (ABI safety); only adjust internal sizing/alignment calculations if validated.
- [x] Reduce hard failures on very large enumerations:
  - grow up to a cap in a single enumeration pass (resume-on-grow; no restart)
- [x] Add a higher hard-cap fallback path when the soft cap is exceeded (current hard cap: 2GB).
- [x] Ensure large buffers trim back after use (`MaybeTrimBuffer`) so memory returns to baseline.

Acceptance:

- Large directories no longer fail in typical workloads; worst-case behavior is graceful (fallback or clear error).
- Post-enumeration memory does not permanently retain huge allocations.

#### Phase 4 — Parallel Batch Operations (plugin)

Files:

- `Plugins/FileSystem/FileSystem.FileOps.cpp`

Design requirements:

- Bounded concurrency (copy/move low; delete higher with ordering safety).
- Serialized callbacks per operation (host callback must never be invoked concurrently).
- Host pause must still work: blocking inside `FileSystemProgress` must eventually stall all workers at checkpoints.
- Speed limit must be task-global once work is parallel.
- Copy/Move progress callbacks must include per-file `currentSourcePath` + `currentItem*Bytes` frequently enough that the popup can keep multiple in-flight lines up to date.

Tasks:

- [x] Refactor `OperationContext` into:
  - shared thread-safe state (atomics, limiter state, stop flag)
  - per-item/per-worker state (paths, per-file progress bookkeeping)
- [x] Implement bounded worker queue for `CopyItems/MoveItems` (start conservative: 2–4).
- [x] Implement ordered parallel delete:
  - schedule parent/child deletes using a dependency graph (children before parents)
  - allow overlapping inputs without falling back to sequential
  - support Recycle Bin deletes with bounded concurrency
- [x] Implement a task-global bandwidth limiter used by all workers:
  - honor dynamic updates to `options->bandwidthLimitBytesPerSecond`
  - ensure effective throughput approximates the cap even with N workers
- [x] Throttle progress callback frequency (but keep pause/cancel responsive).
- [x] Decide + implement delete bytes semantics (Decision #3).

Acceptance:

- Copy/move improves throughput on multi-item workloads without regressing UI responsiveness.
- Pause/resume is reliable (workers converge to paused quickly).
- Cancellation is responsive and consistently returns `ERROR_CANCELLED`.

#### Phase 5 — Host File-Operation Engine Alignment

Files:

- `RedSalamander/FolderWindow.FileOperations.State.cpp`
- `RedSalamander/FolderWindow.FileOperationsInternal.h`
- `RedSalamander/FolderWindow.FileOperations.cpp`

Tasks:

- [x] Make task progress counters monotonic (completed items/bytes should never go backwards even if a plugin reports out-of-order / regressive updates).
- [x] Verify the host callback implementation remains correct if plugin completions arrive out-of-order:
  - host derives UI state primarily from `FileSystemProgress` counters (not `itemIndex` ordering)
  - `itemIndex` is treated as informational (not used for sequencing assumptions)
- [x] Expose an effective applied speed limit to the popup bandwidth graph (use plugin-reported `FileSystemOptions::bandwidthLimitBytesPerSecond` when available; fall back to the host desired limit).
- [x] Validate pre-calc skip/cancel edge cases:
  - covered by `--fileops-selftest` (Phase 5: pre-calc cancel, pre-calc skip, cancel queued)
  - optional manual checks remain valuable for validating the popup UX (“Skip/Cancel” feel) on real trees.
- [x] Ensure default Wait-mode queuing is race-free when multiple tasks are started rapidly (treat newly-created tasks as “active” for `ShouldQueueNewTask`).
- [x] Choose the queue-pause “oldest active task” using operation-enter time (not task ID) for deterministic Wait-mode serialization.
- [x] Confirm Wait/Parallel switching behavior matches this spec under:
  - multiple active tasks in pre-calc (covered by `--fileops-selftest` Phase 5 switch tests)
  - multiple active tasks mid-copy/move (manual validation still recommended)

Acceptance:

- Host state remains thread-safe, monotonic where expected, and matches the popup’s assumptions.

#### Phase 6 — Popup Behavior & UX Finalization

Files:

- `RedSalamander/FolderWindow.FileOperations.Popup.cpp`
- `Specs/FileOperationsSpec.md`

Tasks:

- [x] Align popup close semantics with spec (Decision #5).
- [x] When Copy/Move runs with multiple files in-flight, display multiple source lines with per-file progress (MRU; up to max concurrency lines) (Decision #6).
- [x] Ensure per-file source text never renders underneath the right-side mini progress indicator (reserve width + clip).
- [x] Clamp per-file in-flight `completedBytes <= totalBytes` in popup snapshots (avoid “empty” mini bars on overshoot/races).
- [x] Align rainbow bandwidth graph sample colors with progress-bar colors (use the stored per-sample hue + theme value; treat empty path as accent).
- [ ] Validate popup rendering expectations under parallel ops:
  - smoke coverage: `--fileops-selftest` Phase 6 (resize + pause while copy is running)
  - graph scaling remains readable
  - “Paused/Waiting/Calculating” overlays remain correct
  - rainbow mode remains stable per sample (no blinking)
  - Suggested manual checks:
    - With Copy/Move concurrency > 1, verify multiple MRU “From:” lines render (no overlap with the right-side mini bars) and update as files change.
    - Toggle Wait/Parallel while 2+ tasks are running and while 2+ tasks are in pre-calc; verify only one task continues in Wait mode and others show “Waiting” overlays.
    - Pause/Resume a running task; verify graph/history freezes while paused and resumes smoothly (no sudden spikes from dt accumulation).
    - Enable `menu.rainbowMode`; verify per-sample colors persist as the graph scrolls (older samples retain their hue) and match per-file progress colors.
    - Resize the popup narrower; verify path truncation uses middle ellipsis and never renders under the mini progress indicator.
- [x] If delete bytes are implemented, validate the delete card’s “size progress” line is meaningful (covered by `--fileops-selftest` Phase 6).
- [x] Wire up the destination selector UX (click the destination chevron next to `To:` before task start to open `ShowDestinationMenu(...)`).

Acceptance:

- Popup remains responsive and visually stable with many concurrent tasks.
- All strings remain localized and all colors remain theme-derived.

#### Phase 7 — Verification & Stress Matrix (gates)

Build gates:

- [x] `.\build.ps1 -ProjectName FileSystem`
- [x] `.\build.ps1` (full solution)

Stress scenarios:

- [x] Watcher churn: rapid create/rename/delete under a watched folder (covered by `--fileops-selftest` Phase 7).
- [ ] Large directory listing: directories with many entries/long names.
  - baseline covered by `--fileops-selftest` Phase 7
  - knob coverage: set FileSystem `enumerationSoftMaxBufferMiB` / `enumerationHardMaxBufferMiB` lower/higher and verify behavior (no long-lived huge buffers; expected `ERROR_INSUFFICIENT_BUFFER` only when hard cap is hit).
- [ ] Parallel copy/move:
  - baseline knob coverage covered by `--fileops-selftest` Phase 7 (`copyMoveMaxConcurrency` 1/4/8 + speed limit toggle + MRU in-flight lines)
  - many small files and several large files
  - knob coverage: set FileSystem `copyMoveMaxConcurrency` to `1`, `4`, `8` and verify:
    - throughput scales on many-small-file workloads
    - popup shows multiple in-flight file lines (MRU; up to the configured concurrency / host UI max)
  - speed limit set/unset mid-flight
  - pause/resume repeatedly
- [ ] Parallel delete:
  - baseline knob coverage covered by `--fileops-selftest` Phase 7 (`deleteMaxConcurrency` 1/8)
  - deep directory trees
  - recycle bin on/off
  - knob coverage: set FileSystem `deleteMaxConcurrency` / `deleteRecycleBinMaxConcurrency` and validate correctness vs throughput
- [ ] Wait/Parallel switching while:
  - multiple tasks are in pre-calc
  - multiple tasks are in operation

Instrumentation (recommended, debug-only):

- [x] Watcher: queue depth, overflow count, callback latency (`FileSystem.Watch`).
- [x] File ops: cancel latency, limiter target vs achieved throughput, progress callback frequency (`FileOps.PreCalc`, `FileOps.Operation`, `FileOps.CancelLatency`).
- [x] Enumeration: peak buffer size, fallback usage, trim events (`FileSystem.DirectoryOps.Enumerate`, `FileSystem.DirectoryOps.TrimBuffer`).
- [x] Debug-only end-to-end self-test runner:
  - run: `.\.build\x64\Debug\RedSalamander.exe --fileops-selftest`
  - log: `%TEMP%\\RedSalamander.FileOpsSelfTest.log`
  - exit code: `0` pass / `1` fail

### Rollout / Risk Notes

- Prefer landing in small PRs aligned with phases (watch → enum → ops → UI polish).
- Parallel batch ops and global speed limiting are the highest-risk changes; gate behind conservative defaults first (small concurrency) and scale up later.
- Keep contracts explicit in `Common/PlugInterfaces/FileSystem.h` so other filesystem plugins can follow the same expectations.

#### Phase 8 — Cross-Filesystem Operations + Drag/Drop + Properties (new scope)

Files:

- Contracts:
  - `Common/PlugInterfaces/FileSystem.h`
  - `Specs/PluginsVirtualFileSystem.md`
  - `Specs/FileOperationsSpec.md`
- Host:
  - `RedSalamander/FolderWindow.FileOperationsInternal.h`
  - `RedSalamander/FolderWindow.FileOperations.State.cpp`
  - `RedSalamander/FolderWindow.FileOperations.cpp`
  - `RedSalamander/FolderView.DragDrop.cpp`
  - `RedSalamander/FolderViewInternal.h` (IDataObject formats)
  - `RedSalamander/FolderView.FileOps.cpp` (Alt+Enter properties routing)
- Plugins (at least):
  - `Plugins/FileSystem/*`
  - `Plugins/FileSystemDummy/*`
  - (optionally) `Plugins/FileSystem7z/*`, `Plugins/FileSystemCurl/*`, `Plugins/FileSystemS3/*`

Tasks:

- [x] Define and implement cross-filesystem bridge contracts:
  - Add `IFileWriter` + `IFileSystemIO::CreateFileWriter` for streaming writes.
  - Add `IFileSystem::GetCapabilities` for explicit operation support + cross-filesystem import/export policy.
  - Add `IFileSystemIO::GetItemProperties` for themed properties UI on non-Win32 paths.

- [x] Implement host cross-filesystem Copy/Move:
  - Allow cross-pane Copy/Move when plugin contexts differ if both sides explicitly allow the operation (capabilities policy) and required interfaces are present.
  - Prefer in-memory streaming `IFileReader` → `IFileWriter`.
  - Preserve Pause/Cancel semantics and speed limit (host-global limiter shared across per-task worker threads).
  - Integrate with existing conflict prompting (overwrite/read-only/access denied/etc.).
  - Move is Copy + Delete (delete only after successful copy).

- [x] In-app drag/drop routing:
  - Add an internal data-object format for pane→pane drag/drop that carries plugin id + mount context + per-item paths (internal paths).
  - Drop handler MUST prefer this internal format when present, so pane→pane drag/drop uses the same host queue/popup and cross-filesystem bridge behavior as `F5`/`F6`.
  - External drag/drop to Explorer MUST remain supported for filesystem-backed items (continue exporting `CF_HDROP`).

- [x] Properties UX:
  - Replace/extend the `Alt+Enter` / Properties command so non-Win32 filesystem plugins can surface properties via `IFileSystemIO::GetItemProperties`.
  - Display properties in a themed dialog; include richer per-bucket/unknown text for errors where possible.

- [x] Self-test coverage:
  - Add `--fileops-selftest` steps that exercise cross-filesystem Copy/Move on a small tree:
    - local ↔ dummy (sizes, counts, cancel, overwrite prompt where applicable)
  - Add deterministic test coverage for “apply to all similar” caching behavior in the prompt flow (UI-layer observable effect, not just internal cache state) when possible.

Acceptance:

- Cross-pane copy/move works for supported plugin pairs without requiring “same filesystem context”.
- Pane→pane drag/drop uses the same popup/queue/conflict UX as keyboard/menu operations.
- External drag/drop to Explorer does not regress for local filesystem-backed items.
- Properties (Alt+Enter) works for non-Win32 plugins and remains localized + theme-respecting.

## Appendix B — Conflict Handling Plan & Status (Informative)

_Merged from the former `ExtremelyLongAndDeliberatelyStupidNameForFutureConflictHandlingUserPromptsPlan.md` (removed after merge; last updated: 2026-02-04)._

### Scope

This note captures the agreed plan to add user-driven conflict handling (overwrite/read-only/permission etc.) for file operations, using the existing popup instead of new modal dialogs.

### Status (2026-02-04)

- Implemented end-to-end (host per-item orchestration + inline popup prompt + decision cache + retry cap).
- Automated coverage added to `--fileops-selftest` (Phase9_* conflict prompt steps).
- Remaining test gap: deterministic `PermanentDelete` (Recycle Bin failure) repro in selftest is environment-dependent.
- Follow-up requirements implemented:
  - FolderView entry points (commands + drag/drop) route through host queue/popup when available.
  - Per-task multi-item concurrency supports serialized prompts and apply-to-all caching.
  - Conflict prompt shows the specific in-flight item paths (not just the bucket).
  - UI-layer apply-to-all caching verified via popup selftest invoke messages.

### Current gaps (2026-02-03)

- Host ignores `status` in `FileSystemItemCompleted`; no prompts ever show.
- Default flags in `FolderWindow.FileOperations.cpp` pre-authorize overwrite/replace-readonly/continue-on-error, so conflicts are auto-resolved or hidden.
- Bulk `CopyItems/MoveItems/DeleteItems` are used; there is no per-item retry/skip orchestration.
- Popup has no interaction surface for conflicts (progress only).

### Target behavior

- Conflicts trigger an inline prompt in the popup (Direct2D), not a modal HWND.
- User options per conflict type: Overwrite, Replace read-only, Permanent delete (when Recycle Bin fails), Skip, Skip All, Cancel, Retry (single-item only).
- Retry is capped to one attempt per item per bucket; second failure removes Retry and offers Skip/Skip All/Cancel.
- “Apply to all similar” applies only to non-retry actions.
- Close button on popup continues to mean “Cancel All” (with existing confirmation).

### Implementation plan (completed)

1) [x] Tighten defaults
   - Remove `ALLOW_OVERWRITE`, `ALLOW_REPLACE_READONLY`, `CONTINUE_ON_ERROR` from default flag sets in `FolderWindow.FileOperations.cpp`. Keep `USE_RECYCLE_BIN` for delete.

2) [x] Per-item orchestration
   - In `Task::ExecuteOperation`, drive items individually (or small batches) using `CopyItem/MoveItem/DeleteItem` instead of `*Items`.
   - Maintain aggregate progress counters so popup totals stay accurate; keep pause by blocking inside callbacks.
   - Progress aggregation details (discovered during implementation):
     - `CopyItem/MoveItem` may not emit a final `FileSystemProgress` with `completedItems == totalItems` (it *does* emit `FileSystemItemCompleted`), so the host must advance `completedItems` using item-completed callbacks (monotonic by callback count).
     - Per-item calls reset plugin-reported totals (`totalItems == 1`, `totalBytes == itemBytes`), so the host must map progress into a task-global view (`plannedItems`, and pre-calc totals when available).
   - Path semantics:
     - Destination paths must be formed without assuming Windows-only separators; join folder + leaf using the folder’s observed separator style (`/` vs `\\`) so virtual file systems aren’t broken by `std::filesystem::path` normalization.
   - Rollout note:
     - Keep bulk/batched execution as the default until conflict UI is ready (to avoid regressing existing plugin parallelization); enable per-item orchestration when conflict handling needs host-driven per-item decisions.

3) [x] Conflict detection
   - In per-item execution, classify failures into buckets and pause the item loop until a user decision arrives.
   - Bucket set: exists, read-only, access denied, sharing violation, disk full, path too long, recycle bin failure, network/offline, unknown.
   - Recycle Bin nuance (discovered during implementation):
     - When `FILESYSTEM_FLAG_USE_RECYCLE_BIN` is set, `ALLOW_REPLACE_READONLY` does not apply (shell delete path), so failures are bucketed as `RecycleBinFailed` to offer the `PermanentDelete` fallback first.

4) [x] Retry rules
   - Retry is per-item only; never cached or applied to all.
   - One retry per item per bucket, with optional 500–1000 ms backoff for transient buckets.
   - If the retry fails with the same bucket, re-prompt with “Retry failed” and only Skip/Skip All/Cancel.

5) [x] Decision cache
   - Add `ConflictAction` + `ApplyToAll` enums and a per-task cache keyed by bucket (mutex-protected). Cache excludes Retry.

6) [x] Popup prompt panel
   - Extend `TaskSnapshot`/`PopupHitTest` to carry a prompt payload and button hits.
   - Render panel above task list in `FolderWindow.FileOperations.Popup.*`; route clicks via posted messages; worker waits on `wil::unique_event`.
   - Localize strings and reuse theme/high-contrast/rainbow handling.

7) [x] Threading + RAII
   - Keep serialized callbacks; avoid holding locks while waiting for user input.
   - Callback pointers stay raw and cookie is passed back verbatim.

8) [x] Tests
   - Unit: HRESULT→bucket mapping, retry guard, aggregate progress math for per-item loop.
   - Manual: conflict prompts inline; retry capped; apply-to-all works for non-retry; pause/cancel responsiveness retained; large batches run without pre-scan.
   - Add/extend `--fileops-selftest` coverage:
     - No-overwrite copy fails with `HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS)` and destination stays unchanged.
     - Per-item copy keeps `totalItems == plannedItems` (not `1`) and aggregates `completedBytes` to pre-calc totals.
     - Conflict prompt coverage:
       - Exists → Overwrite (and follow-up ReadOnly → ReplaceReadOnly).
       - Skip All caches decision and returns `ERROR_PARTIAL_COPY`.
       - Retry cap (SharingViolation): second failure removes Retry and sets retryFailed.
     - FolderView routing coverage:
       - Assert both `RedSalamanderFolderView` instances have a file-op request callback installed (so drag/drop and F5/F6/Delete won't bypass host popup/queue in normal UI).
      - UI-layer apply-to-all coverage:
        - Use `WndMsg::kFileOpsPopupSelfTestInvoke` (see `Common/WindowMessages.h`) to toggle "Apply to all" and click the chosen action, verifying UI wiring and bucket decision caching.

### File touchpoints

- `RedSalamander/FolderWindow.FileOperations.cpp` (flag defaults)
- `RedSalamander/FolderWindow.FileOperationsInternal.h` (decision enums/cache/state)
- `RedSalamander/FolderWindow.FileOperations.State.cpp` (per-item loop, conflict queue, retry guards)
- `RedSalamander/FolderWindow.FileOperations.Popup.h/.cpp` (prompt panel UI + routing)
- `.rc` + `Resources.h` (strings/buttons)
