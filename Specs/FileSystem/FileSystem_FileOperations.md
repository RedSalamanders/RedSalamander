# File Operations Specification

Last updated: 2026-06-16

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
- Standalone captioned File Operations windows MUST follow the shared app-owned tool-window chrome contract, including the persisted `ui.windowBackdrop` preference.

## Performance Validation Contract

File Operations is a performance-sensitive subsystem. Any new feature or optimization that can affect pre-calculation, queueing, progress updates, pause/cancel responsiveness, popup refresh churn, cross-filesystem bridging, throughput, or lock hold time MUST:

- identify the protected scenario before implementation is considered complete,
- add or reuse measurable instrumentation,
- add deterministic `--fileops-selftest` coverage or another deterministic harness,
- archive validation runs under `Specs/TestRuns/`,
- support any claimed improvement with archived before/after evidence.

This requirement is mandatory from the first landing of the feature or optimization, including baseline-only instrumentation work.

File Operations work SHOULD prefer the existing `FileOps.*` metric families when they already cover the scenario; otherwise the change MUST add the missing metrics with the feature.

Correctness selftests for concurrency, conflict routing, and destructive safety MUST prove the vulnerable behavior they claim to protect. A test that only proves eventual completion is not enough for liveness-sensitive paths; it MUST also assert a relevant progress/dispatch shape. A conflict-routing test MUST use distinguishable decisions and verify the final state of each affected child independently. A UI-sampler test MUST either drive the sampler deterministically or prove the live window/timer prerequisite it relies on.

Progress-display investigations MUST keep UI cadence and operation cadence separable:

- `FileOps.Progress.FirstCallbackDelayMs` records delay from task start to the first progress callback.
- `FileOps.Progress.MaxCallbackGapMs` / `FileOps.Progress.CallbackGapMs` record aggregate progress-callback gaps.
- `FileOps.Progress.MaxCallbackGapBytes` / `FileOps.Progress.CallbackGapBytes` record current-item bytes associated with the largest gap and aggregate gaps, so a run can distinguish "no copy progress" from "progress arrived in large callback bursts".
- `FileOps.Progress.Stream.*` records the same callback-gap and byte-gap measurements per `(cookie, progressStreamId)`.
- `FileOps.Popup.Rate.UpdateUs`, `FileOps.Popup.Rate.MaxCallbackSilenceMs`, and `FileOps.Popup.Rate.MaxDisplayGapMs` record popup-side rate-update work, callback silence observed by the popup, and display-timer gaps.

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

- Plugin operation and callback contracts: `Specs/Plugins/Plugins_VirtualFileSystem.md`
- Theme key list (including file ops keys): `Specs/Core/Core_SettingsStore.md`

## Archive Operations (Normative)

`cmd/pane/pack` creates an archive from the active local pane selection. `cmd/pane/unpack` extracts selected or focused local archives to a destination selected by the Unpack prompt. Unsupported providers MUST keep the pane in place and show localized pane feedback.

Pack uses the selected items, or the focused item when nothing is selected. The built-in `ZIP (Plugin)` packer writes deterministic stored ZIP archives with `/` separators, preserved selected empty directories, and entries sorted by archive path. Other writable archive formats are discovered from the bundled `7zip.dll` and are created through `IOutArchive::UpdateItems`.

Unpack supports stored ZIP entries through the built-in reader and delegates compressed ZIP entries, 7-Zip archives, and other formats supported by the bundled `7zip.dll` to the 7-Zip extraction path. Both extraction paths preserve the same destination, overwrite, mask, and safe-entry-path contract. Unsupported/encrypted methods fail with localized pane feedback.

Archive delete-after options are destructive cleanup operations. When `delete after packing` or `delete after unpacking` is enabled, the host MUST show the permanent-delete confirmation prompt with default Cancel before removing selected sources or selected archive files. Cancelling that prompt MUST prevent the cleanup delete from running.

Archive entry names MUST be relative `/`-separated paths. Both extraction paths MUST reject empty names, absolute paths, drive-qualified paths, UNC-style prefixes, backslashes, `.` / `..` components, colons, embedded NULs, and reserved DOS device names such as `CON`, `PRN`, `AUX`, `NUL`, `COM1`-`COM9`, and `LPT1`-`LPT9` even when followed by an extension.

After combining an entry name with the chosen destination, both extraction paths MUST verify that the normalized target remains inside the destination directory before creating directories, opening files, or moving temp files into place. 7-Zip extraction MUST reject symbolic-link and hard-link entries with `ERROR_NOT_SUPPORTED`.

ZIP names MUST be decoded as UTF-8 when general-purpose flag bit 11 is set, and as CP437 otherwise. Extraction should create parent directories before file entries and should commit file contents via same-directory temp files. Transient `ERROR_SHARING_VIOLATION` / `ERROR_LOCK_VIOLATION` failures while replacing the final output MAY be retried briefly before reporting failure.

## Settings And Ownership (Normative)

File Operations settings are split between host-owned global defaults and plugin-owned behavior knobs.

### Host-owned global settings

The host-wide `fileOperations.*` settings live in `SettingsStore` and in `Preferences -> File Operations`.

- `fileOperations.preCalcEnabled`
- `fileOperations.preCalcMaxWorkers`
- `fileOperations.autoDismissSuccess`
- `fileOperations.crossFsBridgeBufferSizeKB`
- `fileOperations.defaultBandwidthLimitBytesPerSecond`

These settings:

- MUST apply to newly created tasks only.
- MUST be snapshotted at task creation time.
- MUST NOT retune already-running tasks when the user changes Preferences.

### Plugin-owned settings

Plugin-specific file-operation knobs live in `Preferences -> Plugins -> [plugin]`.

For the built-in FileSystem plugin, the current settings include:

- `concurrencyMode` (`auto` / `manual`)
- `copyMoveMaxConcurrency` (`1..16`, default `4`)
- `deleteMaxConcurrency` (`1..64`, default `8`)
- `deleteRecycleBinMaxConcurrency` (`1..16`, default `2`)
- `recycleBinBatchSize` (`1..1000`, default `500`)
- `searchMaxDirectoryWalkers` (`1..8`, default `4`)

The File Operations host pane MUST NOT duplicate those plugin-owned controls. It MAY show explanatory text directing the user to the plugin page.

For the built-in local FileSystem plugin, `copyMoveMaxConcurrency` is an operation-level worker/transfer budget for Copy/Move work, not only a selected-root count. A recursive directory Copy/Move with one selected folder SHOULD still use the configured worker budget when there is independent file or child-directory work inside that folder. A batch Copy/Move with multiple selected folders SHOULD keep selected roots progressing concurrently while lending spare budget to recursive subtrees; uneven folder trees MUST NOT force the dominant subtree to remain serial when `copyMoveMaxConcurrency` has spare capacity.

### Settings precedence

For plugin-owned concurrency:

1. Per-task override (when the operation surface exposes one)
2. Per-connection override (`Connection Manager -> extra`, when non-zero)
3. Plugin-wide default (`Preferences -> Plugins -> File System`)
4. Plugin hardcoded default

For host-owned bandwidth defaults:

1. Per-task speed-limit override in the progress popup
2. Global default `fileOperations.defaultBandwidthLimitBytesPerSecond`
3. Hardcoded default (`0` = unlimited)

## Issues Pane Contract

The file-operations subsystem also exposes a dedicated top-level Issues pane for warnings and errors collected from active and completed tasks.

- The Issues pane MUST open as an independent top-level window.
- The visible body MUST use the shared DX grid path.
- The pane MUST show issue rows built from file-operation issue state, including at least time, task, operation, severity, HRESULT, status, category, message, source path, and destination path when present.
- The pane MUST preserve stable-row selection when the selected issue is still present after a refresh.
- A no-change refresh MUST NOT rebuild or disturb the applied grid state.
- Sorting, wheel scrolling, and persisted window placement MUST remain part of the pane contract.
- The pane follows the shared `DxUi` accessibility and bounded-work validation rules from `Specs/UI/UI_DxUiSharedGrid.md`.

## Window Backdrop Contract

The standalone captioned File Operations HWNDs are:

- the progress popup (`FileOperationsPopup`),
- the Issues pane (`FileOperationsIssuesPane`),
- the speed-limit prompt (`FileOperationsSpeedLimitPromptWindow`).

These windows MUST apply the persisted `ui.windowBackdrop` setting through the shared window chrome/backdrop helper path with tool-window target semantics. They MUST gracefully resolve to no system backdrop when the OS or high-contrast accessibility mode requires it. Activation handling MUST update title-bar active/inactive state without reapplying DWM backdrop work on every activation message.

Backdrop regression coverage belongs in `--commands-selftest` cases for the Issues pane and speed-limit prompt/progress popup. Performance validation for backdrop-related changes SHOULD reuse existing File Operations popup metrics such as `FileOps.InfoTask.EnsurePopupVisible.*`, `FileOps.Popup.WmPaintUs`, and `FileOps.Popup.Render.TotalUs`.

## Popup Windowing And Locking Contract

The progress popup may receive synchronous Win32 callbacks while it is being shown, moved, resized, or invalidated. Calls such as `ShowWindow`, `SetWindowPos`, and similar HWND-affecting APIs MUST NOT be made while holding the file-operation state mutex when the target window procedure can re-enter `FileOperationState` (for example to persist the last popup rectangle from `WM_MOVE`). Code that needs to re-show or reposition an existing popup MUST snapshot the HWND and any required state under the mutex, release the mutex, validate the HWND, then perform the Win32 windowing calls.

`--fileops-selftest --selftest-case=Phase14_PopupHostLifetimeGuard` owns regression coverage for this contract, including the hidden/offscreen popup re-entry case that starts a new operation while the popup is being made visible again.

## Execution Model (Normative)

### Threading

- The host MUST execute each Task (including pre-calc and `IFileSystem::*`) on a background worker thread (one per task).
- When a Task uses per-item execution with per-item concurrency (`maxConcurrency > 1`), the host SHOULD schedule per-item work using a **shared worker pool** across all active Tasks (especially in Parallel mode) so the total worker thread count stays bounded and workers can be reassigned between Tasks after each item.
- When a file-system plugin uses internal parallelism (e.g. plugin max concurrency knobs for Copy/Move/Delete), the plugin SHOULD schedule that work using a **shared bounded worker pool** across all in-flight operations (rather than spawning per-operation thread pools) so worker threads can be reused/reassigned between operations as items finish.
- The built-in local FileSystem plugin MUST use the same recursive copy engine for `CopyItem(...)`, `CopyItems(...)`, and the copy/delete fallback used by cross-volume `MoveItem(...)` / `MoveItems(...)`.
- `CopyItems(..., count == 1, ...)` MUST NOT bypass recursive directory parallelism when the selected item is a recursive, non-reparse directory and the effective `copyMoveMaxConcurrency` is greater than `1`.
- Recursive copy workers SHOULD enumerate nested directories progressively and enqueue discovered file/directory work under the same operation budget so deep or uneven folder trees can keep available workers busy.
- Batch Copy/Move MUST bound actual file transfers by the operation's effective `copyMoveMaxConcurrency` even when several selected roots each start recursive worker jobs. Recursive worker jobs MAY expose a larger per-root worker allowance to improve queue fairness, but file-transfer admission MUST remain gated by the shared operation state.
- Cross-volume Move copy-delete fallback MUST use the same recursive copy engine and operation-level transfer budget as Copy. Debug selftests MAY force the `ERROR_NOT_SAME_DEVICE` fallback path on a same-volume move, but production Release behavior MUST only take that path after the OS reports the cross-device move error.
- Serial copy fallback is still intentional for max concurrency `1`, non-recursive directory operations, root reparse points that are not followed, non-directory sources, scheduler unavailability, cancellation/error convergence, and operations where parallel setup is not useful.
- Recursive directory copy and multi-item delete MUST dispatch each unit of work as a **short scheduler work item** (the worker returns to the shared pool between units) rather than running a long-lived per-worker consumer loop. A single large operation MUST NOT pin every shared-pool worker for its whole duration; concurrent file operations share the pool fairly and each continues making progress (no zero-progress stall).
- Recursive copy saturation coverage MUST assert more than eventual completion: it MUST prove that dispatch work continues beyond the configured transfer concurrency (for example with `FileOps.CopyRecursiveParallel.WorkItemDispatches`) and/or that at least two distinct selected roots make in-flight progress in the same sample window when the concurrency budget permits it.
- **Storage-adaptive concurrency**: the built-in Local FileSystem MUST probe the destination volume's medium (seek-penalty / bus type) and adapt copy/move and delete concurrency to it. Seek-penalty/rotational (HDD) media MUST clamp copy/move concurrency to `1`; NVMe MAY queue deeply; SSD uses a moderate degree; removable media clamps low. On probe failure the host MUST fall back to the historical default concurrency.
- Storage-adaptive probing MUST resolve the real volume root before opening a device for IOCTLs. Mounted-volume and SUBST-style aliases MUST probe the backing volume, not the lexical drive letter. If seek-penalty classification fails but bus type succeeds, an NVMe bus type MUST still select the NVMe/deep-queue profile.
- The host MUST drive progress via `IFileSystemCallback` and forward updates to the UI thread.
- The host MAY surface background work as Informational Tasks. Informational Tasks:
  - MUST NOT participate in Wait/Parallel queueing rules.
  - MUST NOT block or pause file-operation Tasks.
  - MUST be read-only in the UI (no conflict prompts; no overwrite/replace-readonly decisions).

### Built-in Local FileSystem Quiet Points

The built-in local FileSystem plugin has worker paths that can outlive the initiating UI stack (file operations, search workers, and directory watching). Those paths MUST:

- pin module lifetime or provide an equivalent quiet point before work can outlive the caller;
- stop producers and request cancellation before callback pointers or cookies are released;
- join or drain worker/threadpool callbacks before returning from teardown APIs;
- invoke host callbacks serially for a single logical operation/search/watch delivery unless the ABI explicitly allows concurrency.

`IFileSystemDirectoryWatch::UnwatchDirectory(...)` MUST cancel outstanding directory-change I/O, drain threadpool I/O/work callbacks, and return only after no further callbacks for that watch can run. Rename/delete/recreate churn under a watched directory is validated by `--fileops-selftest` Phase 7 and must keep the overflow/resync contract intact.

Focused teardown coverage is split across `--fileops-selftest --selftest-case=Phase14_PopupHostLifetimeGuard`, `--fileops-selftest --selftest-case=Phase7_WatcherChurn`, and `--commands-selftest --selftest-case=filesystem_local_watch_unwatch_drains_inflight_callback`.

### Preconditions (cross-pane Copy/Move)

- The host MUST reject cross-pane **Copy**/**Move** when both panes point to the **same effective destination folder** (same normalized path text), because this is almost always user error (accidental self-copy / self-move).
  - The host MUST show a localized error (see `IDS_MSG_PANE_OP_REQUIRES_DIFFERENT_FOLDER`).

- The host MUST reject **Copy**/**Move** when any source item would be copied/moved onto itself or into its own subtree (destination folder overlaps the source item), because this can recurse indefinitely or produce confusing no-op operations.
  - The host MUST show a localized error (see `IDS_FMT_FILEOPS_INVALID_DESTINATION_OVERLAP`).
  - The overlap check MUST compare both raw normalized paths and resolved/canonicalized candidates when canonicalization is available. A failed or degraded canonicalization result MUST NOT create an alias bypass that allows copy/move into the source itself or its subtree.

- **Same-context copy/move**:
  - If both panes are operating on the same effective file system context (same filesystem plugin id + same per-instance mount context when the plugin uses `IFileSystemInitialize`), the host SHOULD execute Copy/Move using that plugin instance directly.
  - Before creating a same-context Copy/Move/Delete task, the host MUST read `IFileSystem::GetCapabilities()` and reject the operation when capabilities fail, are missing/empty/invalid, omit or malform mandatory `pathIdentity`, advertise unstable path text (`pathIdentity.pathTextStableIdentity == false`), or advertise the corresponding operation as unsupported (`operations.copy`, `operations.move`, or `operations.delete` is `false`). Rejection MUST happen before task creation, popup allocation, or worker-thread start, and MUST show localized pane feedback.
  - `pathIdentity` is mandatory provider capability data for every built-in provider. Identity-sensitive planners such as Batch Rename and the shared same-context Copy/Move/Delete gate MUST fail closed before mutation when `pathIdentity` is missing, malformed, unstable, or unsupported.
  - A failed `GetCapabilities()` HRESULT (including `ERROR_NOT_SUPPORTED`/`E_NOTIMPL`), a null/empty document, or invalid/unparseable JSON is a provider contract violation: the host applies the fail-closed handling above and logs the violation via `Debug::Error` once per provider instance (not once per call).
  - If an unsupported provider API is reached directly despite the host guard, the provider MUST return `HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED)` without mutating provider state.

- **Cross-context (cross-filesystem) copy/move**:
  - If the effective contexts differ, the host MUST NOT silently fall back to passing “foreign” paths to an arbitrary `IFileSystem` instance.
  - Instead, the host MUST either:
    1) Execute the operation via a host-driven **cross-filesystem bridge** (see “Cross-filesystem bridge”), or
    2) Reject the operation and show a localized error (see `IDS_MSG_PANE_OP_REQUIRES_COMPATIBLE_FS`).

### Built-in Provider Capability Matrix (Normative)

`IFileSystem::GetCapabilities()` is mandatory and is the source of truth for host enablement, provider path identity, and same-provider mutation planning. Failed (`ERROR_NOT_SUPPORTED`/`E_NOTIMPL`/other), empty, or invalid capability responses are provider contract violations and host-side rejections (fail-closed, logged once per provider instance). Same-provider API support below applies to both singular and bulk APIs: `CopyItem` + `CopyItems`, `MoveItem` + `MoveItems`, and `DeleteItem` + `DeleteItems`. The `Path identity` column summarizes the mandatory `pathIdentity` profile from `Specs/Plugins/Plugins_VirtualFileSystem.md`; host planners such as Batch Rename MUST use that profile instead of local ad hoc comparison helpers, and same-context Copy/Move/Delete MUST reject missing, malformed, unsupported, or unstable profiles before task creation.

| Provider | Same-provider Copy | Same-provider Move | Same-provider Delete | Cross-FS export copy/move | Cross-FS import copy/move | Advertised concurrency | Path identity | Progress-stream contract |
|----------|--------------------|--------------------|----------------------|---------------------------|---------------------------|------------------------|---------------|--------------------------|
| Local FileSystem (`builtin/file-system`) | yes | yes | yes | `*` / `*` | `*` / `*` | copy/move `1..16` (default `4`), delete `1..64` (default `8`), recycle delete `1..16` (default `2`) | stable, `ordinalIgnoreCase`, `\` preferred, case-only rename supported | Recursive copy/move uses stable nonzero stream IDs for concurrent workers; callbacks for one logical operation are serialized. |
| FileSystemDummy (`builtin/file-system-dummy`) | yes | yes | yes | `*` / `*` | `*` / `*` | copy/move `4`, delete `8`, recycle delete `2` | stable, `ordinalIgnoreCase`, `\` preferred, case-only rename supported | Deterministic offline conformance provider; writable recursive copy/move MUST merge existing destination directories and report item/progress callbacks with stable stream IDs. |
| 7-Zip (`builtin/file-system-7z`) | no | no | no | `*` / none | none / none | `1` / `1` / `1` | stable, `ordinalCaseSensitive`, `/` preferred, mutation not applicable | Same-provider mutation APIs return `ERROR_NOT_SUPPORTED`; export-copy uses the host bridge and host stream IDs. |
| Google Drive (`builtin/file-system-gdrive`) | no | no | no | none / none | none / none | `1` / `1` / `1` | `ordinalCaseSensitive`; `pathTextStableIdentity = false` (duplicate display names) unless paths encode stable item IDs — fail-closed comes from that flag, never from `componentComparison: "unknown"` | File-operation mutation/IO APIs return `ERROR_NOT_SUPPORTED`; no live credentials are required for capability checks. |
| Microsoft Drive (`builtin/file-system-onedrive-personal`, `builtin/file-system-onedrive-business`, `builtin/file-system-sharepoint`) | no | yes | yes | `*` / none | `*` / none | copy/move `1`, delete `4`, recycle delete `1` | stable, `ordinalIgnoreCase`, `/` preferred, case-only rename supported | Same-provider copy is disabled by the host; move/delete progress is provider-owned with copy/move effectively single-stream. |
| Curl FTP (`builtin/file-system-ftp`) | yes | yes | yes | `*` / `*` | `*` / `*` | copy/move `1..8` from plugin settings (default `4`), delete `1..8` (default `4`), recycle delete `1` | stable; `ordinalCaseSensitive` conservative default (the proven instance relation when determinable). MUST be concrete — never `unknown` | Bulk copy/move/delete use bounded provider schedulers; concurrent transfers use stable stream IDs. |
| Curl SFTP/SCP (`builtin/file-system-sftp`, `builtin/file-system-scp`) | yes | yes | yes | `*` / `*` | `*` / `*` | copy/move `1..8` from plugin settings (default `4`), delete `1..8` (default `4`), recycle delete `1` | stable, `ordinalCaseSensitive`, `/` preferred, case-only rename supported | Bulk copy/move/delete use bounded provider schedulers; concurrent transfers use stable stream IDs. |
| Curl IMAP (`builtin/file-system-imap`) | no | no | yes | `*` / `*` | none / none | `1` / `1` / `1` | stable, `ordinalCaseSensitive`, `/` preferred, mutation not applicable | Copy/move are export-only bridge candidates; same-provider copy/move APIs return unsupported. |
| S3 (`builtin/file-system-s3`) | yes | yes | yes | `*` / `*` | `*` / `*` | copy/move `1`, delete `8`, recycle delete `1` | stable, `ordinalCaseSensitive`, `/` preferred, case-only rename supported | Same-provider transfers are effectively single-stream and MAY report stream `0`; delete may run higher provider concurrency. |
| S3 Table (`builtin/file-system-s3table`) | no | no | no | `*` / none | `*` / `*` | `1` / `1` / `1` | stable, `ordinalCaseSensitive`, `/` preferred, mutation not applicable | Same-provider mutations return unsupported; cross-FS import/export is host-bridge mediated. |

Writable recursive providers (Local FileSystem, FileSystemDummy, Curl FTP/SFTP/SCP, S3, and Microsoft Drive where the operation is supported) MUST apply the directory-merge rule from Conflict Handling: an existing destination directory is a merge target, not an overwrite conflict. Providers that cannot satisfy the operation contract MUST advertise it as unsupported and return `ERROR_NOT_SUPPORTED` if called directly.

Provider conformance coverage:

- `--fileops-selftest --selftest-case=FileOps_ProviderCapabilityMatrix` validates the offline capability matrix for Local FileSystem, FileSystemDummy, and 7-Zip, verifies direct 7-Zip unsupported APIs, verifies host rejection before task creation, and checks FileSystemDummy recursive copy/move directory-merge integrity without live network credentials.
- Remote/network provider file-operation smoke coverage remains sandbox-gated under Phase 16 selftests; these tests MUST skip rather than touch live user data when credentials or dedicated selftest roots are unavailable.
- Recursive progress-stream and cancellation contracts are also protected by Phase 7 copy/move parallelism cases and Phase 11 bridge cases.

Object-store / Graph provider declarations (conformance-checked offline):

- S3 declares copy + move + delete (object-store folder merge with per-object conflicts).
- Microsoft Drive declares move + delete but NOT same-provider copy; the host bridges copy, and directory-onto-directory move merges server-side.
- Microsoft Drive Graph mutations that are safe to replay, such as PATCH child moves and DELETE, MAY retry throttled/transient transport responses within the provider's bounded retry policy. Ambiguous POST create operations MUST NOT use blind transport retry; they MUST reconcile by re-querying the parent before deciding whether to replay or fail.
- S3 per-object merge/copy conflict retries MUST be bounded per object and MUST poll cancellation between retry/reprobe/delete steps so a remote prefix operation cannot become an unbounded or uncancelable loop.
- The provider-specific merge/conflict EXECUTION for S3 and Microsoft Drive is validated by code review and the shared host-side merge engine, with deterministic debug Graph/object-store selftests for retry, reconciliation, depth-cap, and type-mismatch cases that can run without live credentials.

### Pause / Cancel

- **Cancel**:
  - Host sets a cancel flag; `IFileSystemCallback::FileSystemShouldCancel` returns `TRUE`.
  - Plugins MUST abort promptly and return `HRESULT_FROM_WIN32(ERROR_CANCELLED)` (or `E_ABORT`, normalized by the host).
- **Pause**:
  - Host pauses by blocking inside progress callbacks (`IFileSystemCallback::FileSystemProgress` and `IFileSystemDirectorySizeCallback::DirectorySizeProgress`) at progress checkpoints.
  - Host MAY also pause “non-primary” tasks via **queue pause** when Wait mode is enabled while multiple tasks are already active.

### Conflict Handling (Normative)

Conflict handling covers per-item failures that require a user decision (overwrite, replace read-only, permissions, etc.).

- The host MUST provide conflict handling for all in-app entry points that can trigger Copy/Move/Delete (keyboard shortcuts, menus/context menus, pane -> pane drag/drop, and Find Files result commands).
- The host MUST NOT silently auto-resolve conflicts by default (no implicit overwrite, replace-readonly, or continue-on-error without user intent).
- FolderView entry points, including clipboard paste and folder-picker move, MUST route Copy/Move/Delete through the File Operations host queue when hosted by `FolderWindow`. If that callback is missing in a normal UI path, the view MUST fail visibly and log a host-wiring error instead of calling the plugin directly with a second set of progress/conflict semantics. Direct plugin fallback is reserved for explicit no-host/test scenarios.
- Find Files result Copy/Move/Delete commands MUST enter this same pipeline with resolved plugin/context/path selections. Copy/Move to other pane MUST infer the source pane from the selected result paths and use the opposite pane destination, so conflicts, progress, cancellation, and completion notifications behave like pane-originated operations.

#### Defaults

- Copy/Move MUST start without allowing overwrite and without allowing replace-readonly (conflicts must surface).
- A retry/re-run Copy/Move MAY suppress an `already exists` prompt only when the existing destination file is byte-identical to the current source. Same size and last-write time alone are insufficient proof and MUST still surface the normal conflict.
- A source directory whose destination path already exists as a directory MUST be merged by recursing into the existing destination directory. Directory-vs-directory existence is not an overwrite conflict; only file-vs-existing-file, file-vs-existing-directory, and directory-vs-existing-file collisions raise the `already exists` conflict. Same-volume Move MUST apply the same merge rule, falling back to copy/delete when the platform rename API cannot move a directory onto an existing directory.
- Directory reparse-point copy under copy-reparse policy MUST treat an existing destination directory as a valid merge target instead of requiring overwrite permission before recreating the reparse point.
- Delete SHOULD start by using Recycle Bin when supported.
- A Delete operation that cannot be guaranteed to use the local Recycle Bin MUST show the permanent-delete confirmation prompt with default Cancel before any file operation task is created. This guard belongs at the host operation boundary, so commands, shortcuts, context menus, drag/drop routing, and future callers cannot bypass it by omitting a separate confirmation flag.
- Continue-on-error MUST be user-driven (via per-conflict Skip/Skip All decisions), not a default behavior.

#### Destructive correctness and partial state

- Built-in local FileSystem delete and overwrite-cleanup paths MUST decide file/directory/reparse status from a handle opened with `FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS`, not only from a prior path attribute snapshot. Reparse points MUST be deleted as links and MUST NOT be recursively traversed by delete or overwrite cleanup.
- Recursive delete implementations MUST re-open each queued child/frame no-follow before deciding whether to recurse. Enumeration attributes are advisory only and MUST NOT be the final destructive decision when another process could have swapped the path.
- Empty-only replacement of an existing real directory by a reparse-point copy MUST use a single non-recursive remove attempt on the destination directory. It MUST NOT pre-scan with `IsDirectoryEmpty` and then perform a separate destructive action, because a concurrent child creation between those steps would turn an empty-only grant into a stale decision.
- Cross-volume Move copy/delete fallback MUST report `HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY)` when the copy succeeded but the source delete or cleanup cannot finish. The completed task diagnostics and Issues pane MUST make the outcome explicit with "source preserved" and "partial copy left" wording so users know both sides may need review.
- Cross-filesystem Move MUST NOT delete a source file unless the destination file has been byte-proven complete. When the source reader reports a size, copied bytes MUST match that size before source deletion is eligible. If the source size cannot be obtained, the bridge MUST either prove the destination against another reliable source-size signal or preserve the source and report `HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY)`.
- Copy-reparse retargeting MAY remap targets that point inside the copied source root into the copied destination root. After remapping, the normalized target MUST still be inside the destination root; otherwise the provider MUST fail the reparse copy instead of creating an escaped destination target.
- Deterministic destructive selftests MUST keep out-of-tree sentinel content and assert it remains untouched after injected file/dir/reparse swaps.

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
- Conflict-routing regression coverage for serialized parallel prompts MUST use mixed per-child decisions in one task (for example Overwrite for some colliding children and Skip for others) and assert the resulting destination content for each child. Repeating the same answer for every prompt is not sufficient evidence that decisions are routed to the intended child.
- Directory merge is the default: copying/moving a folder onto an existing folder of the same name MERGES children rather than raising a top-level `already exists`. When a merged folder has a CHILD that collides, the conflict prompt MUST name the colliding CHILD by its leaf name (e.g. `collide.txt`), not the whole top-level directory.
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

One-shot grant scope:

- **Overwrite** and **Replace read-only** granted from a conflict prompt are ONE-SHOT: they apply only to the single item that was answered. The host MUST clear them at child-loop exit and at each directory-recursion entry, so a grant never leaks to a sibling or descendant item.
- The ONLY way to broaden such a grant to many items is the explicit **Apply to all similar** toggle.
- This holds identically for sequential and parallel execution: per-worker grants MUST NOT leak across items.
- The emitted action set never includes a separate **Skip all** broadening for these grants — **Apply to all** combined with **Skip** covers it.

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
- The host MUST snapshot `fileOperations.crossFsBridgeBufferSizeKB` at task creation and allocate the bridge buffers from that frozen value.
- The host MAY adapt the active bridge buffer size from that default using `IFileSystem::GetTransferHints(...)` on the source and destination endpoints.
- When both endpoints return transfer hints, the host SHOULD use the larger hinted buffer size that still satisfies host-side bounds.

**Single top-level directory copy** under the bridge:

- The host MUST create destination directories **sequentially**.
- The host SHOULD copy the contained files **in parallel** subject to the effective task concurrency budget (and any per-connection caps), so the popup can show multiple in-flight lines even when the user copied only one folder.
- Directory creation MUST NOT require a full tree pre-pass before file transfer admission. As soon as a destination directory has been ensured, file work under that parent SHOULD be admitted to bridge workers, subject to the task concurrency budget and cancellation/pause state.
- When the bridge runs parallel workers, it MUST:
  - use a unique `progressStreamId` per concurrent worker (stable within the operation),
  - serialize `IFileSystemCallback` invocations (no concurrent callback entry).

Per-file conflicts under the bridge:

- When the bridge copies/moves a DIRECTORY across providers and a CHILD collides with a pre-existing destination item, the bridge MUST raise a per-FILE conflict naming that child (not the whole directory). **Skip** preserves the existing child and continues copying non-colliding siblings; the operation then ends as `ERROR_PARTIAL_COPY`.
- The bridge MUST NOT fail the whole directory closed on a single child collision.
- Temp staging files (when the materialization fallback is used) MUST use unpredictable (CSPRNG) names, must be created with exclusive-create/no-follow semantics, and cancellation MUST leave no temp residue. Overwrite/replace grants for the final destination MUST NOT authorize overwriting a temp staging path.

Move semantics under the bridge:

- Cross-filesystem **Move** SHOULD be implemented as Copy + Delete (delete after successful copy).
- For file Move, the bridge MUST revalidate the promoted destination before deleting the source when source-size verification was uncertain or provider I/O can report a successful short transfer. If destination validation fails or is unavailable, the item MUST finish as `ERROR_PARTIAL_COPY` with the source preserved.
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

- Pre-calc runs as part of the same task, before invoking `IFileSystem::*`.
- In Wait mode, pre-calc MUST execute while holding the queue slot (sequential with respect to other queued tasks).
- The host MAY parallelize pre-calc across source roots and recursive child directories within one task, bounded by `fileOperations.preCalcMaxWorkers`. Cancellation and Skip MUST remain responsive while worker fan-out is active.

**Early admission (Copy overlap)**:

- For **Copy** operations, pre-calculation runs CONCURRENTLY with the transfer: bytes begin moving before the recursive size scan finishes, so deep-tree first-byte latency drops. Once the transfer has started (gated on the operation's start tick), the UI MUST show a **Running** status rather than the blocking **Calculating** status; totals and ETA remain `estimating` until pre-calc publishes final totals, which then reconcile in place.
- **Move** and **Delete** MUST keep the SERIAL order (pre-calculation completes before execution) because they modify/remove the source — a concurrent scan would size a tree that is being deleted.
- The serial path retains the cancel-during-pre-calc fast exit (see "Cancel button behavior"). For the Copy overlap, cancellation flows through the transfer's result instead.

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
- Task status glyphs in the client area and the popup caption status glyph MUST render through Direct2D/DirectWrite text paths. They MUST NOT create/select native fonts or use GDI text drawing as a glyph fallback.
- The UI updates on a timer (~100ms) by reading the latest progress snapshot collected via callbacks.
- Displayed speed, ETA, and graph history MUST be timer-driven display estimates, not raw callback-cadence visualizations.
- Callback bursts or brief callback silence SHOULD be smoothed so speed/ETA text and graph samples feel fluid; short silence MAY hold the previous smoothed rate, then decay it when silence continues.
- The graph MUST append samples at popup-display cadence and MUST NOT backfill a long callback interval as a visible trough when progress bytes arrive later in a burst.
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
- A global status summary formatted from the same task-status model as the cards: `N running, M waiting, K need attention`.

The footer MUST NOT expose persistent preferences as always-visible controls. Auto-dismiss for successful or canceled completed tasks is the host-owned `fileOperations.autoDismissSuccess` preference and belongs under `Preferences -> File Operations`.

### Operation cards

Each task card MUST support collapse/expand and MUST adapt to task state:

**Header**
- Every file-operation card MUST resolve exactly one `TaskSnapshot::StatusKind` before rendering. The header label, task status glyph, graph overlay, popup caption severity, and footer summary MUST derive from that resolved status rather than re-running separate state ladders.
- Status precedence is terminal result first, then active conflict, waiting/queued, paused, pre-calc/calculating, preparing, and running.
- During active conflict: `IDS_FILEOPS_STATUS_NEEDS_ATTENTION`
- During pre-calc: `IDS_FILEOPS_CALCULATING` / `IDS_FMT_FILEOPS_CALCULATING_TIME`
- During waiting/queue-paused: `IDS_FILEOPS_GRAPH_WAITING`
- During paused: `IDS_FILEOPS_GRAPH_PAUSED`
- During normal running: `IDS_FMT_FILEOPS_OP_COUNTS` (e.g. `Copy: 3/12`)
  - For Copy/Move, when pre-calc totals (files + folders) are available, the host SHOULD use those totals for the `X/Y` counts; if per-entry completion counts are not available, `X` MAY be estimated from byte progress for display.
- Terminal results use `IDS_FILEOPS_STATUS_COMPLETED`, `IDS_FILEOPS_STATUS_CANCELED`, `IDS_FILEOPS_STATUS_PARTIAL`, or `IDS_FMT_FILEOPS_STATUS_FAILED`.
- A per-task collapse/expand chevron rendered through the shared icon-glyph path (Segoe Fluent Icons preferred, Unicode fallback); operation-card chevrons MUST NOT be hand-drawn with ad hoc stroke geometry.

**Body (expanded)**
- During pre-calc: display the currently accumulated **item totals** (files + folders) and total bytes so far (as they are discovered).
- When pre-calc is skipped (totals unknown), Copy/Move cards SHOULD display a best-effort breakdown of completed **top-level** items by type (files vs folders) when that classification is available.
  - The counts MUST be monotonic.
  - When classification is complete for all top-level items, the host SHOULD keep the breakdown consistent with `completedItems` (`completedFiles + completedFolders == completedItems`).
  - Throughput:
    - Copy/Move: smoothed bytes/sec
    - Delete: smoothed items/sec
- ETA (Copy/Move):
  - Shown only when total bytes are known and the smoothed display speed is > 0.
  - ETA SHOULD be smoothed separately from raw instantaneous rate changes to avoid jitter.
- Paths:
  - Copy/Move: `From:` and `To:` lines
    - Before a Copy/Move task starts, the UI MAY offer a destination selector menu (other panel + history) on the `To:` line.
    - The destination selector menu MUST use the shared DxUI popup-menu contract rather than a native `TrackPopupMenu` surface.
  - Delete: `Deleting:` line
  - During pre-calc, the host SHOULD display the current directory being scanned (from `IFileSystemDirectorySizeCallback::DirectorySizeProgress.currentPath`) when available; otherwise it SHOULD display the first planned source path.
  - Parallel Copy/Move (multi-file in-flight):
    - When a single Task is executing Copy/Move with multiple files actively copying at once (plugin internal parallelism **or host bridge parallelism**), the UI MUST display multiple `From:` file lines with per-file progress (instead of only one “current item”).
    - The UI SHOULD display one stable line per active progress stream (up to max concurrency lines).
      - A stream is identified by `(cookie, progressStreamId)` from `IFileSystemCallback::FileSystemProgress(...)`.
      - A line SHOULD be updated in-place as the stream advances to new items (so completed items are replaced by the next item rather than lingering at `100%`).
      - When a stream completes an item and becomes idle, its line SHOULD disappear promptly (so “done” lines don’t linger at `100%` for long-running tasks).
    - Entries MAY still show `100%` briefly (e.g., end-of-file), optionally with a short grace window (~300ms), before disappearing or advancing to the next item.
- Progress bars:
  - Pre-calc: indeterminate marquee bar
  - Copy/Move: current-item bar (primary) + overall bar (secondary)
  - Delete: overall item bar; MAY additionally show size progress when pre-calc data is available
  - Finished successful tasks MUST render their progress bar as complete even when the last advisory byte counters lagged the completion event.
- Bandwidth graph (Copy/Move):
  - Shows recent throughput history.
  - Samples represent the smoothed display throughput at popup timer cadence.
  - The graph MUST render one colour band per active transfer stream by default in all themes, using a theme-harmonized palette (the Rainbow theme keeps its existing look). Each graph sample MUST attribute its filled area by active progress-stream byte share for that timer bucket. Each stream's band height reflects its cumulative byte share, so equal-rate parallel streams render as visually equal bands within tolerance; the latest progress callback MUST NOT recolor the whole sample by itself.
  - Graph-band fairness coverage MUST drive the rate-history/sampler deterministically when validating band attribution. It MUST NOT depend on popup visibility, `ShowWindow` timing, or live timer cadence unless the test explicitly asserts those prerequisites.
  - Bands engage only once history carries ≥2 concurrent streams; a single-file copy keeps the classic single theme-colored fill.
  - When speed limit is active, shows a horizontal line at the effective limit.
  - Y-axis MUST auto-scale with headroom so the graph remains readable.
  - Overlay text MUST have a drop shadow for visibility against colored graph backgrounds.
  - Overlays:
    - Pre-calc: `IDS_FILEOPS_GRAPH_CALCULATING` + animation
    - Paused: `IDS_FILEOPS_GRAPH_PAUSED` (graph frozen)
    - Waiting/queued: `IDS_FILEOPS_GRAPH_WAITING` (graph frozen)

**Controls**
- During pre-calc:
  - Copy/Move: **Skip**, **Speed Limit**, **Cancel**
  - Delete: **Skip**, **Cancel**
- During operation:
  - Copy/Move: **Pause/Resume**, **Speed Limit**, **Cancel**
  - Delete: **Pause/Resume**, **Cancel**
- Completed tasks:
  - **Dismiss** MUST remain the primary completed-task action.
  - When diagnostics are available, **Show log** and **Export issues** MUST be reachable through one **More...** menu affordance rather than rendered as separate flat buttons beside Dismiss.
- Menus launched from operation-card controls (destination selector, speed limit, conflict **More...**, and completed-task **More...**) MUST use the standard DxUI drop-down menu-button chrome, including a glyph-rendered chevron, and MUST NOT render as split buttons unless the control has a separate primary action beside the menu action. These controls MUST anchor to the invoking button using the shared DxUI popup-menu placement callbacks. These menus MUST use the shared non-modal DxUI context-menu session (`ContextMenu::ShowAsync` or equivalent), not a nested modal menu loop, so the File Operations popup continues to process timers, animation, progress painting, and other owner-window messages while the menu is open. Menu-result callbacks MUST revalidate the target HWND/task before applying changes. A completed-task **More...** menu near the trailing edge MUST open right-aligned above that button, not centered over another task's live graph or progress display.

**Conflict prompts (inline)**
- When a task is blocked on a conflict decision, the popup MUST display an inline prompt associated with that task (not a separate modal dialog).
- The prompt MUST include:
  - bucket-specific message text (localized)
  - the relevant item path(s) (`From` and `To` for Copy/Move; `Deleting` for Delete)
  - an optional compact “All similar” toggle (only for non-retry actions), placed adjacent to the action buttons and clearly visible
  - at most three primary action buttons for the bucket, plus one **More...** menu affordance when additional actions are available
- Primary conflict actions SHOULD be:
  - Existing destination: **Overwrite**, **Skip**, **Cancel**
  - Recycle Bin failed: **Delete permanently**, **Skip**, **Cancel**
  - Transient or unknown failures: **Retry**, **Skip**, **Cancel**
- Rarer available actions such as **Replace read-only** and **Skip All** MUST remain reachable from the **More...** menu instead of expanding the prompt into additional rows of flat buttons.
- While a task is blocked on a prompt, its progress UI SHOULD appear paused/waiting (frozen counters/graph overlays) until the decision is applied.

### Path truncation

When paths do not fit, the UI MUST truncate using a **middle ellipsis** (`…`) so the most important portion remains visible:

- Source line: preserve the file/folder name at the end.
- Destination line: destination is a folder path; do not show a filename.
- When showing a per-file mini progress indicator on the right (parallel in-flight display), the UI MUST reserve space and clip text so the filename/path never renders underneath the progress indicator.

## Completion Event Consumers (Normative)

Consumers such as Compare Directories and Find Files MAY subscribe to file-operation completion notifications, but completion fan-out MUST NOT extend a destroyed consumer's lifetime or call back into invalid UI state. A subscription that targets a window/object with independent lifetime SHOULD include a weak lifetime guard; the multicast dispatcher MUST skip expired guarded subscriptions and remove explicitly unsubscribed tokens.

Find Files result commands that start Move/Delete operations MUST defer row removal until the matching file-operation task completes successfully. Pending removals MUST be keyed by the stable task id, not by operation plus source path alone, so overlapping or repeated operations cannot remove rows for the wrong task. Pending removals SHOULD be age-reaped to avoid unbounded retention after missing or abandoned completion notifications; a failed, cancelled, or partial task MUST leave the original result rows visible.

## Speed Limit (Normative)

### Semantics

- The per-task speed limit applies to Copy/Move via `FileSystemOptions::bandwidthLimitBytesPerSecond`.
- When passing `FileSystemOptions` across the host↔plugin boundary (including via callbacks), the creator MUST set `FileSystemOptions::sizeBytes = sizeof(FileSystemOptions)` and the consumer MUST validate it (treat mismatch as `E_INVALIDARG` in this in-repo ABI sweep).
- `0` means unlimited.
- New copy/move tasks MUST seed their initial desired speed limit from `fileOperations.defaultBandwidthLimitBytesPerSecond` when the caller did not provide an explicit task value.
- Plugins MAY clamp the host-provided limit and report an effective applied limit by writing back to `FileSystemOptions::bandwidthLimitBytesPerSecond` before progress callbacks.
- Presets (bytes/sec):
  - 1 MiB/s, 5 MiB/s, 10 MiB/s, 50 MiB/s, 100 MiB/s, 1 GiB/s
- The speed-limit preset menu MUST use the shared non-modal DxUI popup-menu contract rather than a native `TrackPopupMenu` surface or blocking nested menu loop.
- Copy/Move cards MUST expose the **Speed Limit** menu button during pre-calc/preflight so the user can adjust the per-task limit before transfer bytes start moving.
- `Custom...` opens an owned DirectX prompt surface. It MUST NOT regress to a visible native dialog template or visible child-control fallback.
- The custom speed-limit prompt is task-scoped: it MUST target the selected live/unfinished Copy/Move task, be owned by the file-operations popup, and preserve the surrounding navigation shell focus/ownership when it is opened, canceled, accepted, or reopened.
- Command selftests for this prompt MUST keep the target task observable until the prompt cycle completes so broad all-Commands sweeps exercise the same task-scoped path without racing against an already-completed dummy operation.

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
- The custom speed-limit prompt MUST expose:
  - a DX label describing the field
  - a DX editable value field
  - DX validation text for invalid custom values
  - DX `OK` / `Cancel` command buttons
- The custom speed-limit prompt MUST reject invalid throughput text without closing, keep the validation visible until corrected input is provided, and apply the accepted limit immediately to the target task once the prompt closes successfully.

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
- Status text: `IDS_FILEOPS_STATUS_COMPLETED`, `IDS_FILEOPS_STATUS_CANCELED`, `IDS_FILEOPS_STATUS_PARTIAL`, `IDS_FILEOPS_STATUS_NEEDS_ATTENTION`, `IDS_FMT_FILEOPS_STATUS_FAILED`, `IDS_FMT_FILEOPS_GLOBAL_STATUS_SUMMARY`
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

- Directory-size result caching is intentionally not part of the current contract. A future cache plan MUST include watcher-backed invalidation, short TTLs, stale-total selftests, and proof that cached totals are never used for destructive transfer decisions when freshness cannot be proven.
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

- Filesystem plugin improvements: `Specs/Plans/WIP/FileSystem_PluginImprovementPlan.md`
- Plugin interface definitions: `Common/PlugInterfaces/FileSystem.h`
- VFS contract notes: `Specs/Plugins/Plugins_VirtualFileSystem.md`
- Theme key list: `Specs/Core/Core_SettingsStore.md`

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
    - configurable via `copyMoveMaxConcurrency` (max 16)
    - resolved through `concurrencyMode`: manual uses the configured cap directly; auto resolves from `IFileSystem::GetStorageCharacteristics(...)`
  - `DeleteItems` parallelizes delete with ordering safety across overlapping inputs (children before parents) and supports Recycle Bin deletes with bounded concurrency:
    - default max concurrency: 8
    - default max concurrency (Recycle Bin): 2
    - configurable via `deleteMaxConcurrency` / `deleteRecycleBinMaxConcurrency`
  - Recycle Bin delete batching is supported via `recycleBinBatchSize` (default `500`, configurable `1..1000`), using grouped `IFileOperation::DeleteItems()` batches with single-item fallback for unsupported/error cases.
  - `CopyProgressRoutine` enforces a task-global bandwidth cap across workers (best-effort) and serializes callback delivery (host never sees concurrent callback calls).
  - Delete populates `completedBytes` best-effort using deleted file sizes (to support size progress when pre-calc totals exist).
  - Recursive name-only scan search uses `searchMaxDirectoryWalkers` (default `4`, configurable `1..8`) for the bounded parallel directory walk.
- Batch Rename preview planning is host-owned and MUST produce leaf-only targets before provider mutation. Execution uses `IFileSystem::RenameItems` for same-provider leaf renames, skips no-op rows, revalidates the preview snapshot before mutation, and reports long-running progress through the File Operations informational-task path when needed. Local Batch Rename execution MUST preflight changed sources and external destination conflicts before provider dispatch so a stale preview cannot partially rename earlier rows. Host callers that batch leaf renames SHOULD use `RedSalamander/FileSystemRenameBatch.*` to marshal `FileSystemRenamePair` arrays and strings into one `FileSystemArena` allocation; if `RenameItems` reports unsupported with `E_NOTIMPL`, `ERROR_CALL_NOT_IMPLEMENTED`, or `ERROR_NOT_SUPPORTED`, that helper MAY fall back to one `RenameItem` call per row. Batch Rename executes parent/child selections in deepest-first depth groups before notifying `DirectoryInfoCache::NotifyPathMoved` for successful rows. The built-in local FileSystem plugin's case-only rename handling remains the provider-side responsibility.

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
   - A successful completed delete MUST display as complete even if its final advisory byte counter is below the pre-calc total.

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
- `Specs/FileSystem/FileSystem_FileOperations.md`
- `Specs/Plans/WIP/FileSystem_PluginImprovementPlan.md`

Tasks:

- [x] Clarify in `Common/PlugInterfaces/FileSystem.h` that `overflow==TRUE` means “resync required” (not just “dropped/coalesced”).
- [x] Document callback expectations that affect pause/responsiveness:
  - progress callbacks may be blocked by the host
  - plugins must avoid holding locks that would deadlock if callbacks block
- [x] Align file-op progress semantics across layers:
  - out-of-order `itemIndex` completion is allowed
  - `totalBytes` may be unknown/0 (host pre-calc provides totals when available)
  - delete bytes policy (see Decision #3)
- [x] Update `Specs/FileSystem/FileSystem_FileOperations.md` to explicitly define popup close behavior (Decision #5).

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
- `Specs/FileSystem/FileSystem_FileOperations.md`

Tasks:

- [x] Align popup close semantics with spec (Decision #5).
- [x] When Copy/Move runs with multiple files in-flight, display multiple source lines with per-file progress (MRU; up to max concurrency lines) (Decision #6).
- [x] Ensure per-file source text never renders underneath the right-side mini progress indicator (reserve width + clip).
- [x] Clamp per-file in-flight `completedBytes <= totalBytes` in popup snapshots (avoid “empty” mini bars on overshoot/races).
- [x] Align rainbow bandwidth graph sample colors with progress-bar colors (use the stored per-sample hue + theme value; treat empty path as accent).
- [x] Keep popup rate display fluid by smoothing speed/ETA and sampling graph history on the popup timer instead of raw progress-callback cadence (covered by `--fileops-selftest` Phase 6 smoothing checks).
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
- [x] Replace the native custom speed-limit dialog with an owned DX prompt surface and remove the `IDD_FILEOP_SPEED_LIMIT_CUSTOM` dialog template.

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
  - baseline knob coverage covered by `--fileops-selftest` Phase 7 (`copyMoveMaxConcurrency` 1/4/16 + speed limit toggle + DX custom speed-limit prompt validation + MRU in-flight lines)
  - recursive single-folder coverage: `Phase7_CopyItemsSingleFolderRecursiveParallelism` verifies `CopyItems(count == 1)` and direct `CopyItem(...)` each fan a selected folder with one child directory across multiple progress streams
  - recursive multi-root coverage: `Phase7_CopyItemsMultiRootUnevenRecursiveParallelism` verifies three selected folders with one dominant nested subtree lend spare recursive budget to that dominant subtree while keeping the operation-level transfer limit in force
  - recursive matrix coverage: `Phase7_CopyRecursiveParallelismMatrix` verifies many shallow files, several child directories with one dominant subtree, mixed folder/file selections, copied reparse work items under `CopyReparse`, the spare-budget-zero `nestedConcurrency == 1` edge, forced `ERROR_NOT_SAME_DEVICE` move fallback, an optional real cross-volume move when a second writable fixed volume is available, `continueOnError` partial-copy behavior, and cancellation while workers are active
  - many small files and several large files
  - knob coverage: set FileSystem `copyMoveMaxConcurrency` to `1`, `4`, `16` and verify:
    - throughput scales on many-small-file workloads
    - popup shows multiple in-flight file lines (MRU; up to the configured concurrency / host UI max)
  - slower/virtual provider and bounded-queue coverage: `Phase7_CopyMoveConcurrency16Perf`, `Phase7_SharedPerItemScheduler`, and `Phase11_BridgePipelineDummyToDummyPerf`; future real-device tuning must open a new plan with same-machine `Specs/TestRuns/` evidence before making user-facing performance claims
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
- [x] Recursive Copy/Move: queued files/directories/reparse points, selected-root nested budget, focused single-folder stream count, multi-root dominant-subtree stream count, recursive matrix shape timings/stream counts including copied reparse items and nested concurrency 1, optional real cross-volume move coverage, and debug-forced move fallback counter (`FileOps.CopyRecursiveParallel.*`, `FileOps.CopyItems.NestedConcurrencyBudget`, `FileOps.MoveItems.NestedConcurrencyBudget`, `FileOps.Move.DebugForceCopyFallback`, `FileOps.SelfTest.CopyItemsSingleFolder*`, `FileOps.SelfTest.CopyItemSingleFolder*`, `FileOps.SelfTest.CopyItemsMultiRoot*`, `FileOps.SelfTest.CopyRecursiveMatrix*`).
- [x] Clearflow parallelism/status: pre-calc worker occupancy, bridge streaming admission, single-folder worker occupancy, and one-status popup snapshots (`FileOps.SelfTest.ClearflowPreCalcMultiRootWorkers`, `FileOps.SelfTest.ClearflowPreCalcSingleRootFanOutWorkers`, `FileOps.Bridge.FileAdmissionCount`, `FileOps.Bridge.FileStartedBeforeProducerDone`, `FileOps.SelfTest.BridgeWideShallowEarlyFileStarts`, `FileOps.SelfTest.ClearflowSingleDeepFolderWorkerOccupancy`, `cmd_pane_fileops_conflict_prompt_compacts_actions`).
- [x] Destructive correctness: partial move status, reparse retarget containment, and debug-injected delete TOCTOU swap coverage (`FileOps.SelfTest.CrossVolumeMovePartialFailureStatus`, `FileOps.SelfTest.ReparseRetargetDestinationContainment`, `FileOps.Delete.DebugToctouSwapInjected`, `FileOps.SelfTest.DeleteToctouSwapGuard`).
- [x] Enumeration: peak buffer size, fallback usage, trim events (`FileSystem.DirectoryOps.Enumerate`, `FileSystem.DirectoryOps.TrimBuffer`).
- [x] Debug-only end-to-end self-test runner:
  - run: `.\.build\x64\Debug\RedSalamander.exe --fileops-selftest` (or `--selftest`)
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
  - `Specs/Plugins/Plugins_VirtualFileSystem.md`
  - `Specs/FileSystem/FileSystem_FileOperations.md`
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
  - The properties dialog follows the same ownership/z-order behavior as other application popups; it is not permanently topmost over the main window.
  - `Esc` closes the dialog. `Ctrl+C` copies the full property text for the dialog, and the dialog footer contains a short, unobtrusive hint for that shortcut.
  - Long property values MUST wrap within the value column and grow their row/card so long names and paths remain readable after open and resize.
  - `Alt+Enter` / Properties MUST open the Properties window immediately. If `IFileSystemIO::GetItemProperties` is slow, the dialog shows a loading state with an animated indicator and localized loading text until the provider returns or fails.
  - Common property keys from different plugins MUST be normalized to the same display labels when they carry the same meaning (`name`, `path`, `type`, `sizeBytes`, and standard timestamp keys).
  - When a properties payload contains both `General` and `Timestamps`, the host MUST display `Timestamps` immediately after `General`, regardless of the order supplied by the plugin.
  - Properties views MUST omit unavailable optional metadata instead of displaying placeholder zero values. In particular, timestamp fields whose value is `0` are not shown.
  - The built-in local filesystem General section SHOULD stay focused on name, full path, type, and file size; parent/root/extension rows are omitted because they duplicate those values.
  - Built-in local filesystem file sizes SHOULD use folder-view compact units and include the exact byte count in parentheses for sizes of 1 KB and larger.
  - Built-in local filesystem properties MUST add target metadata for link-like items when available: `.lnk` files expose a `Shortcut` section with `Target`; `.url` files expose an `Internet Shortcut` section with `URL` and a local `Target` when the URL resolves to one; junctions, mount points, and directory symbolic links expose a `Reparse Point` section with `Kind` and `Target`.
  - Built-in local filesystem target metadata MUST be omitted rather than shown as placeholders when target resolution fails or the reparse tag is unsupported.
  - The dialog MUST show an item-stream section only when `GetItemProperties` JSON contains named streams. The stream section lists stream name and size. Each stream row has a View action that opens that stream in ViewerText, where the user can inspect it in text or hex mode. When the active filesystem also exposes `IFileSystemItemStreams`, each removable stream gets an enabled Remove action that deletes the stream and refreshes the dialog; if the refreshed item has no remaining streams, the stream section disappears. Stream Remove actions do not need a tooltip when the visible command text already says `Remove`.

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
