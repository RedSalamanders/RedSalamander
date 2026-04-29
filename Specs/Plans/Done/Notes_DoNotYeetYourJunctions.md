# Do Not Yeet Your Junctions

Critical workflow report + implementation plan for **multi-item copy/move/delete** (recursive) with **transient status**, **failure handling**, and a new **reparse-point policy** defaulting to **“create link/reparse at destination”**.

This spec is intentionally named like this so it’s hard to ignore.

> Status (`2026-02-06`): Phases 1-6 are implemented, plus the post-phase hardening follow-ups (in-app failed-items pane + exported issue-report retention cleanup). Extra hardening: the `followTargets` warning is no longer guarded by process-global state, pre-calc non-cancel failures are logged, diagnostics flush no longer drops entries on `WriteFile` failure, delete progress UI no longer looks "stuck" for single-folder deletes, local pre-calc sizing no longer builds an unbounded in-memory directory queue, and progress accounting clamps completed bytes after skip so summaries never report `completed > total`. Cross-FS bridge capability gaps were also closed: `CreateFileWriter` + `GetFileBasicInformation` now exist for FTP/SFTP/SCP (Curl) and S3, and `GetFileBasicInformation` exists for 7z (read-only).
>
> Closeout audit (`2026-04-25`): Done. The durable reparse-point policy and file-operation behavior are represented in the FileSystem code, selftest coverage, and `Specs/Testing/Testing_TestCoverage.md`; the remaining "next hardening steps" are optional follow-ups, not blockers for this plan.

## Why this workflow is fragile today

RedSalamander’s file ops stack is already sophisticated (queueing, pause/cancel, pre-calculation, per-item concurrency, conflict prompts, cross-filesystem bridge), but **reparse points (symlinks/junctions/mount points)** are currently a correctness cliff:

- **Pre-calc directory sizing** skips reparse points (good).
- **Actual copy/delete recursion** (local FS plugin) does **not** skip reparse points (danger).
- **Cross-filesystem bridge recursion** does **not** skip reparse points (danger).

That mismatch can produce “looks safe” pre-calc/progress UX, followed by an operation that:
- Recurses forever (junction loop),
- Copies far outside the intended selection,
- Deletes contents outside the intended selection.

## Current architecture (short map)

- **Host orchestration**: `RedSalamander/FolderWindow.FileOperations.State.cpp`
  - Task lifecycle, queueing/wait-for-others, pause/cancel, transient progress snapshots
  - Conflict prompting and “apply-to-all” caching
  - Cross-filesystem copy/move via `CrossFileSystemBridge` when plugin IDs differ and capabilities allow
- **Local filesystem implementation**: `Plugins/FileSystem/FileSystem.FileOps.cpp`
  - Recursive copy/delete implemented via `FindFirstFileExW` enumeration
  - Does not special-case `FILE_ATTRIBUTE_REPARSE_POINT` during recursion today
- **Pre-calc sizing**: `Plugins/FileSystem/FileSystem.DirectoryOps.cpp`
  - Skips reparse points explicitly to avoid loops

## Failure modes and UX clarity issues (what to improve)

### 1) Reparse points are a “recursive footgun”
**Severity: critical (data loss / hang).**

If a directory entry is both `DIRECTORY` and `REPARSE_POINT`:
- Recursing into it can escape the selected subtree.
- A junction can point back to an ancestor → infinite recursion.
- Delete recursion can erase unrelated folders (junction target).

### 2) Conflict “Overwrite” can spin if overwrite cannot resolve
**Severity: high (hang / “stuck retry”).**

Host conflict logic retries indefinitely when the cached action is **Overwrite** and the retried call keeps returning an **Exists** bucket. Any implementation that returns `ERROR_ALREADY_EXISTS` even with overwrite enabled risks an infinite retry loop.

### 3) Cross-filesystem bridge lacks safe semantics for reparse points
**Severity: high.**

The bridge can follow a local junction/symlink during directory recursion. Even when “copy as link” is desired, most non-local file systems cannot represent NTFS reparse points, so the bridge must choose a clear and safe behavior.

### 4) “Skip” during pre-calc is not “skip operation”
**Severity: medium (user confusion).**

In the popup, **Skip** during pre-calculation means “skip size counting”, not “skip the job”. This is OK internally, but should be labeled clearly (future UX tweak).

### 5) Completed tasks are removed immediately (no post-mortem)
**Severity: medium (debuggability).**

This was true before the post-mortem retention pass. Completed-task summaries are now retained in popup UI (with dismiss/clear actions) and backed by diagnostics logging.

## New setting: Reparse point policy (default: create link/reparse at destination)

### Where it lives
- **Preferences → Plugins → File System → Configuration**
  - Implemented as a plugin configuration field, so it:
    - Persists under `plugins.configurationByPluginId["builtin/file-system"]`,
    - Automatically appears in Preferences via the existing schema-driven plugin config UI,
    - Is included in the aggregated schema export next to the settings file.

### Key and values
- Key: `reparsePointPolicy`
- Values:
  - `copyReparse` (**default**): treat reparse points as *links* (do not follow); recreate reparse at destination when possible.
  - `followTargets`: treat reparse points like normal files/folders (follow/expand). Dangerous for recursion.
  - `skip`: do not follow; do not copy/move reparse points encountered during recursion.

### Behavior by operation
- **Copy/Move (recursive)**:
  - `copyReparse`: do **not** recurse into directory reparse points; recreate the reparse point at destination (symlink/junction).
  - `followTargets`: current behavior (recurse into directory reparse points).
  - `skip`: ignore reparse points during recursion.
- **Delete (recursive)**:
  - Always treat directory reparse points as *leaf items* (delete the link itself, never the target). This is the safest Windows-file-manager behavior.

### Cross-filesystem bridge behavior (initial safe stance)
- Reparse points cannot generally be preserved across arbitrary virtual file systems.
- The bridge must **not** follow directory reparse points by default (loop/out-of-tree risk).
- Initial behavior: treat directory reparse points as non-recursable entries; do not traverse. (If preservation is unsupported, skip or mark partial in a future iteration with clearer UI.)

## Plan (nice, stupid, and testable)

1. **“Name the Monster”**: add `reparsePointPolicy` to the File System plugin configuration schema (default `copyReparse`).
2. **“Stop Digging Holes”**: update local recursive copy/delete to detect `FILE_ATTRIBUTE_REPARSE_POINT` and avoid traversal.
3. **“Copy The Link, Not The Universe”**: implement best-effort reparse recreation at destination for directory symlinks/junctions.
4. **“Don’t Delete The Neighbor’s House”**: ensure recursive delete never traverses directory reparse points.
5. **“Bridge Troll Toll”**: update CrossFileSystemBridge recursion to not blindly follow directory reparse points.
6. **“Prove It Or It Didn’t Happen”**: add self-test phases for:
   - a junction/symlink loop inside a directory tree (copy must not hang),
   - delete behavior (must delete the link, not the target),
   - overwrite retry sanity (no infinite retry).
7. **“Build It, Don’t Guess”**: run `.\build.ps1` (Debug x64) and record results.

## Tests (a.k.a. “The Infinite Junction of Doom”)

Self-test additions in `RedSalamander/FolderWindow.FileOperations.SelfTest.cpp`:
- **Test A**: Create a directory `src` with a junction `src\\loop` → `src`. Copy `src` recursively.
  - Expected: operation completes; destination contains a reparse point named `loop`; no infinite recursion/hang.
- **Test B**: Create a junction `src\\linkToTarget` → `target`. Delete `src` recursively.
  - Expected: junction removed; `target` still exists with its contents.

## Implementation results (completed)

- [x] Setting added to plugin schema + Preferences UI
- [x] Local FS copy reparse payload handling implemented (in-tree retarget + out-of-tree preserve)
- [x] Protected/localized junctions that deny `List folder / read data` no longer break `copyReparse` (reparse metadata is queried with minimal access)
- [x] Local FS file-reparse copy no longer silently falls back to data-copy dereference
- [x] Local FS delete never traverses directory reparse points
- [x] CrossFileSystemBridge now has explicit policy outcomes:
  - `copyReparse`: unsupported reparse preservation fails with promptable `ERROR_NOT_SUPPORTED`
  - `skip`: skipped reparse entries are tracked and surfaced as partial result (`ERROR_PARTIAL_COPY`)
  - cross-FS **move** with skipped root reparse does not delete source
- [x] Pre-calc root directory reparse points are treated as leaf links (no target traversal in sizing)
- [x] One-time `followTargets` warning prompt before execution
- [x] Overwrite cached modifier retry cap implemented + tested
- [x] Conflict/pre-calc UX wording clarified (`Skip preflight`, `Skip item`, `Skip all similar conflicts`)
- [x] `IFileSystemIO` metadata extension added and implemented in all in-repo plugins
- [x] FileSystemS3 `GetAttributes` is now existence-aware (file vs directory-prefix) so cross-FS copies classify roots correctly and missing items return `ERROR_FILE_NOT_FOUND`
- [x] Completed-task post-mortem retention in popup (dismiss per-task, clear completed in footer)
- [x] Task diagnostics pipeline: bounded in-memory info/warning/error + periodic disk persistence in top-level `Logs/` + log cleanup policy
- [x] Failed-task summaries now always surface diagnostics context (no `Warnings: 0, Errors: 0` on failed result)
- [x] Completed-task cards now expose `Show log` for warning/error outcomes and global `Auto-dismiss success` toggle (default `false`, persisted)
- [x] Conflict prompt now uses a dedicated `Unsupported reparse` bucket/message for non-preservable directory reparse cases
- [x] Completed-task cards now expose `Export issues` for structured per-task warning/error report export
- [x] In-app **Failed Items** pane added (`View` menu + `Ctrl+J`), with theme/rainbow-aware row rendering and persisted/restored window placement
- [x] Exported issue-report file retention cleanup added (`FileOperations-Issues-*.txt`)
- [x] Capabilities-driven per-item concurrency is now parsed from plugin `GetCapabilities()` (`concurrency.*`) across in-repo plugins
- [x] Cross-filesystem bridge concurrency now scales safely when both source/destination advertise support (bounded by host guardrails)
- [x] Recycle Bin delete path migrated from `SHFileOperationW` to `IFileOperation`
- [x] Recycle Bin delete now captures per-item shell failure HRESULT/path (not just aggregate operation HRESULT)
- [x] Pre-calc now runs with bounded parallel workers (up to 4) with deterministic skip/cancel gating
- [x] `GetDirectorySize` file-root semantics normalized across in-repo directory-ops plugins; host Win32 file-size fallback removed
- [x] Self-tests expanded (payload correctness, move reparse scenarios, bridge reparse outcomes, metadata preservation)
- [x] Build passing (Debug x64)

### What changed (code pointers)

- **File System plugin setting**
  - `Plugins/FileSystem/FileSystem.h`: added `reparsePointPolicy` schema field (default `copyReparse`).
  - `Plugins/FileSystem/FileSystem.cpp`: parses `reparsePointPolicy`; persists it in `GetConfiguration()` output.
- **Local FS recursive safety**
  - `Plugins/FileSystem/FileSystem.FileOps.cpp`: copy recursion detects `FILE_ATTRIBUTE_REPARSE_POINT` and:
    - `copyReparse`: recreates directory reparse points at destination without traversal.
      - In-tree link targets retarget into destination tree.
      - Out-of-tree targets are preserved.
      - Unsupported tags return `ERROR_NOT_SUPPORTED`.
    - `skip`: does not traverse / does not copy those entries.
    - `followTargets`: old behavior (traverse).
    - Note: directory reparse metadata is read using `FILE_READ_ATTRIBUTES` (not `GENERIC_READ`) so protected junctions (e.g. localized system junctions under `C:\Program Files`) can still be preserved as links.
  - `Plugins/FileSystem/FileSystem.FileOps.cpp`: file reparse copy now returns explicit failure when link-copy is unsupported (no silent normal-copy fallback).
  - `Plugins/FileSystem/FileSystem.FileOps.cpp`: delete recursion treats **directory reparse points as leaf items** (never traverses them).
  - `Plugins/FileSystem/FileSystem.DirectoryOps.cpp`: root directory reparse points in pre-calc are treated as leaf links (status `S_OK`, zero-byte sizing).
- **Cross-filesystem bridge safety**
  - `RedSalamander/FolderWindow.FileOperations.State.cpp`: bridge policy semantics are explicit:
    - `copyReparse`: directory reparse preservation across bridge returns `ERROR_NOT_SUPPORTED` (promptable).
    - `skip`: directory reparse entries are skipped and counted; operation returns partial.
    - cross-FS move avoids source delete when bridge skip occurred.
- **Host conflict/UX hardening**
  - `RedSalamander/FolderWindow.FileOperations.State.cpp`: cached modifier retry cap (1 attempt per bucket per item) prevents overwrite spin loops.
  - `RedSalamander/FolderWindow.FileOperations.State.cpp`: one-time `followTargets` warning before recursive copy/move.
  - `RedSalamander/FolderWindow.FileOperations.State.cpp`: `UnsupportedReparse` conflict bucket classification for directory reparse `ERROR_NOT_SUPPORTED` paths (including cross-FS bridge unsupported-reparse hints), with retry disabled for that bucket.
  - `RedSalamander/FolderWindow.FileOperations.Popup.cpp`, `RedSalamander/RedSalamander.rc`, `RedSalamander/Resource.h`: apply/skip wording clarity updates.
- **Cross-FS metadata/timestamps**
  - `Common/PlugInterfaces/FileSystem.h`: added `FileSystemBasicInformation` + `GetFileBasicInformation` / `SetFileBasicInformation`.
  - `RedSalamander/FolderWindow.FileOperations.State.cpp`: bridge copies metadata best-effort after writer commit.
  - Implemented across plugins:
    - `Plugins/FileSystem/FileSystem.cpp` (real Win32 metadata read/write),
    - `Plugins/FileSystemDummy/FileSystemDummy.cpp` (in-memory metadata read/write),
    - `Plugins/FileSystem7z/FileSystem7z.cpp` / `Plugins/FileSystemCurl/FileSystemCurl.DirectoryOps.cpp` / `Plugins/FileSystemS3/FileSystemS3.IO.cpp` (explicit unsupported return).
  - `Plugins/FileSystemS3/FileSystemS3.IO.cpp`: `GetAttributes` now probes exact-object vs prefix existence to return correct file/directory attributes (and proper not-found errors).
- **Capabilities-driven concurrency (Phase 4)**
  - `Common/PlugInterfaces/FileSystem.h`: documents optional `concurrency` capabilities object (`copyMoveMax`, `deleteMax`, `deleteRecycleBinMax`).
  - `RedSalamander/FolderWindow.FileOperations.State.cpp`: `DeterminePerItemMaxConcurrency(...)` now reads plugin `GetCapabilities()` JSON and defaults to `1` when absent/invalid.
  - `RedSalamander/FolderWindow.FileOperations.State.cpp`: bridge no longer hard-forces concurrency `1`; it now uses `min(sourceCap, destinationCap)` with host bounds.
  - Plugin capabilities updated:
    - `Plugins/FileSystem/FileSystem.cpp`, `Plugins/FileSystem/FileSystem.h`, `Plugins/FileSystem/FileSystem.Watch.cpp` (dynamic capabilities from current config),
    - `Plugins/FileSystemDummy/FileSystemDummy.h`,
    - `Plugins/FileSystem7z/FileSystem7z.h`,
    - `Plugins/FileSystemCurl/FileSystemCurl.h`,
    - `Plugins/FileSystemS3/FileSystemS3.h`.
- **Modernization (Phase 5)**
  - `Plugins/FileSystem/FileSystem.FileOps.cpp`: `DeleteToRecycleBin(...)` now uses `IFileOperation` + `IShellItem` with non-interactive shell flags and abort mapping.
    - Adds a per-item `IFileOperationProgressSink` to capture specific failed-item HRESULT/path and stream delete progress back to the host (current path + item counts). Uses `UpdateProgress(workTotal, workSoFar)` so progress advances even during long shell “prepare/enumerate” phases where per-item callbacks may be delayed.
    - Returns item-level failure HRESULT when available instead of only aggregate operation status.
  - `RedSalamander/FolderWindow.FileOperations.State.cpp`: recycle-bin bucket failures are recorded as item diagnostics (`delete.recycleBin.item`) before conflict resolution.
  - `RedSalamander/FolderWindow.FileOperations.State.cpp`: pre-calc now supports bounded parallel scanning (up to 4 workers) using per-worker progress cookies and shared aggregation.
    - skip/cancel/pause remain deterministic,
    - late worker updates are suppressed once skip/cancel is observed,
    - final pre-calc completion/skipped state is finalized exactly once.
  - `RedSalamander/FolderWindow.FileOperations.State.cpp`: removed host-side Win32 fallback for `GetDirectorySize(...)=ERROR_DIRECTORY` file roots.
  - File-root `GetDirectorySize` contract now implemented in in-repo directory-ops plugins:
    - `Plugins/FileSystem/FileSystem.DirectoryOps.cpp`,
    - `Plugins/FileSystemDummy/FileSystemDummy.cpp`,
    - `Plugins/FileSystem7z/FileSystem7z.cpp`,
    - `Plugins/FileSystemCurl/FileSystemCurl.DirectoryOps.cpp`.
  - `Common/PlugInterfaces/FileSystem.h`: `GetDirectorySize` contract now explicitly documents file-root `S_OK` behavior (`fileCount=1`, `directoryCount=0`).
- **Post-mortem + diagnostics**
  - `RedSalamander/FolderWindow.FileOperations.State.cpp`:
    - keeps a bounded completed-task summary list independent from live task lifetime,
    - records per-task warning/error diagnostics with bounded in-memory retention,
    - captures bounded per-task structured issue entries (`warning`/`error`) for completed summaries,
    - supports `info` + `debug` diagnostics levels for non-problem events (verbosity controlled by settings),
    - flushes diagnostics every ~5s (or on completion) as **JSONL** (one JSON object per line) to `Logs/FileOperations-YYYYMMDD.jsonl` (sibling of `Settings` / `Crashes`),
      - fields include: `ts`, `level`, `task`, `op`, `category`, `hr`, `hrName`, `hrText`, `memWorkingSetBytes`, `memPrivateBytes`, `message`, `src`/`dst`, `srcLeaf`/`dstLeaf`,
      - `src`/`dst` are best-effort and prefer the most specific in-flight item paths (so conflict/error diagnostics identify the exact file/folder when possible),
    - cleans up old diagnostics log files (retains newest 14),
    - synthesizes minimal summary diagnostics for failed results that had no detailed entries (`0/0` guard),
    - enriches failed-task diagnostics with operation, HRESULT text, progress counters, and latest source/destination paths,
    - resolves and opens the freshest available diagnostics log for `Show log`,
    - exports per-task structured issue reports to `Logs/FileOperations-Issues-Task<ID>-YYYYMMDD-HHMMSSmmm.txt` (includes an explicit `Status text` column),
    - applies retention cleanup to exported issue-report files.
  - `RedSalamander/FolderWindow.FileOperationsInternal.h`: new state contracts for completed summaries, diagnostics recording, and dismiss APIs.
  - `RedSalamander/FolderWindow.FileOperations.IssuesPane.cpp` + `.h`: new modeless failed-items pane with list view of warning/error diagnostics from completed tasks.
    - `View -> File Operations Failed Items` + default shortcut `Ctrl+J`
    - row colors respect theme/rainbow mode
    - pane is an unowned (normal) top-level window so it does not stay above the main window
    - list view header and rows are themed, and row separators only draw for actual items (no "empty grid" look)
    - includes a `Status text` column to show the system error text for HRESULTs
    - pane size/position/state persist under `settings.windows["FileOperationsIssuesPane"]` and are normalized to stay visible on current displays.
  - `RedSalamander/FolderWindow.FileOperations.Popup.cpp` + `.h`: popup now renders completed cards with result/warning/error context and `Dismiss` action, adds `Show log` + `Export issues` for diagnostic outcomes, includes a theme-colored title-bar status glyph (check/warn/error), persists window placement under `settings.windows["FileOperationsPopup"]` (normalized to stay visible), and includes a global `Auto-dismiss success` footer checkbox; footer switches to `Clear completed` when no active operations.
  - `RedSalamander/FolderWindow.FileOperations.cpp`: success auto-dismiss is applied when toggle is enabled; failed-items pane toggle is exposed to app command routing.
  - `RedSalamander/RedSalamander.cpp`, `RedSalamander/CommandRegistry.cpp`, `RedSalamander/ShortcutDefaults.cpp`, `RedSalamander/RedSalamander.rc`, `RedSalamander/Resource.h`: menu/command/shortcut/resource wiring for failed-items pane.
- **Self-test coverage**
  - `RedSalamander/FolderWindow.FileOperations.SelfTest.cpp`: added `Phase12_ReparsePointPolicy`
    - Copy test validates not only reparse tag presence but reparse **target payload correctness**:
      - in-tree loop retargets to destination tree,
      - out-of-tree target remains external target.
    - Copy test also applies a "protected junction" ACL (deny `List folder / read data` to `Everyone`) on the source loop link to match Windows localized/system junction behavior; reparse metadata reads must succeed with minimal access.
    - Local move reparse test verifies move preserves link target semantics and does not traverse.
    - Bridge move (`skip`) test verifies partial result + source preservation.
    - Bridge copy (`copyReparse`) test verifies unsupported flow reaches conflict prompt, classifies into `UnsupportedReparse`, and can be skipped to partial.
    - Delete test verifies deleting tree containing junction removes link but preserves junction target.
  - `RedSalamander/FolderWindow.FileOperations.SelfTest.cpp`: `Phase9_ConflictPrompt_OverwriteAutoCap` validates overwrite retry cap behavior.
  - `RedSalamander/FolderWindow.FileOperations.SelfTest.cpp`: `Phase10_PermanentDeleteWithValidation` now additionally validates:
    - local + dummy file-root `GetDirectorySize` return `S_OK` with correct file-root counters,
    - recycle-bin locked-file delete fails with a specific non-generic HRESULT.
  - `RedSalamander/FolderWindow.FileOperations.SelfTest.cpp`: `Phase11_CrossFileSystemBridge` now validates timestamp metadata preservation on bridge copy.
  - `RedSalamander/FolderWindow.FileOperations.SelfTest.cpp`: `Phase11_CrossFileSystemBridge` now validates bridge per-item concurrency (>1 in-flight observed, outputs validated).
  - `RedSalamander/FolderWindow.FileOperations.SelfTest.cpp`: `Phase13_PostMortemDiagnostics` validates:
    - failed summaries do not regress to `Warnings: 0, Errors: 0`,
    - persisted log-file creation,
    - global `Auto-dismiss success` behavior (`true` auto-removes successful completed cards, `false` retains them),
    - per-task `Export issues` report generation (path exists and non-empty),
    - self-test restoration of the previous auto-dismiss setting value.

### Build/verify notes

- Built on `2026-02-07` with: `.\build.ps1 -Configuration Debug`

## Verification commands

- Build: `.\build.ps1` (Debug, x64)
- Self-test (Debug): `.\.build\x64\Debug\RedSalamander.exe --fileops-selftest` (expects exit code `0`, closes itself when done; can take a couple minutes)

## Resume later checklist (so this doesn’t rot)

- **Setting/UI**: Preferences → Plugins → File System → Configuration → `Reparse points (symlinks/junctions)`
  - Default: `Create link/reparse at destination` (`copyReparse`)
  - Other values: `followTargets`, `skip`
- **Settings JSON**: `plugins.configurationByPluginId["builtin/file-system"].reparsePointPolicy`
  - Example: `{"reparsePointPolicy":"copyReparse"}`
- **Global completed-task behavior**:
- Popup footer checkbox: `Auto-dismiss success` (default off)
  - Settings JSON path: `fileOperations.autoDismissSuccess`
  - Example: `{"fileOperations":{"autoDismissSuccess":false}}`
- **Diagnostics retention setting (UI exposed)**:
  - Preferences → Advanced → `File Operations` → `Max diagnostics log files`
  - Settings JSON path: `fileOperations.maxDiagnosticsLogFiles`
  - Example: `{"fileOperations":{"maxDiagnosticsLogFiles":30}}`
- **Diagnostics verbosity settings (UI exposed)**:
  - Preferences → Advanced → `File Operations` → `Diagnostics: Info` / `Diagnostics: Debug`
  - Settings JSON paths: `fileOperations.diagnosticsInfoEnabled`, `fileOperations.diagnosticsDebugEnabled`
  - Example: `{"fileOperations":{"diagnosticsInfoEnabled":true,"diagnosticsDebugEnabled":false}}`
- **Diagnostics retention settings (JSON-only advanced overrides)**:
  - `fileOperations.maxIssueReportFiles`
  - `fileOperations.maxDiagnosticsInMemory`
  - `fileOperations.maxDiagnosticsPerFlush`
  - `fileOperations.diagnosticsFlushIntervalMs`
  - `fileOperations.diagnosticsCleanupIntervalMs`
- **Failed-items pane**:
  - View menu: `File Operations Failed Items`
  - Default shortcut: `Ctrl+J`
  - Window placement key: `windows["FileOperationsIssuesPane"]`
- **File-operations popup**:
  - Window placement key: `windows["FileOperationsPopup"]`
- **What the self-test proves**: no junction-loop hang on copy; delete removes link (not target)

## Known limitations / follow-ups (if we keep hardening)

- **Non-(symlink|junction) directory reparse tags** currently return `ERROR_NOT_SUPPORTED` under `copyReparse` (consider “skip + report” UX).
- **Diagnostics UI vs advanced overrides**: Preferences exposes `maxDiagnosticsLogFiles` + verbosity toggles (`diagnosticsInfoEnabled`, `diagnosticsDebugEnabled`); `maxIssueReportFiles` and timing/buffer knobs remain JSON-only advanced overrides.
- **S3 "directories" are virtual**: `CreateDirectory()` is a no-op; copying empty directories into S3 via bridge may not preserve the empty directory (no marker object is created).
- **Curl/S3 bridge upload progress is staged**: destination writers may buffer/stage and do remote upload work during `Commit()`. Bridge progress reflects bytes written to the writer, so `Commit()` can still take time; for small files the UI switches the per-item bar to indeterminate during `Commit()` to avoid appearing stuck at 100%.
- **S3 GetDirectorySize is best-effort** and can be slow/expensive for large prefixes; sizing `/` (all buckets) is intentionally `ERROR_NOT_SUPPORTED`.
- **7z remains read-only**: archive mutation (write) is intentionally not supported; cross-FS export is the supported path.

---

## Plan refresh (2026-02-06): global implementation plan

This section supersedes ambiguity in earlier notes and locks behavior before further coding.

### Confirmed decisions

- For **in-tree links** under `copyReparse`, destination links must retarget into the **destination tree** (not back to source).
- For **unsupported directory reparse tags** under `copyReparse`, return an error that lands in conflict UX (user can Skip/Cancel).
- `followTargets` requires an explicit **one-time warning** before first execution.
- Interface extension for metadata/timestamps is acceptable even if ABI-breaking, provided all in-repo plugins are updated together.
- No migration/backward-compat layer is required for the new global file-operations settings; `fileOperations.*` is the canonical path.

### Normative behavior matrix

| Operation | Item type | Policy | Local FS | Cross-FS bridge | Expected UX / result |
|---|---|---|---|---|---|
| Copy/Move | Directory reparse (in-tree target) | `copyReparse` | recreate link, retarget into destination subtree, no traversal | unsupported for preserving link | local: success; bridge: promptable unsupported |
| Copy/Move | Directory reparse (out-of-tree target) | `copyReparse` | recreate link, preserve external target, no traversal | unsupported for preserving link | local: success; bridge: promptable unsupported |
| Copy/Move | Directory reparse (unsupported tag) | `copyReparse` | no traversal | no traversal | conflict prompt (Skip/Skip all/Cancel) |
| Copy/Move | Directory reparse | `followTargets` | traverse target | traverse target | one-time warning required before run |
| Copy | Directory reparse encountered in subtree | `skip` | skip reparse entry, continue tree copy | skip reparse entry, continue copy | task result partial (`ERROR_PARTIAL_COPY`) |
| Move | Root directory reparse | `skip` | move link item normally (no traversal) | skip root item (do not copy, do not delete source) | task result partial (`ERROR_PARTIAL_COPY`) |
| Move | Subtree directory reparse encountered | `skip` | skip link entry during recursive processing | skip link entry; source delete suppressed for safety | task result partial (`ERROR_PARTIAL_COPY`) |
| Delete (recursive) | Directory reparse | any | delete link itself, never traverse target | delete link itself, never traverse target | safe leaf delete semantics |
| Copy/Move | File reparse | `copyReparse` | copy-as-link only; explicit error on unsupported/privilege/policy failures | plugin-dependent | no silent unsafe fallback |
| Copy/Move | File reparse | `followTargets` | normal file semantics | normal bridge stream semantics | success/failure per normal file ops |
| Copy/Move | File reparse | `skip` | skip item | skip item | visible skipped/partial outcome |

### UX contract (conflict clarity)

- `Apply selected action to all similar conflicts` means: cache for the same conflict bucket.
- `Skip item` means: skip current item only.
- `Skip all similar conflicts` means: cache skip for that bucket for the rest of task.
- Pre-calc action text must explicitly say `Skip preflight` to avoid confusion with operation skip.
- Bridge/internal reparse skips must surface as partial/skipped outcomes (not silent).

### Overwrite retry-cap acceptance rule

- Cached modifier actions (`Overwrite`, `Replace read-only`, `Permanently delete`) may auto-retry at most **once per bucket per item**.
- If still unresolved, stop auto-retrying, clear cached action for that bucket, and prompt again.
- This prevents silent infinite loops while preserving user control.

### Why `IFileOperation` recycle-bin migration is still medium risk (and how to de-risk)

- It is not inherently blocked by current centralized file-op threads.
- Risk is behavior drift (shell callback cadence, cancellation semantics, error mapping), not thread model incompatibility.
- De-risk path:
  - run behind feature flag or staged switch,
  - match current `HRESULT` mapping and conflict buckets,
  - verify cancel/pause responsiveness and no UI-thread blocking.

### Why parallel pre-calc conflicts with current skip/cancel flow, and solution

- Current pre-calc assumes a mostly linear sequence with immediate skip/cancel observability.
- Naive parallel scans can keep workers running after user hit skip, causing stale updates and confusing totals.
- Proposed safe design:
  - bounded worker pool (2-4),
  - shared cancellation token + callback checks per worker,
  - atomic aggregation with monotonic clamping for totals,
  - stop accepting late worker updates once skip/cancel is observed,
  - deterministic finalization step that marks pre-calc complete/skipped exactly once.

### Metadata/timestamp preservation (cross-FS)

- Implemented: CrossFileSystemBridge now uses `IFileSystemIO::{Get,Set}FileBasicInformation` after file writer `Commit()`.
- Source-side support (in-repo):
  - Local FS + Dummy: full basic info.
  - Curl (FTP/SFTP/SCP), S3, 7z: file `lastWriteTime` is reported and used for `creation/lastAccess` fallback. If a plugin cannot provide a non-zero `lastWriteTime`, it returns `ERROR_NOT_SUPPORTED` to avoid propagating 1601 timestamps into Win32 destinations.
- Destination-side support:
  - Local FS + Dummy implement `SetFileBasicInformation`.
  - Curl/S3/7z return `ERROR_NOT_SUPPORTED` (no attempt is made to force remote timestamps/attrs).

### Cross-plugin routing table (native vs bridge)

Terminology:

- "Same plugin": source + destination panes have the same plugin ID (example: Local FS -> Local FS, S3 -> S3).
- "Cross-plugin": panes have different plugin IDs (example: Local FS -> S3, FTP -> Local FS, FTP -> SFTP).

Host routing (high level):

| User operation | Same plugin | Cross-plugin (different plugin IDs) |
|---|---|---|
| Pre-calc (size scan) | plugin `GetDirectorySize()` on each source item | plugin `GetDirectorySize()` on each source item (destination is irrelevant) |
| Copy/Move | plugin-native `IFileSystem::{CopyItems,MoveItems}` | host `CrossFileSystemBridge` (per-item mode only) |
| Delete | plugin-native `IFileSystem::DeleteItems` | n/a (delete always targets the source plugin only) |

CrossFileSystemBridge requirements (copy/move):

| Bridge step | Source requirements | Destination requirements | Notes |
|---|---|---|---|
| Enumerate directories | `IFileSystem::ReadDirectoryInfo()` | n/a | Used for recursive traversal. |
| Copy file bytes | `IFileSystemIO::CreateFileReader()` | `IFileSystemIO::CreateFileWriter()` | Reader/writer are host-driven; `Commit()` finalizes. |
| Create directories | n/a | `IFileSystemDirectoryOperations::CreateDirectory()` | Required for recursive copy/move. |
| Preserve timestamps/attrs (best-effort) | `IFileSystemIO::GetFileBasicInformation()` | `IFileSystemIO::SetFileBasicInformation()` | Non-fatal: failures are warnings. |
| Move: delete source after successful copy | `IFileSystem::DeleteItems()` | n/a | Suppressed when the task had skips/partials to avoid data loss. |

In-repo plugin bridge support matrix (what exists today):

| Plugin IDs | Bridge as source | Bridge as destination | Writer semantics when destination | Directory semantics when destination |
|---|---|---|---|---|
| `builtin/file-system` (Local FS) | yes | yes | direct streaming to destination path | real directories created |
| `builtin/file-system-dummy` (Dummy) | yes | yes | direct in-memory writer | in-memory directories created |
| `builtin/file-system-ftp`, `builtin/file-system-sftp`, `builtin/file-system-scp` (Curl) | yes | yes | staged: writes to a temp file, uploads on `Commit()` | creates remote directories (best-effort; protocol dependent) |
| `builtin/file-system-s3`, `builtin/file-system-s3table` (S3) | yes | yes | staged: writes to a temp file, uploads on `Commit()` | "directories" are virtual; `CreateDirectory()` is a no-op |
| `builtin/file-system-7z` (7z) | yes | no (read-only) | n/a | n/a |

Why Curl/S3 `CreateFileWriter()` stages to a temp file (instead of pure streaming through RAM):

- It keeps the `IFileWriter` contract simple: sequential `Write()` then `Commit()` (abort is "release without commit").
- It allows robust cancel/cleanup behavior without needing protocol-specific abort/resume plumbing in the bridge.
- It avoids unbounded memory use on large files; the bridge never buffers whole files in RAM.
- Trade-off: the progress bar can reach 100% at the end of the staging write, then "sit" during the final upload (tail latency).

### Phased execution order

1. **Safety + UX now**: reparse semantics completion, one-time follow-target warning, overwrite retry cap, explicit skip/partial reporting.
2. **Test hardening now**: reparse payload target correctness, move + cross-FS reparse tests, deterministic overwrite-cap test.
3. **Host/plugin contract next**: metadata/timestamp API extension + plugin implementations.
4. **Performance next**: capabilities-driven concurrency policy (replace plugin-ID gate), then safe bridge concurrency expansion.
5. **Modernization later**: recycle-bin `IFileOperation` migration, parallel pre-calc rollout, additional refactors.
6. **Post-mortem observability now**: completed-task retention + bounded diagnostics logging persisted to disk with cleanup.

### Execution status (`2026-02-06`)

- [x] Phase 1 complete
- [x] Phase 2 complete
- [x] Phase 3 complete
- [x] Phase 4 complete
- [x] Phase 5 complete
- [x] Phase 6 complete

### Next hardening steps

1. Optional UX refinement: add per-column sort and copy-selected-rows actions in the failed-items pane.
2. Optional diagnostics UX refinement: expose `maxIssueReportFiles` in Preferences Advanced (currently JSON-only).
3. Continue non-behavioral cleanup items (duplication removal, separator heuristic polish, optional arena growth) after current behavior stabilizes.
