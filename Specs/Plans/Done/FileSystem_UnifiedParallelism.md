# Unified File-Op Parallelism Model (Copy/Move/Delete) + Connection Overrides

## Progress
- [x] Add Connection Manager overrides (`copyMoveMaxConcurrency`, `deleteMaxConcurrency`) to schema + UI
- [x] Add `Dummy` protocol to Connection Manager UI
- [x] Add Dummy `/@conn:<name>` resolution in `FileSystemDummy`
- [x] Host: apply per-connection caps to per-item concurrency
- [x] Host: CrossFileSystemBridge single-folder parallel copy (multi stream IDs, serialized callbacks)
- [x] Host: global per-connection gate for bridge transfers
- [x] Curl: parse overrides + enforce per-connection global cap for copy/move/delete
- [x] Local: parallel delete within a single directory (recursive)
- [x] S3: update spec + concurrency behavior (delete)
- [x] Selftest: bridge single+multi-folder parallelism + override clamp + global gate tests
- [x] New baseline run under `Specs/TestRuns/4cb089111a23/FileOps/2026-02-28_100446/`
- [x] (Optional) Add `Tools/CompareTestRuns.ps1` and record diffs (compare: `2026-02-27_141202` → `2026-02-28_000800`)

---

## Latest artifacts

- FileOps run: `Specs/TestRuns/4cb089111a23/FileOps/2026-02-28_100446/`
- Compare vs `Specs/TestRuns/4cb089111a23/FileOps/2026-02-28_000800/`:
  - added cases: `Phase11_BridgeMultiFolderParallelCopyInFlightLines`, `Phase11_ConnectionOverrideGlobalGate`
  - Δduration_ms: `+66985`

## Model (slots / budget)

The goal is one consistent behavior across:
- Host per-item scheduling (`ExecutionMode::PerItem`)
- Cross-filesystem bridge (host-driven `IFileReader` → `IFileWriter`)
- Filesystem plugins (local + remote)

### Copy/Move

- **Single top-level folder**: destination folder creation is sequential; **file transfers inside the folder run in parallel**.
- **Multiple items**: parallelize across top-level items. If top-level parallelism doesn’t fully use the configured budget, use remaining budget for within-folder parallelism.

### Delete

- **Single top-level folder**: traversal is sequential; deletion of contained items uses parallelism when allowed by budget (post-order for directories).
- **Multiple items**: parallelize across top-level items; apply the same budget rules and global per-connection caps.

### Progress streams

- Parallel work MUST use **unique `progressStreamId` per concurrent worker** (within a single top-level item) so the popup can show multiple in-flight lines.
- Callback invocations MUST be **serialized** (even when work is parallel).

---

## Concurrency precedence rules (locked)

### Copy/Move

`effectiveCopyMoveMax = min(hostUiMax, sourcePluginCap.copyMoveMax, destPluginCap.copyMoveMax (if bridged), connectionOverride(copyMove) if nonzero)`

### Delete

`effectiveDeleteMax = min(hostUiMax, pluginCap.deleteMax (or deleteRecycleBinMax), connectionOverride(delete) if nonzero)`

Where:
- `hostUiMax` is `Task::kMaxInFlightFiles` (8).
- Plugin caps come from `IFileSystem::GetCapabilities().concurrency.*`.
- Connection overrides come from `ConnectionProfile.extra`.

**Global rule:** Connection overrides apply **globally per connection profile across the whole app** (so concurrent tasks don’t exceed a picky server’s cap).

---

## Implementation notes (by component)

### Settings / Connection Manager

- Schema: add known `ConnectionProfile.extra` keys:
  - `copyMoveMaxConcurrency` (0 = inherit, clamp 1..8)
  - `deleteMaxConcurrency` (0 = inherit, clamp 1..64)
- UI: add numeric fields and persist/remove keys (`0` removes key to keep JSON clean).
- Add `Dummy` protocol entry to the Connection Manager list.

### Dummy plugin

- Accept `IHost*` at creation time and store `IHostConnections`.
- Support `/@conn:<name>/...` by fetching connection JSON and resolving to `initialPath + suffix`.

### Host (FolderWindow file operations)

- Keep `_perItemMaxConcurrency` for **top-level scheduling**, but preserve a separate **budget** value for within-folder bridge parallelism.
- Apply per-connection overrides to the effective budget by scanning `/@conn:` in source paths and destination folder.
- Implement a process-wide per-connection limiter keyed by `ConnectionProfile.id` (GUID string) for bridge transfers.
- Cross-filesystem bridge: implement a parallel directory copy path that:
  - enumerates sequentially and creates directories,
  - enqueues file copy work items,
  - runs workers with unique `progressStreamId`,
  - serializes callbacks,
  - honors continue-on-error via `ERROR_PARTIAL_COPY`.

### FileSystemCurl (FTP/SFTP/SCP/IMAP)

- Parse per-connection overrides from `ConnectionProfile.extra`.
- Enforce a global per-connection limiter keyed by `connectionId` (or stable endpoint key when not using Connection Manager).
- Add `deleteMaxConcurrency` plugin setting + capabilities.
- Implement parallel recursive delete when effective concurrency > 1.

### Local filesystem plugin

- Implement parallel recursive delete for a single directory (post-order) using the shared job scheduler.

### S3

- Update delete concurrency cap to allow host per-item scheduling for many-object deletes.
- Update spec to reflect delete support and any limitations.

---

## Acceptance criteria

- FTP/SFTP→local single-folder copy shows multiple in-flight file progress lines and transfers are concurrent.
- Setting connection `copyMoveMaxConcurrency=1` forces truly sequential transfers even with many files.
- Debug `--fileops-selftest` passes with the new tests.
- Specs are updated and consistent with implementation.

---

## Test artifacts

| Profile | Area | Where | Notes |
|---|---|---|---|
| Tiny | FileOps | `Specs/TestRuns/Tiny/FileOps/<timestamp>/` | Hardware-dependent timings |
| 4cb089111a23 | FileOps | `Specs/TestRuns/4cb089111a23/FileOps/<timestamp>/` | Baseline for current machine |
