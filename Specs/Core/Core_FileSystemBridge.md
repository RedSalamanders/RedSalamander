# Cross-File-System Bridge Specification

## Overview

The cross-file-system bridge is the host-side engine that copies or moves data between two **different** `IFileSystem` plugin instances (e.g. local disk ↔ S3, 7z archive → local, FTP ↔ OneDrive). A plugin's native `CopyItem`/`MoveItem` can only operate inside its own backend; when source and destination live in different plugins (or different connection contexts of the same plugin), no single plugin can perform the transfer. The bridge fills that gap: the host pumps file bytes itself through the `IFileReader`/`IFileWriter` streaming contract (`Common/PlugInterfaces/FileSystem.h`), staging each file to a temporary destination name and promoting it only after the byte count checks out.

**Design goals:**
- No partial or corrupt file is ever visible at the final destination path (staged temp + rename-promote).
- Cross-FS MOVE never deletes a source file unless the destination copy has been re-verified.
- Backend-agnostic: the bridge depends only on `IFileSystemIO` (`CreateFileReader`/`CreateFileWriter`), `GetTransferHints`, and the regular `IFileSystem` operations (`MoveItem` for promote, `DeleteItem` for cleanup, `ReadDirectoryInfo` for tree walks).
- Adaptive throughput: buffer size from plugin transfer hints, double-buffered pipelining for large files, per-connection concurrency budgets, live bandwidth limiting.

Implementation lives in `RedSalamander/FolderWindow.FileOperations.State.cpp` (`CrossFileSystemBridge`, ~:6794-9279). This document describes the code as it is.

## When the Bridge Engages

Two gates, both of which MUST pass:

1. **UI-time capability gate** (`CanCrossFileSystemCopyMove`, `FolderWindow.FileOperations.cpp:287`). When the two panes differ in plugin id **or** instance context, a COPY/MOVE is only offered if:
   - Both plugins publish valid `GetCapabilities` JSON (v1 with all four sections `operations`/`concurrency`/`crossFileSystem`/`pathIdentity`; any violation is **fail-closed** — the operation is disabled).
   - Source has `operations.read`, destination has `operations.write`; MOVE additionally requires source `operations.delete`.
   - The source's `crossFileSystem.export.copy|move` id list allows the destination plugin AND the destination's `crossFileSystem.import.copy|move` list allows the source plugin (`*` wildcard supported).
2. **Task-time switch** (`State.cpp:6459`): `useCrossFileSystemBridge = (_destinationFileSystem != nullptr) && (COPY || MOVE)`. The command handler passes a destination file system only when contexts differ. Both sides MUST answer the `IFileSystemIO` QI, otherwise the task fails with `ERROR_NOT_SUPPORTED`.

When the gate does not fire, the host delegates to the source plugin's native `CopyItem`/`MoveItem`. The bridge only runs in `ExecutionMode::PerItem`; there is no batch (`CopyItems`) bridge path — each top-level item gets its own `CrossFileSystemBridge` instance and `CopyPath` call.

Note: not every plugin participates. GoogleDrive does not implement `IFileSystemIO` at all and declares empty `crossFileSystem` lists (bridge fully disabled both directions). 7z is export-copy only (read-only backend, `CreateFileWriter` returns `ERROR_NOT_SUPPORTED`). MicrosoftDrive allows cross-FS copy but not move.

## Data Flow — One File

`CopyFileWithBuffer` runs this sequence per file:

```
  1 acquire per-connection copy/move permits (both endpoints, ascending-id order)
  2 destination overwrite policy check (prompt BEFORE any bytes move)
  3 CreateFileReader(source); snapshot source basic information (best effort)
  4 reader->GetSize(); failure is fatal (ERROR_PARTIAL_COPY, nothing written)
  5 query destination IFileSystemAtomicWriter for this exact final path
  6 create either a staged sibling writer or an explicitly atomic final-path writer
  7 if the writer exposes IFileWriterExpectedSize, provide the source size before Write
  8 pump bytes while computing the copy-time FNV-1a hash
  9 require written bytes == source size; Commit; promote staged output when applicable
 10 re-open final destination and require its size == source size for COPY and MOVE
 11 COPY: hash final destination and require it to match; mismatch deletes the bad final
 12 SetFileBasicInformation (best effort); record size/hash/source snapshot in the manifest
```

The bridge never infers atomic-final behavior from staging-name syntax. `IFileSystemAtomicWriter::SupportsAtomicWriterCommit` is the only opt-in; otherwise the writer receives a CSPRNG-named sibling and the destination plugin's `MoveItem` promotes it. S3 opts in because multipart completion is its atomic publication point.

**Serial pump:** loop of `Read(buffer, bufferBytes)` → inner write loop that tolerates short writes (`Write` returning 0 bytes → `ERROR_WRITE_FAULT`); `Read` returning 0 bytes is EOF. Short reads are valid. Successful provider counts are untrusted: `bytesRead > bytesToRead` or `bytesWritten > bytesToWrite` fails with `ERROR_INVALID_DATA` before counters, hashes, or buffer offsets consume the value.

**Pipelined pump:** selected when `fileTotalBytes > bufferBytes`. Exactly two buffer slots; one `std::jthread` reader fills slots round-robin, the calling thread drains them as the writer, synchronized by one mutex + one condition variable. Allocation/thread-start failure degrades to the serial pump. A writer-side stop sets the pipeline stop flag, wakes task pause waiters and the pipeline condition variable, and then joins, so a reader parked in the host pause gate cannot strand shutdown.

Both pumps poll cancellation and pause per chunk, update atomic task-wide counters with overflow checks, and apply bandwidth delays outside the throttle-state mutex. Progress cadence is resolved from transfer hints (200 ms minimum). Progress callbacks are serialized and round-trip `FileSystemOptions`; a failed callback aborts the copy.

## Directory Trees

`CopyPath` (`:9214`) dispatches per top-level item:

- Root and child reparse points, including reparse-point files, honor `ReparsePointPolicy`: Skip records/skips the entry, FollowTargets dereferences it, and CopyReparse records `ERROR_NOT_SUPPORTED`. The bridge never silently materializes a reparse target under Skip/CopyReparse.
- Every source-supplied child component is validated before any source or destination path join. Empty names, `.`, `..`, separators, embedded NUL/C0/DEL, duplicates, and destination-invalid components are rejected per item. Windows-style destinations additionally reject `:*?"<>|`, trailing dot/space, DOS device names, components over 255 UTF-16 code units, and case-insensitive sibling collisions. Invalid entries produce `ERROR_INVALID_NAME`; valid siblings may continue and the directory result becomes `ERROR_PARTIAL_COPY`.
- Directories go through `EnsureDestinationDirectory` (`:7232`; 3-attempt probe/create loop, dest-is-file collision prompts then deletes) and then either:
  - **Sequential** (`CopyDirectorySequential` :8649): one `ReadDirectoryInfo` enumeration, direct recursion per subdirectory. An empty directory — plugin returns `(nullptr, S_OK)` — is success.
  - **Parallel** (`CopyDirectoryParallel` :8807): the calling thread produces work items from an explicit stack (creating directories itself), files flow through an admission queue capped at 16 entries with condition-variable backpressure; workers from the process-wide `PerItemTaskScheduler` each allocate a private `bufferBytes` buffer and run `CopyFileWithBuffer`. Chosen when `ComputeWithinFolderBudget()` (task budget divided by active top-level items) exceeds 1.
- `FILESYSTEM_FLAG_CONTINUE_ON_ERROR` makes child failures non-fatal; the tree then finishes with `ERROR_PARTIAL_COPY`. Any per-file conflict Skip also makes the item result `ERROR_PARTIAL_COPY`.
- For endpoints advertising `FILESYSTEM_TRANSFER_HINT_HIGH_METADATA_COST`, the host suppresses the separate pre-calculation traversal. The copy walk remains authoritative; COPY lists each source directory once and MOVE lists it once for copy plus once for cleanup.

## Buffering & Adaptive Sizing

- Setting: `fileOperations.crossFsBridgeBufferSizeKB`, default 4096 KB, clamped 512–16384 KB at load (`Common/SettingsStore.cpp:3307`), snapshotted into the task at start (`State.Runtime.Part.cpp:684`).
- Provider hints remain active even when the user changed the configured buffer setting. The source and destination are queried for `FileSystemTransferHints`; the maximum nonzero preferred buffer is clamped to 512 KiB–16 MiB and wins. Without an explicit buffer hint, WAN latency or `PREFERS_LARGE_BUFFERS` raises the configured fallback to at least 8 MiB.
- The bridge merges the maximum latency class, ORs flags, and uses the maximum requested progress period. `HIGH_METADATA_COST` also controls pre-calculation as described above.
- The resolved tuning is frozen for the task and emitted in perf details; there is no mid-flight retune. Curl advertises 1 MiB to match its ring; local/7z/cloud providers retain provider-appropriate hints.
- Host bridge-buffer reservations are governed by one cancellation-aware process-wide 256 MiB budget. Each pipelined file reserves two buffers; directory workers release idle primary reservations and acquire their own. Provider-internal buffers are separate from this host budget, but the S3 multipart writer has its own process-wide four-part/256 MiB admission budget. Therefore the bridge plus S3 writer buffers have a documented combined ceiling of 512 MiB, excluding SDK/HTTP implementation overhead.

## Integrity Guarantees & MOVE Semantics

**Guaranteed by the host for all backends:**
- A source whose size cannot be read is never copied (fatal pre-check).
- The writer sees the source size before its first write when it exposes `IFileWriterExpectedSize`. This lets cloud writers stream without whole-file local spooling while preserving ABI compatibility for older writers.
- The final destination is accepted only after exact pump byte count, successful `Commit`, and a fresh final-path size check. COPY additionally re-hashes the final destination against the FNV-1a accumulated while bytes were copied. A bad COPY final is best-effort deleted and the item fails `ERROR_PARTIAL_COPY`.
- Overwrite/replace-readonly grants are stripped from ordinary temp writers and re-applied at promote. An explicitly atomic final writer receives only the grants established by the same pre-transfer conflict decision.
- MOVE source deletion remains a separate pass after the whole tree copied. COPY never records this cleanup manifest. A file is deleted only when it can be extracted from the MOVE manifest, the source basic-info snapshot still matches, destination size matches, and a destination-only hash matches the copy-time hash. Each manifest node is erased before its cleanup decision continues, so retained proof memory decreases through the cleanup walk and reaches zero even when a changed source shell is preserved. The cleanup phase does not reopen and re-read the source. Directories are removed only after verified children; changing only a source root's basic information preserves that root shell, deletes still-matching copied children, and yields `ERROR_PARTIAL_COPY`.
- The native local plugin takes the strong-consistency branch permitted by the MOVE contract. A successful `CopyFileEx` (or completed staged `CopyFile`) is the byte-copy proof; the plugin records no-follow source and destination snapshots, requires the source snapshot to remain stable across the copy, and revalidates both snapshots before deletion. It performs no post-copy content reread. The existing-destination resume path is the exceptional case: because this invocation did not copy the bytes, it hashes both files once before accepting them as identical.

**Backend-dependent (NOT guaranteed by the host):**

| Backend | Reader short reads | GetSize freshness | Abort-safety of writer | Notes |
|---|---|---|---|---|
| Local (Win32) | only at EOF | open-time (file opened share-write/delete) | SAFE — own temp-sibling staging + `FlushFileBuffers` on Commit | durable commit |
| 7z | routine (streaming modes) | archive-index-time | n/a (read-only) | decode/CRC failures are terminal and remain latched across reads/seeks |
| Curl (FTP/SFTP/HTTP) | fills up to the 1 MiB ring or EOF | targeted SIZE/stat probe, otherwise unknown/fatal | final path protected by host staging | unknown size is never fabricated as zero; post-replace backup cleanup is warning-only |
| S3 | fills requests; exact range and `Content-Range` required | HEAD snapshot, or ranged-GET total when HEAD is denied; pinned by versionId or ETag | SAFE — multipart completion publishes atomically; abort on release | async one-part overlap; final-key atomic writer removes bridge CopyObject promotion |
| MicrosoftDrive | fills requested range up to 16 MiB | metadata snapshot pinned by ETag | SAFE — expected-size upload session streams during Write | cross-FS move disabled; no whole-file `%TEMP%` spool for known nonempty sizes |
| GoogleDrive | — | — | — | no `IFileSystemIO`; bridge disabled |
| Dummy | never short | exact | SAFE (RAM) | test backend |

S3 and MicrosoftDrive range requests pin the object revision; a same-size concurrent replacement fails closed instead of producing a torn stream. Timestamps are copied best-effort (`SetFileBasicInformation`, warning-only). Reparse payloads, ACLs, alternate data streams, and extended attributes are never transferred.

## Failure Handling & Cleanup

- A `wil::scope_exit` armed when a staged writer is created deletes its temp on every pre-promote failure. An atomic-final writer has no host temp. A COPY that fails final size/hash verification best-effort deletes the bad final. Cleanup failure is diagnostic and cannot rewrite the already-determined transfer status.
- Cancellation (`task._cancelled` or jthread stop token) is polled per chunk in both pumps, in throttle sleeps, in condition-variable predicates, and by the progress callback. Cancel mid-file → temp deleted, source untouched. Cancel mid-MOVE-cleanup after some deletes is logged (`bridge.move.cleanup.cancelled`); already-verified-and-deleted sources stay deleted.
- Transient MOVE-cleanup probe errors (sharing violation, busy, timeout, network) are retried 3× with backoff.
- A circuit breaker wraps whole top-level bridge calls (never chunks), keyed by connection-profile GUIDs. Its transient vocabulary includes the HRESULTs actually emitted by Curl/S3/network providers (`ERROR_TIMEOUT`, `ERROR_SEM_TIMEOUT`, connection abort/refusal, `ERROR_BAD_NET_RESP`, `ERROR_UNEXP_NET_ERR`, and related network failures). Cancellation/auth failures do not count; local disk is exempt.

## Concurrency Model

- One `std::jthread` per task; per-item fan-out via the process-wide `PerItemTaskScheduler` (≤16 workers, round-robin fairness across tasks). When a scheduler worker waits for a nested job, it helps execute that target job; saturating all workers with directory producers cannot deadlock nested file work.
- Bridge parallelism budget = `min(source budget, destination budget)`, each derived from the plugin's capabilities JSON (`concurrency.copyMoveMax`) or, in Auto mode, from `GetStorageCharacteristics().preferredCopyMoveConcurrency`; capped at 16 and further clamped by per-connection-profile overrides.
- Per-connection copy/move semaphore permits (`ConnectionConcurrencyLimiter`) bound concurrent transfers per profile across all tasks; dual-connection transfers acquire permits in ascending connection-id order to avoid deadlock.
- Within one file, at most two threads touch data (pipelined reader + calling-thread writer), never the same stream concurrently. Shared perf counters are atomic. Progress callbacks and `FileSystemOptions` snapshots use the callback mutex; parallel workers use stable stream IDs.
- Conflict prompts are serialized: one prompt per task, other workers wait in `WaitWhilePaused`; cacheable actions use the `All similar` toggle, and skip-all behavior is represented only by `Skip` plus that toggle. The first producer/worker failure is published before stop and is preferred over synthetic `ERROR_CANCELLED` unless cancellation was actually requested.

## Test Hooks (`ENABLE_TESTS` builds only; absent from production)

- The Causeway bridge I/O decorator can independently inject source over-report, premature EOF, writer under-consumption/over-report, destination size faults, and file-reparse metadata. Successful read/write counts remain pass-through by default.
- Hostile source enumeration injects separators, traversal, embedded NUL/control, ADS syntax, DOS device names, and case collisions while retaining a valid sibling.
- `FailNextBridgeFileCopyForSelfTest` injects inside the shared `CopyFileWithBuffer` worker path so both sequential and parallel failure propagation are covered.
- `SetFileOpsBridgePipelineModeForSelfTest` forces the pipelined pump on/off.
- Pause points cover MOVE cleanup, MOVE-manifest extraction, and writer-failure/paused-reader shutdown; all hooks have bounded bailout. MOVE manifest peak/current counters support deterministic shrink-to-zero assertions, and runtime perf rows emit peak and remaining entry counts.
- Env-var-driven destination mutators: create a file at a pending destination-directory path (CreateDirectory race), or rewrite a promoted destination before MOVE cleanup (corruption-detection tests).
- Scheduler saturation, buffer-budget peak, high-metadata enumeration counts, normalized transient HRESULTs, and transfer-hint resolution are observable deterministically.
- Provider-local debug tests cover MicrosoftDrive range/request-count and streaming upload, S3 version pinning/async multipart/atomic capability, Curl unknown-size/read-fill/EOF-tail/cleanup semantics, and 7z persistent corrupt-entry failure.
- Coverage lives in `RedSalamander/SelfTest/FileOperations/` (Causeway/Fairstream/Phase 11 families) plus `PluginContractTests`. New cases MUST be registered in `kFileOpsFamilyDefinitions` or Full-suite runs skip them silently.

## Known Limitations

- Reparse points cannot be copied cross-FS (Skip or FollowTargets only).
- Metadata transfer is timestamps/attributes best-effort only; no ACLs, ADS, or sparse/compressed semantics.
- `GetSize` remains mandatory. Curl tries a targeted SIZE/stat probe; a genuinely sizeless source still fails closed rather than streaming to an unverifiable final.
- Files not larger than one buffer (default 4 MiB) always use the serial pump.
- Bridge COPY destination hashing adds one destination read pass. Bridge MOVE cleanup adds one destination hash pass but never re-reads the source. Native local cross-volume MOVE normally adds zero content-read passes after `CopyFileEx` because stable source/destination snapshots are its strong-consistency proof.
- Host pump buffers are capped at 256 MiB process-wide. S3 multipart payload buffers are independently capped at 256 MiB process-wide, for a 512 MiB combined application-owned transfer-buffer ceiling; provider SDK/HTTP internals are outside this accounting.

## Security Posture

The host is the trust boundary between remote/plugin namespaces and the destination filesystem. It validates every length-delimited child component before path construction, distrusts provider byte counts, pins cloud object revisions, verifies final COPY content, and preserves MOVE sources unless the destination-only proof matches the copy-time hash. These checks are host obligations and do not depend on third-party plugins sanitizing names or reporting well-formed counts.

The remaining deliberate limits are fail-closed: unknown source size, unsupported reparse preservation, unprobeable final destination, or revision mismatch prevents successful completion. See `Specs/Reviews/FileSystemBridge-2026-07-07-Findings.md` and the completed Causeway plan for item-level rationale and evidence.
