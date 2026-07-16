# Operation Causeway — Cross-File-System Bridge Security & Performance Remediation

**Status:** Complete — implemented, validated, and closed 2026-07-13
**Date:** 2026-07-07
**Execution review:** 2026-07-13 at `ff3f62572` on `codex/causeway-closeout`
**Scope owner:** the host cross-filesystem copy/move engine (`CrossFileSystemBridge`,
`RedSalamander/FolderWindow.FileOperations.State.cpp` ~:6794-9279), the `IFileReader`/`IFileWriter`
plugin contract (`Common/PlugInterfaces/FileSystem.h`), and the reader/writer paths of the seven
filesystem plugins.
**Companion docs:** spec `Specs/Core/Core_FileSystemBridge.md`; full itemized findings
`Specs/Reviews/FileSystemBridge-2026-07-07-Findings.md`.
**Provenance:** multi-agent review (4 mappers, 6 review dimensions, adversarial verification — 61
agents). 40 findings raised, 38 confirmed, 2 refuted. CW-1 (path traversal) was additionally
verified end-to-end by hand (S3 key / FTP `LIST` name → `JoinFolderAndLeaf` → local
`MakeAbsolutePath` collapse).

## 2026-07-13 execution-readiness audit

The current source was re-read against every row and the companion findings ledger before implementation.
All Causeway runtime findings remain present. The original ledger is executable with these corrections and
scope clarifications:

- **CW-1 trust-boundary scope is broader than the original examples.** The host must reject embedded NULs,
  C0 controls, and DEL in addition to empty/dot names and separators. Otherwise a length-delimited
  `FileInfo::FileName` can be truncated when later passed as a NUL-terminated plugin path. Structural
  component validation applies at every bridge enumeration consumer (pre-calculation, copy, and MOVE
  cleanup), while destination-specific Windows rules and case-collision detection apply when the destination
  declares `ordinalIgnoreCase`/the built-in local identity. This is tracked as **CW-1A** and closes with CW-1.
- **CW-2 has a symmetric writer-output contract gap.** A destination writer can report
  `bytesWritten > bytesToWrite`; although this does not expose bytes beyond the host buffer in the same way as
  an over-reported read, it corrupts progress/integrity accounting and can overflow counters. This is tracked
  as **CW-2A** and closes with the same provider-output clamps and contract tests.
- The CW-1 text's old “Fold in **CW-15**” reference meant **review finding 15** (foreign-name/NTFS
  collisions), not plan row CW-15 (redundant tree enumeration). The row below now says this explicitly.
- **CW-T1 does not reuse FIR-1/FG-A1.** FIR-1 is specifically the destination-reader `GetSize` seam.
  Premature EOF, over-reported reads, writer under-consumption, writer failure, and cleanup-reader failure are
  Causeway-owned decorator modes and must stay independently armable.
- **CW-9 has one implementation owner here.** Firebreak 6.2 and the archived Clearwater CW-P1/CW-T-P1
  rows are routing/evidence ledgers. Causeway implements the shared copy-time hash once, updates those ledgers,
  and must not introduce a second checksum scheme.
- **CW-16 must preserve the bridge's atomic-commit rule.** It may use an explicit optional destination-writer
  capability/extension for “commit atomically to final path”; it must never infer the final key by parsing a
  `.rs_tmp_` filename or weaken overwrite/conflict handling.
- The requested TSan checks in CW-22/CW-23 are not available in the supported Windows/MSVC toolchain.
  Completion proof is deterministic concurrent stress plus code inspection/source-contract coverage proving
  atomic or mutex ownership; ASan remains required for memory-safety paths.
- The pre-change deterministic bridge baseline is archived at
  `Specs/TestRuns/4cb089111a23/FileOps/2026-07-13_084840` (Debug, same-machine diagnostic baseline).
  Final throughput claims still require a test-enabled Release candidate run and the scenario-specific
  before/after records required below.

## 2026-07-13 completion matrix

The detailed tables below preserve the review's intake state. This final matrix is the authoritative
execution state for closeout.

| Final state | IDs | Closure |
|-------------|-----|---------|
| [x] | CW-1, CW-1A | Host validation rejects structurally hostile provider names before every join, applies Windows destination rules and case-collision checks, reports `ERROR_INVALID_NAME`, and never echoes the hostile name in diagnostics. |
| [x] | CW-2, CW-2A | Both bridge pumps reject reader and writer counts outside the caller-owned buffer; the bounds are part of the public interface contract. |
| [x] | CW-3 | A worker waiting for a nested scheduler job participates in dequeue, cancellation wakes waiters, and worker-count saturation completes deterministically. |
| [x] | CW-4 | Curl preserves unknown size as `ERROR_NOT_SUPPORTED`, probes the single object when possible, fills caller reads, and returns the buffered EOF tail. |
| [x] | CW-5, CW-6, CW-7 | S3 and Graph ranged reads pin object identity; `412`/`416` fail as revision mismatches; COPY and MOVE verify final size and COPY verifies the copy-time hash before accepting the promoted destination. |
| [x] | CW-8 | Reparse-point files and roots now pass through the same fail-closed policy as reparse directories. |
| [x] | CW-9 | Bridge MOVE reuses its copy-time FNV-1a and reads only the destination at cleanup. Native cross-volume MOVE uses successful copy plus stable no-follow source/destination snapshots as zero-reread byte proof; only an exceptional pre-existing-destination resume hashes content. |
| [x] | CW-10, CW-11 | Graph writers stream expected-size content directly to an upload session, abort unfinished sessions, and readers issue one pinned range up to the requested buffer (including the covered 8 MiB request). |
| [x] | CW-12, CW-14, CW-16 | S3 reuses destination validation, asynchronously overlaps one multipart upload with pumping, and exposes an explicit atomic-final-writer capability so the no-conflict path avoids temp-key promotion. |
| [x] | CW-13, CW-15, CW-21, CW-26 | Throttle sleeps occur outside the mutex; high-metadata walks suppress redundant pre-scan; all transfer hints remain active with non-default settings; host and S3 payload budgets are independently capped at 256 MiB (512 MiB combined application-owned ceiling, excluding SDK internals). |
| [x] | CW-17–CW-20 | 7z terminal decode/CRC failures remain latched; pipeline teardown wakes paused readers; Curl post-replacement cleanup is success-with-warning; Curl's preferred buffer matches and fills its 1 MiB ring. |
| [x] | CW-22–CW-25 | Shared perf/options state has atomic or mutex ownership, the first real worker error wins over synthetic cancellation, and transient classification matches provider HRESULTs. |
| [x] | CW-T1 | Independently armable hostile-provider decorators cover early EOF, count violations, writer failure/under-consumption, and cleanup-reader failure; all cases are registered in the FileOps family. |
| [>] | CW-X1 | Intentionally remains routed to the Compare owner. `Specs/Core/Core_CompareDirectories.md` now records that decision I/O failure is distinct from `Different`; no Compare runtime change is claimed by Causeway. |

## Additional findings discovered and closed during implementation

| ID | Finding and resolution |
|----|------------------------|
| CW-4A | Curl discarded a buffered EOF tail after the transport signalled end-of-stream; `Read` now returns the tail before clean EOF. |
| CW-5A | S3 did not map HTTP `412`/`416` into the bridge's revision-mismatch vocabulary; both now fail closed as `ERROR_REVISION_MISMATCH`. |
| CW-9A | The first native MOVE proof revision still re-read both files after copying; it was replaced with stable metadata snapshots and zero post-copy content reads. |
| CW-9B | The pre-existing-identical resume path hashed twice; its single comparison now also records the byte count used by cleanup proof. |
| CW-10A | Abandoned Graph upload sessions were not explicitly cancelled; uncommitted writers now issue best-effort session deletion. |
| CW-17A | A 7z reader could publish the declared final byte before the extractor published a terminal CRC error; final-byte reads now await extraction completion. |
| CW-26A | S3 multipart buffers initially sat outside the host budget; a separate four-slot 256 MiB process budget now bounds them and the authoritative spec records the 512 MiB combined ceiling. |

## Validation and performance evidence

| Evidence | Result |
|----------|--------|
| Debug build | Full test-enabled build completed with 0 warnings and 0 errors; `.build/logs/msbuild-20260713_120237_256.log`. |
| Optimized build | Full test-enabled Release build completed with 0 errors; `.build/logs/msbuild-20260713_120600_127.log`. Its 9 warnings are pre-existing `ViewerPETests` unused-symbol diagnostics outside Causeway. |
| Causeway Release selftests | 7 passed, 0 failed, 0 skipped; `Specs/TestRuns/4cb089111a23/FileOps/2026-07-13_120925`. |
| Floodgate cross-FS Release guards | 6 passed, 0 failed, 0 skipped; `Specs/TestRuns/4cb089111a23/FileOps/2026-07-13_121001`. |
| Full Debug FileOperations suite | 108 passed, 0 failed, 20 conditional skips; `Specs/TestRuns/4cb089111a23/FileOps/2026-07-13_115548`. |
| Provider contracts | FileSystem 42/0, 7z 13/0, MicrosoftDrive 75/0, S3 128/0, Curl 76/0; all plugin enumeration/schema/capability checks also passed. |
| Memory and metadata metrics | COPY/MOVE bridge peak 128 MiB against the 256 MiB host ceiling; source listings exactly 2/2 for COPY and 4/4 for MOVE; `Specs/TestRuns/4cb089111a23/FileOps/2026-07-13_111014`. |
| Pipeline before/after | Same-machine Release deterministic scenario: 130,766,000 us baseline versus 422,000 us candidate for four 32 MiB files with 30 ms chunk latency; `Specs/TestRuns/4cb089111a23/FileOps/2026-07-13_110340`. |
| Sanitizer | Test-enabled ASan build completed with 0 warnings/errors and focused provider contracts passed 3/3; `.build/logs/msbuild-20260713_105540_725.log`. |
| Archive contract | The 18 files in the final Release Causeway/Floodgate evidence archives pass `Get-RSTestRunArchiveViolations` with repository limits. `Test-TestRunArchive.ps1` itself rejects an empty changed-path array when all archives are already ignored/clean; that tooling defect is outside Causeway and did not prevent explicit validation. |

The supported Windows/MSVC toolchain has no TSan. The concurrency acceptance proof therefore uses
deterministic scheduler saturation, paused-reader teardown, async upload overlap tests, and source inspection
of every shared field's atomic/mutex ownership, plus ASan for memory safety.

Live cloud and SMB credentials/endpoints were not available. No live multi-GB wall-clock claim is made.
The acceptance evidence instead uses test-enabled Release scenarios and deterministic fake transports that
measure exact range count, API count, byte re-read count, upload overlap, listing count, and memory ceilings;
these are the behavior contracts affected by the implementation and are repeatable without external service
variance.

The broad diagnostic Full suite run
`.build/TestSandbox/runs/20260713T092222Z-79864-75feb16f0fde4a2db38e25b722f59583`
completed 1,132 passed, 4 failed, 54 skipped. All Causeway-owned tests and provider contracts were green.
The four failures were unrelated existing stabilization items: Compare Search
Service legacy `auto_vacuum` status timing, Commands navigation shell focus timing,
`RedSalamanderMonitorEtwLatency`, and the subsequent deployment overwrite of its still-locked executable.
They remain owned by the active `HOLD-I19` stabilization workstream and are not concealed as Causeway passes.

## Single-owner notes (do not duplicate work already owned elsewhere)

- **CW-9** (MOVE cleanup re-reads every byte to verify) overlaps the *intent* of Clearwater
  **CW-P1/CW-T-P1** (add a copy-time FNV-1a/Crc32 into `CopiedEntry`, re-hash only the destination
  at cleanup). If that lands first, CW-9 collapses to "confirm the full dual-reader `memcmp`
  (`State.cpp:7396`) is gone." Coordinate; do not implement a second hashing scheme.
- **CW-4** (Curl reader-path sizeless `GetSize`) is a *distinct, still-open* gap from Floodgate
  `FG-P0-3-FALLBACK-SIZE-ZERO`, which fixed only the **native staged-copy** path
  (`CurlProbeRemoteFileSize`, `FileSystemCurl.Shared.cpp:2879`). The `IFileSystemIO`
  reader the bridge uses (`CreateFileReader` → `CurlStreamingReader::GetSize`) was never wired to it.
- **CW-6/CW-7** (plain COPY has no post-promote verify; eventual-consistency re-stat) extend Floodgate
  `FG-P0-2-RESTAT-FALSEPOS`, which is scoped to MOVE only. Keep the MOVE re-stat work in Floodgate;
  Causeway owns the **COPY-side** gap.
- Destination-side `GetSize` fault injection is owned by `Operation_FileOperations_FaultInjectionRatchet`
  (**FIR-1 / FG-A1**); Causeway's test rows reuse that seam, they do not re-build it.

## P0 — Security (data-security holes; fix first)

| State | ID | Work item | Required proof |
|-------|----|-----------|----------------|
| [ ] | CW-1 | **CRITICAL Zip-Slip / path traversal.** Cross-FS tree copy joins the *source-plugin-supplied* child name into the local destination path filtering only exact `.`/`..` (`State.cpp:8723-8729` sequential, `:9083-9089` parallel; validator `TryGetValidatedFileInfoName:1933` is buffer-bounds only, `JoinFolderAndLeaf:2407` appends verbatim). Embedded `\`,`/`, and `..\..\` survive; the local dest plugin's `MakeAbsolutePath` (`Plugins/FileSystem/FileSystem.Path.cpp:234`) *collapses* dot-segments (own self-test `:340`), so the write escapes the chosen folder. Reachable from a malicious/compromised S3 bucket (key leaf split only on `/`, `FileSystemS3.S3.cpp:399-402`) or FTP/HTTP listing (raw `LIST` remainder, `FileSystemCurl.Shared.cpp:869-893`) → arbitrary local file write / RCE via e.g. the Startup folder. **Fix:** reject at the host any enumerated child name that is empty, `.`/`..`, contains `\` or `/`, contains embedded NUL/control/DEL, or violates the destination's declared component contract. For a Windows-style destination also reject `:*?"<>|`, trailing dot/space, reserved DOS device names, and case-colliding siblings. Treat it as a per-item error (`ERROR_INVALID_NAME`), not a silent skip. This is the trust boundary; do not rely on any plugin to sanitize. Fold in **review finding 15** (not plan row CW-15) in the same validation layer. | RED-first selftest: a fake source plugin (Dummy variant) enumerates a child named `..\escape.txt` (plus `a\b`, embedded-NUL/control, `x:stream`, `CON`, and a case-colliding pair); bridge copy/move returns `ERROR_INVALID_NAME` for those entries, writes **nothing** outside the destination root, copies valid siblings, and the op ends `ERROR_PARTIAL_COPY`. Removing the guard makes the escape file appear outside the root (test goes RED). |
| [ ] | CW-2 | **Host never clamps plugin-reported `bytesRead` to the buffer size** (serial pump `State.cpp:8192`, pipelined `:8316`). A misbehaving/hostile plugin returning `bytesRead > bufferBytes` drives an out-of-bounds heap over-read that is then written into the destination. **Fix:** after every `Read`, fail with `ERROR_INVALID_DATA` if `bytesRead > bufferBytesIn`. | Fake reader returns `bytesRead` larger than the buffer once; pump returns `ERROR_INVALID_DATA`, temp deleted, no OOB (clean under ASan). Removing the clamp trips ASan / writes wrong length. |
| [ ] | CW-2A | **Destination writer output counts are also untrusted.** Reject `bytesWritten > bytesToWrite` in both pumps and state the reader/writer count bounds in `IFileReader`/`IFileWriter`. | Fake writer over-reports once; bridge returns `ERROR_INVALID_DATA`, deletes staging output, and leaves the final destination untouched. |

## P1 — Availability & functional breaks

| State | ID | Work item | Required proof |
|-------|----|-----------|----------------|
| [ ] | CW-3 | **HIGH process-wide deadlock.** `CopyDirectoryParallel` runs its producer *on* a `PerItemTaskScheduler` worker (`StartJob:9897`), then starts a **nested** file-copy job and blocks in `WaitJob(job)` (`:9165`), which is a bare `condition_variable` wait with no work-stealing and no cancel predicate (`:3882-3891`). Enough concurrently-active producers (e.g. 4 cross-FS dir-copy tasks on a 4-core box) occupy every worker; nested jobs never get dequeued. Cancel does not recover (cleanup runs only on a free worker's dequeue path); `Shutdown()` joins workers before finishing jobs so app-exit hangs too. **Fix:** remove the nested-blocking-job pattern — run file workers on the same job/level, or make `WaitJob` help-execute (participate in dequeue) and honor cancellation. | Stress selftest: N = worker-count concurrent cross-FS directory copies each with a subtree; all complete (or cancel cleanly) with no thread parked in `WaitJob` after cancel; a watchdog fails the test if the scheduler is still busy after the deadline. Reproduce the hang against the pre-fix build. |
| [ ] | CW-4 | **HIGH Curl bridge broken for sizeless `LIST` dialects.** `FileSystemCurl::CreateFileReader` builds `CurlStreamingReader(conn, path, entry.sizeBytes)` discarding `entry.sizeKnown` (`DirectoryOps.cpp:1339`; `GetSize:416-425` then returns `S_OK`+0). The host reserves `GetSize` **failure** as the only unknown-size signal (`State.cpp:8098-8115`) and treats `S_OK`+0 as an authoritative 0-byte file, so the size gate (`:8485`) fails every non-empty file; the MOVE post-promote re-stat (`:8547-8557`) also breaks (dest size 0 ≠ source). **Fix:** wire the existing `CurlProbeRemoteFileSize` single-object probe (FTP `SIZE` / SFTP fstat) into `CreateFileReader` when `!sizeKnown`; if still unknown, return a genuine `GetSize` **failure** so the host applies its unknown-size policy rather than fabricating 0. | Fake FTP server with a sizeless `LIST`: a non-empty bridge COPY promotes and a MOVE deletes source only after re-stat; a genuinely short upload still fails `ERROR_PARTIAL_COPY`. RED against current reader path. |

## P1 — Performance (large-transfer & cloud efficiency)

| State | ID | Work item | Required proof |
|-------|----|-----------|----------------|
| [ ] | CW-9 | **HIGH MOVE cleanup re-transfers every byte — in TWO code paths.** (a) Bridge: before deleting each moved source, `CopiedFileStillMatchesDestination:7491` opens fresh readers on both endpoints and runs a full byte-for-byte `memcmp` (`ReadersHaveEqualContent:7396`), single-streamed through the one primary buffer (`:7442`) even when the copy ran N-way parallel. Local→cloud MOVE re-downloads the entire just-uploaded object (billed egress); cloud→local re-reads the source twice, non-overlapped. (b) **Native local plugin, cross-volume move (this hits every local→network-drive / local→UNC MOVE, no bridge involved):** `MoveFileWithProgressW` is issued WITHOUT `MOVEFILE_COPY_ALLOWED`, so a cross-volume target returns `ERROR_NOT_SAME_DEVICE` (`Plugins/FileSystem/FileSystem.FileOps.cpp:5544`) and the plugin copy+deletes (`:5549`); the delete phase runs `DestinationMatchesSourceFile:3572` → `FileContentsEqual:3364`, re-reading the FULL source and FULL destination (dest read is over the network) before deleting each source file. Plain local→network COPY is unaffected (single `CopyFileEx` pass, `:3785`). **Fix:** adopt the Clearwater CW-P1 copy-time hash (hash while copying, verify only the destination at cleanup) — or, where a strong-consistency destination byte-proved the copy, skip the re-read entirely — and apply it to BOTH paths so they don't drift. **Coordinate with Clearwater CW-P1 — one hashing scheme only.** | Before/after transfer-byte + wall-clock metric on (a) a MOVE of a multi-GB tree local↔cloud and (b) a local→SMB MOVE: cleanup verification transfers ≤ the destination once (hash) or zero extra bytes (skip), never source+destination in full, in both the bridge and the native cross-volume path. Archive under `Specs/TestRuns/`. |
| [ ] | CW-10 | **HIGH OneDrive writer defeats pipelining.** `MicrosoftDriveFileWriter::Write` is `WriteFile` to a delete-on-close local `%TEMP%` file (`FileSystemMicrosoftDrive.cpp:4653-4666`); nothing hits Graph until `Commit` (`:6279`). Wall time = full source read **then** full upload (no overlap), plus a whole extra local disk write+read and a transient free-disk requirement equal to the file size (a 20 GB copy fails outright on a system drive with < 20 GB free). **Fix:** stream to an upload session during `Write` (chunked `PUT` as bytes arrive) so read and upload overlap, or at minimum document + guard the free-space precondition. | Metric: for a large FTP→OneDrive copy, upload starts before the source read completes (overlap > 0); peak `%TEMP%` usage bounded well below file size. Before/after wall-clock archived. |
| [ ] | CW-11 | **HIGH OneDrive reader 1 MiB serial ranges.** `MicrosoftDriveRangedFileReader::FetchRange` clamps every range GET to ≤1 MiB (`:4487-4488`) and `Read` (`:4391-4441`) issues them synchronously back-to-back, so an 8 MiB bridge `Read` (its own advertised `preferredBufferBytes`, `:5741`) becomes 8 sequential RTT-bound request/response cycles with no prefetch. **Fix:** honor the caller's requested size (single ranged GET up to the buffer), and/or prefetch the next range concurrently. | Throughput metric on a high-RTT link: effective MB/s rises toward link capacity; request count per file drops ~8×. Archived before/after. |
| [ ] | CW-12 | **HIGH S3 small-file round-trip storm.** Per bridged file to S3: host `GetAttributes` (HEAD) + `CreateFileWriter`→`EnsureWritableS3Target` (per-ancestor summary probes + `ListObjectsV2`, `IO.cpp:627-698`) + `Commit` **re-runs** `EnsureWritableS3Target` (`:833`) + host promote = server-side `CopyObject`+delete + (MOVE) a HEAD — ~15 API calls where a `PutObject` is 1. **Fix:** cache the writable-target check between create and commit (don't re-probe); drop redundant per-ancestor probes; see CW-16 for the promote copy. | Call-count metric copying a deep small-file tree to S3: ≤ ~3 S3 ops per file (create/put/promote) with no per-ancestor re-probe and no duplicated `EnsureWritableS3Target`. Archived. |

## P2 — Data integrity (narrow races / verification gaps)

| State | ID | Work item | Required proof |
|-------|----|-----------|----------------|
| [ ] | CW-5 | **Cloud torn-copy under concurrent overwrite.** S3 (`FileSystemS3.IO.cpp:567`, size cached once `:504-528`) and MSDrive (`:4410`, size refreshed only on 401/403) ranged readers issue independent per-chunk GETs with **no ETag/`If-Match`/versionId pin**. A concurrent same-size replacement of the source splices two versions; the host size gate passes and (for COPY, see CW-6) it commits silently. **Fix:** capture the object ETag/version on the first response (or the size HEAD) and send `If-Match`/`If-Range`/versionId on every subsequent range → `412`/`416` fails closed. | Fake S3/Graph transport swaps the object to a same-size version between chunk fetches: reader returns a precondition failure → bridge `ERROR_PARTIAL_COPY`, temp deleted. RED without the pin. |
| [ ] | CW-6 | **Plain COPY has no destination verification.** The post-promote re-stat and byte-compare are gated to `FILESYSTEM_MOVE` (`State.cpp:8544`); COPY trusts written-byte count + `Commit` + promote only. This is the enabler that turns CW-4/CW-5/CW-7 into *silent* corruption. **Fix:** run the post-promote destination size re-stat (and, when a copy-time hash exists per CW-9, a hash check) for COPY as well, at least on non-strong-consistency backends. | Fake destination whose promoted file mis-reports/mismatches size: COPY returns `ERROR_PARTIAL_COPY` and removes the bad destination. RED with today's MOVE-only gate. |
| [ ] | CW-7 | **Stale cloud `GetSize` truncates COPY.** A cloud reader's cached `GetSize` (OneDrive `:4410`) can be smaller than the live object; the pump stops at the stale size, the size gate matches, and COPY commits a truncated file. Largely subsumed by CW-5+CW-6; track until both land. | Covered by the CW-5/CW-6 tests with a *shrinking-then-growing* stale-size variant. |
| [ ] | CW-8 | **Reparse-point FILES bypass `ReparsePointPolicy`.** The tree walk applies the reparse policy to directory entries only (`:8733`, `:9093`); a reparse *file* (symlink/junction-to-file, mount surrogate) falls through to `CopyFile` and is dereferenced/copied regardless of Skip/CopyReparse. **Fix:** apply the same policy branch to files. | Selftest with a reparse-point file under a copied tree: Skip records+skips it, CopyReparse → `ERROR_NOT_SUPPORTED`; neither silently copies target bytes. |

## P2 — Performance (secondary)

| State | ID | Work item | Required proof |
|-------|----|-----------|----------------|
| [ ] | CW-13 | **Throttle serializes all workers.** `ThrottleThreadSafe` (`:7068-7077`) takes `throttleMutex` and then *sleeps* inside `Throttle` (`:7029-7066`); with a bandwidth limit set, every pump worker blocks behind the sleeping one. **Fix:** compute the per-worker delay under the lock, release, then sleep — or use a lock-free token bucket. | With a limit set and >1 worker, aggregate throughput approaches the configured cap (not cap/Nworkers); micro-metric shows no worker parked on `throttleMutex` during a peer's sleep. |
| [ ] | CW-14 | **S3 writer uploads 64 MiB parts synchronously inside `Write`.** `FileSystemS3.IO.cpp:947` flushes each part on the pump thread one at a time, stalling the reader during upload. **Fix:** overlap part uploads (background/async multipart) so `Write` returns while the part uploads. | Metric: reader wait during S3 writes drops; overlap > 0 on a multi-part file. |
| [ ] | CW-15 | **Source tree enumerated 2× (COPY) / 3× (MOVE).** Pre-scan for byte totals + copy walk + (MOVE) cleanup walk each re-`ReadDirectoryInfo` (`:5116` pre-scan; copy `:8649/:8807`; cleanup `:7713`) — expensive on high-metadata-cost cloud backends. **Fix:** capture the tree once into a worklist and reuse it across phases where consistency allows. | Enumeration-call metric on a cloud tree: ≤1 listing per directory for COPY, ≤2 for MOVE. Archived. |
| [ ] | CW-16 | **S3 promote is a full server-side copy+delete.** `PromoteTempToFinalPath` → S3 `MoveItem` = `CopyObject` temp-key→final-key + delete (`FileSystemS3.S3.cpp:845`), doubling per-object work. **Fix:** write directly to the final key when no conflict prompt is pending (skip the temp-then-copy for S3), or use the multipart complete as the atomic point. | Call-count/byte metric: no `CopyObject` per file on the no-conflict path. |

## P3 — Robustness / correctness (LOW; batchable)

| State | ID | Work item | Required proof |
|-------|----|-----------|----------------|
| [ ] | CW-17 | **7z reader EOF-contract violations:** a mid-item decoder failure is reported once then subsequent `Read`s at the EOF position return clean `S_OK`+0 (`FileSystem7z.cpp:2991`), and `Read` can return `S_OK` with `bytesRead==0` *before* true EOF (`:3013`) — both let the host mistake a decode failure/short stream for a complete file (masked today only by the size gate). **Fix:** latch the decoder error and keep returning it; never signal EOF before the known size. | Selftest with a corrupt 7z entry: bridge copy fails, never commits a short file. |
| [ ] | CW-18 | **Pipeline shutdown join can hang.** After a writer-side failure the reader `jthread` is `join()`ed (`:8472`) but a reader parked in `WaitWhilePaused` or a long plugin `Read` isn't signalled to abort promptly. **Fix:** set the stop flag + wake pause/cancel before join; ensure the reader checks it in `WaitWhilePaused`. | Selftest: fail the writer mid-file while the reader is paused; the copy unwinds within the deadline. |
| [ ] | CW-19 | **Curl overwrite-promote reports failure after the point of no return** (`FileSystemCurl.CopyMove.cpp:1061`): once the destination is replaced, a backup-delete failure surfaces as a failed promote, so the host may retry/clean up a destination that is actually correct. **Fix:** treat post-replacement backup-cleanup failure as success-with-warning. | Fake transport where backup delete fails after replace: promote returns `S_OK`, destination intact. |
| [ ] | CW-20 | **Curl advertises 8 MiB `preferredBufferBytes` but the reader ring holds 1 MiB** (`DirectoryOps.cpp:354`); the bridge sizes an 8 MiB buffer the reader can never fill in one call. **Fix:** align the advertised hint with the ring, or enlarge the ring. | Hint vs actual ring size consistent; per-`Read` fill matches the advertised buffer. |
| [ ] | CW-21 | **Transfer hints are dead except `preferredBufferBytes`, and any non-default user buffer setting silently disables adaptation** (`ResolveAdaptiveCrossFsBridgeBufferBytes:2046-2089`; only adapts when the configured value equals the default). `latencyClass`/`flags`/`preferredProgressPeriodMs` ignored. **Fix:** consume latency/flags for pump tuning; make adaptation independent of whether the user changed the buffer size (or document the coupling in the spec + Preferences). | Adaptation applies with a non-default buffer setting; a latency hint measurably changes pump behavior. Spec updated. |
| [ ] | CW-22 | **Data race on `Task::_perf` plain `uint64_t` counters** updated from concurrent bridge threads (`:6366`). **Fix:** make them `std::atomic` (as the bridge-local perf already is) or accumulate per-thread and merge. | TSan-clean under a parallel bridge copy. |
| [ ] | CW-23 | **`PromptDestinationCollision` reads shared `FileSystemOptions` without `callbackMutex`** that writers hold (`:7155`). **Fix:** take the same lock (or snapshot under it). | TSan-clean; no torn read of options during a concurrent progress update. |
| [ ] | CW-24 | **Genuine worker failure misreported as `ERROR_CANCELLED`** in `CopyDirectoryParallel` (`:8880`) — the not-ready/stopped branch collapses a real failure into cancellation, hiding the true error. **Fix:** preserve the first recorded producer/worker HRESULT and prefer it over `ERROR_CANCELLED` unless cancel was actually requested. | Selftest: inject a worker failure with no cancel; task result carries the real HRESULT, not `ERROR_CANCELLED`. |
| [ ] | CW-25 | **Circuit-breaker HRESULT vocabulary mismatch** (`:2962`): the transient-classifier expects codes Curl/S3 don't actually emit for network failures, so the breaker under-triggers. **Fix:** reconcile the classifier with the real HRESULTs each plugin maps network errors to. | Unit test feeds each plugin's real network-failure HRESULTs; classifier counts them transient. |
| [ ] | CW-26 | **Unbounded-by-design memory footprint:** `workers × 2 buffers` (pipeline) `+` S3 writer RAM parts can exceed 1 GiB (`:8858`). **Fix:** bound total bridge buffer memory (global budget) and document the ceiling in the spec. | Peak-RSS metric under max concurrency stays under a documented ceiling. |

## Coverage & cross-references (not bridge fixes)

| State | ID | Work item | Required proof |
|-------|----|-----------|----------------|
| [ ] | CW-T1 | **No end-to-end self-test exercises the size-integrity gate** (`:8485`) via premature EOF or writer under-consumption. Add independently armable Causeway reader/writer decorator modes and assert `ERROR_PARTIAL_COPY` + temp deleted. FIR/FG-A1 remains the separate destination-`GetSize` seam. | Removing the size-mismatch block (`:8485`) makes the test RED. |
| [>] | CW-X1 | **CompareDirectoriesEngine maps reader-create/`Read` failure to `Different`** (`CompareDirectoriesEngine.cpp:3010`), contradicting the bridge's fail-hard `IFileReader` stance — a contract ambiguity surfaced by divergent consumers, **not** a bridge defect. Routed to the Compare owner (`Specs/Core/Core_CompareDirectories.md`); record the contract decision (I/O error ≠ content difference) there and align both consumers. | Decision recorded in the Compare spec; Compare reports I/O errors distinctly from `Different`. |

## Closeout Rule

Each correctness/security item MUST land with deterministic selftest or fake-provider coverage that is
RED before the fix and GREEN after, and MUST update `Specs/Core/Core_FileSystemBridge.md` where it
changes durable behavior. Each performance item MUST follow
`Specs/Testing/Testing_PerformanceValidation.md` with archived before/after evidence under
`Specs/TestRuns/`. New fileops selftest cases MUST be registered in `kFileOpsFamilyDefinitions` (and the
right `kFileOpsFamily*` array) or Full-suite runs skip them silently.
