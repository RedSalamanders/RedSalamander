# Cross-File-System Bridge — Implementation Today & Performance Redesign Thinking

**Status:** COMPLETE / ARCHIVED — R-1b/S3 is shipped and validated; all residual redesign rows are done, rejected, or explicitly deferred
**Date:** 2026-07-07; reconciled 2026-07-20 against `1c8abc0318adc3530e846afb7f00f62e4c2dd43c`
**Purpose:** describe the bridge **as it is implemented today** and capture the *design-level*
weaknesses — where the shape of the implementation (not a single line) limits performance — so we can
decide what to rework.
**Companions:**
- Spec (normative behavior): `Specs/Core/Core_FileSystemBridge.md`
- Itemized findings: `Specs/Reviews/FileSystemBridge-2026-07-07-Findings.md`
- Completed point-fix ledger (bugs/security): `Specs/Plans/Done/Operation_Causeway_CrossFsBridgeSecurityAndPerformanceRemediation_2026-07-07.md`

**Single-owner note:** Causeway owns the discrete correctness/security/perf *fixes* (CW-1..CW-26). This
doc owns the *architectural* rethink — the pump model, the verify model, the per-file model, the
concurrency model. Where a redesign here would subsume a Causeway row, that row is noted.

## 2026-07-20 current-state reconciliation and closeout scope

Causeway changed the implementation materially after this design document was written. The current code and
authoritative bridge spec already contain the following results; they are closed inputs, not work to repeat:

- the host consumes buffer, latency, flag, and progress-period transfer hints even when the user changed the
  configured buffer size;
- host bridge buffers have a cancellation-aware 256 MiB process budget and the S3 writer has a separate four-part /
  256 MiB process budget;
- the two-slot host pump has deterministic same-machine Release evidence: 130,766,000 us serialized baseline versus
  422,000 us pipelined candidate for four 32 MiB files with 30 ms injected chunk latency;
- bridge MOVE hashes bytes while copying and re-reads only the destination during cleanup; native cross-volume MOVE
  normally uses stable source/destination snapshots with zero content re-read;
- scheduler workers waiting on nested jobs help execute that job, so directory producers cannot deadlock the worker
  pool;
- S3 pins ranged reads to an object revision, validates the destination once, writes directly to the final key through
  the explicit atomic-writer capability, and overlaps one 64 MiB multipart upload with the host pump;
- Microsoft Drive issues one pinned range up to the requested buffer and streams known-size writes through a Graph
  upload session instead of spooling the whole file to `%TEMP%`.

### Last finding: the remaining S3 writer window was one part

At reconciliation, `MultipartS3FileWriter` owned one `_uploadThread`. At the start of every subsequent `Write`, it
called `CollectAsyncPart()` before appending more data, and `FlushBufferedParts()` rejected a second scheduled part
while that thread was joinable. The result was exactly one outstanding S3 `UploadPart` request per file. The debug
contract proved only that one part overlapped the caller; it did not prove concurrent part requests.

Adding host pump depth without concurrent provider requests is not the next throughput lever. With one sequential
reader and one sequential writer, a deeper queue mainly absorbs jitter; it cannot hide S3 request latency or multiply
multipart throughput. This agrees with Appendix A's existing R-1b boundary: genuine cloud request overlap requires
concurrent ranged reads or multipart writes.

**Implementation checkpoint (2026-07-20, validated):** the writer now owns a FIFO of independently owned
pending parts and admits two `UploadPart` workers per writer. Every assembling or submitted payload owns one lease from
the existing four-slot process semaphore; workers are `std::jthread`s joined before their payload/session/owner can be
released; results are collected in admission order and explicitly sorted by part number before completion. The first
worker failure is atomically published, new admission observes it, all started workers drain, and the session abort path
remains single-shot. Test-only transport coverage now includes a serialized baseline, the two-worker candidate,
out-of-order completion, exact first-error propagation, abort retry and unload quiet point, permit release, and the
small-object `PutObject` path. The focused Debug and test-enabled Release suites are green, including a dual-failure case
that proves the first failure by observation time is retained even when FIFO collection joins a lower-numbered part first.

**Build checkpoint:** the first contamination-recovery Debug build exposed that FileSystemS3 did not reference the
canonical `Common` runtime implementing `Debug::Perf` JSONL output (`WritePerfJsonl`/`HasPerfJsonlOutput`). The plugin
now uses the same `Common.vcxproj` reference as the other metric-emitting plugins. The required full Debug recovery
rebuild then passed with 0 warnings/0 errors, and the focused Debug plugin-contract run passed all S3 multipart,
lifecycle/unload, and cross-plugin curl-survivor checks.

**Release checkpoint:** the first test-enabled Release attempt exposed an existing configuration mismatch: a
Debug-only curl-survivor step unconditionally required `_DEBUG` probe exports from Release plugins. That step and its
helper/type are now correctly gated by `_DEBUG`; Release explicitly reports the skip while the production S3 lifecycle
step and the S3 `ENABLE_TESTS` performance suite remain active. The corrected run is all green:
`PluginContractTests.exe` exited 0, the multipart suite passed 21/21, eight S3
runtime-refresh cycles reached the unload quiet point, and Release explicitly skipped only the absent Debug-curl probes.
The final archived comparison is 694,050 us serialized versus 347,994 us bounded (49.86% reduction), in-flight
high-water 2, and payload high-water 2 of 4. The earlier failed enclosing attempt is retained only under
`.build/codex-runs/20260720_180600_s3_multipart_release_attempt1_failed`, outside durable evidence.

**Broad closeout checkpoint:** Full run `20260720T162702Z-9092-2fe094418c414776ab29aa4aa8a59641` rebuilt Debug and
proved the in-scope surfaces: Compare Directories passed 226/0, File Operations passed 108/0 including Phase 11's
bridge concurrency/performance/correctness cases, and `PluginContractTests` passed with S3 debug 150/0 and multipart
21/0. The run initially found that the rewrite had removed the reviewed Track-11 source marker for asynchronous retry
of a failed destructor abort; the marker was restored at the retry queue and that exact Pester case passes. The Full
wrapper as a whole remains red (1,132 passed, 29 failed, 53 skipped) from 27 Commands UI/settings failures,
ViewerPE's outer fresh-harness failure after its inner run printed `ViewerPETests passed`, and the unrelated Track-14
settings source-marker failure. None touches the bridge/S3 implementation or its authoritative contract, so those
separate owners do not leave residual work in this plan. After the Track-11 marker repair, the final focused
FileSystemS3 Debug build passed with 0 warnings/0 errors and `PluginContractTests` again passed S3 debug 150/0,
multipart 21/0, all eight unload-refresh cycles, and the cross-plugin curl survivor proof.

### Completed implementation slice: R-1b/S3

Implement a bounded S3 multipart-upload window while leaving the public `IFileWriter` ABI sequential:

- permit at least two `UploadPart` requests for one writer to be outstanding concurrently;
- keep each submitted part's number, payload, result, and ETag independently owned until collection;
- complete the multipart upload with parts ordered by part number even when requests finish out of order;
- charge the existing four-slot / 256 MiB S3 budget per live 64 MiB payload, including the assembling buffer, rather
  than once per writer;
- stop admitting new parts after the first failure, join every started worker, preserve the first real HRESULT, and
  abort the multipart session exactly once on failure or abandoned release;
- keep plugin unload safe: no detached threads, and every instance-owned worker reaches a quiet point before the
  writer can be destroyed;
- keep small objects on the existing single `PutObject` path and avoid concurrency overhead for one-part uploads.

**Protected scenario:** a large cross-filesystem upload to S3 containing at least four 64 MiB multipart parts while
each deterministic upload request has fixed latency. The serialized baseline and bounded-concurrency candidate must
use the same payload, delay, build flavor, and machine.

**Metric family:** `FileOps.S3.Multipart.*`, including total upload duration, part count, in-flight high-water mark,
capacity-wait time, uploaded bytes, and first failure. Retain `FileOps.Bridge.Copy` as the end-to-end host metric.

**Go/no-go performance gate:** in a test-enabled Release build, the bounded-concurrency candidate must reduce the
deterministic multi-part wall clock by at least 30% versus the serialized baseline, report an in-flight high-water mark
of at least two, remain within the 256 MiB S3 payload budget, and show no material regression for a one-part object.
Live S3 evidence is optional and directional; no live-cloud claim is permitted without credentials and a same-endpoint
baseline/candidate comparison.

### Progress ledger

Update this table after each code, test, or evidence checkpoint. The plan cannot move to `Specs/Plans/Done/` while any
row is unchecked or while any residual redesign item below lacks an explicit done/rejected/deferred disposition.

| Checkpoint | State | Evidence / result |
|------------|-------|-------------------|
| Reconcile plan with current Causeway code and record the last finding | [x] | This section; HEAD `1c8abc031` |
| Add deterministic serialized baseline and required metrics before changing concurrency | [x] | Same four-part/150 ms test transport with max-in-flight 1 versus 2; `Baseline`, `Candidate`, `Improvement`, and writer total/high-water/wait/upload/failure metrics are present in the archived Release JSONL |
| Implement bounded concurrent part ownership and per-payload admission | [x] | FIFO of two `PendingPart` owners, per-payload four-slot semaphore leases, `std::jthread` quiet point, ordered collection/sort, and atomically published first failure |
| Prove overlap, ordered completion, first-error propagation, abort-once, release quiet point, and memory ceiling | [x] | Dedicated S3 suite: 21 passed/0 failed; deterministic part 2-before-1 still commits `[1,2]`; part 2's earlier `0x800704CF` remains exact even after lower-numbered part 1 later fails, and only two parts schedule; abort is once plus one queued retry in abandonment; workers/permits drain to zero; peak 2 of 4 payloads |
| Focused Debug build and S3/PluginContract tests green | [x] | Mandatory full Debug rebuild: 0 warnings/0 errors; `PluginContractTests.exe`: exit 0 including S3 debug 150/0, multipart 21/0, module quiet point, eight S3 lifecycle refresh cycles, and cross-plugin curl survivor proof |
| Test-enabled Release baseline/candidate meets the 30% gate | [x] | Final all-green run: 694,050 us baseline vs 347,994 us candidate = 49.86% reduction; required ≥30%; in-flight high-water 2; payload high-water 2/4; multipart suite 21/0 and quiet point green |
| Archive and validate `results.json`, `trace.txt`, and `perf_metrics.jsonl` under `Specs/TestRuns/` | [x] | `Specs/TestRuns/4cb089111a23/FileOps/20260720_181000_s3_multipart_concurrency_release/`; explicit three-file archive validation passed; 74,709 bytes total |
| Update `Specs/Core/Core_FileSystemBridge.md` with the lasting concurrency/metric/budget contract | [x] | Provider-local two-worker exception, per-payload budget, failure/quiet-point semantics, Release perf gate, metric family, and small-object path are normative |
| Reconcile `Specs/Plans/WIP/README.md` and disposition every residual R-1/R-2/R-3/R-4 item | [x] | WIP I9 removed; this plan is listed as an archived reference; R-1a rejected, R-1b done, parallel reads deferred to a future accepted design, R-2/R-4 done, R-3 batch ABI deferred |
| Broad bridge/FileOps and source-contract gate green; plan moved to `Specs/Plans/Done/` | [x] | Full Debug build and in-scope gates green: Compare 226/0, FileOps 108/0, PluginContract S3 150/0 + multipart 21/0, Track-11 Pester passed after marker repair; unrelated Full failures are dispositioned in the broad checkpoint |

---

## 1. What the bridge is for

When you copy or move between two different plugins (local disk → S3, 7z → local, FTP → OneDrive), no
single plugin can do the transfer — each one only knows its own backend. The bridge is the host acting
as the middleman: it reads bytes from the source plugin and writes them to the destination plugin
itself, through the `IFileReader`/`IFileWriter` streaming contract. Its safety promise is "never leave a
partial or corrupt file at the final path," which it keeps through temp-name promotion or an explicitly
capability-gated atomic final-key writer commit after the byte count checks out.

---

## 2. How it is implemented today

**Engagement.** The bridge runs only when source and destination are different plugins/contexts, and
only in per-item mode. Each top-level item you selected gets its *own* `CrossFileSystemBridge` object
and its own `CopyPath` call. There is **no batch path** — the native `CopyItems` bulk entry point is
never used across the bridge. Everything funnels into one function per file, `CopyFileWithBuffer`
(`RedSalamander/FolderWindow.FileOperations.State.cpp:8029`).

**The per-file sequence is a straight line:**

```
acquire per-connection permits (both ends)
 └ check overwrite policy (prompt BEFORE any bytes move)
   └ CreateFileReader(source)
     └ reader->GetSize()          ← FAILURE IS FATAL, nothing is written
       └ CreateFileWriter(<final-or-temp>)   [final only for explicit atomic-writer capability]
         └ PUMP bytes  (serial OR 2-slot pipeline)
           └ writtenBytes == GetSize ?   else ERROR_PARTIAL_COPY, delete temp
             └ writer->Commit()          else delete temp
               └ non-atomic writer only: MoveItem(temp → final)  [promote]
                 └ re-open dest, GetSize must match; COPY also verifies the copy-time hash
                   └ SetFileBasicInformation (timestamps, best-effort)
```

A `wil::scope_exit` deletes a host-owned temp on every failure path unless promote already succeeded. An
`IFileSystemAtomicWriter` publishes the final path only at `Commit`; the host therefore skips the temp/promote
round-trip for that explicit capability. S3 opts into this path because multipart completion is its atomic point.

**The pump has two modes.**
- *Serial* (`:8175`): one thread does `Read` into a buffer, then a `Write` loop drains it, then repeat.
- *"Pipelined"* (`:8256`): **exactly two** buffer slots and **one** `std::jthread` reader — the reader
  fills a slot while the calling thread writes the other, coordinated by a single mutex + condition
  variable. It engages only when `fileTotalBytes > bufferBytes`; a file that fits in one buffer always
  runs serial.

**Buffer sizing.** Default 4 MiB, clamped 512 KiB–16 MiB. `ResolveAdaptiveCrossFsBridgeTuning` queries both
endpoints' `GetTransferHints`, takes the maximum provider-preferred buffer, maximum latency class, ORed flags, and
maximum progress period, then freezes the result for the task. Provider hints remain active when the user changes the
configured fallback. Host primary/secondary buffers reserve against one cancellation-aware 256 MiB process budget.

**Directory trees** go sequential (plain recursion) or parallel (a stack-based producer feeds a 16-entry admission
queue drained by process-wide scheduler workers, each with its own buffer). A worker waiting for nested directory work
helps execute that job. MOVE remains a two-pass operation, but the copy pass records a streaming hash and stable source
snapshot; cleanup re-reads only the destination before deleting the unchanged source.

**Per-plugin reality (where the abstraction leaks):**

| Plugin | Reader | Writer | Promote |
|--------|--------|--------|---------|
| Local (Win32) | short only at EOF | temp-sibling + `FlushFileBuffers` | true rename |
| S3 | pinned 8 MiB ranged GET | 64 MiB multipart parts with at most two submitted parts outstanding under a four-payload process budget | atomic final-key multipart completion |
| OneDrive | pinned range up to the requested buffer | known-size Graph upload session streams during `Write` | Graph move (cross-FS move disabled) |
| Curl (FTP/SFTP/HTTP) | fills a 1 MiB ring; unknown size remains unsupported | host-staged sequential writer | delete-backup + rename; post-replace cleanup is warning-only |
| 7z | routine short reads; read-only | n/a | n/a |
| Google Drive | — (no `IFileSystemIO`) | — | — |

---

## 3. Design-level weaknesses & redesign directions

The correctness bugs are fixable point-by-point in Causeway. The items below are different: the *shape*
of the implementation limits performance, and they warrant a rework rather than a patch.

### R-1 — CLOSED: S3 request concurrency shipped; generic depth-N rejected; parallel reads deferred

**Current design.** The host has two slots, one outstanding read, and one writer call; it engages for files larger
than one buffer. That double buffer already overlaps the two host legs effectively in the deterministic Causeway
scenario. S3 now admits two independently owned `UploadPart` workers per writer under a four-payload process budget,
preserves the first observed failure, orders completion input by part number, and joins all workers before release.

**Measured result.** The final test-enabled Release run reduced the protected four-part fixed-latency path from
694,050 us serialized to 347,994 us bounded (49.86%), with request high-water 2 and payload high-water 2/4.

**Disposition.** R-1a generic host depth-N is **rejected as the next standalone optimization** unless a future jitter
benchmark proves material value. R-1b/S3 bounded multipart-write concurrency is **done** with the archived evidence
above. Capability-gated parallel ranged reads remain **deferred** and require a separately accepted ABI/design plan;
they are not residual work in this plan.

### R-2 — DONE: MOVE verification no longer re-transfers the source

**Current design.** Causeway closed this row. Bridge MOVE records a copy-time hash and re-reads only the destination;
native cross-volume MOVE normally uses stable no-follow source/destination snapshots without a content re-read.

The superseded analysis correctly identified both the bridge and native cross-volume paths. Causeway implemented the
shared outcome: hash during copy, retain stable source/destination identity, and re-read only the destination when
verification is required. There is no residual R-2 implementation work in this plan.

### R-3 — CLOSED: atomic final writes and validation reuse landed; batch API deferred

**Current design.** Every file still creates a reader and writer, but the host and plugin already fan files out under
bounded connection/scheduler concurrency. S3 reuses destination validation and the explicit atomic-final capability
removes temp-key CopyObject promotion. A new cross-plugin batch ABI remains unproven and is deferred.

The direct-final-key, reused-validation, and bounded cross-file fan-out portions of the proposed redesign are now the
current implementation. A provider batch ABI remains a future option only if fresh request-count evidence justifies
its additional contract surface; it is not residual work in this plan.

### R-4 — DONE: nested scheduler waits help execute their target job

**Current design.** Causeway closed this row. A scheduler worker waiting for a nested job participates in dequeue, and
cancellation wakes waiters; deterministic worker-saturation coverage is green.

The superseded deadlock risk is now covered by the help-execute wait and cancellation contract. There is no residual
R-4 implementation work in this plan.

### Smaller design smells — DONE by Causeway

- Provider hints remain active with non-default settings; buffer, latency, flags, and progress cadence are consumed.
- The host and S3 payload budgets are explicit and independently capped.

---

## 4. Final disposition sequence

1. **R-1b/S3 bounded multipart-write concurrency is complete** with the active performance and correctness gates.
2. Keep **R-2** and **R-4** closed; do not reopen their superseded implementation text.
3. Close **R-3** by recording the landed atomic-final/validation/fan-out behavior and explicitly deferring a new batch
   ABI until request-count evidence demonstrates that it is worth its contract cost.
4. Reject **R-1a depth-N** as a standalone closeout requirement; require a fresh measured plan to reopen it.

## Closeout Rule

Each redesign that ships MUST update `Specs/Core/Core_FileSystemBridge.md` to describe the new behavior,
carry before/after performance evidence per `Specs/Testing/Testing_PerformanceValidation.md` archived
under `Specs/TestRuns/`, and mark the Causeway rows it subsumes as satisfied.

---

# Appendix A — R-1 Detailed Design: depth-N decoupled pump

**Disposition (2026-07-20): historical design, not an active closeout requirement.** Causeway landed the transfer-hint
and memory-budget prerequisites and proved the existing two-slot pump. R-1a depth-N is rejected unless a fresh jitter
benchmark demonstrates material benefit. Only Appendix A.7's provider-side request concurrency remains relevant; its
S3 multipart-write subset is the completed slice defined at the top of this document. Parallel ranged reads require a
separate accepted ABI/design plan.

Scope: replace the fixed 2-slot pump (`State.cpp:8256-8478`) with a depth-N bounded-queue pump whose
depth and buffer size are chosen from the endpoints' `latencyClass`. This is the highest-return, lowest-
structural-risk change; it is local to the per-file copy loop and keeps the `IFileReader`/`IFileWriter`
contract unchanged. It splits into **R-1a** (this appendix — depth-N single-reader/single-writer) and a
later **R-1b** (parallel ranged I/O), separated because R-1b needs contract changes.

## A.1 Before — what exists today

`ShouldUseBufferedPipeline` (`:7079`) turns on a **2-slot** pipeline only when `fileTotalBytes >
bufferBytes`. One `std::jthread` reader fills `slots[readIndex]`; the calling thread writes
`slots[writeIndex]`; both walk a shared pair of indices `mod 2` under one mutex + one condition variable.

```
 reader jthread                writer = calling thread
      │  fill slot[readIndex]        │  drain slot[writeIndex]
   ┌──▼──┐   ready flag   ┌──────────▼──┐
   │slotA│ ◄───────────►  │  write out  │
   │slotB│                └─────────────┘
   └─────┘  round-robin mod 2, shared indices, single mutex/cv
   read-ahead depth ≈ 1 chunk
```

Limitations that cap throughput:
- **Depth ≈ 1.** The reader can be at most one chunk ahead of the writer, so the read leg and the write
  leg barely overlap. On a WAN/cloud link where *both* legs are high-latency, most of each leg's latency
  is still paid serially.
- **Tight coupling.** Reader and writer contend on the same indices/cv; a stalled writer (slow cloud
  upload) blocks the reader after one chunk of slack — there is no buffer to absorb latency jitter.
- **Mode-gated.** Files ≤ one buffer never pipeline at all — the entire small-file workload runs serial.
- **Depth is hard-wired to 2** and ignores `latencyClass` (which the hint struct already carries but the
  bridge discards, `:2088`).

## A.2 After — depth-N bounded-queue pump

One contiguous allocation of `N × bufferBytes`, managed as two FIFO queues. A single reader thread and a
single writer thread (the calling thread) meet **only at the queues** — no shared indices, minimal lock
hold. The reader can run up to `N-1` chunks ahead; a stalled writer is absorbed by up to `N` buffers of
slack, and vice-versa.

```
                    free queue  (buffers the reader may fill)
        ┌──────────────────────◄──────────────────────────┐
        │                                                  │
   READER thread                                    WRITER thread (calling)
    pop free ─► Read() ─► push filled            pop filled ─► Write() ─► push free
        │                                                  ▲
        └──────────────────────►──────────────────────────┘
                    filled queue (FIFO — preserves byte order)
   read-ahead depth ≈ up to N-1 chunks; legs overlap and jitter is absorbed
```

**Data structures.**

```cpp
struct PumpBuffer { std::byte* data; unsigned long length; uint64_t offset; };

struct PumpChannel {                 // single-reader / single-writer today; MPMC-ready
    std::vector<std::byte> storage;  // ONE allocation: N * bufferBytes
    std::deque<PumpBuffer> freeQ;    // seeded with all N buffers
    std::deque<PumpBuffer> filledQ;  // reader -> writer, FIFO
    std::mutex m;
    std::condition_variable freeCv, filledCv;
    bool eofReached = false;         // set by reader after its last real chunk
    bool aborted    = false;         // set by either side on error/cancel; wakes both
    HRESULT firstError = S_OK;       // the error that caused the abort
};
```

**Reader loop (spawned jthread).**

```
loop:
  task.WaitWhilePaused()
  if CancelRequested(): abort(ERROR_CANCELLED); break
  buf = take_free()                 // waits on freeCv until freeQ nonempty || aborted
  if aborted: break
  n = reader->Read(buf.data, bufferBytes)
  if FAILED(n): abort(hr); break
  if n > bufferBytes: abort(ERROR_INVALID_DATA); break   // CW-2 clamp, host-side
  if n == 0: { eofReached = true; notify filledCv; break }
  buf.length = n
  push_filled(buf)                  // notify filledCv
end
// on exit: notify both cvs so the writer can never hang
```

**Writer loop (calling thread).**

```
loop:
  task.WaitWhilePaused()
  if CancelRequested(): abort(ERROR_CANCELLED); break
  wait until (!filledQ.empty() || aborted || eofReached)   // filledCv
  if aborted: hr = firstError; break
  if filledQ.empty() && eofReached: hr = S_OK; break        // drained cleanly
  buf = pop_filled()
  offset = 0
  while offset < buf.length:                                 // honor short writes
      w = writer->Write(buf.data + offset, buf.length - offset, &written)
      if FAILED(w): abort(w); break
      if written == 0: abort(ERROR_WRITE_FAULT); break
      offset += written
      fileCompletedBytes += written                          // overflow-checked (unchanged)
      overallCompletedBytes.fetch_add(written)               // overflow-checked (unchanged)
      maybeReportProgress(...)                                // 200ms throttle (unchanged)
      ThrottleThreadSafe(...)                                 // bandwidth limit (unchanged)
  push_free(buf)                     // notify freeCv
  if aborted: break
end
// after loop: signal aborted (if not already), then readerThread.join()
```

Ordering is guaranteed with zero extra work: one sequential reader + a FIFO `filledQ` ⇒ the writer sees
chunks in file order, so the sequential `IFileWriter` contract is honored.

## A.3 How `latencyClass` selects the depth

Both endpoints already answer `GetTransferHints` (`FileSystemTransferHints.latencyClass`, values
`LOCAL/LAN/WAN/CLOUD`, `Common/PlugInterfaces/FileSystem.h:86-93`). The **slower endpoint dominates** the
pipe, so take the max. Extend the existing `ResolveAdaptiveCrossFsBridgeBufferBytes` (`:2046`) to also
return the merged latency class and derive a depth:

| `max(src, dst)` latencyClass | depth N | rationale |
|------------------------------|---------|-----------|
| LOCAL / UNKNOWN | 2 | disk-bound; minimal overlap, minimal RAM (or stay serial for local↔local) |
| LAN | 4 | hide LAN RTT + fill jitter |
| WAN | 8 | several chunks in flight to cover RTT |
| CLOUD | 8 (budget-capped) | writer is bursty (multipart / staged); keep the reader well ahead |

**Selection algorithm.**

```
lc  = max(srcHints.latencyClass, dstHints.latencyClass)      // CLOUD>WAN>LAN>LOCAL
N   = depthForLatency(lc)                                    // table above
N   = min(N, perFileBudgetBytes / bufferBytes)              // memory clamp (A.4)
N   = (fileTotalBytes > bufferBytes) ? max(N, 2) : 1        // single-chunk files: serial
```

**Bandwidth-delay-product sanity.** To saturate a link you need in-flight bytes ≥ bandwidth × RTT. With
4 MiB buffers, depth 8 = 32 MiB in flight, which covers e.g. 250 ms × 1 Gbps (~31 MiB) — right for WAN.
Honest limit: on **object stores the bottleneck is per-request TTFB, not bandwidth** — a single
sequential reader issuing one ranged GET at a time cannot hide that latency no matter how deep the queue.
Depth-N (R-1a) buys read/write-leg overlap and jitter absorption; hiding cloud request latency needs
**concurrent** ranged requests → R-1b below. Setting N high on cloud without R-1b mainly deepens the
buffer against the writer's multipart bursts, so cap CLOUD depth by the memory budget rather than raising
it further.

`flags` also feed in: `PREFERS_LARGE_BUFFERS` biases buffer size up (existing knob); `HIGH_METADATA_COST`
does not affect pump depth. Fix the "adaptation only when the setting is default" coupling here too — the
user's buffer size and the auto depth should compose, not be mutually exclusive (Causeway CW-21).

## A.4 Memory budget

Per-file footprint is exactly `N × bufferBytes` (one allocation). Concurrent files multiply it, so a
global cap is mandatory (today's implicit `workers × 2 × buffer` is already the CW-26 risk; depth-N makes
it explicit and bounded).

- New setting `fileOperations.crossFsBridgeMaxBufferMemoryMB` (default e.g. 256 MiB).
- `perFileBudgetBytes = min(globalBudget / max(1, activeBridgeWorkers), hardPerFileCap)`.
- A global atomic reservation: a worker reserves `N × bufferBytes` before allocating; if the reservation
  fails, it **degrades** (halve N, retry; floor at serial). Release on file completion.
- Because depth is clamped by the budget, total bridge buffer RAM never exceeds the cap regardless of
  concurrency — directly closing CW-26.

## A.5 Invariants preserved (must not regress)

- **Byte-count integrity gate** (`:8485`) unchanged — still compares `fileCompletedBytes` to `GetSize`.
- **Order** — FIFO guarantees in-order writes.
- **Overflow checks** on `fileCompletedBytes` / `overallCompletedBytes` stay in the writer.
- **Short reads tolerated, short writes looped** — reader accepts `n < bufferBytes` (not EOF unless
  `n==0`); writer drains via `written` and treats `written==0` as `ERROR_WRITE_FAULT` (unchanged logic,
  new location).
- **CW-2 clamp** added: reader rejects `n > bufferBytes`.
- **Progress + bandwidth throttle** remain in the single writer thread ⇒ progress stays monotonic and
  single-sourced; no new locking.
- **Single writer** ⇒ the destination `IFileWriter` is only ever touched by one thread (contract-safe).

## A.6 Failure & cancellation semantics

| Event | Handling |
|-------|----------|
| Reader `Read` fails | set `aborted`+`firstError`, notify both cvs; writer returns `firstError`; `scope_exit` deletes temp |
| Reader returns `n > bufferBytes` | `ERROR_INVALID_DATA` via the abort path (CW-2) |
| Writer `Write` fails / `written==0` | set `aborted`+error, notify both; reader wakes on `freeCv`, sees abort, exits; then `join()` |
| Cancel / pause | both loops poll `CancelRequested()`/`WaitWhilePaused()`; cancel sets `aborted` so neither side can block — this also fixes the CW-18 join-hang **by construction** (abort is always signalled before `join()`) |
| Allocation / reservation failure | degrade depth (halve, floor at serial); never fail the copy for RAM reasons |
| EOF | reader sets `eofReached` after its last real chunk; writer drains `filledQ` then returns `S_OK` |

The one residual (same as today): a reader parked *inside* a long blocking plugin `Read` cannot be
interrupted mid-call; `aborted` is honored as soon as `Read` returns, and `join()` waits for that. This
is a plugin-contract limitation, not a pump defect; R-1b's shorter concurrent requests reduce the worst-
case parked time.

## A.7 R-1b (later) — parallel ranged I/O

To actually hide cloud/WAN **request** latency, issue multiple reads concurrently:
- **Multi-reader:** a small pool of reader threads, each doing a ranged read at a distinct offset;
  results reassembled in order before the writer (or written at offset). Depth then means "outstanding
  requests," which is the real lever for S3/OneDrive TTFB.
- **Multi-writer:** parallel multipart part uploads (S3) / parallel chunk PUTs (OneDrive).
- **Contract impact:** needs offset-addressed reads (`ReadAt(offset,…)`) and either offset-addressed
  writes or an in-order reassembly buffer — hence a separate, larger change gated behind R-1a landing and
  the capability being advertised per plugin. Backends that cannot do ranged/parallel I/O (7z stream,
  Curl single stream) stay on R-1a automatically.

## A.8 Test plan

- **Fakes with injected latency:** Dummy reader/writer that sleep a configurable per-call latency and
  per-call jitter; assert wall-clock for a multi-chunk transfer approaches `max(readTime, writeTime) +
  oneLeg` (overlap) rather than `readTime + writeTime` (serial). Parameterize by depth.
- **Backpressure:** slow writer + fast reader ⇒ reader blocks at ≤ N filled buffers (never unbounded);
  slow reader + fast writer ⇒ writer blocks, no spin.
- **Correctness:** short-read reader (returns < requested mid-stream) still produces a byte-exact
  destination; `n > bufferBytes` reader ⇒ `ERROR_INVALID_DATA`, temp deleted (CW-2 RED-first).
- **Cancel/pause:** cancel mid-transfer at each depth ⇒ prompt unwind, `join()` returns within deadline,
  temp deleted (also the CW-18 regression guard).
- **Memory cap:** N concurrent large-file copies ⇒ peak bridge RSS ≤ the configured cap; depth degrades
  under pressure instead of over-allocating (CW-26 guard).
- **Depth selection:** unit-test `depthForLatency` and the max-of-endpoints merge for every class pair.
- **Perf evidence:** before/after matrices for local↔local, local↔S3, FTP↔OneDrive archived under
  `Specs/TestRuns/` per `Testing_PerformanceValidation.md`.

## A.9 Integration points

- Replace the 2-slot block `State.cpp:8256-8478` with the `PumpChannel` reader/writer loops.
- Replace `ShouldUseBufferedPipeline` (`:7079`) with `ResolvePumpDepth(...)` returning `N` (1 = serial).
- Extend `ResolveAdaptiveCrossFsBridgeBufferBytes` (`:2046`) to also surface the merged `latencyClass`
  (rename to `ResolveCrossFsBridgeTransferPlan` returning `{bufferBytes, depth}`), and drop the
  "default-only" adaptation gate (CW-21).
- Add the global memory reservation (atomic + setting) and wire the degrade path.
- Keep the existing serial `copySerial` (`:8175`) as the `N==1` path and the allocation-failure fallback.

**Subsumes on landing:** CW-11 (OneDrive serial ranges — via R-1b), CW-14 (S3 synchronous parts — via
R-1b decoupling), CW-18 (join-hang — by construction), CW-21 (dead hints / adaptation gate), CW-26
(unbounded memory — explicit budget). Update `Specs/Core/Core_FileSystemBridge.md` §"Buffering & adaptive
sizing" and §"Concurrency model" accordingly.
