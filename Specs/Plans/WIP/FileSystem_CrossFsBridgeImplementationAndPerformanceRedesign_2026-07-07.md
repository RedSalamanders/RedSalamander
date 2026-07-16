# Cross-File-System Bridge — Implementation Today & Performance Redesign Thinking

**Status:** WIP — design rationale / redesign backlog (thinking doc, not a fix ledger)
**Date:** 2026-07-07
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

---

## 1. What the bridge is for

When you copy or move between two different plugins (local disk → S3, 7z → local, FTP → OneDrive), no
single plugin can do the transfer — each one only knows its own backend. The bridge is the host acting
as the middleman: it reads bytes from the source plugin and writes them to the destination plugin
itself, through the `IFileReader`/`IFileWriter` streaming contract. Its safety promise is "never leave a
partial or corrupt file at the final path," which it keeps by staging to a temp name and renaming into
place only after the byte count checks out.

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
       └ CreateFileWriter(<final>.rs_tmp_<128-bit CSPRNG>)   [temp name]
         └ PUMP bytes  (serial OR 2-slot pipeline)
           └ writtenBytes == GetSize ?   else ERROR_PARTIAL_COPY, delete temp
             └ writer->Commit()          else delete temp
               └ MoveItem(temp → final)  [rename inside destination FS = "promote"]
                 └ MOVE only: re-open dest, GetSize must match
                   └ SetFileBasicInformation (timestamps, best-effort)
```

A `wil::scope_exit` deletes the temp on every failure path unless the promote already succeeded, so the
*final path* is protected on all shipped backends.

**The pump has two modes.**
- *Serial* (`:8175`): one thread does `Read` into a buffer, then a `Write` loop drains it, then repeat.
- *"Pipelined"* (`:8256`): **exactly two** buffer slots and **one** `std::jthread` reader — the reader
  fills a slot while the calling thread writes the other, coordinated by a single mutex + condition
  variable. It engages only when `fileTotalBytes > bufferBytes`; a file that fits in one buffer always
  runs serial.

**Buffer sizing.** Default 4 MiB, clamped 512 KB–16 MiB. `ResolveAdaptiveCrossFsBridgeBufferBytes`
(`:2046`) queries both endpoints' `GetTransferHints`, takes the **max** of the two
`preferredBufferBytes`, and freezes it for the task — but only *if the user left the setting at
default*; any manual override disables adaptation entirely. Only `preferredBufferBytes` is consumed;
`latencyClass`, `flags`, and `preferredProgressPeriodMs` are collected and then ignored.

**Directory trees** go sequential (plain recursion) or parallel (a stack-based producer feeds a
16-entry admission queue drained by process-wide scheduler workers, each with its own buffer). MOVE is a
*two-pass* operation: copy the whole tree first, then a second cleanup pass deletes each source file —
but only after re-opening both sides and doing a **full byte-for-byte `memcmp`** (`:7396`).

**Per-plugin reality (where the abstraction leaks):**

| Plugin | Reader | Writer | Promote |
|--------|--------|--------|---------|
| Local (Win32) | short only at EOF | temp-sibling + `FlushFileBuffers` | true rename |
| S3 | per-range GET, no version pin | 64 MiB multipart parts, uploaded synchronously **inside `Write`** | server-side `CopyObject` + delete |
| OneDrive | every range GET capped at **1 MiB**, back-to-back | **whole file staged to local `%TEMP%`**, network untouched until `Commit` | Graph move (cross-FS move disabled) |
| Curl (FTP/SFTP/HTTP) | 1 MiB ring (advertises 8 MiB); sizeless `LIST` → `GetSize` S_OK+0 | overwrite deletes dest at create time | delete-backup + rename |
| 7z | routine short reads; read-only | n/a | n/a |
| Google Drive | — (no `IFileSystemIO`) | — | — |

---

## 3. Design-level weaknesses & redesign directions

The correctness bugs are fixable point-by-point in Causeway. The items below are different: the *shape*
of the implementation limits performance, and they warrant a rework rather than a patch.

### R-1 — The "pipeline" is depth-1 and mode-gated; it barely overlaps anything

**Current design.** Two slots, one outstanding read, and it only engages for files larger than one
buffer. At any instant there is *one* chunk being read and *one* being written.

**Why it's weak.** On a high-latency link — exactly where overlap should pay off — a depth-1 pipeline
hides at most one leg's latency; you are still fundamentally ping-ponging read↔write. And it doesn't
engage at all for files ≤ one buffer, so the entire small-file cloud workload runs fully serial. The
read side and the write side are also coupled: a slow S3 multipart upload directly stalls the reader.

**Proposed redesign.** A bounded read-ahead queue with **configurable depth** (N in-flight read buffers,
driven by the endpoint `latencyClass` hint), and **decouple read concurrency from write concurrency** so
the reader keeps filling while the writer drains at its own rate. Engage it for *all* transfers on
high-latency endpoints, not just large files. This is the single biggest throughput lever.
*Subsumes/relates to:* Causeway CW-11 (OneDrive 1 MiB ranges), CW-14 (S3 synchronous parts).

### R-2 — MOVE verification re-transfers the whole file

**Current design.** Before deleting each moved source, the cleanup pass re-opens *both* endpoints and
runs a full byte-for-byte compare, single-streamed through one buffer (`:7491` → `:7396`).

**This is NOT bridge-only.** The exact same verify-re-read exists a second time in the **local
FileSystem plugin's own cross-volume move**: a MOVE whose destination is on a different volume (which
includes every local→network-drive and local→UNC move) calls `MoveFileWithProgressW` *without*
`MOVEFILE_COPY_ALLOWED`, gets `ERROR_NOT_SAME_DEVICE` (`Plugins/FileSystem/FileSystem.FileOps.cpp:5544`),
and falls into a plugin copy+delete fallback (`:5549`). The delete phase runs `DestinationMatchesSourceFile`
(`:3572`) → `FileContentsEqual` (`:3364`), which seeks both handles to 0 and re-reads **the full source
and the full destination** in 64 KiB chunks before deleting each source file. So a plain local→SMB MOVE
— no bridge involved — already re-reads every byte of the destination back over the network. (Plain
local→network COPY is unaffected: it is a single `CopyFileEx` pass, `:3785`, with no verify re-read.)

**Why it's weak.** Local→cloud MOVE (bridge) re-downloads everything you just uploaded (and pays cloud
egress); cloud→local re-reads the source a second time; local→SMB MOVE (native plugin) re-reads the
destination back over the network. That is O(2× data) of pure overhead for a safety check,
non-overlapped, even when the copy ran N-way parallel — and it exists in two separate code paths.

**Proposed redesign.** Compute a **streaming hash during the copy** and compare it against a destination
hash — ideally one the backend already returns (S3 gives an MD5/ETag on upload), otherwise a single
destination-side read. Never re-read the source. Apply it in **both** implementations — the bridge
(`State.cpp:7396`) *and* the local plugin's cross-volume move fallback (`FileSystem.FileOps.cpp:3364`) —
or they will drift. **Coordinate with Clearwater CW-P1**, which already proposes a copy-time
`CopiedEntry` hash — adopt that one scheme, not a third. *Subsumes:* Causeway CW-9.

### R-3 — Strictly per-file with per-file setup cost — punishing on cloud + small files

**Current design.** Every file independently creates a reader and a writer; there is no batch path and
no cross-file pipelining. On S3 a single small file costs ~10–15 API round-trips (ancestor probes, a
re-probe at commit, a server-side copy+delete to promote).

**Why it's weak.** For a large tree of small files on an object store, the request overhead dwarfs the
byte transfer entirely — the actual data is a rounding error. The temp-then-promote model is itself a
poor fit for stores that have **no rename**: "atomic promote via `MoveItem`" becomes CopyObject+delete,
doubling server-side work.

**Proposed redesign.** A cloud-aware path that (a) writes **directly to the final key** — multipart
complete *is* the atomic point, no temp-then-copy — when no conflict prompt is pending; (b) reuses the
writable-target check between create and commit instead of re-probing; and (c) allows **concurrent file
uploads** (fan-out across files) rather than only depth-within-one-file. Consider surfacing a real batch
interface for object stores.
*Subsumes/relates to:* Causeway CW-12 (round-trip storm), CW-16 (server-side copy promote).

### R-4 — A producer runs *on* a worker that then blocks waiting for other workers

**Current design.** `CopyDirectoryParallel` starts a nested scheduler job and blocks in `WaitJob`
(`:9165`), which has no work-stealing and no cancel predicate.

**Why it's weak.** This is both the process-wide deadlock (Causeway CW-3) *and* a fragile structure even
when it doesn't deadlock — worker threads are a finite pool, and occupying one to wait on the same pool
is self-throttling.

**Proposed redesign.** Don't nest blocking jobs. Run file workers at the same scheduler level as the
producer, or make the wait *participate* in dequeuing (help-execute) and honor cancellation.
*Subsumes:* Causeway CW-3.

### Smaller design smells (fold into the above)

- **Adaptive buffer switches off the moment a user touches the setting** (`:2046`) — a surprising,
  probably unintended coupling. Adaptation and the user's max should compose, not be mutually exclusive.
  *(Causeway CW-21.)*
- **The transfer-hint struct is ~80% dead** — `latencyClass` would tell you exactly when to go
  deep-pipeline vs. serial (R-1), and it is ignored. Wiring it up is a prerequisite for R-1 doing the
  right thing automatically. *(Causeway CW-21.)*

---

## 4. Suggested sequence

1. **R-1 (pump redesign)** first — highest performance return for the least structural risk; it is local
   to the pump and testable with a fake high-latency reader/writer. Wire up `latencyClass` (the smells)
   as part of it.
2. **R-2 (hash-verify)** next, coordinated with Clearwater CW-P1 — removes the largest single chunk of
   wasted I/O on MOVE.
3. **R-4 (scheduler)** — must happen regardless (it is also a deadlock); do it before leaning harder on
   parallel directory copy.
4. **R-3 (cloud batch / direct-final-key)** last — the biggest surface area and the most backend-specific
   work; sequence it after the pump and verify models are settled.

Correctness/security fixes in Causeway (CW-1 path traversal, CW-2 buffer clamp, CW-4 Curl sizeless
`GetSize`) are independent of this rework and should land first — they are small and data-safety-
critical.

## Closeout Rule

Each redesign that ships MUST update `Specs/Core/Core_FileSystemBridge.md` to describe the new behavior,
carry before/after performance evidence per `Specs/Testing/Testing_PerformanceValidation.md` archived
under `Specs/TestRuns/`, and mark the Causeway rows it subsumes as satisfied.

---

# Appendix A — R-1 Detailed Design: depth-N decoupled pump

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
