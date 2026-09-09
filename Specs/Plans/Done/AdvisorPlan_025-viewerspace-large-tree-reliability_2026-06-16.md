# Advisor Plan 025 - ViewerSpace bounded scan model

> **NON-NORMATIVE COMPLETE RECORD.** Archived from root `plans/025-viewerspace-large-tree-reliability.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** COMPLETE
- **Archived:** 2026-08-25
- **Original path:** `plans/025-viewerspace-large-tree-reliability.md`
- **Live owner:** Specs/Plans/Done/Operation_Farsight_ViewerPluginsRemediation_ExportSafetyDecodeReliabilityWebSecurityAndTextGeometry_2026-06-16.md
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 025: ViewerSpace — bound live scan-model memory on huge trees (and reclaim arena waste)

## Status

- **Status**: DONE — implementation, focused archive, consolidated Farsight archive, and repository-wide Full gate green (2026-07-12)
- **Priority**: P2 (reliability — unbounded working set scanning a drive root)
- **Effort**: L (memory cap) / S (arena reclaim)
- **Risk**: MED (changes the in-memory model; must not corrupt size roll-ups or the treemap)
- **Depends on**: none
- **Category**: perf / reliability
- **Planned at**: commit `b274022d9`, 2026-06-16 — **anchors refreshed at `a72512919`, 2026-06-21** (reconcile): the only change to `ViewerSpace.cpp` since planning was a 3-line monotonic root-total guard in `DrainUpdates()` (~`:7118`), unrelated to this finding. Line anchors below `:7118` are shifted +3 below; the unbounded-live-model finding is unchanged.

### 2026-07-11 implementation reconciliation

Operation Farsight F-S2-1 expanded this original model-memory plan after re-reading the provider ABI and traversal lifecycle. The implemented slice adds a single checked scan policy, exact aggregation under retained ancestors, bounded outstanding/traversal/path state, two-pass hostile-buffer validation before model mutation, ancestor identity cycle rejection where nonzero provider ids exist, saturating accounting, whole-candidate item-id preflight, and reusable child-arena blocks. The original optional hardlink-dedup idea was not pursued because it is a distinct semantic/memory tradeoff and is not required for bounded itemization. Scope necessarily includes `ViewerSpace.ScanPolicy.h`, the real `ViewerPETests` harness, the authoritative ViewerSpace spec, and focused archived performance evidence; the superseded “only cpp/h modified” criterion no longer applies. Scoped zero-warning builds and deterministic regressions close the focused F-S2-1 slice.

## Why this matters

ViewerSpace is a disk-space scanner. Its **live** model grows one `Node` per scanned directory (plus
synthetic "Other" buckets), one arena slot per child, and one `FileRecord` per retained top file, with
**no upper bound**. The only size guard in the file — `MaxScanCacheSnapshotBytes` (64 MiB) — gates only
whether a *reusable snapshot is persisted*; it does not bound the live model. Scanning a volume with
millions of directories (e.g. `focusedPath` = a drive root) can grow `_nodes` to millions of entries
plus arenas, consuming hundreds of MB to GBs of working set and, under memory pressure, pushing the
process toward `bad_alloc`/termination. The user can never see most of those nodes (the treemap/draw
layer already does level-of-detail), so the cost is paid for nothing. This is the one ViewerSpace
finding worth blocking on for "scan C:\".

A secondary, cheap win: the children arena permanently leaks roughly half its capacity via abandoned
grow blocks.

## Current state

```cpp
// ViewerSpace.cpp:7281-7309  — the ONLY size guard gates the persisted snapshot, not the live model
const uint64_t maxSnapshotBytes = MaxScanCacheSnapshotBytes();
uint64_t estimatedSnapshotBytes = static_cast<uint64_t>(
    _childrenArena.size() * sizeof(uint32_t) + _nodes.size() * sizeof(ScanResultCacheNode) +
    _fileRecords.size() * sizeof(ScanResultCacheFileRecord));
// ... sums name bytes ...
if (estimatedSnapshotBytes > maxSnapshotBytes) { /* skip persisting snapshot */ return; }
```

- `_nodes` grows via `_nodes.resize(node.id + 1)` per `AddChild` (around `:7000-7021`) with `node.id`
  monotonically increasing (`:6478`/`:6575`); `_childrenArena` grows per child (`AddRealNodeChild`,
  `~:7663-7733`); `_fileRecords` grows per retained top file. None is capped.
- Bytes are rolled up to parent nodes and to the root total during the scan (`~:6630-6724`) — so an
  aggregate "rolled-up but not individually modeled" representation is already partially supported.
- The arena's grow logic abandons the old block instead of reclaiming/reserving (the ~2× waste,
  `~:7681-7723`).

## Commands you will need

| Purpose | Command | Expected |
|---|---|---|
| Build plugin | `.\build.ps1 -ProjectName ViewerSpace` | exit 0, 0 warnings/errors |
| Full build | `.\build.ps1` | exit 0 |
| Tests | `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` | all pass |

## Scope

**In scope**: `Plugins/ViewerSpace/ViewerSpace.cpp`, `Plugins/ViewerSpace/ViewerSpace.h` (model +
scan accounting). **Out of scope**: the treemap/draw LOD; the reparse-point loop guard (already
correct); the worker-thread/PendingUpdate machinery (already race-free).

## Git workflow

- Branch: `advisor/025-viewerspace-large-tree`
- Message: `perf(ViewerSpace): cap live scan-model memory; reclaim children arena`
- Do NOT push/PR unless instructed.

## Steps

### Step 1: Add a live-model memory/node cap with graceful aggregation

Introduce a bound on the live model (a node-count cap and/or an estimated-bytes cap; reuse the same
estimate formula already at `:7282-7284`). When the bound is reached during a scan, **stop creating
fine-grained child nodes for further-descended directories** and instead keep rolling their byte totals
up into their nearest modeled ancestor (the roll-up at `~:6630-6724` already accumulates sizes), marking
that ancestor as "aggregated/truncated" so the UI/treemap can show a "+N more (not itemized)" affordance
rather than silently dropping data. The reported **total** size must remain correct (bytes still roll up
to root); only per-node *itemization* is bounded.

- Make the cap a named constant (and ideally surface it in the existing config schema if Space has one;
  if not, a constant is fine for v1). Pick a default that keeps the worst case in the low hundreds of MB
  (e.g. cap on `_nodes` count derived from a target byte budget).
- `log()`/`Debug::Perf::EmitValue` when the cap engages (mirror the existing
  `viewer.space.model.cache_skipped_large` perf counter with a new `viewer.space.model.itemization_capped`).

**Verify**: `.\build.ps1 -ProjectName ViewerSpace` → exit 0.

### Step 2: Reclaim the children-arena waste

Fix the arena grow path (`~:7681-7723`) so it does not permanently abandon ~half its capacity on each
grow — e.g. reserve geometric capacity up front, or move/compact existing entries into the new block
(or switch to a `std::vector` with `reserve` growth if the arena is just a manual vector). Keep the
external indices stable (children are referenced by arena index elsewhere) — if indices would move,
STOP and report rather than corrupting references.

**Verify**: `.\build.ps1` → exit 0.

### Step 3 (optional, decide with maintainer): hardlink double-counting

Hardlinked files are counted once per link, inflating reported folder/total sizes (`~:6497-6500`,
`6573-6586`). Deduping requires tracking `(volumeSerial, fileIndex)` per file — which itself costs
memory, in tension with Step 1's cap. If pursued, gate it behind a config option and a bounded de-dup
set, and document the memory cost. If the tradeoff is unclear, **leave this out** and record it in the
backlog — do not let it block Steps 1–2.

## Test plan

- Add a scan-accounting test to `Tests/ViewerPETests/ViewerPETests.cpp` (ViewerSpace harness cases
  exist there): scan a synthetic deep/wide directory tree (created in the test) larger than the cap and
  assert (a) the process memory stays bounded / the cap perf-counter fires, and (b) the **reported total
  bytes equal the true total** (itemization may be aggregated, totals must be exact). If creating a tree
  large enough to trip a production-sized cap is impractical in CI, add a test hook to set a tiny cap for
  the test run and assert aggregation + exact totals at that small cap.
- Verification: `.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild` → all pass.

Consolidated proof is archived at
`Specs/TestRuns/4cb089111a23/Viewers/2026-07-12_105903_farsight_viewer_closeout/`: all 13 focused
cases passed with zero conditional skips and 1,426 performance metric records, including the blocked-
provider close/post-update race.

## Done criteria

ALL must hold:

- [x] The live model has an enforced upper bound (node count and/or estimated bytes); exceeding it
      aggregates further descent instead of growing unbounded.
- [x] Reported total/rolled-up sizes remain exact when the cap engages (test asserts it).
- [x] The children arena reuses/coalesces prior blocks and does not strand growth capacity.
- [x] Accepted/rejected/capped/resource perf counters and Debug snapshot fields are archived.
- [x] Scoped `ViewerSpace` and `ViewerPETests` builds exit 0 with 0 warnings/errors.
- [x] Focused helper, hostile-provider runtime, and existing ViewerSpace window regressions pass; the hostile runtime passes three consecutive runs.
- [x] Reconciled implementation scope includes `ViewerSpace.cpp`/`.h`, `ViewerSpace.ScanPolicy.h`, tests, authoritative spec, and archived evidence.
- [x] `plans/README.md` updated.
- [x] The consolidated Farsight closeout archive is complete.
- [x] The repository-wide `.\Tools\Run-AllTests.ps1 -Suite Full` passes in the final closeout workspace (`Specs/TestRuns/4cb089111a23/Continuation/2026-07-12_205636_farsight_full_green/`).

## STOP conditions

- Excerpts don't match live code (drift).
- The arena reclaim would invalidate stable child indices used elsewhere — land Step 1 only and report.
- You cannot keep totals exact under aggregation without a larger model redesign — report; an accurate
  total with capped itemization is the hard requirement.

## Maintenance notes

- The bounded-itemization + exact-totals contract is the thing to preserve; any future model change
  must keep totals correct independent of the cap.
- **Resolved and archived by Operation Farsight F-S1-10 (2026-07-12):** ViewerSpace
  no longer joins an incomplete scan worker during window teardown. Close invalidates the generation,
  requests stop, reaps only completion-signaled workers, and abandons an untrusted-provider-blocked
  worker behind an owning self-reference. Normal scan unload gates release only after `join()` proves
  return; abandonment intentionally retains its gate so `RedSalamanderPluginCanUnloadNow` quarantines
  runtime reload for the process. Scheduler shutdown permanently blocks recreation/new permits. The
  deterministic blocked-provider close/stale-post regression and its consolidated metrics evidence
  are green.
- Reviewer: verify totals correctness at the cap boundary and that arena indices stay valid.
