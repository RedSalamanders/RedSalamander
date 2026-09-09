# Advisor Plan 009 - Investigate per-file icon extraction parallelism

> **NON-NORMATIVE RETIRED RECORD.** Archived from root `plans/009-icon-pipeline-investigation.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** RETIRED
- **Archived:** 2026-08-25
- **Original path:** `plans/009-icon-pipeline-investigation.md`
- **Live owner:** Specs/Plans/WIP/Operation_FolderView_WarpDrive_RemainingCloseout_2026-08-25.md (I2 remainder; Task 9 is complete in the Done WarpDrive archive)
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 009: Investigate (and only then parallelize) per-file icon extraction in FolderView

> **2026-06-28 consolidation note**: Superseded as an execution source by
> `Specs/Plans/WIP/Operation_FolderView_WarpDrive_RemainingCloseout_2026-08-25.md`. Live code
> already parallelizes extension and per-file icon-index lookup in
> `RedSalamander/FolderView.Enumeration.cpp`; the remaining question is the
> grouped icon bitmap extraction/conversion queue in `FolderView.Icons.cpp`, and
> only if current metrics show it is material.
>
> **Frozen historical instructions (do not execute)**: This is an INVESTIGATE-FIRST plan. The deliverable
> is a measurement-backed report plus a go/no-go recommendation; implementation
> happens only if the measurements confirm the bottleneck AND the change stays
> inside the stated boundaries. Follow steps in order; honor STOP conditions.
> Historical: status-row updates no longer apply in `plans/README.md`.
>
> **Drift check (run first)**: `git diff --stat d6bfccc42..HEAD -- RedSalamander/FolderView.Icons.cpp RedSalamander/IconCache.cpp Tests/PerformanceTests2`
> On drift, re-verify the excerpts below before proceeding.

## Status

- **Priority**: P3
- **Effort**: M (investigation S; implementation M if green-lit)
- **Risk**: MED (threading change in icon pipeline if implemented)
- **Depends on**: none (008 is unrelated; can run in parallel)
- **Category**: perf
- **Planned at**: commit `d6bfccc42`, 2026-06-11

## Why this matters

Audit signal (MED confidence — NOT yet verified, hence investigate-first): folders heavy in per-file-icon types (.exe, .lnk, .ico) fill icons slowly because per-file icon extraction appears to run serially on a single worker. Shell icon extraction is ~milliseconds per file, so 500 .exe files ≈ seconds of placeholder icons. If the serial claim is confirmed, batching across a small MTA worker pool is a visible UX win. If it is not confirmed, the correct outcome of this plan is a short "no-go, here's why" report — that is a success, not a failure.

## Current state (verified excerpts)

- `RedSalamander\FolderView.Icons.cpp:855-905` — the thumbnail pipeline has a cancel mechanism (`CancelThumbnailLoading`, batch ids, `_thumbnailLoadQueue` guarded by `_enumerationMutex`, condition variable `_enumerationCv`) and a per-worker-session WIC factory (`EnsureThumbnailWicFactory`, which caches into a caller-owned `wil::com_ptr` and counts creations via `_thumbnailLoadStats.wicFactoryCreate`). `ProcessThumbnailLoadQueue()` begins at :907.
- Per-item icon state: `FolderView.h:799` — `int iconIndex = -1; // System image list icon index from SHGetFileInfo`.
- The per-FILE icon queue (distinct from thumbnails) is in the same file/area — locate `ProcessIconLoadQueue` or equivalent (`Grep -n "IconLoadQueue\|PerFile" RedSalamander\FolderView.Icons.cpp`).
- `RedSalamander\IconCache.cpp` — shell icon extraction; repo rule (AGENTS.md): `IconCache::Initialize` is UI-thread/STA; worker threads calling `IconCache::ExtractSystemIcon()` MUST init COM as MTA (`wil::CoInitializeEx(COINIT_MULTITHREADED)`).
- Existing perf tests targeting exactly this area: `Tests\PerformanceTests2\FolderIconEnumerationPerfTest.cpp` and `FolderIconEnumerationDuplicatePathPerfTest.cpp` — read both before measuring; they are your measurement harness.
- Mandatory process: `.github\skills\perf-validation\SKILL.md` + `Specs\Testing\Testing_PerformanceValidation.md` (scenario, instrumentation, archived runs under `Specs\TestRuns\`).

## Commands you will need

| Purpose | Command | Expected on success |
|---------|---------|---------------------|
| Build | `.\build.ps1` | exit 0 |
| Perf tests | vstest (locate per ci.yml:150-167) `& $vstestPath .\.build\x64\Debug\PerformanceTests2.dll` | all pass |
| Full selftests | Start-Process `--selftest` pattern | exit 0 |

## Scope

**In scope**:
- Reading anything; measuring; a written report `plans/009-findings.md`
- IF green-lit by your own measurements: `RedSalamander\FolderView.Icons.cpp` (+ `FolderView.h` for queue/worker state), a new or extended perf test in `Tests\PerformanceTests2\`, archived runs in `Specs\TestRuns\`

**Out of scope** (even if green-lit):
- `IconCache.cpp` internals and its STA/MTA contract — consume it per the rules, don't change it.
- The thumbnail pipeline (separate queue) — only the per-file ICON queue.
- Any UI-thread behavior change; enumeration logic (`FolderView.Enumeration.cpp`).

## Git workflow

- Branch: `advisor/009-icon-pipeline` (only needed if implementation proceeds; the report alone can live on the branch too)
- Do NOT push or open a PR unless the operator instructed it.

## Steps

### Step 1: Read and map the per-file icon path

From `FolderView.Icons.cpp`, document: how items enter the per-file icon queue, which thread drains it, whether extraction calls (`IconCache::ExtractSystemIcon` or `SHGetFileInfo`) happen one-at-a-time, how results post back to the UI thread (per AGENTS.md this must be the `PostMessagePayload` mechanism — note which window/message), and how cancellation works. Write this map into `plans/009-findings.md`.

**Verify**: the findings file names the exact function that loops over queued items and the thread it runs on.

### Step 2: Measure

Run `FolderIconEnumerationPerfTest` (and the duplicate-path variant) as the baseline. Then measure the scenario the audit flagged: a folder with many distinct per-file-icon files. If the existing tests don't cover it, create a synthetic fixture (e.g. copy a few hundred distinct small .exe/.lnk files into a temp dir inside the test — check how existing perf tests build fixtures and match that). Capture per-item extraction time and total queue-drain wall time via the existing `_thumbnailLoadStats`-style counters or simple QPC timing in the test.

**Verify**: numbers exist for (a) per-item cost, (b) total drain time, (c) what fraction of drain time is the extraction call itself vs. queue/post overhead. Archive the baseline per the perf-validation skill.

### Step 3: Decide

GO if: extraction is serial AND per-item cost ≥ ~2ms AND total drain for a 500-item scenario ≥ ~1s AND extraction calls are independent (no shared mutable state in the loop other than the results map). Otherwise NO-GO: finish `plans/009-findings.md` with the numbers and the reason, update plans/README.md status to DONE (investigated, no-go), and stop here.

### Step 4 (GO only): Implement bounded parallelism

- Use the Windows thread pool (`TrySubmitThreadpoolCallback` or a `wil` threadpool wrapper — AGENTS.md prefers threadpool over detached threads) with a small cap (e.g. min(4, hardware/2)) draining the same queue; each worker: MTA COM init per the IconCache contract, extract, then hand results to the existing UI-thread posting path unchanged.
- Cancellation: the existing batch-id pattern (`FolderView.Icons.cpp:855-866`) must short-circuit workers; teardown must wait for in-flight workers before the FolderView is destroyed (find the existing worker-shutdown path and extend it; an orphaned callback touching a dead `this` is the failure mode to design against — document your ownership answer in the findings file).
- Per AGENTS.md, any cross-thread payload change requires reviewing: payload ownership, UI-thread boundary, cancellation path, teardown drain, and a teardown stress selftest if the host can queue payloads — add/extend that selftest if you touch the posting path.

**Verify**: build exit 0; full `--selftest` exit 0; run the app against the fixture folder if a desktop session exists (icons fill visibly faster, no crashes on rapid navigation away mid-load — navigate in/out of the folder 20× as a manual stress).

### Step 5 (GO only): After-evidence

Re-run Step 2 measurements; archive after-run under `Specs\TestRuns\`; require ≥2× drain-time improvement on the fixture scenario to keep the change — otherwise revert to the report-only outcome (the complexity isn't paid for).

## Test plan

- Investigation: measurements via `Tests\PerformanceTests2` (existing + possibly one new fixture-based test, modeled on `FolderIconEnumerationPerfTest.cpp`).
- Implementation (GO only): the new/extended perf test becomes the regression guard; full selftest; teardown stress per AGENTS.md.

## Done criteria

- (archived, not live)  `plans/009-findings.md` exists with the thread map + numbers + GO/NO-GO and rationale
- (archived, not live)  NO-GO path: nothing else changed (`git status` clean apart from the findings file); status DONE in plans/README.md
- (archived, not live)  GO path: build + full selftest + PerformanceTests2 green; before/after runs archived; ≥2× improvement documented
- (archived, not live)  `plans/README.md` status row updated

## STOP conditions

- The queue drain turns out to be shared with thumbnail loading in a way that parallelizing one disturbs the other.
- Worker-lifetime/teardown cannot be made safe without restructuring FolderView ownership (anything beyond extending the existing shutdown path).
- `IconCache` proves not thread-safe for concurrent `ExtractSystemIcon` calls (check for shared mutable state inside it FIRST — if unsafe, that's the report's headline, not something to fix here).
- Measured per-item cost is dominated by the UI-thread posting, not extraction — parallel extraction won't help; report that instead.

## Maintenance notes

- If implemented: the worker cap and the batch-id cancellation are the two knobs reviewers should scrutinize; future icon-pipeline changes must preserve the teardown drain.
- The findings file should be deleted or folded into `Specs\` docs once acted on (repo convention: durable contracts move to `Specs\<Domain>\`, per AGENTS.md spec-closeout).
