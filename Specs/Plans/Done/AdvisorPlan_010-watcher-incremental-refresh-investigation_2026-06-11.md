# Advisor Plan 010 - Investigate incremental folder refresh

> **NON-NORMATIVE RETIRED RECORD.** Archived from root `plans/010-watcher-incremental-refresh-investigation.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** RETIRED
- **Archived:** 2026-08-25
- **Original path:** `plans/010-watcher-incremental-refresh-investigation.md`
- **Live owner:** Specs/Plans/WIP/Operation_FolderView_WarpDrive_RemainingCloseout_2026-08-25.md (I2 remainder; Task 8 is complete in the Done WarpDrive archive)
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 010: Investigate incremental folder refresh on FolderWatcher events (design spike — no implementation)

> **2026-06-28 consolidation note**: Superseded as an execution source by
> `Specs/Plans/WIP/Operation_FolderView_WarpDrive_RemainingCloseout_2026-08-25.md`. Live code
> already has DirectoryInfoCache refresh-post coalescing, FolderView dirty-cache
> debounce, and same-folder refresh state preservation. The remaining work is to
> measure and close or extend that existing path, not to design incremental
> refresh from scratch.
>
> **Frozen historical instructions (do not execute)**: This plan produces MEASUREMENTS and a DESIGN
> DOCUMENT only. Do not implement incremental refresh — the audit's confidence
> here is MED and the implementation risk is HIGH; the design doc is the
> deliverable a maintainer reviews before any code happens. Honor STOP
> conditions. Historical: status-row updates no longer apply in `plans/README.md`.
>
> **Drift check (run first)**: `git diff --stat d6bfccc42..HEAD -- RedSalamander/FolderWatcher.cpp RedSalamander/FolderView.Enumeration.cpp RedSalamander/FolderView.h`
> On drift, re-verify against live code before proceeding.

## Status

- **Priority**: P3
- **Effort**: M (spike only)
- **Risk**: LOW (read/measure/write-doc only)
- **Depends on**: none
- **Category**: perf (investigate)
- **Planned at**: commit `d6bfccc42`, 2026-06-11

## Why this matters

Audit signal (MED confidence): when the folder watcher reports a change, FolderView appears to replace the entire item list — full re-enumeration, re-layout (icons, text layouts, sort), and full repaint — even for a single file creation. During bulk file operations (copy of 1,000 files into a watched folder of 5,000), that would mean repeated full rebuilds and visible jank. A file manager's perceived quality lives exactly here. But incremental diff/refresh is a genuinely hard change (sorting, selection preservation, icon reuse, in-flight enumeration races), so the cheap, correct first step is: measure how bad it actually is, map the current pipeline, and produce a reviewable design with explicit risks — following the repo's own planning convention (`Specs\Plans\WIP\`).

## Current state (verified anchors — the spike fills in the rest)

- `RedSalamander\FolderWatcher.cpp` — the change-notification source (read it; document its event granularity: per-change vs. coalesced, what payload it delivers).
- `RedSalamander\FolderView.Enumeration.cpp` — enumeration pipeline. Verified at plan time: payload-message draining exists (`DrainPendingEnumerationPayloadMessages` around :249-270 per audit; confirm), enumeration produces a full item set, and the extension→icon cache uses a case-insensitive transparent hash (`WStringViewHash`, :10-32).
- `RedSalamander\FolderView.h:794-810` — per-item cached state that a full rebuild throws away each time: D2D bitmaps (icon, thumbnail), `IDWriteTextLayout`s, metrics, `iconIndex`.
- Existing related perf tests: `Tests\PerformanceTests2\FolderViewRefreshDuplicatePathPerfTest.cpp` (read it — it measures something adjacent: refresh with duplicate paths).
- Process docs: `.github\skills\perf-validation\SKILL.md`, `Specs\Testing\Testing_PerformanceValidation.md`, and the repo planning convention: durable design docs live in `Specs\Plans\WIP\` and get moved to `Specs\Plans\Done\` when implemented (AGENTS.md "Spec Closeout is Mandatory").
- Relevant UI spec to stay consistent with: `Specs\UI\UI_FolderView.md`.

## Commands you will need

| Purpose | Command | Expected on success |
|---------|---------|---------------------|
| Build | `.\build.ps1` | exit 0 |
| Perf tests | vstest per ci.yml:150-167 against `PerformanceTests2.dll` | all pass |

## Scope

**In scope**:
- Reading; temporary local instrumentation for measurement IF needed (must be reverted before finishing — `git status` clean at the end except the deliverables);
- Deliverable 1: measurements summary inside the design doc;
- Deliverable 2: `Specs\Plans\WIP\UI_FolderView_IncrementalRefresh.md` (the design doc);
- A status row + pointer in `plans/README.md`.

**Out of scope**:
- ANY production code change that ships (instrumentation must be reverted).
- Changing FolderWatcher semantics.
- The icon pipeline internals (plan 009's territory) beyond describing the interaction.

## Git workflow

- Branch: `advisor/010-incremental-refresh-spike`
- The only committed changes: the design doc + plans/README.md row.
- Do NOT push or open a PR unless the operator instructed it.

## Steps

### Step 1: Map the pipeline

Trace, with file:line references: watcher event → (coalescing? debounce?) → message to FolderView → enumeration trigger → item-list replacement → what per-item state is rebuilt (sort, icons, layouts, selection?) → invalidation/repaint extent. Answer explicitly: (a) is there ANY existing reuse of previous items (e.g. icon index carry-over)? (b) how is the user's selection/focus/scroll position preserved across a refresh today (there must be some mechanism — find it; it's the seed of any diff design)? (c) is enumeration async, and what happens when a second watcher event arrives mid-enumeration?

**Verify**: each arrow in the chain has a file:line in your notes.

### Step 2: Measure the cost

Scenario per the perf-validation skill: watched folder with 5,000 items; create/delete/rename one file; measure wall time from watcher event to paint-complete, plus how much per-item state was rebuilt (counters or temporary QPC logging). Repeat for a burst (100 changes in 2s — simulates a copy into the folder). Use or extend `FolderViewRefreshDuplicatePathPerfTest.cpp` style harness where possible; otherwise temporary instrumentation (reverted afterwards). Record whether coalescing already bounds the burst case.

**Verify**: numbers for single-change and burst scenarios exist (before-state archived per skill convention if you used the perf-test harness).

### Step 3: Write the design doc

`Specs\Plans\WIP\UI_FolderView_IncrementalRefresh.md`, containing:
1. Current-pipeline map (Step 1) and measured costs (Step 2).
2. Verdict: is incremental refresh warranted? If coalescing + cheap rebuilds make single-change cost < ~30ms paint-to-paint, recommend NO and stop the doc at a short rationale — that outcome is fully acceptable.
3. If warranted — a phased design (sketch level, for maintainer review):
   - Phase A (cheapest): preserve per-item expensive state across rebuilds by keying old items by path and carrying over `iconIndex`/bitmaps/layouts for unchanged (name, size, mtime, attrs) items — list replacement stays, only the caches survive.
   - Phase B: true add/remove/update diff with stable sort insertion and partial invalidation.
   - For each phase: what can break (selection, sort stability, in-flight enumeration races, plugin filesystems whose enumeration isn't cheap), test strategy (which selftest suite covers it; new perf scenario), and an effort estimate.
4. Open questions for the maintainer (e.g. is Phase A enough? does the DirectoryInfoCache (`RedSalamander\DirectoryInfoCache.cpp`) already play a role here?).

**Verify**: the doc exists, ≤ ~3 pages, every claim in it carries a file:line or a measured number.

### Step 4: Clean up

Revert all instrumentation. `git status` shows only the design doc (+ plans/README.md).

## Test plan

None (spike). The design doc must DEFINE the test plan for the eventual implementation.

## Done criteria

- (archived, not live)  `Specs\Plans\WIP\UI_FolderView_IncrementalRefresh.md` exists with pipeline map, numbers, verdict, phased design or no-go rationale, open questions
- (archived, not live)  No production code changes remain (`git diff --stat d6bfccc42..HEAD -- RedSalamander/ Common/` shows nothing from this plan)
- (archived, not live)  `plans/README.md` status row updated with a one-line verdict

## STOP conditions

- You cannot trigger watcher events deterministically in a measurable way after two approaches — report the obstacle.
- The pipeline turns out to already do incremental reuse (audit was wrong) — write the short no-go doc citing where, and finish early; that IS the deliverable.
- Measurement requires changes you can't cleanly revert.

## Maintenance notes

- If the maintainer green-lights a phase, that becomes a NEW plan (with this doc as its "Current state") — do not extend this one.
- The doc follows the repo's spec-closeout rule: if implemented later, it moves to `Specs\Plans\Done\` and its contracts merge into `Specs\UI\UI_FolderView.md`.
