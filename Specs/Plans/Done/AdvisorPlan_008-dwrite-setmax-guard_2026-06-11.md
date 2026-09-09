# Advisor Plan 008 - Skip redundant DWrite SetMaxWidth/Height calls

> **NON-NORMATIVE RETIRED RECORD.** Archived from root `plans/008-dwrite-setmax-guard.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** RETIRED
- **Archived:** 2026-08-25
- **Original path:** `plans/008-dwrite-setmax-guard.md`
- **Live owner:** Specs/Plans/Done/FolderView_UpdateItemTextLayouts_Optimization_2026-06-19.md
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 008: Skip redundant IDWriteTextLayout SetMaxWidth/SetMaxHeight calls in FolderView item layout

> **Frozen historical instructions (do not execute)**: This plan is archived; do not follow these steps. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. Historical: status-row updates no longer apply for this plan
> in `plans/README.md`.
>
> **Drift check (run first)**: `git diff --stat a72512919..HEAD -- RedSalamander/FolderView.h RedSalamander/FolderView.Layout.cpp RedSalamander/FolderView.Rendering.cpp Tests/PerformanceTests2`
> If any in-scope file changed since this plan was written, compare the
> "Current state" excerpts against the live code before proceeding; on a
> mismatch, treat it as a STOP condition.

## Status

- **Priority**: P2
- **Effort**: S (code) + perf-evidence overhead (mandatory in this repo)
- **Risk**: LOW-MED (stale-constraint bugs if invalidation paths are missed)
- **Depends on**: none
- **Category**: perf
- **Planned at**: commit `d6bfccc42`, 2026-06-11 — **refreshed at `a72512919`, 2026-06-21** (reconcile): the 2026-06-19 "MetricPilot" layout decomposition (`Specs/Plans/Done/FolderView_*MetricPilot_2026-06-19.md`) reworked `FolderView.Layout.cpp`. The finding is unchanged — the cached-layout `SetMaxWidth/SetMaxHeight` calls are still unconditional and there are still **no** cached-constraint guard fields — but the calls now live in **two** functions and line anchors moved by ~50–160 lines. "Current state" below is rewritten against `a72512919`.

## Why this matters

`EnsureItemTextLayout` runs for visible items during painting. When an item's label layout already exists, the code unconditionally calls `SetMaxWidth`/`SetMaxHeight` on the cached `IDWriteTextLayout` — DirectWrite treats a constraint change as layout invalidation, so even no-op-valued calls are not guaranteed free, and during window resize every visible item takes this path repeatedly. Guarding with "only call when the value actually changed" is a classic cheap win for large folders in Details/Thumbnails modes. **This repo mandates perf evidence for hot-path changes** (AGENTS.md "Performance Validation is Mandatory") — half of this plan is producing that evidence, not just the 10-line code change.

## Current state

**Two** functions in `RedSalamander\FolderView.Layout.cpp` create-then-refresh the per-item text layouts, and BOTH refresh the cached layout's constraints unconditionally. The fix (cache the last-applied constraints, guard the Set calls, reset the cache on layout recreate) applies identically to both. Verified unconditional at `a72512919` via `grep -n "SetMax" RedSalamander\FolderView.Layout.cpp` (6 call pairs; no guards present — `grep -rn "labelMaxWidth\|cachedMaxWidth" RedSalamander\FolderView.*` returns nothing).

1. **`FolderView::UpdateItemTextLayouts()`** (`FolderView.Layout.cpp:426-632`) — the batch layout pass added by the MetricPilot decomposition. Cached-layout refresh blocks (each guarded by `if (item.*Layout)` *after* a create-if-missing block):
   - label: `:520-524` — `item.labelLayout->SetMaxWidth(constrainedWidth); item.labelLayout->SetMaxHeight(constrainedHeight);`
   - details: `:578-582` — `…->SetMaxWidth(constrainedWidth); …->SetMaxHeight(constrainedDetailsHeight);`
   - metadata: `:622-626` — `…->SetMaxWidth(constrainedWidth); …->SetMaxHeight(constrainedMetadataHeight);`
2. **`FolderView::EnsureItemTextLayout(FolderItem& item, float labelWidth)`** (`:796-928`) — the per-item path (this is the function the original plan cited at `:747`; it now starts at `:796`). Same three constraints, in `else`/`else if (item.*Layout)` branches:
   - label: `:836-840`; details: `:880-884`; metadata: `:916-920`.

   The `const float constrained{Width,Height,DetailsHeight,MetadataHeight}` locals are computed at `:812-815` (and analogously inside `UpdateItemTextLayouts`).

- Both functions now create layouts through the helper **`FolderView::CreateInstrumentedItemTextLayout(ItemTextLayoutKind kind, …)`** (`:766-794`), which already does `++_frameTextLayoutCreateCount` on success — i.e. **the layout-create counter the original Step 1 told you to "go find" now exists** (added by the MetricPilot work). Reuse `_frameTextLayoutCreateCount` (and any sibling frame counters near it) as the perf metric; do NOT invent a parallel one. The idle-creation paths `ScheduleIdleLayoutCreation()` (`:930`) / `ProcessIdleLayoutBatch()` (`:973`) funnel through these same two functions, so guarding both covers the idle path — but still grep `SetMaxWidth` across all of `RedSalamander\` (Step 2) to confirm no new external call site appeared.
- `RedSalamander\FolderView.h:777` — `struct FolderView::FolderItem`, holding the layout members at `:802-809`:
  ```cpp
  wil::com_ptr<IDWriteTextLayout> labelLayout;      // :802
  DWRITE_TEXT_METRICS labelMetrics{};               // :803
  // … detailsText …
  wil::com_ptr<IDWriteTextLayout> detailsLayout;    // :805
  DWRITE_TEXT_METRICS detailsMetrics{};
  // … metadataText …
  wil::com_ptr<IDWriteTextLayout> metadataLayout;   // :809
  DWRITE_TEXT_METRICS metadataMetrics{};
  ```
  This is where the cached-constraint fields go (read `:777-815` for the exact member ordering/naming style before adding).
- Existing perf tests: `Tests\PerformanceTests2\FolderViewColumnLayoutTests.cpp`, `FolderViewRefreshDuplicatePathPerfTest.cpp`, `FolderIconEnumerationPerfTest.cpp` (vstest-based: `vstest.console.exe .\PerformanceTests2.dll`).
- Mandatory process reading: `.github\skills\perf-validation\SKILL.md` and `Specs\Testing\Testing_PerformanceValidation.md` — scenario definition, deterministic selftest coverage, archived runs under `Specs\TestRuns\`.
- Layout invalidation sites: the layouts are rebuilt when `item.labelLayout` is reset; constraint caches MUST be reset wherever layouts are reset/recreated (find them: `Grep -n "labelLayout" RedSalamander\FolderView*.cpp` — every `.reset()`/assignment site).

## Commands you will need

| Purpose | Command | Expected on success |
|---------|---------|---------------------|
| Build | `.\build.ps1` | exit 0 |
| Perf tests | locate vstest as in ci.yml:150-167, then `& $vstestPath .\.build\x64\Debug\PerformanceTests2.dll` | all pass |
| Full selftests | Start-Process pattern with `--selftest` | exit 0 |
| Format | `.\format-all.ps1` | exit 0 |

## Suggested executor toolkit

- Read `.github\skills\perf-validation\SKILL.md` BEFORE starting — it defines the evidence you must produce.
- Read `.github\skills\direct2d-rendering\SKILL.md` for D2D/DWrite conventions.

## Scope

**In scope**:
- `RedSalamander\FolderView.h` (add cached-constraint fields to the item struct at ~:794-810)
- `RedSalamander\FolderView.Layout.cpp` (the guards; reset of cached constraints where layouts are created)
- Whichever FolderView partial files contain layout-reset sites that must also reset the cached constraints (likely `FolderView.Layout.cpp` itself and possibly `FolderView.Enumeration.cpp`/`FolderView.Rendering.cpp` — only the reset lines)
- `Specs\TestRuns\` (new archived perf evidence per the skill)
- `Tests\PerformanceTests2\` ONLY if a new measurement hook is required by the skill workflow (prefer reusing existing tests)

**Out of scope** (do NOT touch):
- Any other perf idea in this area (metadata-provider batching, icon pipeline, incremental enumeration) — separate plans (009/010) exist; scope creep here is the main risk.
- `Common\DxUi\**` — this is FolderView-local.

## Git workflow

- Branch: `advisor/008-dwrite-setmax-guard`
- Subject style: "Skip redundant text layout constraint updates in FolderView".
- Do NOT push or open a PR unless the operator instructed it.

## Steps

### Step 1: Baseline perf evidence

Follow the perf-validation skill: build Debug, run the FolderView-related PerformanceTests2 tests, and capture the relevant counters/numbers (the skill defines where runs are archived — create the "before" record under `Specs\TestRuns\` named per its convention). If the existing tests don't surface a layout-call counter, record wall-time numbers from `FolderViewColumnLayoutTests` as the baseline and note the limitation.

**Verify**: a "before" artifact exists under `Specs\TestRuns\`.

### Step 2: Add cached constraints + guards

In the item struct (`FolderView.h:777`, layout members at `:802-809`), add:
```cpp
float labelMaxWidth = 0.0f;  float labelMaxHeight = 0.0f;
float detailsMaxWidth = 0.0f; float detailsMaxHeight = 0.0f;
float metadataMaxWidth = 0.0f; float metadataMaxHeight = 0.0f;
```
(match the struct's existing naming style — read neighboring members and follow them). In `FolderView.Layout.cpp`, apply the guards in **both** `UpdateItemTextLayouts()` (`:426-632`) and `EnsureItemTextLayout()` (`:796-928`) — they each have label/details/metadata refresh sites (six pairs total):
- On every `CreateInstrumentedItemTextLayout` success path, store the constraints used.
- In every cached-layout refresh block (the `if (item.*Layout)` / `else if (item.*Layout)` branches), replace unconditional Set calls with:
  ```cpp
  if (item.labelMaxWidth != constrainedWidth) { item.labelLayout->SetMaxWidth(constrainedWidth); item.labelMaxWidth = constrainedWidth; }
  if (item.labelMaxHeight != constrainedHeight) { item.labelLayout->SetMaxHeight(constrainedHeight); item.labelMaxHeight = constrainedHeight; }
  ```
  Exact float equality is correct here (the values come from the same deterministic computation; no epsilon).
- At every site where a layout com_ptr is reset or reassigned outside the create path, reset its cached constraints to 0 (grep per Current state; list the sites you touched in your report).
- IMPORTANT: if any code path calls `SetMaxWidth` on these layouts elsewhere (grep the whole `RedSalamander\` dir for `SetMaxWidth`), those sites must also maintain the cache — if there are more than ~3 such external sites, STOP (the cache belongs in a tiny wrapper instead; report).

**Verify**: `.\build.ps1` → exit 0.

### Step 3: Correctness verification

Full `--selftest` (the commands suite drives real UI through these paths) → exit 0. Then a manual smoke if a desktop session is available: run the app, open a large folder (e.g. `C:\Windows\System32`), switch display modes, resize the window — labels must re-wrap/ellipsize correctly at every width (stale-constraint bugs show up as clipped/overflowing text after resize). If no desktop session is available, state that the manual smoke was skipped.

**Verify**: selftest exit 0.

### Step 4: After perf evidence + archive

Re-run the Step 1 measurements; archive the "after" record under `Specs\TestRuns\` per the skill convention, with a one-paragraph summary (calls avoided / time delta). If the delta is in the noise, SAY SO in the artifact — honest evidence is the requirement, not a win.

**Verify**: the archived run exists and the summary states before/after numbers.

### Step 5: Format and finish

`.\format-all.ps1`, rebuild, rerun PerformanceTests2 → all pass.

## Test plan

- Regression: full `--selftest`, `PerformanceTests2.dll` via vstest (existing layout tests must still pass).
- Evidence: before/after archived runs (Steps 1, 4).
- No new unit tests required (rendering-path logic; selftests + perf tests are the repo's vehicle here).

## Done criteria

- (archived, not live)  All cached-layout `SetMaxWidth/SetMaxHeight` call sites in `FolderView.Layout.cpp` are guarded — both functions (`UpdateItemTextLayouts` and `EnsureItemTextLayout`), all six pairs (grep shows no unguarded `item.*Layout->SetMax` refresh lines)
- (archived, not live)  Cached constraints reset at every layout reset/recreate site (sites listed in report)
- (archived, not live)  `.\build.ps1` exit 0; full `--selftest` exit 0; PerformanceTests2 all pass
- (archived, not live)  Before AND after perf runs archived under `Specs\TestRuns\`
- (archived, not live)  `plans/README.md` status row updated

## STOP conditions

Stop and report back if:

- `SetMaxWidth` on these layouts is called from more than ~3 sites outside `FolderView.Layout.cpp`.
- Any text-rendering selftest or perf test fails after the change.
- The perf-validation skill demands instrumentation that does not exist and would require building new counter plumbing bigger than the fix itself — report the mismatch; the maintainer may accept wall-time-only evidence.
- Visual smoke shows label wrapping artifacts after resize (stale constraints) and one fix attempt doesn't resolve it.

## Maintenance notes

- Anyone adding a new per-item text layout (e.g. a new display-mode line) must follow the same cache-on-create / guard-on-update / reset-on-recreate pattern.
- Reviewer focus: the reset sites — a missed reset produces layouts stuck at an old width after font/DPI/mode changes, which tests may not catch but users will.
- Deferred: metadata/details text *content* recomputation batching (audit finding PERF-06) — separate, unplanned; recorded in plans/README.md.
