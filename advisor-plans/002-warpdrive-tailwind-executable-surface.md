# Plan 002: Make WarpDrive/Tailwind/thumbnail the only live FolderView-perf contract

> **SUPERSEDED IN PART (2026-08-25).** The operator overrode “keep WarpDrive as one SIZE-WAIVER master.” The diary is frozen at `Specs/Plans/Done/Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md`. Live closeout is `Specs/Plans/WIP/Operation_FolderView_WarpDrive_RemainingCloseout_2026-08-25.md` (I2). Retired continuation batons live in Done, not WIP. Do not merge the remainder back into one master. Tailwind/I1 pointer hygiene below remains useful if those files still name the old WIP path.

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `advisor-plans/README.md`.
>
> **Drift check (run first)**: `git diff --stat b5c3b0eab..HEAD -- Specs/Plans/WIP/Operation_FolderView_WarpDrive_RemainingCloseout_2026-08-25.md Specs/Plans/Done/Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md Specs/Plans/WIP/Operation_Tailwind_FolderViewWarpDrivePreMergeReviewRemediation_2026-06-30.md Specs/Plans/WIP/FolderView_ThumbnailBackgroundEnrichmentFollowup_2026-07-04.md Specs/Plans/Done/DxUi_Uia_ContinuationBaton_2026-06-29.md Specs/Plans/Done/FolderView_WarpDrive_ContinuationBaton_2026-06-29.md Specs/Plans/WIP/README.md Specs/UI/UI_FolderView.md`
> If those files drifted, re-read the live Master Checklist and Tailwind sequencing before editing.

## Status

- **Priority**: P0
- **Effort**: M
- **Risk**: LOW for docs; MED if an executor starts implementing TW items or Task 10 scroll-rects — this plan forbids that
- **Depends on**: `advisor-plans/001-wip-lifecycle-and-status-truth.md`
- **Category**: docs
- **Planned at**: commit `b5c3b0eab`, 2026-08-25

## Why this matters

FolderView WarpDrive remaining closeout is the compact I2 remainder (`Specs/Plans/WIP/Operation_FolderView_WarpDrive_RemainingCloseout_2026-08-25.md`). The ~277 KB diary is frozen in Done. Do not keep the master in WIP.

## Current state

### WarpDrive remaining executable surface

`Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md`:

- Master roll-up Tasks 0–9 are `[x]` (`:37-62`).
- Task 10 roll-up is still `[ ]` (`:66`) even though Steps 1–3 are `[x]` and Steps 2–3 record **gate not met — not implemented** (`:1752-1781`). Only Step 4 Verify is `[ ]` (`:1783-1786`).
- Task 11 Closeout is `[ ]` (`:70`, `:1799-1830`).
- Acceptance `[x] Durable contracts are merged` (`:1849`) **contradicts** Task 11 Steps 1–3 still `[ ]`.
- `UI_FolderView.md:298-305` already requires cached-only visible thumbnails and names the enrichment follow-up. Device-loss, brush reuse, paste-shortcut, refresh/icon metrics are also in that spec (search `cached-only`, `paste-shortcut`, `folder.refresh`).
- Thumbnail enrichment is **not** in production: `FolderView.Icons.cpp:123` sets `SIIGBF_INCACHEONLY`; `ExtractProviderAllowedThumbnailWithDeadline` (`:950`) is reached from `ForceProviderAllowedProbe` (`:1224-1229`), not a background enrichment queue.

### Tailwind ownership rot

`Operation_Tailwind_FolderViewWarpDrivePreMergeReviewRemediation_2026-06-30.md`:

- HOLD in the index; body still reads like a pre-merge execution queue.
- TW-5/23/24 routed to `DxUi_Uia_ContinuationBaton` / GR-A1 (`:113-114`, `:182`, `:232-233`). That baton is RETIRED (`DxUi_Uia_ContinuationBaton_2026-06-29.md:1-9`) and forbids reactivation.
- Sequencing step 3 still says wire `result.generation` (`:240-242`) while routing says GR-3 done (`:232`). Live code gates paste-shortcut completion on generation (`FolderView.FileOps.cpp` near `OnPasteShortcutComplete`; `UI_FolderView.md:522-524`).
- TW-2/TW-3/TW-D1 belong to `FolderView_ThumbnailBackgroundEnrichmentFollowup_2026-07-04.md:25-26,102-110`.
- Open Tailwind-only leftovers at review time (re-grep before editing): TW-6, TW-9/10, TW-11, TW-13/14, TW-15, TW-16/17, TW-18, TW-20, TW-21, TW-25/26/27, TW-29, credential-toggle investigation (`:265-268`).

### Retired batons

Keep as compact tombstones. Do not paste diary back into them.

### Repo conventions

WIP is non-normative. Durable FolderView behavior belongs in `Specs/UI/UI_FolderView.md`. Do not add checkboxes to that spec (`NORMATIVE_PLAN_MARKER`).

## Commands you will need

| Purpose | Command | Expected on success |
|---------|---------|---------------------|
| Inventory | `.\Tools\Get-SpecInventory.ps1 -FailOnFindings` | exit 0 |
| Confirm CHK-1/enrichment not secretly done | `rg -n "ExtractProviderAllowedThumbnailWithDeadline|ForceProviderAllowedProbe" RedSalamander/FolderView.Icons.cpp` | deadline helper exists; production visible path still CachedOnly |
| Confirm TW-7 generation gate | `rg -n "OnPasteShortcutComplete|result\.generation" RedSalamander/FolderView.FileOps.cpp` | generation compared before applying |
| Confirm batons still retired in Done | `rg -n "Frozen history: do not resume" Specs/Plans/Done/DxUi_Uia_ContinuationBaton_2026-06-29.md Specs/Plans/Done/FolderView_WarpDrive_ContinuationBaton_2026-06-29.md` | both match |
| Line endings | `git diff --check` | clean |

Do not run Full as part of this plan.

## Scope

**In scope:**
- `Specs/Plans/WIP/Operation_FolderView_WarpDrive_RemainingCloseout_2026-08-25.md`
- `Specs/Plans/WIP/Operation_Tailwind_FolderViewWarpDrivePreMergeReviewRemediation_2026-06-30.md`
- `Specs/Plans/WIP/FolderView_ThumbnailBackgroundEnrichmentFollowup_2026-07-04.md` (Status/closeout pointers only unless a sentence is factually wrong)
- `Specs/Plans/WIP/README.md` (I1/I2/H0 routing sentences only)
- `Specs/UI/UI_FolderView.md` **only** if Task 11 Step 1 list in WarpDrive still claims contracts that are already there — then tick WarpDrive boxes and add a one-line “already merged, see UI_FolderView.md §…” note in the **plan**, not new normative checkboxes in the spec

**Out of scope:**
- Implementing dirty-region V2, DXGI scroll rects, background thumbnail enrichment, DxUi a11y coalescing, TW-9 UIA fixes, or any C++
- Deleting the WarpDrive Done archive
- Restoring the WarpDrive master or continuation batons to WIP
- Editing I11
- Expanding retired batons

## Git workflow

- Branch: `advisor/002-warpdrive-tailwind-executable-surface`
- Commits: WarpDrive live banner + checkbox sync; Tailwind re-home; README routing. Imperative messages.
- Do NOT push unless asked.

## Steps

### Step 1: Add a live-work banner to WarpDrive

Immediately after the existing File Operations coordination banner (`:1-3`), insert:

```markdown
> **LIVE WORK (2026-08-25).** Executable remainder is Task 10 Step 4 (verify or explicitly waive Full-as-Task-10-gate) and Task 11 closeout. Tasks 0–9 are done. Task 10 Steps 2–3 are measured no-ops (gate not met). The Implementation Findings Log / pause archives below are historical; do not resume from 2026-06-29 Full-suite flakes. Background thumbnail completeness is **not** this file: `FolderView_ThumbnailBackgroundEnrichmentFollowup_2026-07-04.md`.
```

The diary already lives in Done (SIZE WAIVER on the archive only). Do not copy it back into WIP.

**Verify**: the remainder plan names Task 10 Step 4 + Task 11 only. Diary is in Done (`rg "Pause archive refreshed" Specs/Plans/Done/Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md` still hits).

### Step 2: Sync Task 10 roll-up and Acceptance with recorded decisions

After confirming Steps 1–3 text still says gate-not-met / implemented no-op:

- Check the Task 10 **roll-up** box (`:66`) **only if** you also either (a) check Step 4 with an explicit note “Verify deferred to Task 11 Full gate; not a second product task”, or (b) leave Step 4 `[ ]` and add one sentence under Step 4 that Task 11 Step 3 is the same Full/perf matrix. Prefer (b) so two boxes do not claim Full twice — then leave roll-up `[ ]` until 003/closeout. **Required:** stop implying scroll-rects/V2 are still to be built.
- Uncheck or annotate Acceptance `:1849` if Task 11 Step 1 is still `[ ]`. Replace with: “Most Task 11 Step 1 contracts already live in `Specs/UI/UI_FolderView.md` (cached-only thumbnails, device-loss, brush reuse, async paste-shortcut, refresh/icon metrics). Remaining spec gap: background enrichment (owned by I1). Testing-spec sample/budget/archive rules: coordinate with I5 / plan 003.”
- Add one bullet under Task 11 Step 1: background enrichment is **out of this plan**; I1 owns it.

**Verify**: `rg -n "gate not met — not implemented" Specs/Plans/Done/Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md` still present. Remainder must not instruct implementing DXGI scroll rects.

### Step 3: Rebuild Tailwind as a HOLD routing ledger

At the top (after any 001 Status block), add:

```markdown
> **HOLD.** Do not pick items from this file independently. Current FolderView perf closeout is WarpDrive (I2). TW-2/TW-3/TW-D1 cleanup is the thumbnail follow-up (I1). Items below are either DONE (do not reopen), routed to a **live** owner, or still unowned and must be promoted to a new ACTIVE plan before code changes.
```

Then:

1. Mark TW-7 **DONE** with evidence: generation gate in `OnPasteShortcutComplete` + `UI_FolderView.md` paste-shortcut contract. Strike sequencing step 3 (`:240-242`) or mark it done.
2. Re-point TW-5/TW-23/TW-24 away from the retired DxUi baton. Options allowed by this plan (pick one, document in the file):
   - **Orphan OPEN — needs new ACTIVE owner** (preferred if you cannot prove a live Granite/DxUi plan still owns them), or
   - Route to WarpDrive Task 11 only if the work is purely a11y-snapshot perf on FolderView closeout — **STOP** if that would expand WarpDrive beyond FolderView perf closeout.
   - Do **not** write “execute via DxUi_Uia_ContinuationBaton”.
3. Explicitly route TW-2/TW-3/TW-D1 to `FolderView_ThumbnailBackgroundEnrichmentFollowup_2026-07-04.md`.
4. Leave remaining open IDs listed in a short table: ID, sev, live owner or `UNOWNED — do not execute from Tailwind`. Do not expand fix directions with stale line numbers; say “re-grep symbol before any future ACTIVE plan”.
5. Keep “Considered and refuted” sections (they prevent re-audit).

**Verify**: `rg "DxUi_Uia_ContinuationBaton" Specs/Plans/WIP/Operation_Tailwind_FolderViewWarpDrivePreMergeReviewRemediation_2026-06-30.md` → only historical mentions plus “retired, do not resume”. Sequencing no longer lists TW-7 as open work.

### Step 4: Thumbnail follow-up pointer hygiene

In `FolderView_ThumbnailBackgroundEnrichmentFollowup_2026-07-04.md` Closeout (`:102-110`):

- Keep “update `Specs/UI/UI_FolderView.md`” (still required; spec currently only *names* the follow-up at `:299`, it does not specify enrichment metrics).
- Change “update or retire the TW-2/TW-3/TW-4 rows in the Tailwind plan” to: TW-4 already DONE (Granite GR-18); TW-2/TW-3 retire in Tailwind when this plan moves to Done.
- Drop or rephrase “keep the Granite GR-D1 decision row marked as signed off” if that Granite file is already in Done (it is) — replace with “TW-D1 remains signed off; do not reopen cached-only visible path”.

Do **not** implement enrichment.

**Verify**: Closeout still forbids weakening `folderView_thumbnail_cached_only_no_close_stall`.

### Step 5: WIP README routing

Update I1/I2/H0 sentences so they match the remainder + Done archive. Do **not** restore I2 SIZE WAIVER on the new remainder. H0 must say “do not execute TW-5/23/24 via the retired DxUi baton in Done.”

**Verify**: `.\Tools\Get-SpecInventory.ps1 -FailOnFindings` exit 0. Index file set unchanged.

## Test plan

- No new runtime tests. This plan is documentation routing.
- Inventory + `git diff --check` are the gates.
- If you touch `UI_FolderView.md`, run `.\Tools\Get-SpecInventory.ps1 -FailOnFindings` and do **not** add `- [ ]` checkboxes there.

## Done criteria

- [ ] WarpDrive opens with a LIVE WORK banner naming Task 10 Step 4 + Task 11 only
- [ ] Acceptance/Task 11 no longer claim both “specs merged” and “Step 1 not started” without explanation
- [ ] Tailwind does not assign work to retired batons
- [ ] TW-7 is not in the open sequencing list
- [ ] TW-2/TW-3/TW-D1 point at the thumbnail follow-up
- [ ] Retired batons unchanged in size class
- [ ] `Get-SpecInventory.ps1 -FailOnFindings` exit 0
- [ ] No `.cpp` changes
- [ ] `advisor-plans/README.md` row 002 updated

## STOP conditions

- WarpDrive Tasks 0–9 boxes have been unchecked in the live file (product regression — do not “fix” by re-opening R1).
- Thumbnail enrichment is already implemented (queue/apply metrics exist). Then STOP and report; I1 may be closeout-ready instead of “clean open plan”.
- A live ACTIVE owner besides I1/I2 already owns TW-9/10/11 (search `Specs/Plans/WIP/README.md` and Done). Do not create a third owner; link it.
- Editing `UI_FolderView.md` would require describing unimplemented enrichment as current behavior.

## Maintenance notes

- Reviewer: the diary must remain; only the live banner and checkbox honesty matter.
- Any future TW-9 UIA fix needs its own ACTIVE plan with planned-at SHA and current `DxUi.Accessibility.cpp` anchors — not a Tailwind drive-by.
- Plan 003 will further split WarpDrive Task 11 Full-suite wait from FolderView product completeness.
