# Plan 005: Route WhimFiles, Parallax, Atlas, and unique Notes_Scratch items

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `advisor-plans/README.md`.
>
> **Drift check (run first)**: `git diff --stat b5c3b0eab..HEAD -- Specs/Plans/WIP/Product_WhimFilesRemainingDecisions_2026-08-25.md Specs/Plans/WIP/Operation_Atlas_RemainingSpecificationDecisions_2026-08-04.md Specs/Plans/WIP/Operation_Parallax_DeferredArchitectureSecurityAndSimplificationDecisionReview_2026-07-14.md Specs/Plans/WIP/Notes_Scratch.md Specs/Plans/WIP/README.md Specs/FileSystem/FileSystem_FileOperations.md Specs/UI/UI_FolderView.md Specs/UI/UI_NavigationView.md`

## Status

- **Priority**: P2
- **Effort**: M
- **Risk**: LOW for routing text; HIGH if Notes_Scratch is treated as authority or deleted; HIGH if Parallax SUPERSEDED is marked without fresh evidence
- **Depends on**: `advisor-plans/001-wip-lifecycle-and-status-truth.md`
- **Category**: docs
- **Planned at**: commit `b5c3b0eab`, 2026-08-25

## Why this matters

The index already classifies WhimFiles as DECISION (D1), Atlas as DECISION (D0), Parallax as HOLD (H1), and Notes_Scratch as a human-managed note. The **bodies** still invite implementation: WhimFiles is “Draft for review” with P0 G2 undo; Parallax has twelve `PENDING REVIEW` rows and zero recorded outcomes; Notes contains unique product/bug intent that is not indexed. Agents will build from the wrong queue or lose human notes.

Atlas itself is mostly healthy: CLIP-01 / BR-01 / ABI-01 are accepted and absorbed into domain specs. Do not re-open them.

## Current state

### WhimFiles (D1)

`Product_WhimFilesRemainingDecisions_2026-08-25.md` is the live D1 routing table. The competitive analysis is frozen in `Specs/Plans/Done/Product_WhimFilesGapAnalysisAndImprovementPlan_2026-07-08.md`.
Banner `:3-4` already settles G2 **authority** under File Operations (exact identities / Proven Engine artifacts only). The G2 table row was retargeted on 2026-08-25 so it no longer cites `plans/013` as a live owner.
`Specs/FileSystem/FileSystem_FileOperations.md` (search `Undo/G2`) defers undo implementation.
README `:49` forbids D0/D1/H1 from changing production until promotion.

Overlap (do not merge files; annotate):

| Gap | Live owner |
|---|---|
| G2 undo | I11 File Operations contract + FileOps spec (deferred implementation) |
| G3 tabs / G9 hover preview / extra FolderView modes | Atlas `ATLAS-DEC-FV-01` (`UI_FolderView.md:1380-1383`) |
| G8 bookmarks / places | Atlas `ATLAS-DEC-NAV-01` |
| G1 live filter dimensions | no ACTIVE owner — remains D1 until promoted |
| Other G4–G24 | D1 until a maintainer promotes **one** slice to a new ACTIVE plan |

Root `plans/013-undo-file-operations-spike.md` was archived to `Specs/Plans/Done/AdvisorPlan_013-undo-file-operations-spike_2026-06-11.md` on 2026-08-25; do not revive it.

### Atlas (D0)

Open rows remain correctly deferred by domain specs (VFS paging `Plugins_VirtualFileSystem.md`, FolderView optional modes, NavigationView optional features, Monitor extras, placement debounce, etc.).
Resolved table `:41-47`: CLIP-01, BR-01, ABI-01 ACCEPTED. Keep them. Optional: move resolved rows below a “do not re-litigate” heading (already the intent of `## Resolved decisions`).

### Parallax (H1)

All PAR-1–12 are `PENDING REVIEW` (`:69-82`). Done criteria `:367-380` unchecked.
FO banner `:3-4` says PAR-7 feeds destination containment and PAR-10 remains provider-concurrency — that is **coordination**, not a recorded `SUPERSEDED`/`KEEP REJECTED` outcome.

PAR-3 (`DestroyWindow` on unique_hwnd) is likely superseded by Win32 CHK-1 **after plan 004 marks CHK-1 done**. This plan may mark PAR-3 SUPERSEDED **only if** 004 is DONE and the live grep still shows no production unique_hwnd `DestroyWindow(_hWnd.get())`.

Do not mark PAR-1 (named-pipe impersonation) or PAR-5 (secrets) SUPERSEDED without the evidence refresh commands already in the Parallax file (`:354-365`).

### Notes_Scratch.md

Banner `:1-6` and README `:14-16`: do not use as documentation; do not index as a plan; human-managed.

Unique **non-DONE** notes still present at survey time (route, do not implement):

1. Cut clipboard consumption (`:12-13`) — largely absorbed by Atlas CLIP-01 + FileOps; confirm wording only.
2. Connection Manager password field dead (`:16-17`).
3. All windows resizable (`:24-26`) → Atlas WIN-02.
4. FO preflight-skip counts / shared parallel pool (`:29-32`) → I11 (do not edit I11; add a routing sentence in Notes or README only if the human wants it — **prefer README H1/D0/I11 pointers in this plan’s README cell**, not a new ACTIVE plan).
5. FTP `.`/`..`, silent copy progress, bursty pcopy (`:38-41`) → Curl/FileSystemCurl owner (no current WIP index row; record as “needs ACTIVE promotion or reject”).
6. CM Connect vs Windows Hello timing (`:44-46`); nav emoji color (`:49-50`).
7. `UpdateWindow` audit (`:56-58`).
8. USB/network disconnect messaging (`:66-68`).
9. Multi-item copy confirmation source-folder wording (`:73-78`) → I11.
10. Typed directory-watch + focus rules (`:84-107`) → Atlas VFS-02 adjacency or FolderView owner.
11. FO status flash / speed-limit during calculate (`:109-112`) → I11.
12. Cross-pane directory cache (`:117-118`).
13. Broad IHost pane ABI (`:123-132`) → Atlas HOST-01.
14. Mount-root Backspace (`:1359-1362`).

The large DONE paste (~`:211` onward) is historical noise. README already allows removing obsolete unreferenced scratch **after unique items are routed or rejected**. Default of this plan: **route unique items in a short table inside Atlas or a new subsection of WIP README “Scratch routing (non-authoritative)” — do not delete Notes_Scratch.md.**

## Commands you will need

| Purpose | Command | Expected on success |
|---------|---------|---------------------|
| Inventory | `.\Tools\Get-SpecInventory.ps1 -FailOnFindings` | exit 0; `Notes_Scratch.md` still HumanManagedNote |
| Confirm Notes class | `rg -n "HumanManagedNote" Tools/Tests/SpecInformationArchitecture.Tests.ps1` | existing test |
| Confirm CLIP-01 in FileOps spec | `rg -n "cut sequence|Copied; source kept" Specs/FileSystem/FileSystem_FileOperations.md` | matches |
| Confirm FV-01 deferral | `rg -n "ATLAS-DEC-FV-01|marquee selection" Specs/UI/UI_FolderView.md` | matches |
| Line endings | `git diff --check` | clean |

## Scope

**In scope:**
- `Specs/Plans/WIP/Product_WhimFilesRemainingDecisions_2026-08-25.md`
- `Specs/Plans/WIP/Operation_Atlas_RemainingSpecificationDecisions_2026-08-04.md` (routing/annotations only; no new decisions invented)
- `Specs/Plans/WIP/Operation_Parallax_DeferredArchitectureSecurityAndSimplificationDecisionReview_2026-07-14.md`
- `Specs/Plans/WIP/README.md` (D0/D1/H1 sentences + optional scratch-routing appendix)
- `Notes_Scratch.md` **only** if you add a single pointer line at the top after the existing banner: “Unique open notes were routed on 2026-08-25 in Specs/Plans/WIP/README.md §Scratch routing.” Do not delete content. Do not implement from it.

**Out of scope:**
- Implementing any G\* feature, PAR remediation, or Notes bug
- Editing I11
- Deleting Notes_Scratch.md
- Re-opening CLIP-01 / BR-01 / ABI-01
- Filling all twelve Parallax outcomes without running that file’s evidence commands
- Creating new ACTIVE product plans unless the operator explicitly promotes a slice in the same instruction that dispatched you (this plan’s default: **do not create new ACTIVE files**)

## Git workflow

- Branch: `advisor/005-decision-hold-scratch-routing`
- Commits: WhimFiles status/G-routing; Parallax PAR-3-only if evidence allows; README scratch table.
- Do NOT push unless asked.

## Steps

### Step 1: WhimFiles is DECISION, not a build queue

- Status: DECISION (if 001 already set this, keep it).
- Add a routing table for G1–G24: `defer to Atlas ID` / `defer to I11` / `D1 only until promoted`.
- G2 row: “Authority settled in FileOps spec; implementation deferred; `plans/013` was archived to Done on 2026-08-25 and is not an owner.”
- Strike or qualify Waves 1–5 (`~:960-991`) with: “Not an execution schedule. Promotion required.”

**Verify**: `rg "plans/013" Specs/Plans/WIP/Product_WhimFilesRemainingDecisions_2026-08-25.md` does not call it a live owner.

### Step 2: Atlas — do not re-litigate resolved rows

- Keep open ledger.
- Add one sentence under Resolved: FileOps + FolderView/Batch Rename specs own CLIP-01/BR-01; Plugin API specs own ABI-01.
- If a domain spec still says “undecided” for CLIP/BR, STOP and report (that would be a spec bug, out of default scope). At survey time FileOps already records the accepted policy.

### Step 3: Parallax — record outcomes only with evidence

For each PAR you **do not** have time to refresh: leave PENDING REVIEW.

Minimum required in this plan:

- Run the DestroyWindow grep from Parallax `:357`. If production unique_hwnd owner pattern is gone **and** plan 004 marked CHK-1 done, record PAR-3 as `SUPERSEDED` with date, commit `b5c3b0eab` or HEAD, outcome, owner `Win32_Inventory.md` (residual test-harness hits only), reopen trigger: “new `DestroyWindow(_hWnd.get())` on a `wil::unique_hwnd` owner.”
- If CHK-1 is not done, leave PAR-3 PENDING and point at I8.

Do **not** mark PAR-7 SUPERSEDED solely because the FO banner mentions containment. That needs `Core_SharedHelpers` / FileOps taxonomy evidence.

**Verify**: decision table uses one of the allowed outcomes (`KEEP REJECTED`, `REOPEN INVESTIGATION`, `ACCEPT POLICY CHANGE`, `ROUTE REMEDIATION`, `SUPERSEDED`) for any row you change. No production source edits (`:378`).

### Step 4: Scratch routing appendix (non-authoritative)

In `Specs/Plans/WIP/README.md`, after Lifecycle rules, add a short section:

```markdown
## Scratch routing (non-authoritative)

`Notes_Scratch.md` remains human-managed and excluded from specification authority.
Unique open notes as of 2026-08-25 were routed as follows (not implementation approval):
| Note | Route |
|---|---|
| Cut clipboard after paste | Absorbed: Atlas CLIP-01 / FileOps spec |
| Resizable windows | Atlas WIN-02 |
| IHost pane ABI | Atlas HOST-01 |
| FO parallel pool / confirmation wording / status flash | I11 (do not duplicate I11) |
| … |
```

Fill the rest from the unique-item list. Mark “needs ACTIVE promotion” where there is no owner.

Optional one-line pointer in Notes_Scratch under the banner. Do not delete the DONE megablock in this plan (size reduction is a human choice).

**Verify**: `Get-SpecInventory.ps1 -FailOnFindings` still treats Notes as HumanManagedNote; no `WIP_UNINDEXED` for Notes; no new index row for Notes.

## Test plan

- Existing `SpecInformationArchitecture.Tests.ps1` HumanManagedNote cases must stay green.
- No new product tests.

## Done criteria

- [ ] WhimFiles cannot be mistaken for an ACTIVE implementation plan
- [ ] G2 does not cite `plans/013` as live
- [ ] Atlas resolved FO/ABI rows not reopened
- [ ] Parallax either still PENDING or has a dated allowed outcome with grep evidence
- [ ] Unique Notes items appear in a non-authoritative routing table
- [ ] `Notes_Scratch.md` not deleted; still unindexed
- [ ] I11 file untouched
- [ ] `Get-SpecInventory.ps1 -FailOnFindings` exit 0
- [ ] `advisor-plans/README.md` row 005 updated

## STOP conditions

- You would need to edit I11 to “route” FO notes — put them in the README table pointing at I11 instead.
- Parallax evidence grep contradicts the FO banner in a way you cannot classify — leave PENDING.
- Operator asks to promote G1 (filter) to ACTIVE — that is a new plan, not a silent status flip inside D1.
- Inventory wants Notes indexed because you added plan-like Status fields to Notes_Scratch — revert those fields.

## Maintenance notes

- Reviewer: reject new ACTIVE WIP files from this plan unless the dispatch message explicitly promoted a slice.
- D1 remains a competitive backlog; do not merge it into Atlas (different audience: product gaps vs spec-undecidable questions).
- After a human deletes Notes_Scratch, remove the scratch-routing appendix or mark it historical.
