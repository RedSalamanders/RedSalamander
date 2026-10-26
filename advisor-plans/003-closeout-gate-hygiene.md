# Plan 003: Separate implementation-complete plans from Fresh Full / perf-enforcement leftovers

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `advisor-plans/README.md`.
>
> **Drift check (run first)**: `git diff --stat b5c3b0eab..HEAD -- Specs/Plans/WIP/Operation_ReviewFollowup_FOTerminalFocus_2026-08-19.md Specs/Plans/WIP/Tooling_ToolsInventoryAndGovernance_2026-08-13.md Specs/Plans/WIP/Operation_PerfMeasurementContract_2026-07-06.md Specs/Plans/WIP/Operation_FolderView_WarpDrive_RemainingCloseout_2026-08-25.md Specs/Plans/WIP/README.md Specs/Testing/Testing_PerformanceValidation.md Specs/Plans/Done/Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md Specs/Plans/Done/Operation_FolderView_FocusSelectionReviewRemediation_2026-08-21.md`

## Status

- **Priority**: P1
- **Effort**: M
- **Risk**: LOW — documentation; HIGH if the executor interprets “closeout” as license to run multi-hour Full and “greenwash” unrelated failures
- **Depends on**: 001; WarpDrive Task 11 wording after 002
- **Category**: docs
- **Planned at**: commit `b5c3b0eab`, 2026-08-25

## Why this matters

Three ACTIVE plans are mostly **implemented** and remain in WIP only because Full / perf-contract enforcement did not finish. Mixing “write the feature” with “wait for Fresh Full” causes agents to re-implement finished work or to treat Tools/I12 as blocked on unrelated Commands failures.

Separately, the perf-measurement WIP still tells executors to edit a flake-convergence plan that **already moved to Done**, and Task 1 still wants a Pester field list that does not match the authoritative `Perf Gate Template` already in `Testing_PerformanceValidation.md`.

## Current state

### I12 Terminal/FO follow-up — implementation complete, Full blocked by scope

`Operation_ReviewFollowup_FOTerminalFocus_2026-08-19.md`:

- Progress (`:16-23`): both newly confirmed High gaps fixed; next is Fresh Full.
- Checklist items 1–9 `[x]`; item 10 focused evidence `[x]`; Fresh Full `[blocked by scope]` (`:115`).
- Item 7 still says “The later I13 FolderView remediation owns replacement ordering” (`:33`). That work is **COMPLETE** in `Specs/Plans/Done/Operation_FolderView_FocusSelectionReviewRemediation_2026-08-21.md` (State COMPLETE, outcome 4: canonical successor/predecessor resolver). I12 must not speak of I13 as future.

Done criteria (`:61-63`) currently require Fresh Full before moving to Done. That is consistent with README; the doc must say **implementation is complete** and the only remaining outcome is an operator-authorized Fresh Full, not more code.

### I10 Tools governance — implementation complete, Full blocked by unrelated Commands

`Tooling_ToolsInventoryAndGovernance_2026-08-13.md`:

- Outcomes 1–5 `[x]`; closeout Full `[ ]` (`:119-124`).
- STOP already says: if Full reveals unrelated pre-existing failure, stop rather than guess (`:251-255`).
- Listed Commands failures (`:162-180`) are application selftests, not Tools. `Specs/Testing/Testing_ToolingGovernance.md` already exists as the durable contract (`:1-18`).

The plan should state: **do not keep this file ACTIVE for Tools implementation**; remaining move-to-Done requires either (a) a later Fresh Full that is green, or (b) an explicit maintainer waiver that Tools closeout is satisfied by the 521/521 tooling Pester + `Get-SpecInventory` already recorded, with Commands failures owned elsewhere.

This plan does **not** authorize the waiver. It authorizes writing both options into the WIP file so a human can choose.

### I5 Perf measurement contract — Task 0 done; later tasks stale vs spec

`Operation_PerfMeasurementContract_2026-07-06.md`:

- Status (`:11`): “short-root contract merged into authoritative testing/archive docs; contract-enforcement tasks not started”.
- File map still modifies `Specs/Plans/WIP/Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md` (`:25`, `:128-129`, `:336`, `:365`). That file is now `Specs/Plans/Done/Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md`.
- Task 1 (`:360-412`) wants a Pester test matching exact labels `Metric keys:`, `Analyzer command:`, `Sample-quality result:`, `Environment matrix:`, `Baseline run:`, `Candidate run:`.
- Authoritative spec already has **Perf Gate Template** (`Specs/Testing/Testing_PerformanceValidation.md:242-259`) with `Scenario:`, `Metric keys:`, `Deterministic validation:`, `Build flavor:`, `Environment/root matrix:`, `Archived evidence:`, `Analyzer and sample quality:`, `Before/after delta:`.

Blindly adding Task 1’s Pester test will **fail** against the live spec, or force a duplicate second template. Reconcile: treat the Gate Template as the landed record; retarget remaining tasks to enforcement (runner metadata, closeout lint) without requiring a second heading named “Perf Measurement Record” unless you add that heading as an alias of the Gate Template.

Task 6 (`:785-810`) wants WIP README + Pester lint that perf-sensitive plans mention the record. That is still valid work **inside this WIP file**, not a reason to edit Done flake-convergence history.

Protected Done history: `Get-SpecInventory.ps1` uses `ProtectedDoneBaseCommit = '6aecfde8e'`. **Do not edit** the Done flake-convergence file unless inventory says that path is already mutable; default STOP: no Done edits.

### I2 WarpDrive Task 11

After plan 002, Task 11 still requires Full + Debug/Release FolderView perf matrix (`:1820-1830`). Rewrite remaining outcomes to:

1. Spec migration leftover (enrichment → I1; testing sample/budget rules → I5).
2. Operator Full + current-HEAD perf matrix citation.
3. Move to Done only after (1)+(2).

Do not keep “resume from textinputframework.dll crash” as current next action.

## Commands you will need

| Purpose | Command | Expected on success |
|---------|---------|---------------------|
| Inventory | `.\Tools\Get-SpecInventory.ps1 -FailOnFindings` | exit 0 |
| Confirm flake plan is Done | `Test-Path Specs/Plans/Done/Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md` | True |
| Confirm I13 Done | `rg -n "State:\\*\\* COMPLETE" Specs/Plans/Done/Operation_FolderView_FocusSelectionReviewRemediation_2026-08-21.md` | match |
| Confirm Gate Template | `rg -n "^## Perf Gate Template" Specs/Testing/Testing_PerformanceValidation.md` | match at ~242 |
| Line endings | `git diff --check` | clean |

Do not run `Run-AllTests.ps1 -Suite Full` in this plan.

## Scope

**In scope:**
- `Specs/Plans/WIP/Operation_ReviewFollowup_FOTerminalFocus_2026-08-19.md`
- `Specs/Plans/WIP/Tooling_ToolsInventoryAndGovernance_2026-08-13.md`
- `Specs/Plans/WIP/Operation_PerfMeasurementContract_2026-07-06.md`
- `Specs/Plans/WIP/Operation_FolderView_WarpDrive_RemainingCloseout_2026-08-25.md` (Task 11 / remaining-outcomes text only)
- `Specs/Plans/WIP/README.md` (I5, I10, I12 next-action sentences)
- `Specs/Testing/Testing_PerformanceValidation.md` **only** if you add a single alias sentence that “Perf Measurement Record” in older WIP text means the Perf Gate Template — no second competing template

**Out of scope:**
- Running or “fixing” Commands selftests listed in the Tools plan
- Editing `Specs/Plans/Done/**` (including flake-convergence)
- Implementing Show-PerfRuns or Run-AllTests metadata (that remains I5 **product** work after the WIP file is truthful)
- I11
- Moving I10/I12/I2/I5 to Done

## Git workflow

- Branch: `advisor/003-closeout-gate-hygiene`
- Commits: I12 I13 pointer; Tools remaining-outcomes; Perf retarget; WarpDrive Task 11 remainder; README.
- Do NOT push unless asked.

## Steps

### Step 1: I12 — implementation complete / I13 past tense

In `Operation_ReviewFollowup_FOTerminalFocus_2026-08-19.md`:

- Remaining outcome 7: replace “later I13” with a past-tense pointer to `Specs/Plans/Done/Operation_FolderView_FocusSelectionReviewRemediation_2026-08-21.md`.
- Add a one-line **Remaining outcomes** after the checklist: “No further code. Operator must authorize `.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh` on a stable snapshot and archive it; then move this file to Done and merge any last durable sentences already listed under Authoritative specs.”
- Keep STOP: do not move the 2026-08-19 Done defect-fixes plan back to WIP (`:67`).

**Verify**: `rg "later I13" Specs/Plans/WIP/Operation_ReviewFollowup_FOTerminalFocus_2026-08-19.md` → no match.

### Step 2: Tools — remaining closeout is policy, not implementation

In `Tooling_ToolsInventoryAndGovernance_2026-08-13.md` section 6 and “Closeout remains” (`:182-184`):

- State that outcomes 1–5 and the follow-up hardening list are complete.
- Remaining: Full suite **or** documented maintainer waiver.
- List the 18 Commands names as **out-of-owner** (do not attempt as part of I10).
- Do not delete the failure list; it is evidence.

Update README I10 next-action to the same split.

**Verify**: no new Tools command or inventory file is added.

### Step 3: Retarget Perf WIP off the moved flake plan

In `Operation_PerfMeasurementContract_2026-07-06.md`:

- Replace every path `Specs/Plans/WIP/Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md` with the Done path, and mark those steps **do not edit the Done file**. Representative records belong in **this** WIP file or in `Testing_PerformanceValidation.md`, not in protected Done history.
- Task 1: rewrite to “confirm Perf Gate Template fields remain in `Testing_PerformanceValidation.md`; add Pester against **those live labels** (`Scenario:`, `Metric keys:`, `Deterministic validation:`, `Build flavor:`, `Archived evidence:`, `Before/after delta:`). Do not require `Analyzer command:` as an exact substring if the spec says `Analyzer and sample quality:`.”
- Task 0 Step 2 “link the flake-convergence plan” → mark superseded (file moved); keep the link in Related as Done.
- Keep Tasks 2–6 as still-open **enforcement** work, but fix file paths.

Optional one sentence in `Testing_PerformanceValidation.md` under Perf Gate Template:

```markdown
Older WIP text may call this block a Perf Measurement Record; the fields above are the required content.
```

Do not duplicate the whole template.

**Verify**: `rg "Specs/Plans/WIP/Operation_TestSuiteStabilization" Specs/Plans/WIP/Operation_PerfMeasurementContract_2026-07-06.md` → no match. `Get-SpecInventory.ps1 -FailOnFindings` exit 0 (Done history unchanged).

### Step 4: WarpDrive Task 11 remainder

Under Task 11, replace “resume from last pause archive” language in the **live** remainder (not the historical log) with:

- Spec: I1 enrichment; I5 testing-contract enforcement; FolderView contracts already in `UI_FolderView.md` as of 002.
- Evidence: operator Full + current-HEAD Debug and test-enabled Release FolderView perf matrix, cited under `Specs/TestRuns/`.
- Do not treat 2026-06-29 `textinputframework.dll` traces as the current blocker unless a new archive at HEAD reproduces them.

**Verify**: first 80 lines of WarpDrive still contain the 002 LIVE WORK banner and do not tell the agent to open a 2026-06-29 continuation folder as the next edit.

### Step 5: README I5/I10/I12

Align next-action cells with Steps 1–4. I5: “enforcement tasks remain; flake-convergence is Done — do not edit it.” I10: “implementation complete; Full or waiver.” I12: “implementation complete; Fresh Full.”

## Test plan

- No Full suite.
- If you add the optional Gate Template alias sentence, `Get-SpecInventory.ps1 -FailOnFindings` must still pass (no new checkboxes in the normative spec).
- If you add Pester for Gate Template fields, put it in `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` **only if** I5’s remaining work is in scope — prefer leaving the actual Pester implementation as I5 product work listed as remaining outcomes, not as a surprise Tools change in this docs plan. **This plan’s default is docs-only.** Do not add the Pester test unless you have time and follow tooling-governance (existing test file, no new Tools file). Default: **docs-only**.

## Done criteria

- [ ] I12 does not call I13 “later”
- [ ] I10 remaining outcomes do not include Tools implementation checkboxes that are already `[x]`
- [ ] Perf WIP does not instruct editing a WIP path for flake-convergence
- [ ] WarpDrive live remainder does not resume from 2026-06-29 diary as current work
- [ ] `Get-SpecInventory.ps1 -FailOnFindings` exit 0
- [ ] No Done-folder edits
- [ ] `advisor-plans/README.md` row 003 updated

## STOP conditions

- Inventory reports `DONE_HISTORY_CHANGED` after your edits — revert Done touches.
- Tools inventory/governance product work is actually unfinished (`Tools/README.md` missing, or `Get-ToolInventory.ps1 -FailOnFindings` fails). Then STOP; I10 is not closeout-hygiene-only.
- I12 checklist items 1–9 are unchecked in the live file.
- Operator asks you to run Full as part of this plan — that is a different task; this plan only documents the gate.

## Maintenance notes

- Reviewer: reject Full-suite “fixes” of Commands cases from this branch.
- A human still must authorize Fresh Full before I10/I12/I2 move to Done.
- I5 remaining enforcement (runner metadata, closeout lint) is real product/tooling work; this plan only makes the WIP file point at live files.
