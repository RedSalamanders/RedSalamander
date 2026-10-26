# Plan 001: Stamp README lifecycle fields and honest status on every in-scope WIP file

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `advisor-plans/README.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat b5c3b0eab..HEAD -- Specs/Plans/WIP Tools/Modules/Tooling/SpecInformationArchitecture.psm1 Tools/Tests/SpecInformationArchitecture.Tests.ps1 Specs/Testing/Testing_ToolingGovernance.md`
> If any in-scope file changed since this plan was written, compare the
> "Current state" excerpts against the live files before proceeding; on a
> mismatch, treat it as a STOP condition.

## Status

- **Priority**: P0
- **Effort**: M
- **Risk**: LOW — docs and inventory tests only; no production C++
- **Depends on**: none
- **Category**: docs
- **Planned at**: commit `b5c3b0eab`, 2026-08-25

## Why this matters

`Specs/Plans/WIP/README.md` already requires every **active** plan to name priority/status, planned-at commit and drift command, ownership boundary, remaining outcomes, authoritative specs, implementation/tests, verification, done criteria, and STOP conditions. Several indexed ACTIVE files still say `Status: WIP`, omit planned-at/drift/STOP, or use `IN PROGRESS` which is not a valid index state. `Get-SpecInventory.ps1` currently enforces index coverage and retired-tombstone size, **not** those lifecycle fields, so the drift returns.

Without this pass, later executors cannot tell which files are executable, blocked on Full, or decision-only, and they will follow stale "next action" sentences.

## Current state

- `Specs/Plans/WIP/README.md:7-15` defines `ACTIVE` / `HOLD` / `DECISION` / `RETIRED`. `Notes_Scratch.md` is excluded from the index (`:14-16`).
- `Specs/Plans/WIP/README.md:55` is the lifecycle rule this plan operationalizes.
- Inventory enforcement today (`Tools/Modules/Tooling/SpecInformationArchitecture.psm1:577-683`) checks `WIP_UNINDEXED`, `WIP_DUPLICATE_INDEX`, `WIP_STALE_INDEX`, `WIP_AUTHORITY_CLAIM`, and compact RETIRED tombstones. It does **not** require planned-at/STOP on ACTIVE files.
- Good compact exemplars (copy this shape): `Specs/Plans/WIP/CI_FormatAutocommitRace_2026-08-25.md` and `Specs/Plans/WIP/FileOperations_RemainingStressValidation_2026-08-04.md`. The 2026-08-04 CI write-up is frozen in Done.
- Convention: WIP files are **non-normative**. Do not claim they override `Specs/<Domain>/`. Keep a `> **NON-NORMATIVE …**` banner on ACTIVE files that lack one. Do **not** use a heading or sentence matching `(?im)^\s*(?:this\s+)?plan\s+(?:is\s+)?(?:the\s+)?(?:authoritative|source of truth)` — that trips `WIP_AUTHORITY_CLAIM`.

### Per-file field gaps at `b5c3b0eab`

Use this table. Do not invent remaining work; later plans own content truth.

| File | Index | Body status string today | Missing / wrong |
|---|---|---|---|
| `FileSystem_GoogleDrivePluginPlan.md` | ACTIVE I0 | `Status: WIP` (`:8`) | planned-at, drift, ownership boundary, named STOP; State must become ACTIVE |
| `FolderView_ThumbnailBackgroundEnrichmentFollowup_2026-07-04.md` | ACTIVE I1 | no Status block | State ACTIVE, planned-at, drift, ownership, STOP |
| `Operation_FolderView_WarpDrive_RemainingCloseout_2026-08-25.md` | ACTIVE I2 | remainder after 2026-08-25 archive | Status block already present; do not reopen the Done master |
| `CI_FormatAutocommitRace_2026-08-25.md` | ACTIVE I9 | remainder after 2026-08-25 archive | leave body; index already ACTIVE |
| `Product_WhimFilesRemainingDecisions_2026-08-25.md` | DECISION D1 | remainder after 2026-08-25 archive | keep DECISION routing only |
| `FileOperations_RemainingStressValidation_2026-08-04.md` | ACTIVE I3 | already good | optional: rename nothing; leave for plan 004 |
| `Operation_Astrolabe_TestContractArchitectureMigration_2026-07-17.md` | ACTIVE I4 | `Status \| WIP` table (`:8`) | planned-at, drift, STOP, verification command; State ACTIVE |
| `Operation_PerfMeasurementContract_2026-07-06.md` | ACTIVE I5 | `Status: WIP` (`:11`) | planned-at, drift, STOP; State ACTIVE. Do not rewrite Task checkboxes here (plan 003) |
| `Operation_RosettaLantern_RedConfigureSatelliteLocalizationReview_2026-07-17.md` | ACTIVE I6 | `Status \| WIP` (`:5`) | planned-at, drift, STOP; State ACTIVE. Do not change the review ledger here (plan 004) |
| `Core_HResultStatusFormattingCleanup_2026-06-19.md` | ACTIVE I7 | `Status: WIP` (`:8`) | planned-at, drift, STOP; State ACTIVE |
| `Win32_Inventory.md` | ACTIVE I8 | `Status: WIP` (`:5`) | planned-at, drift, STOP; State ACTIVE |
| `Tooling_ToolsInventoryAndGovernance_2026-08-13.md` | ACTIVE I10 | already good | leave body (plan 003) |
| `Operation_FileOperations_GlobalBehaviorDecisionReview_2026-08-16.md` | ACTIVE I11 | **OUT OF SCOPE** | do not open |
| `Operation_ReviewFollowup_FOTerminalFocus_2026-08-19.md` | ACTIVE I12 | `State: IN PROGRESS` (`:10`) | change State to ACTIVE; keep a Progress row. Do not rewrite checklist (plan 003) |
| `Operation_Tailwind_FolderViewWarpDrivePreMergeReviewRemediation_2026-06-30.md` | HOLD H0 | narrative, no Status block | State HOLD, planned-at, drift, STOP. Do not triage TW items here (plan 002) |
| `Operation_Parallax_…_2026-07-14.md` | HOLD H1 | already has Status/STOP | align **Status** label to `State: HOLD` if needed; leave PAR outcomes (plan 005) |
| `Operation_Atlas_RemainingSpecificationDecisions_2026-08-04.md` | DECISION D0 | already good | leave ledger (plan 005) |
| `DxUi_Uia_ContinuationBaton_2026-06-29.md` / `FolderView_WarpDrive_ContinuationBaton_2026-06-29.md` | moved to Done 2026-08-25 | retired history | do not restore them to WIP |
| `Notes_Scratch.md` | excluded | human scratch | do not edit |
| `README.md` | index | last reconciled 2026-08-21 (`:3`) | bump “Last reconciled”; do not change I11 row text except if you must not |

## Commands you will need

| Purpose | Command | Expected on success |
|---------|---------|---------------------|
| Spec inventory | `.\Tools\Get-SpecInventory.ps1 -FailOnFindings` | exit 0, zero blocking findings |
| IA Pester | `Invoke-Pester .\Tools\Tests\SpecInformationArchitecture.Tests.ps1 -PassThru` | FailedCount 0 |
| Line endings | `git diff --check` | no whitespace errors (line-ending notices only are OK) |
| Confirm I11 untouched | `git diff --name-only -- Specs/Plans/WIP/Operation_FileOperations_GlobalBehaviorDecisionReview_2026-08-16.md` | empty |

Do **not** run `.\format-all.ps1` (it reformats unrelated C++). Do **not** run Full.

## Suggested executor toolkit

- Read `Specs/Plans/WIP/README.md` Status meanings and Lifecycle rules before editing.
- Read `Specs/Testing/Testing_ToolingGovernance.md` if you add or change a Tools test (same change must mention inventory/help/spec if you add a **new** Tools file — you should not; only edit existing psm1/tests).
- Localization/skill files are out of scope.

## Scope

**In scope:**
- Every file under `Specs/Plans/WIP/` except I11 and `Notes_Scratch.md`
- `Tools/Modules/Tooling/SpecInformationArchitecture.psm1`
- `Tools/Tests/SpecInformationArchitecture.Tests.ps1`
- `Specs/Testing/Testing_ToolingGovernance.md` (only if you add the ACTIVE-field inventory check; one short paragraph)

**Out of scope:**
- `Specs/Plans/WIP/Operation_FileOperations_GlobalBehaviorDecisionReview_2026-08-16.md`
- `Notes_Scratch.md`
- Root `plans/`
- Any `.cpp` / `.h` / `.rc` / workflow YAML
- Rewriting WarpDrive diary, Tailwind TW tables, WhimFiles G\* bodies, Parallax PAR narratives, Google Drive phases, Rosetta review ledger truth, HRESULT work-item completion, Win32 CHK completion (those are 002–005)

## Git workflow

- Branch: `advisor/001-wip-lifecycle-and-status-truth`
- Commit per logical unit (Status blocks, then inventory test). Message style matches recent history: imperative sentence, e.g. `Stamp WIP lifecycle fields on active plans`.
- Do NOT push or open a PR unless the operator instructed it.

## Steps

### Step 1: Insert the Status block template

Add or replace a **Status** section near the top of each file in the gap table (after any existing `> **NON-NORMATIVE**` / File Operations coordination banner). Use this shape; fill ownership/drift paths from the file’s real topic:

```markdown
## Status

- **State:** ACTIVE   # or HOLD / DECISION — must match Specs/Plans/WIP/README.md
- **Priority:** <keep existing priority sentence>
- **Planned at:** `b5c3b0eab`
- **Ownership boundary:** <one sentence; copy from the README index “Ownership / next action” cell if the file has none>
- **Drift check:** `git diff b5c3b0eab..HEAD -- <topic paths>`
```

Then keep or add (headings may already exist under other names — rename to these exact phrases so the inventory regex can find them):

- `## Remaining outcomes` (or keep an existing equivalent heading **and** mention “Remaining outcomes” in the Status block)
- `## Authoritative specs`
- `## Implementation / tests` (can be one sentence pointing at the checklist)
- `## Verification` with a fenced powershell command
- `## Done criteria`
- `## STOP conditions`

For files that already have these (I3, I9, I10, D0, H1), only fix State string mismatches.

**I12:** change `- **State:** IN PROGRESS` to `- **State:** ACTIVE` and keep the Progress table.

**WarpDrive (I2):** add the Status block **above** “Master Implementation Checklist”. Do not delete Task 10/11 or the diary.

**Tailwind (H0):** add State HOLD at the top. Do not mark TW items done.

**WhimFiles (D1):** replace `Status: Draft for review (no code changes yet)` with State DECISION. Add a one-line rule: this file does not authorize implementation.

**Verify**: `rg -n "^\*\*State:\*\* WIP|Status: WIP|Status \| WIP" Specs/Plans/WIP` → no matches except possibly inside historical quotes. `rg -n "IN PROGRESS" Specs/Plans/WIP/Operation_ReviewFollowup_FOTerminalFocus_2026-08-19.md` → only inside Progress table prose if needed, not as State.

### Step 2: Non-normative banners

Any ACTIVE file missing a non-normative banner gets one line after the title:

```markdown
> **NON-NORMATIVE ACTIVE PLAN.** Current behavior is owned by `Specs/<Domain>/`. This file routes unfinished work only.
```

Do not put this on RETIRED tombstones or on `Notes_Scratch.md`. HRESULT’s existing “FILE OPERATIONS COORDINATION” banner may stay; add NON-NORMATIVE if absent.

HRESULT currently has `## Source Of Truth` (`Core_HResultStatusFormattingCleanup_2026-06-19.md:26`). Rename that heading to `## Shared helpers` or `## Implementation references` so it cannot be read as plan-overrides-spec. Do not change the bullet list.

**Verify**: `.\Tools\Get-SpecInventory.ps1 -FailOnFindings` → still exit 0 (no new `WIP_AUTHORITY_CLAIM`).

### Step 3: Enforce ACTIVE lifecycle fields in inventory

In `Tools/Modules/Tooling/SpecInformationArchitecture.psm1`, inside the `foreach ($file in $directWip)` block that already handles `WIP_AUTHORITY_CLAIM` (`:666-683`), after the RETIRED tombstone check, add an ACTIVE/HOLD/DECISION field check:

- If `$file.WipStatus` is `ACTIVE`, `HOLD`, or `DECISION`, require the raw markdown to match (case-insensitive) all of: `State`, `Planned at`, `Ownership`, `Drift`, `STOP`.
- Also require at least one of `Done criteria` / `Definition of done` / `Exit Criteria` / `Closeout` (WhimFiles and Google Drive use closeout language).
- Emit `WIP_MISSING_LIFECYCLE` Severity Error with the missing tokens listed.
- Do **not** apply this to `RETIRED` (tombstone rules already apply) or `HumanManagedNote`.

Extend `Tools/Tests/SpecInformationArchitecture.Tests.ps1`:

- Existing fixture plan (`:67-79`) already contains Status, Planned at, Ownership, Remaining outcome, Authoritative spec, Verification, Done criteria, STOP — keep those strings so the happy path stays green.
- Add a test that copies the fixture, strips `STOP` from `Specs/Plans/WIP/Plan.md`, and expects `WIP_MISSING_LIFECYCLE` count 1.

If `Specs/Testing/Testing_ToolingGovernance.md` lists inventory finding codes, add `WIP_MISSING_LIFECYCLE` in the same sentence as the other WIP codes (search the file; if codes are not listed, skip the spec edit).

**Verify**:

```powershell
Invoke-Pester .\Tools\Tests\SpecInformationArchitecture.Tests.ps1 -PassThru
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
```

Expected: Pester FailedCount 0; inventory exit 0 against the real tree after Step 1.

### Step 4: Refresh the WIP index timestamp

Set `Specs/Plans/WIP/README.md` “Last reconciled” to `2026-08-25` (or today’s date) and `b5c3b0eab` (or HEAD if you rebased). Do not change I11’s ownership sentence.

**Verify**: `.\Tools\Get-SpecInventory.ps1 -FailOnFindings` exit 0. Index still contains exactly the same file set (I0–I12, H0–H1, D0–D1, R0–R1). `Notes_Scratch.md` still unindexed.

## Test plan

- New case in `Tools/Tests/SpecInformationArchitecture.Tests.ps1`: fixture ACTIVE plan missing `STOP` → `WIP_MISSING_LIFECYCLE`.
- Model after `It 'rejects an unindexed direct WIP file'` (`:175-181`).
- No new C++ tests.

Verification: `Invoke-Pester .\Tools\Tests\SpecInformationArchitecture.Tests.ps1 -PassThru` → FailedCount 0, including the new It.

## Done criteria

- [ ] `.\Tools\Get-SpecInventory.ps1 -FailOnFindings` exits 0
- [ ] `Invoke-Pester .\Tools\Tests\SpecInformationArchitecture.Tests.ps1 -PassThru` FailedCount 0
- [ ] No in-scope file still has `Status: WIP` or `**State:** IN PROGRESS` as the machine status
- [ ] `git diff --name-only` does not include `Operation_FileOperations_GlobalBehaviorDecisionReview_2026-08-16.md` or `Notes_Scratch.md`
- [ ] Retired batons still ≤40 lines and ≤4096 bytes
- [ ] `advisor-plans/README.md` status row for 001 updated

## STOP conditions

Stop and report back (do not improvise) if:

- The code/docs at the locations in “Current state” don’t match the excerpts (the tree drifted).
- `Get-SpecInventory.ps1 -FailOnFindings` fails for a reason other than the field you are adding (e.g. LINK_MISSING_TARGET). Do not “fix” unrelated spec links.
- Making a file pass `WIP_MISSING_LIFECYCLE` would require inventing remaining product work (leave a TODO pointing at plans 002–005 instead of guessing).
- You believe I11 or `Notes_Scratch.md` must be edited for the inventory to pass.
- A RETIRED tombstone would need more than 40 lines to satisfy a new field — do not expand tombstones; exempt them (already exempt if you key off WipStatus RETIRED).

## Maintenance notes

- Future ACTIVE plans must include the Status block on day one or CI/spec inventory will fail.
- Planned-at SHAs in this pass are “hygiene stamps” at `b5c3b0eab`. Plans 002–004 may replace a file’s planned-at when they re-verify anchors.
- Reviewer: reject any C++ change; reject I11 edits; reject WarpDrive diary deletion in this plan.
