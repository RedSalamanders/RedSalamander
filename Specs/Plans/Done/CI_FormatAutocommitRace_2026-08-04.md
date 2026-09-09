> **NON-NORMATIVE RETIRED RECORD.** Moved from `Specs/Plans/WIP/` on 2026-08-25. **No remaining action on this file.** Do not execute it. Remaining CI format-job work lives in `Specs/Plans/WIP/CI_FormatAutocommitRace_2026-08-25.md` (I9).

## Archive status

- **State:** RETIRED
- **Archived:** 2026-08-25
- **Original path:** `Specs/Plans/WIP/CI_FormatAutocommitRace_2026-08-04.md`
- **Live owner:** Specs/Plans/WIP/CI_FormatAutocommitRace_2026-08-25.md (I9; do not execute this archive)
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, remaining-outcome list, STOP condition, and "next action" below is archive text. **No remaining action on this file.** If work continues, use the live owner named in Archive status.

---
# CI Format Autocommit Race Hardening

> **Frozen historical instructions (do not execute).** Migrated from historical root plan `plans/012-format-autocommit-race.md` (archived to `Specs/Plans/Done/AdvisorPlan_012-format-autocommit-race_2026-06-11.md` on 2026-08-25). Remaining work is owned by the 2026-08-25 I9 remainder, not this file.

## Status

- **State:** RETIRED (historical ACTIVE text follows; do not execute)
- **Priority:** P3 workflow safety
- **Planned at:** `e14ba96e30d7f767b4020f271f07a0936fe56b7d`
- **Ownership boundary:** same-repository/fork eligibility, stale-head-safe publication, and concurrency behavior for the `.github/workflows/ci.yml` `format` job
- **Drift check:** `git diff e14ba96e30d7f767b4020f271f07a0936fe56b7d..HEAD -- .github/workflows/ci.yml format-all.ps1 Tools/Tests`

## Current evidence

The current format job checks out `${{ github.head_ref || github.ref }}`, runs `format-all.ps1`, stages with `git add -A`, commits, and executes an unqualified `git push`. It has no fork eligibility guard, no remote-head compare-and-swap check, and no workflow concurrency group. A contributor push during formatting can make publication stale; a fork PR cannot receive the repository token push; `git add -A` can include unrelated generated files.

**2026-08-08 review refresh (A8-CI-01):** the checkout problem occurs before the
credential/push problem as well. For a fork pull request, using only the fork's branch name
against the base repository does not identify the fork repository and exact head SHA, so the
job can fail or inspect the wrong ref even when mutation is disabled later. Operation Cinderstar
package 12 preserves the August 8 evidence but creates no second implementation owner; this file
remains the sole formatter implementation queue.

## Remaining outcomes

1. Decide whether fork PRs skip the mutation job or run a check-only formatting gate; do not silently skip all formatting feedback. A check-only fork job must checkout the exact fork repository and immutable head SHA and must never receive a push path.
2. Capture the pre-format checkout SHA, fetch the exact destination ref before publication, and push only when that remote ref still equals the captured SHA.
3. Stage only the intended modified C/C++ files or a reviewed exact path set; never use `git add -A` in the bot commit.
4. Add or explicitly reject a ref-scoped concurrency group after reviewing its effect on every CI job, not just format.
5. Add deterministic policy tests that parse the workflow and prove checkout repository/ref selection, push eligibility, stale-head compare-and-swap, narrow staging, and concurrency behavior. Use operator CI runs only for behavior GitHub itself must supply.

## Verification

```powershell
Invoke-Pester -Script .\Tools\Tests -PassThru
.\format-all.ps1
git diff --check
```

Operator evidence must cover a same-repository PR with formatting drift, a branch advance during the job, a fork PR, and a push event. The final plan closeout records the exact workflow runs and updates current contributor/CI documentation when behavior changes.

## Done and STOP conditions

Move this plan to Done only after local policy tests and all four operator cases pass. Stop if the format job has changed to a different publication model, the destination ref cannot be identified without ambiguity, required-branch policy forbids the bot push, or the concurrency decision would broaden behavior without maintainer approval.
