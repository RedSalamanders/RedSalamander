# CI Format Autocommit Race Hardening

> **NON-NORMATIVE ACTIVE PLAN.** Sole remaining owner for the `.github/workflows/ci.yml` `format` job. The 2026-08-04 write-up is frozen history at `Specs/Plans/Done/CI_FormatAutocommitRace_2026-08-04.md`. Historical root `plans/012-format-autocommit-race.md` is `Specs/Plans/Done/AdvisorPlan_012-format-autocommit-race_2026-06-11.md`. Do not recreate `plans/`. This file does not authorize a push or workflow change outside the exact format-job boundary.

## Status

- **State:** ACTIVE
- **Priority:** P3 workflow safety
- **Planned at:** `56db9af9d0a19b4b3f375e16989591f29e014a6c`
- **Ownership boundary:** same-repository/fork eligibility, stale-head-safe publication, narrow staging, and concurrency behavior for the `.github/workflows/ci.yml` `format` job
- **Drift check:** `git diff 56db9af9d0a19b4b3f375e16989591f29e014a6c..HEAD -- .github/workflows/ci.yml format-all.ps1 Tools/Tests`
- **Authoritative specs:** contributor/CI documentation only when publication behavior changes; current workflow text is not a product contract

## Remaining outcomes

Re-verified 2026-08-25: checkout still uses `ref: ${{ github.head_ref || github.ref }}` (`.github/workflows/ci.yml` ~`:41`) and the bot still stages with `git add -A` (~`:88`).

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

Operator evidence must cover a same-repository PR with formatting drift, a branch advance during the job, a fork PR, and a push event. Closeout records the exact workflow runs and updates current contributor/CI documentation when behavior changes.

## Done criteria

- Local policy tests and all four operator cases pass.
- Durable workflow behavior is documented in current contributor/CI docs if publication semantics change.
- This plan moves to `Specs/Plans/Done/` only after those gates pass.

## STOP conditions

Stop if the format job has changed to a different publication model, the destination ref cannot be identified without ambiguity, required-branch policy forbids the bot push, or the concurrency decision would broaden behavior without maintainer approval.
