# Advisor Plan 012 - Harden CI auto-format job

> **NON-NORMATIVE RETIRED RECORD.** Archived from root `plans/012-format-autocommit-race.md` on 2026-08-25. **No remaining action.** Do not execute this file. Select current work only through `Specs/Plans/WIP/README.md`.

## Archive status

- **State:** RETIRED
- **Archived:** 2026-08-25
- **Original path:** `plans/012-format-autocommit-race.md`
- **Live owner:** Specs/Plans/WIP/CI_FormatAutocommitRace_2026-08-25.md (I9; do not execute this archive)
- **No remaining action:** yes

> **Frozen history: do not resume.** Every executor instruction, checkbox, STOP condition, and "next action" below is 2026-06 archive text. **No remaining action.** If work continues, use the live owner named in Archive status.

---
# Plan 012: Harden the CI auto-format job against branch races and fork PRs

> **Frozen historical instructions (do not execute)**: This plan is archived; do not follow these steps. CI behavior cannot be
> fully verified locally; Done criteria separate local gates from the operator's
> CI-run verification. Honor STOP conditions. Historical: status-row updates no longer apply
> in `plans/README.md`.
>
> **Drift check (run first)**: `git diff --stat d6bfccc42..HEAD -- .github/workflows/ci.yml format-all.ps1`
> On drift, re-read the format job before proceeding.

## Status

- **Priority**: P3
- **Effort**: S
- **Risk**: LOW-MED (CI-only, but touches a job that pushes commits — a bug here pushes wrong things; the changes below only ADD guards)
- **Depends on**: none (plan 011 also edits ci.yml — coordinate: land 011 first or rebase; the two touch different jobs)
- **Category**: dx
- **Planned at**: commit `d6bfccc42`, 2026-06-11

## Why this matters

The `format` CI job checks out the PR head branch, runs clang-format over the tree, and if anything changed, commits and pushes **directly to the contributor's branch** while other jobs of the same run are still testing the *pre-format* commit. Two concrete failure modes: (1) race — if the author pushes while the job runs, the job's push lands out-of-order or fails mid-flow; (2) forks — `secrets.GITHUB_TOKEN` cannot push to a fork's branch, so the job fails confusingly for outside contributors. The auto-format design itself is the maintainer's deliberate choice (keep it); this plan adds the two missing guards.

## Current state

`.github/workflows/ci.yml:31-73` (read at plan time), abridged to the load-bearing parts:

```yaml
  format:
    runs-on: windows-latest
    steps:
      - name: Checkout
        uses: actions/checkout@v6
        with:
          ref: ${{ github.head_ref }}
          token: ${{ secrets.GITHUB_TOKEN }}
      - name: Install clang-format   # choco install llvm, tolerates 3010
      - name: Run format-all.ps1
      - name: Commit formatting fixes
        shell: pwsh
        run: |
          git diff --quiet
          if ($LASTEXITCODE -ne 0) {
            git config user.name "github-actions[bot]"
            git config user.email "github-actions[bot]@users.noreply.github.com"
            git add -A
            git commit -m "style: auto-format C++ sources`n`nApplied clang-format via format-all.ps1"
            git push
            ...
```

Observations:
- `ref: ${{ github.head_ref }}` is empty on `push` events (the workflow triggers on both `push` to master and `pull_request`) — on push events the job formats a detached checkout of master and `git push` semantics differ; the guards below must handle both event types.
- `git add -A` stages untracked files too — anything a prior step dropped in the tree would be committed. (No prior step currently creates files, but it's fragile.)
- No fork guard, no concurrency guard, no stale-head check before push.

## Commands you will need

| Purpose | Command | Expected on success |
|---------|---------|---------------------|
| YAML sanity | actionlint if available; else careful review + operator CI run | clean |
| Local format script still fine | `.\format-all.ps1` (requires clang-format installed) | exit 0 |

## Scope

**In scope**:
- `.github/workflows/ci.yml` — ONLY the `format:` job.

**Out of scope**:
- `format-all.ps1` itself; `.clang-format` rules.
- Other ci.yml jobs (plan 011 owns its additions).
- Converting the job to check-only mode — that reverses a deliberate maintainer choice; if you believe it's right, write it in the report as a recommendation, don't do it.

## Git workflow

- Branch: `advisor/012-format-job-guards`
- Subject: "Guard CI auto-format against fork PRs and branch races".
- Do NOT push or open a PR unless the operator instructed it.

## Steps

### Step 1: Fork + event guard

Add a job-level condition so the commit-and-push flow only runs where it can work:

```yaml
  format:
    runs-on: windows-latest
    if: github.event_name == 'push' || github.event.pull_request.head.repo.full_name == github.repository
```

For fork PRs the job is skipped entirely (formatting then gets fixed when a maintainer touches the branch, or CI fails the build elsewhere — note this consequence in your report). Same-repo PRs and master pushes behave as today.

**Verify**: YAML parses; expression syntax double-checked against GitHub docs (`github.event.pull_request.head.repo.full_name` is only defined on PR events — the `||` short-circuit on `push` handles that).

### Step 2: Stale-head guard before push

In the "Commit formatting fixes" step, before pushing, verify the remote branch still points at the commit that was checked out and formatted; skip the push (neutral, with a notice) if it moved:

```pwsh
$branch = "${{ github.head_ref }}"
if ([string]::IsNullOrEmpty($branch)) { $branch = "${{ github.ref_name }}" }
git fetch origin $branch --quiet
$remote = git rev-parse "origin/$branch"
$local  = git rev-parse "HEAD~0"   # HEAD before our format commit: capture BEFORE committing
```

Implementation detail: capture `$preFormatSha = git rev-parse HEAD` at the START of the step (before `git commit`); after committing, fetch and compare `origin/$branch` to `$preFormatSha`; if different → `Write-Warning "Branch advanced during CI; skipping auto-format push (will format on next run)"` and exit 0 WITHOUT pushing. If equal → `git push origin HEAD:$branch`.

Also replace `git add -A` with `git add -u` (stage modifications only, not untracked debris).

**Verify**: walk the script mentally for all three cases (no changes / changes+branch unchanged / changes+branch advanced); each path exits 0 and only the middle one pushes. YAML parses.

### Step 3: Add a concurrency group (cheap extra safety)

At the workflow level (top of ci.yml, if not already present — verify), add:

```yaml
concurrency:
  group: ci-${{ github.ref }}
  cancel-in-progress: true
```

This cancels superseded runs when the author pushes again, which shrinks the race window for ALL jobs, not just format. NOTE: this changes CI semantics for the whole workflow (in-progress runs on the same ref get cancelled) — flag it prominently in your report; if the operator dislikes it, it's separable from Steps 1-2.

**Verify**: YAML parses.

### Step 4: Hand off

The operator verifies on real PRs: (a) same-repo PR with formatting drift → bot commit lands, run green; (b) push-during-run → push skipped with the warning; (c) fork PR → format job skipped, rest of CI runs.

## Test plan

No code tests; verification is the case-walk in Step 2 plus the operator's CI runs in Step 4.

## Done criteria

- (archived, not live)  `format:` job has the fork/event `if:` guard
- (archived, not live)  Push happens only after the stale-head check; `git add -u` replaces `git add -A`
- (archived, not live)  Concurrency group added (or explicitly dropped with the operator's agreement noted in the report)
- (archived, not live)  `git diff` on ci.yml touches only the format job + the top-level concurrency block
- (archived, not live)  `plans/README.md` status row updated (BLOCKED-on-CI-verification acceptable)

## STOP conditions

- ci.yml at HEAD shows the format job already rewritten (e.g. someone moved to a check-only model) — reconcile, don't layer guards on a different design.
- You cannot express the fork guard without breaking the `push`-event path (test the expression logic carefully) — report rather than ship a job that never runs.
- Plan 011's changes conflict textually and the rebase isn't clean — coordinate with the operator on ordering.

## Maintenance notes

- If the repo later adopts a merge queue or required-linear-history, the auto-push model needs rethinking — the report's check-only recommendation becomes relevant then.
- Reviewer focus: the three-case behavior of the commit step, and that skipped-push exits 0 (a red format job on a benign race would be worse than today).
