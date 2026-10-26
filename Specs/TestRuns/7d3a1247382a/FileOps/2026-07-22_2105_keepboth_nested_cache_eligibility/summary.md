# EG-FOPS-01 cached Keep Both eligibility closeout

## Claim

A task-level cached Keep Both decision is now reused only when the current Exists
conflict is the exact top-level per-item COPY/MOVE destination. A later child collision
inside a merged directory returns to the ordinary provider conflict flow and cannot
trigger a retry of the partially processed root.

## RED evidence

- Unfixed-code run:
  `Specs/TestRuns/7d3a1247382a/FileOps/2026-07-22_202107/`
- Runner ID:
  `20260722T182102Z-49896-344cc6c22ae74eefb38573285309eae4`
- Result: 2 passed, 1 failed, 0 skipped.
- Exact failure:
  `Cached Keep Both scenario copy-sequential completed without prompting for the ineligible nested collision.`

The deterministic pause was reached only after the earlier merged-directory sibling
had transferred, so the failure reproduced the vulnerable partial-root interleaving.

## GREEN evidence

| Gate | Result | Source run |
|---|---:|---|
| Debug focused cached Keep Both matrix | 3 passed / 0 failed | `2026-07-22_202903` |
| Debug complete Phase 9 | 11 passed / 0 failed | `2026-07-22_203027` |
| Debug complete File Operations | 112 passed / 0 failed / 20 expected skips | `2026-07-22_204841` |
| Test-enabled Release focused matrix | 3 passed / 0 failed | `2026-07-22_205643` |
| Debug existing top-level conflict prompt | 1 passed / 0 failed | runner `20260730T082023Z-37776-d77936d2e4074b41b0cf89be0360931f` |
| Source contracts | 150 passed / 0 failed | `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` |

The matrix covers COPY and MOVE at concurrency 1 and 2. Every scenario asserts two
prompts, no nested Keep Both action, unchanged skipped-child source/destination bytes,
one transferred sibling, correct MOVE source cleanup, and no `tree (2)` root retry.

## Build evidence

- Debug/x64 RedSalamander: 0 warnings, 0 errors.
- Test-enabled Release/x64 RedSalamander: 0 warnings, 0 errors.
- Clean tests-disabled shipping Release/x64 rebuild: 0 errors and 3 existing
  unrelated unused-function warnings. Log:
  `.build/logs/msbuild-20260730_101125_366.log`.

## Performance interpretation

The pre-repair Apply-to-all convergence sample was 1,014,807 us. Fixed-code samples
were 1,068,247 us, 1,123,071 us, and 1,219,868 us. This is a small,
prompt-wait-dominated sample with scheduler variance, so it is evidence of no clear
regression signal rather than a percentile claim. The optimized Release four-scenario
matrix measured 235,000-297,000 us. Selected raw metric rows are retained in
`perf_metrics.jsonl`.

## Curated archive note

The generated full-suite perf stream is intentionally not checked in because it is
about 165 MiB. This archive retains the exact compact results and selected metric rows
needed to review the claim; raw generated run folders remain available locally.
