# C0 Curl partial-tree repair — incorrectly placed control fault

- Source: `58f22d0e` plus C0 repair and expanded tests.
- Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_ -FailFast -RunId c0-curl-partial-tree-verified-debug-20260905`.
- Build receipt: `7175bd529c4994a4f40815ddcdb250ac5b1aced059fb1ae6ad52c6215d6b670c`;
  zero warnings/errors, build 2:58.
- Harness: 9 passed, 1 failed, 1 skipped because fail-fast did not reach the new
  host partial-tree witness. Native partial-tree checks: 33 passed, 4 failed.

All four original partial-tree scenarios now pass with exact mutation counts and
honest completion. All four native Move/Rename collision controls also pass with
zero mutations and safe no-commit proof. Native root guards 29/29 and lost replies
33/33 remain green.

The added two-selected-item control incorrectly expected no mutation attempt after
refusing a descendant with Continue-on-error: the serial walker continues to parent
RMD. Both concurrency settings therefore correctly emitted three total mutation
commands (two for the successful root, one refused parent RMD) and no no-commit
proof for the second root. Its state and receipt assertions failed. The correction
places refusal at the second selected root's first listing, before any mutation
attempt for that item; it does not weaken the assertion or production behavior.

Partial-tree single serial/parallel and bulk serial/parallel timings:
5,425 / 4,354 / 5,510 / 4,173 us. These are bounded loopback observations, not broad
performance qualification. Both invalid controls are retained in the full metric
stream. Raw results, trace, aggregate and complete perf stream were copied intact
from this run's `artifacts/selftest/last_run` under `C:\RedSalamander.Perf\runs`.
The runner reported 39 retained-run disk-audit issues; nothing was deleted.
