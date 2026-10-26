# C0 bounded Curl Delete — first runtime candidate failure

Starting checkpoint `5132d866` plus the bounded cursor/iterative Delete candidate,
wide/deep/parser tests, host registration, and owning specs. Command:
`Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_ -FailFast -RunId c0-curl-streaming-delete-verified-debug-20260905`.

Build: 2:34, zero warnings/errors. Receipt
`1563cbf22a725f803c869620dd8cf3ae2131561fa73aa7ff014e551ab2133e15`,
source snapshot `3645f367a1305088ce46e0f616843bf1330cc281a5808c4261468a5e9b8ddba1`.
Harness: 9 passed, 1 failed, 3 skipped (not reached after failure); native
partial-tree assertions: 32 passed / 5 failed. This is not qualification evidence.

Earlier-child deletion now precedes a later refusal at both concurrency settings.
The independent selected-item controls exposed lost error specificity: a refused
root listing under Continue-on-error became ERROR_PARTIAL_COPY instead of preserving
E_ACCESSDENIED. The next candidate preserves the exact root error, leaving
continuation of other selected items to DeleteItems. Additional diagnostics were
required to distinguish the other three failing state assertions.

The five copied files preserve all results/reasons, traces and the complete ordered
perf stream from the exact run-owned C:\RedSalamander.Perf artifacts. The runner's
45 retained-run disk-audit issues remain explicit; no older evidence was removed.
