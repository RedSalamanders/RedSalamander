# I12 Fresh Full qualification

Verdict: **FAILED** — 2,113 passed / 2 failed / 53 declared skips.
Run: `20260908T183914Z-71172-fa6f9f19305b4ba7b23d571d207daf93`; all 35 entries executed, 33 promoted, 0 reused.

Build receipt: `df8c0f752b86f745aad002a9a9966294c788544fce26b086d29862903151f99c`.
Tested workspace snapshot: `64e601d44e947e758fad64e26d72e5e44bd7fd9cf87848e1c31f60194e691af6`.

Command:
```powershell
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -AllowAlternateVolumeTestRoot -TestRoot Z:\RedSalamander.Perf
```

The repaired ConPTY launch, first floating-terminal opener, shared native input warning, and repeat-safe Find fixture were tested together. The admitted Ghostty pin is unchanged.
The original native warning is larger, centered on the test app/dialog, DPI-aware and click-through. The owner confirmed the presentation after the 138/0/0 replay; this Full also covers its startup-before-activation ordering.

`run-all-tests-results.json` retains both verdict axes, case/entry counts, receipt/snapshot/plan/decision identities and promotion digests. The v1 sidecar and case history retain every case status and skip reason. The immutable plan, final decision and run state retain coverage/provenance.
`commands/`, `compare/`, and `fileops/` retain native suite results and traces. `outputs/` retains the standalone assertion logs available from the run; the plugin log includes terminal selftests.

The selected Commands performance rows preserve exact bytes and relative order for terminal open/render and removal-focus metrics. `perf-compaction.json` records the complete original stream hash, size and row count. This is correctness/qualification evidence, not a general latency improvement claim.

Prior failed Full/replay attempts, corrected test-enabled Release 9/0/0, live terminal entry points, the rejected duplicate warning experiment, and the native 138/0/0 replay remain in [the repair archive](../2026-09-08_180534_i12_terminal_closeout/README.md).

This Fresh Full is failed and cannot close I12. FileOps Phase6_LocalBandwidthThrottle recorded 265 ms UI completion versus a 250 ms limit (worker cancellation metric: 15 ms). ViewerPE interactive failed its ViewerSpace tooltip-paint assertion, also repeated by the combo-host long-run case. Both failures require investigation; no post-gate completion is claimed.

## Follow-up proof

`throttle-replay/` retains five exact repetitions (15 cases including setup/cleanup): all passed with the existing 250 ms threshold, at 32, 31, 16, 31, and 32 ms. The selected original FileOps metrics retain the worker's 15 ms cancellation and the UI witness's 265 ms observation. This is isolated diagnostic evidence; it does not reclassify this failed Full as passed.

`viewer-replays/` retains two passing exact replays with the warning visible, the deterministically failing queued-`WM_MOUSELEAVE` reproduction, and five passing runs after capturing the synchronous paint witness before pumping dismissal. The real paint count and nonempty tooltip state remain required. The renderer and warning presentation are unchanged.

The corrected six-cycle viewer combo-host churn test also passed. The focused documentation, admission, and harness contract checks passed 236/0/0 before the final Fresh Full.

## Subsequent qualification

The [final Fresh Full](../2026-09-08_200340_i12_final_full/README.md) passed on the repaired inputs. This earlier run remains failed, and its original counts and diagnostics are preserved.
