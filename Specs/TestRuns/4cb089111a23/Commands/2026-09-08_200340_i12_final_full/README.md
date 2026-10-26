# I12 Fresh Full qualification

Verdict: **PASSED** — 2,116 passed / 0 failed / 52 declared skips.
Run: `20260908T200723Z-69288-1db8a2860de7474ebbc2f9164aa99216`; all 35 entries executed, 35 promoted, 0 reused.

Build receipt: `f94d01029dc0bbf39e4e1786920704a08df4a05614c8d872d471cb0fd7312956`.
Tested workspace snapshot: `018041afc4cea6efb7d4b0575f7857b4f2c52ff335058c2095bc85ffb01915a1`.

Command:
```powershell
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -AllowAlternateVolumeTestRoot -TestRoot Z:\RedSalamander.Perf
```

The repaired ConPTY launch, first floating-terminal opener, shared native input warning, repeat-safe Find fixture, and queued-dismissal-safe ViewerSpace paint witness were tested together. The admitted Ghostty pin is unchanged.
The original native warning is larger, centered on the test app/dialog, DPI-aware and click-through. The owner confirmed the presentation after the 138/0/0 replay; this Full also covers its startup-before-activation ordering.

`run-all-tests-results.json` retains both verdict axes, case/entry counts, receipt/snapshot/plan/decision identities and promotion digests. The v1 sidecar and case history retain every case status and skip reason. The immutable plan, final decision and run state retain coverage/provenance.
`commands/`, `compare/`, and `fileops/` retain native suite results and traces. `outputs/` retains the standalone assertion logs available from the run; the plugin log includes terminal selftests.

The selected Commands performance rows preserve exact bytes and relative order for terminal open/render and removal-focus metrics. `perf-compaction.json` records the complete original stream hash, size and row count. This is correctness/qualification evidence, not a general latency improvement claim.

Prior failed Full/replay attempts, corrected test-enabled Release 9/0/0, live terminal entry points, the rejected duplicate warning experiment, and the native 138/0/0 replay remain in [the repair archive](../2026-09-08_180534_i12_terminal_closeout/README.md).

The [intermediate failed Full](../2026-09-08_183818_i12_fresh_full/README.md) retains its 265/250 ms FileOps UI-completion timing failure and ViewerSpace tooltip-witness race. Five cancellation repetitions passed at 16–32 ms without changing its limit. Queued mouse-leave reproduced the tooltip failure; capturing its synchronous paint witness passed five repetitions and the six-cycle viewer churn test. The failed run remains failed.

Closeout updates after the gate are limited to plan/index/link Markdown and curated evidence. The post-gate source comparison and admission/archive checks are recorded alongside this result; they are not claimed as inputs executed inside the Full run.
