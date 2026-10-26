# Release File Operations suite and alternate-volume follow-up

The initial noninteractive x64 Release suite reports **191 passed, one failed, 21
skipped**, exit 1, with zero sandbox-audit findings. The sole failure is
`R4A19_DiscoveryIndependentVolumes`: the runner invocation omitted the required
`-AllowAlternateVolumeTestRoot` switch. After adding that switch and selecting the
already marked `C:/RedSalamander.Perf` alternate root, the exact case passes three
repetitions: **9/9 including setup/cleanup, zero failures/skips**, exit 0, clean audit.

The initial run remains failed. The follow-up establishes the corrected case, not a
new complete FileOps or Full pass. No production code, timing threshold, default
authorization policy or test assertion changed to obtain this result.

| Attempt | Run ID | Result |
| --- | --- | --- |
| Initial canonical FileOps | `20260923T090537Z-78784-2907207f3403449cb063bc52d0d63137` | 191 / 1 / 21; 895,653 ms runner duration |
| Corrected alternate-volume case, repeat 3 | `20260923T092317Z-62248-03267df065c742a5b6efb12953af0d7b` | 9 / 0 / 0; 23,735 ms runner duration |

Both use clean detached source `ccebd5af68eb45f2e9deed64a7063780e71eb1ce` at
`C:/RsI26Final`, the existing DxUi pin `78b3de389a189c7f86f611787e0489fb6d474218`, and
test-enabled x64 Release receipt
`09848e19caf0b612e0663dae09d6079c17255b4862d4bd955e7ebed7b86ed1f1`.
The [preceding Release PR packet](../../Validation/2026-09-23_082913_release-pr/README.md)
retains the complete build receipt and build output. Its PR pass is a separate gate.

Both FileOps plans declare `RequiresInteractiveDesktop=false`; the runner launches
`--fileops-selftest --selftest-no-activate`. The primary run root is the previously
initialized `E:/RedSalamander.Perf`. The corrected driver explicitly sets
`REDSALAMANDER_TEST_ALTERNATE_ROOT=C:/RedSalamander.Perf`; the harness validates the
fixed, distinct volume and marked root and owns its generated run child. Historical
C:/D: artifacts are preserved.

The 21 skips retain their reasons in the results: the shared-scheduler shutdown
export is Debug-only, two 7z cases are owned by Compare, and eighteen remote-provider
cases lack their configured connection profiles. No remote credentials were supplied.
The provider-capability matrix no longer exhausts the stack in this Release run.

## Retained evidence

- [Initial aggregate](raw/initial/run-all-tests-results.json), [per-case results](raw/initial/fileops/results.json),
  [trace](raw/initial/fileops/trace.txt), and [exact driver](raw/initial/qualify-fileops.ps1.receipt.txt).
- [Follow-up aggregate](raw/alternate-volume-retry/run-all-tests-results.json),
  [per-case results](raw/alternate-volume-retry/fileops/results.json),
  [trace](raw/alternate-volume-retry/fileops/trace.txt), and
  [corrected driver](raw/alternate-volume-retry/qualify-alternate.ps1.receipt.txt).
- Complete ordered case histories, runner output, workspace identity and plans for
  both attempts. These focused FileOps invocations emit aggregate v1, not Full/PR v2
  promotion certificates. No reusable or composite Full verdict is claimed.
- [SHA256SUMS.txt](SHA256SUMS.txt) binds the 23 retained raw/identity files.
  [External metric identities](external-metrics.json) bind the complete, unmodified
  metric streams retained outside Git. The broad run and targeted repeat are different
  workloads; they are not a paired UI-performance comparison or a new baseline.

The complete original copies remain under
`C:/RedSalamander.Perf/evidence/i26-ui/release-fileops-20260923/`.
The [active I26 plan](../../../../Plans/WIP/UI_FileOperationsReview_2026-09-16.md)
retains the foreground, Full, combined-pin, DPI/accessibility and resource gates.
