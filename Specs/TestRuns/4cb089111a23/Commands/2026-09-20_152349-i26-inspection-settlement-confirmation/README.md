# Pane enumeration settlement events: confirmation

Result: **18 passed, 0 failed, 0 skipped**. x64 Debug, six native Commands cases repeated three times.

Retained pre-change baseline: [inspection selection final](../2026-09-20_144559-i26-inspection-selection-final/README.md), 18/18, same six-case command, unchanged render fixture and machine. Its runner reported 27 warning lines from two C4242 template sites; the earlier zero-warning checkpoint statement is corrected, not treated as evidence. No unchanged rerun was needed before adding the settlement event.

The owner now receives a UI-thread enumeration settlement containing a pane-scoped visit, exact attempted folder and HRESULT. The native fixture selects the conflict item from the successful event, observes exactly one failed event with the same failure status, safely unsubscribes from inside its callback, and verifies no later refresh callback. Legacy completion remains success-only. The event is delivered after queued commands/selections, and superseded results are suppressed. Baseline freshness/failure/reentry/prompt-byte checks and ordinary selection/navigation-shell cases remain in the matrix.

The initial 15/18 candidate is retained: successful processing moves the payload folder into the displayed model, so checking the moved-from field suppressed the success event. The corrected helper snapshots the current accepted model folder only for an interested listener. The second attempt (named final) also failed 15/18 because early failed enumeration receipts lacked their attempted folder; ExecuteEnumeration now records that folder before every early return. Both failed receipts remain diagnostic evidence, not acceptance. The verified run additionally covers reentrant legacy completion navigation suppressing an obsolete event and retaining the new one. The lazy lifetime-marker constructor now uses an exact uint8_t argument, removing its C4242 narrowing diagnostics.

This is event plumbing for R3, not complete inspection controls or a provider/foreground/popup-return transaction. No new worker, timer or default listener allocation is introduced. An active callback is retained across self-unsubscription; only interested panes incur the event work. The new callback metric is candidate-only and is not a matched latency/resource claim.

Common/differing-name French render fixture: 480/640/760 DIP, five batches of eight redraws per width, three repeats, 45 ordered samples per presentation. Baseline median/p95: common 66584/66851 us; differing 66590/66863 us. Compare summary.json; final Release, resource, native-DPI, UIA/AT and full interaction acceptance remain open.

Raw results/metrics/harness captures: `S:/RedSalamander.Perf/evidence/i26-inspection-settlement-confirmation-20260920`. Compaction retains raw hashes and original row numbers. Runner normalization only removes trailing whitespace and normalizes line endings; original log `C:/RedSalamander.Perf/evidence/i26-ui/inspection-settlement-confirmation.log`, SHA256 `59cf78d418018d42adc1e9f0eacaeccc5f0bac8d6a65c6326ee62c257a7b3d51`. Build and source receipt identities match.


The first correctness-passing run (`verified`) retained a common-name p95 spike: 719862 us vs baseline 66851 us, with four consecutive second-repetition samples 3305218/835525/719862/329036 us. No cause is established and these samples were not filtered out. This justified one unchanged-C++ confirmation, also **18/18**, zero warnings/errors/skips/audit issues. All six changed C++ worktree SHA256 values match between the two receipts; intervening documentation changes explain their different workspace identities. Confirmation snapshot: `da4af5cd36c38726a99fd8752c7b6d7986cbc2c6ee873b1cb462735711178efe`.

| Receipt | Common median / p95 (us) | Differing median / p95 (us) |
| --- | --- | --- |
| Retained baseline | 66584 / 66851 | 66590 / 66863 |
| Correctness verified, variance retained | 66601 / 719862 | 66594 / 66799 |
| Unchanged-code confirmation | 66539 / 66835 | 66548 / 66834 |

The confirmation returns near baseline; this does not establish the cause of the spike or waive final Release/resource/native interaction qualification. Do not silently replace the noisy receipt or claim a measured callback performance improvement. [Confirmation receipt](../2026-09-20_152349-i26-inspection-settlement-confirmation/README.md).
