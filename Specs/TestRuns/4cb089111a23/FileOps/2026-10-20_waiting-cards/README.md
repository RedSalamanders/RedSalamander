# Waiting cards and fixed destinations - 2026-10-20

Waiting Copy/Move cards now show their static status, source/destination and applicable actions.
There is no counting text, zero-byte transfer counter, concurrency, percentage, ETA, graph,
whole-task marquee or source-row mini-bar while waiting. The destination is fixed after acceptance;
the destination history menu, queued re-admission transaction, replacement breadcrumb API,
destination-change prompt/refusal presentation, resources and obsolete tests are removed.

## Evidence and scope

- Final review run: 20261020T180343Z-27120-4850d371550a42c5a18c57fadc63c920.
  Setup, FileOps_ReviewRemediationContracts and Cleanup passed. Private admission/interlock probes
  use their own engine; the real stale-token clipboard witness keeps its strict idle-owner check.
  Twenty-four Copy/Move waiting scenarios and 124 broader hierarchy scenarios passed.
- Decision run: 20261020T175848Z-108816-b99a939f16c043988b99a2f091a8c1e7.
  Setup, FileOps_RepeatDecisionContracts and Cleanup passed on its own recorded artifact.
  This is separate evidence, not promoted Full or current-binary coverage.
- Resource/localization contracts: 10 passed, 0 failed, 0 skipped.
- All eight saved English/French, narrow/light and wide/dark captures were inspected.
  These are actual owned-window harness captures with synthetic Studio media paths, no admitted
  operation and no desktop takeover. Disabled actions in presentation fixtures do not prove input
  enablement. The foreground-preservation assertion passed.
- After the review, the sole source change re-finds the queued neighbor after a potentially
  reentrant UIA action in the interactive Phase5 queue journey. The final Debug build compiles
  that test-only freshness repair. Production source bytes match the review snapshot.
- These filtered runs remain repository NOT_EVALUATED. Fresh Full, other native profiles and the
  focus-taking Phase5 journey are pending; they are not inferred from gallery or compile success.

## Timing evidence

The selected metric stream preserves all 288 FileOps.Popup.WaitingCard.FrameUs rows in their original
order. summary.json records nearest-rank statistics, the full raw stream's byte/row counts and SHA256,
selected original line numbers, screenshot digests and preserved sandbox paths. Raw metrics remain
under the marked sandbox. These are candidate-only synchronous redraw timings with reduced motion,
not a throughput benchmark or an accepted performance-budget gate. No before/after improvement is claimed.

## Failed and blocked attempts

failed-documentation-* retains the failed run whose queued screenshot assertion expected old zero
counters. The actual card had no counters and said Waiting; the expectation was corrected.
failed-compile.txt retains the obsolete formatter argument / merged-scope variable compile failure.

Initial baseline reuse was blocked by the full-solution receipt requirement. A subsequent baseline
build lost attestation when the plan file was added during that build; baseline-msbuild.txt alone
does not establish a successful baseline. An early admission-fixture run failed the strict
clipboard idle-owner precondition. Its original canonical artifacts were removed by the runner's
automatic stale-run cleanup before explicit preservation; only the tool transcript retains that
failure (2 setup/cleanup passes, 1 failed phase). It is not reconstructed as a canonical receipt.
The final isolated fixture passes the original assertion. Later failed and passed raw runs were
explicitly preserved under the sandbox before launching another runner.

The original workspace also contained unrelated FolderView/tooling/spec edits. They are included
in the recorded workspace snapshots and were left intact.

Archive validation passed: explicit pre-stage (28 files) and whole-inventory (15,901 files).
