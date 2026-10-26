# Pending conflict inspection context: baseline

Result: **12 passed, 0 failed, 0 skipped**. x64 Debug, four selected Commands cases repeated three times. This is a preparatory data/lifetime slice, not working pane inspection or complete R3 acceptance.

Baseline precedes implementation. Candidate adds assertions for admitted endpoint identity, provider separators/case policy, flattened source index and nested paths; invalid qualification and unmapped operations; fresh token identity, immutable snapshots; metadata-loading continuity, popup hide and cancellation invalidation. Synthetic render-only paths clear their navigation context. Existing conflict decision/metadata/close scenarios remain in the same matrix.

The common/differing-name French renderer fixture is unchanged: 480/640/760 DIP, five batches of eight redraws per width, three repetitions, 45 ordered samples per presentation. All application captures are harness-produced. These Debug timings do not qualify final Release FPS, resources, mixed DPI, UIA/AT or asynchronous pane navigation. Context allocation occurs per prompt, not per paint; this fixture does not separately quantify allocation retention.

Full raw results/metrics/captures: `S:/RedSalamander.Perf/evidence/i26-inspection-context-baseline-20260920`. Selected metrics retain source hash and original line numbers. `runner.txt` only normalizes line endings/trailing whitespace; full original runner log remains `C:/RedSalamander.Perf/evidence/i26-ui/inspection-context-baseline.log` (SHA256 `dbece133f75e6730cf45d36e248e46d734dc3c66c6e544d9d9e79ebfe683e2b2`). Build/source identities match. No required gate is waived.

Paired redraw evidence (microseconds per eight redraws, 45 samples per presentation):

| Run | Common median / p95 | Differing median / p95 |
| --- | --- | --- |
| Baseline | 66580 / 66788 | 66569 / 66764 |
| Candidate | 66581 / 66883 | 66594 / 84154 |
| Candidate confirmation | 66559 / 66845 | 66593 / 66933 |

The first candidate has six >70 ms batches, including 606798/554464 us differing-name outliers, so its tail is not silently accepted. An unchanged-source normal-build rerun was justified by this unexplained variance; it passes all 12 checks and does not reproduce the elevated p95. Both receipts remain. This is bounded fixture evidence, not proof about unmeasured allocation/resource costs or final performance acceptance. An attempted SkipBuild rerun was rejected before testing because the targeted receipt was not full-solution compatible; the governed normal build was used, with no attestation bypass.
