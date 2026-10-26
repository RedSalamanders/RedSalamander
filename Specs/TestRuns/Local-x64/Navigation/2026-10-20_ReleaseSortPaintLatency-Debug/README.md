# I44 Release sort paint latency - Debug

The apparent Extension Ascending regression was dominated by presentation waits in a synthetic warm-render sequence. The old recorder timed the product paint plus an injected invalidation, swap-chain resize/deferred initialization, icon work, message draining and extra renders. Its four-frame requests were not the same quantity as one product sort paint.

The corrected `folderView_perf_sort_toggle_stress` prepares and settles outside timing, then observes the product paint through `Present == S_OK`, without synthetic warming or invalidation inside the measured request. Failed EndDraw/Present and occluded frames cannot satisfy it. A warm no-op does not manufacture a frame. The legacy control retains its original sequence and metric name.

**This is a benchmark protocol correction, not a production speedup.** Windows comparison, the full row-ordering implementation, SetSort, rendering policy, VSync, the 5,000-file corpus and ten requests are unchanged. See the source/fixture checks. The test-only counters add observation, not production behavior.

## Counterbalanced same-executable results

Milliseconds, median [minimum - maximum]. Each protocol has two five-repeat blocks: ten observations per request. Debug runs legacy/product/product/legacy; Release runs product/legacy/legacy/product. The two protocols use one exact candidate executable receipt per profile. Different metric definitions are not pooled or converted into a product speedup percentage.

| Request | Product paint v2 ms | Legacy warm v1 ms |
| --- | --- | --- |
| 1. name ascending | 0.472 [0.324 - 3.192] | 33.005 [28.435 - 37.688] |
| 2. extension ascending | 41.217 [34.838 - 66.408] | 60.666 [57.806 - 64.029] |
| 3. time descending | 52.167 [45.050 - 85.351] | 79.106 [75.829 - 82.508] |
| 4. size descending | 42.963 [40.381 - 75.499] | 68.308 [66.857 - 70.424] |
| 5. none ascending | 35.298 [33.328 - 78.079] | 67.735 [63.793 - 75.323] |
| 6. name descending | 33.883 [32.262 - 49.694] | 63.205 [60.415 - 81.209] |
| 7. extension descending | 22.724 [18.194 - 34.537] | 40.821 [37.905 - 44.920] |
| 8. time ascending | 29.329 [25.612 - 47.221] | 59.932 [48.226 - 68.371] |
| 9. size ascending | 35.484 [32.081 - 53.923] | 56.034 [53.400 - 63.582] |
| 10. none ascending | 19.812 [17.725 - 28.028] | 51.212 [48.097 - 55.424] |

Across 100 product requests, synthetic warm calls total 0; renders total 90; successful presents total 90. Every changed sort has a successful paint; all ten no-op requests have zero renders. The 100 legacy-control requests retain one warm call each. All per-request counters remain in performance-comparison.json and original metric rows.

The initial unchanged-code baseline used the earlier build-only observer. It remains retained as an informational old-protocol reference; it does not establish absence of concurrent native test runtimes. Its Extension Ascending median/range is 58.310 ms [57.054 - 64.229 ms]. It is not substituted for the stronger same-executable counterbalanced controls.

## Runs, receipts and observation

| Phase | Selected | Run ID | Receipt | Overlap polls | Max gap ms | Observer CPU ms |
| --- | --- | --- | --- | --- | --- | --- |
| baseline-debug-01 | True | 20261020T135446Z-105980-c33ae48d89784168bdfb705492a65a52 | 5a3f7b334066b6c7cd1f36c31aa7b5ec4cbbd4a3e9c89c0eab4a011bc3827003 | 0 | 564.632 | 13500 |
| candidate-debug-legacy-block1-02 | False | 20261020T143845Z-37144-aecd545e19814d729a1de3b5d8c742f8 | e1420432e6018bdd929ab7f7e6016759e9977fdafdd823b9fb8ab3fbd8127180 | 2 | 689.927 | 12187.5 |
| candidate-debug-legacy-block1-03 | True | 20261020T144049Z-74524-c3f896fa64464e8ca3f59c2088eaf28d | e1420432e6018bdd929ab7f7e6016759e9977fdafdd823b9fb8ab3fbd8127180 | 0 | 683.581 | 12453.125 |
| candidate-debug-legacy-block2-01 | True | 20261020T144715Z-69960-fb8bca9db2974089b0195477f36ea933 | e1420432e6018bdd929ab7f7e6016759e9977fdafdd823b9fb8ab3fbd8127180 | 0 | 579.398 | 12375 |
| candidate-debug-product-block1-01 | True | 20261020T144255Z-86324-93f676caedde45fab38622ac4f66fd77 | e1420432e6018bdd929ab7f7e6016759e9977fdafdd823b9fb8ab3fbd8127180 | 0 | 731.445 | 12968.75 |
| candidate-debug-product-block2-01 | True | 20261020T144507Z-83864-4fc1f327a87a457a8da83923dbd3da09 | e1420432e6018bdd929ab7f7e6016759e9977fdafdd823b9fb8ab3fbd8127180 | 0 | 639.974 | 12546.875 |

The final v2 observer requires a 30-second quiet gate and monitors the whole canonical runner window, nominally every 200 ms. It detects compilers/build launchers, known DxUi functional/performance executables and test-profile drivers. Exact old idle MSBuild identities require dead-parent and unchanged-CPU verification; resumed CPU is a blocker. The unrelated verified GIGABYTE GCC process has a documented exact identity exclusion.

This is snapshot-based absence of detected build/known-test overlap, not a lossless process-start audit, proof of an idle host, or physical scanout measurement. Present success means the API returned S_OK. Unknown shells, actual gaps and observer CPU are retained. No independent application, build or test was interrupted.

A gate-only v1 controller was retired before any native runner launched so the known concurrent DxUi test profile could become a blocker. A subsequent block passed natively but was conservatively rejected when a metadata-preflight shell mentioning test.ps1 was observed. The metadata run itself was deferred. Every rejected measurement and the retirement remain separate; acceptance never selects by latency.

## Correctness and scope

Both isolated test-enabled x64 Debug/Release solution builds finish with zero warnings/errors. Nine native cases pass per profile: Windows natural-order/reorder contracts, five current/selection regressions, successful product-paint contracts, rendering-error persistence and device-loss recovery. The new paint contract covers changed/no-op sorts, invalid targets, failed EndDraw/Present, occlusion and successful later paint without synthetic warming in observation. All 328 focused tooling/source-contract tests pass on the qualified source; final main metadata validation is recorded separately.

All original ordering/cache/numeric-metadata rules remain in force. No production render scheduling, VSync or comparator optimization is claimed. This does not establish Full, ARM64, ASan or completely idle-host qualification. The shared sandbox audit retains 41 existing unexpected run-directory findings; native case passes do not make that audit clean.

## Evidence retention and reproduction

Every request sample and relevant sorting/layout/presentation metric row is retained in original order. Complete measurement snapshots and the final gate interval, original verdicts/traces/receipts and exact extraction rules are bound in artifact-bindings.json. Larger original streams and full build logs remain under `Z:/RedSalamander.Perf/evidence/natural-sort-release-extension-2026-10-20`; original/extract digests and byte/row counts are retained. Receipt JSON is deduplicated by immutable receipt ID, with an explicit reference per phase.

The observer detects the live concurrent DxUi Accessibility process in all five positive-control observations. The generic probe label says synthetic; probe-context.json records that the already-running real test was observed and no test was launched. The two final observer scripts differ only in case/artifact selectors. Initial v1 and final v2 source/scanner identities stay distinct.

The original I43 packet remains immutable. A separately retained Debug smoke pass overlaps the Release build and is functional evidence only. Core-ordering records include setup/restoration; separate nested stage medians are not additive. Full raw native metric streams, including enumeration/preparation records, remain digest-bound.

Canonical commands for each profile:

```powershell
Tools/Run-AllTests.ps1 -Suite Commands -Configuration Debug -SkipBuild -CaseFilter folderView_perf_sort_toggle_stress -SelfTestRepeat 5
Tools/Run-AllTests.ps1 -Suite Commands -Configuration Debug -SkipBuild -CaseFilter folderView_perf_sort_toggle_legacy_warm_control -SelfTestRepeat 5
```

The archived orchestration scripts are evidence snapshots, not new supported repository commands. The qualified checkout is `Z:/src/RedSalamander-worktrees/debug-sort-row-moves-validation`; main contains the same five C++ inputs, while the independently running main Release instance remains undisturbed.

[Other profile](../2026-10-20_ReleaseSortPaintLatency-Release/README.md) · [Independent verification](independent-verification.json)
