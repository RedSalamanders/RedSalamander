# I44 Release sort paint latency - Release

The apparent Extension Ascending regression was dominated by presentation waits in a synthetic warm-render sequence. The old recorder timed the product paint plus an injected invalidation, swap-chain resize/deferred initialization, icon work, message draining and extra renders. Its four-frame requests were not the same quantity as one product sort paint.

The corrected `folderView_perf_sort_toggle_stress` prepares and settles outside timing, then observes the product paint through `Present == S_OK`, without synthetic warming or invalidation inside the measured request. Failed EndDraw/Present and occluded frames cannot satisfy it. A warm no-op does not manufacture a frame. The legacy control retains its original sequence and metric name.

**This is a benchmark protocol correction, not a production speedup.** Windows comparison, the full row-ordering implementation, SetSort, rendering policy, VSync, the 5,000-file corpus and ten requests are unchanged. See the source/fixture checks. The test-only counters add observation, not production behavior.

## Counterbalanced same-executable results

Milliseconds, median [minimum - maximum]. Each protocol has two five-repeat blocks: ten observations per request. Debug runs legacy/product/product/legacy; Release runs product/legacy/legacy/product. The two protocols use one exact candidate executable receipt per profile. Different metric definitions are not pooled or converted into a product speedup percentage.

| Request | Product paint v2 ms | Legacy warm v1 ms |
| --- | --- | --- |
| 1. name ascending | 0.414 [0.242 - 1.494] | 10.462 [5.377 - 23.065] |
| 2. extension ascending | 9.268 [8.605 - 10.994] | 32.575 [25.266 - 92.511] |
| 3. time descending | 10.410 [9.232 - 12.791] | 64.330 [55.665 - 66.842] |
| 4. size descending | 8.309 [7.957 - 10.963] | 65.794 [63.839 - 66.530] |
| 5. none ascending | 6.991 [6.455 - 8.461] | 66.249 [58.204 - 67.804] |
| 6. name descending | 8.060 [7.336 - 12.203] | 65.402 [56.396 - 66.469] |
| 7. extension descending | 17.200 [3.263 - 23.322] | 66.179 [58.162 - 67.208] |
| 8. time ascending | 8.670 [7.709 - 46.850] | 66.044 [56.980 - 67.199] |
| 9. size ascending | 8.155 [7.738 - 8.892] | 66.059 [58.219 - 67.843] |
| 10. none ascending | 8.038 [6.497 - 8.952] | 66.147 [57.167 - 67.539] |

Across 100 product requests, synthetic warm calls total 0; renders total 90; successful presents total 90. Every changed sort has a successful paint; all ten no-op requests have zero renders. The 100 legacy-control requests retain one warm call each. All per-request counters remain in performance-comparison.json and original metric rows.

The initial unchanged-code baseline used the earlier build-only observer. It remains retained as an informational old-protocol reference; it does not establish absence of concurrent native test runtimes. Its Extension Ascending median/range is 41.424 ms [26.766 - 58.147 ms]. It is not substituted for the stronger same-executable counterbalanced controls.

## Runs, receipts and observation

| Phase | Selected | Run ID | Receipt | Overlap polls | Max gap ms | Observer CPU ms |
| --- | --- | --- | --- | --- | --- | --- |
| baseline-release-01 | True | 20261020T135656Z-5712-e1a53f44ac5948bbac475bc2c85e2b87 | b990d166678ae25d12ab5d321530bcc16453cd775404431d84c7e3c4df03e57b | 0 | 586.463 | 12296.875 |
| candidate-release-legacy-block1-01 | True | 20261020T145123Z-64292-6ca9e7a85a6f4340859eb1a9c19b0370 | 1a8dd44c7ed3da0ad8c2a60638182ec50c8fa9bb6da62defc0c1b1c6b20f92d5 | 0 | 561.739 | 11859.375 |
| candidate-release-legacy-block2-01 | True | 20261020T145324Z-36080-a216e22b6f0e4d1ea14f2f3a7716dcea | 1a8dd44c7ed3da0ad8c2a60638182ec50c8fa9bb6da62defc0c1b1c6b20f92d5 | 0 | 715.456 | 12187.5 |
| candidate-release-product-block1-01 | True | 20261020T144921Z-74436-20ef6ab7270f4421b38d3230d70f1626 | 1a8dd44c7ed3da0ad8c2a60638182ec50c8fa9bb6da62defc0c1b1c6b20f92d5 | 0 | 668.779 | 11750 |
| candidate-release-product-block2-01 | True | 20261020T145527Z-56916-9d340ff114904f6bb4fcecb75eed20d8 | 1a8dd44c7ed3da0ad8c2a60638182ec50c8fa9bb6da62defc0c1b1c6b20f92d5 | 0 | 746.728 | 11578.125 |

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
Tools/Run-AllTests.ps1 -Suite Commands -Configuration Release -SkipBuild -CaseFilter folderView_perf_sort_toggle_stress -SelfTestRepeat 5
Tools/Run-AllTests.ps1 -Suite Commands -Configuration Release -SkipBuild -CaseFilter folderView_perf_sort_toggle_legacy_warm_control -SelfTestRepeat 5
```

The archived orchestration scripts are evidence snapshots, not new supported repository commands. The qualified checkout is `Z:/src/RedSalamander-worktrees/debug-sort-row-moves-validation`; main contains the same five C++ inputs, while the independently running main Release instance remains undisturbed.

[Other profile](../2026-10-20_ReleaseSortPaintLatency-Debug/README.md) · [Independent verification](independent-verification.json)
