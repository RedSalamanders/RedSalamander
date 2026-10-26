# I45 Debug sort writer optimization - Release

The unchanged 5,000-file product-paint scenario attributes this UI thread's synchronous JSONL writer cost before and after the optimization. Debug enables this diagnostic path by default. The old writer copied and converted constant run metadata and formatted timestamp/numeric fields repeatedly, including one record for every drawn item.

The candidate caches escaped metadata per configuration, proves short ASCII into bounded stack scratch, retains the strict Unicode fallback, appends unescaped runs in bulk, formats ASCII numbers/timestamps directly and reuses a row buffer with a 64-KiB retained ceiling. Configuration and append share one mutex. Appends remain synchronous and immediately readable; no metric is sampled, suppressed, deferred or removed. The shared partial-write helper preserves complete successful appends.

Windows natural comparison, listing-local ranks, full row reordering, current/selection, Time/Size numeric keys, layout, drawing, presentation, VSync, checked Debug iterators and graphics validation are unchanged. This is a captured/default-Debug path optimization; it does not establish the same improvement with capture disabled. Both versions use identical test-only per-thread writer observations. Counter reads never acquire the writer lock or hide contention outside the timer. Writer time is inclusive elapsed time, including contention, not CPU time.

The complete pairs localize the large Debug gain to recording: Extension Ascending paint median falls from 34.701 to 19.810 ms and counted writer spans from 17.312 to 2.479 ms. The per-request arithmetic remainder stays about 17 ms (17.467 to 17.205 ms). Debug ordering and drawing retain their ordinary unoptimized/checking costs: candidate OrderItems is 6.527 ms and begin-to-EndDraw 6.728 ms, versus Release 0.648/3.110 ms. These nested scopes include other diagnostic work and are not additive or pure CPU samples.

Release Extension Ascending is a smaller median change, 8.984 to 8.591 ms, with overlapping ranges. Name Descending and the final None request have small higher candidate paint medians, also with overlapping ranges; no uniform Release speedup is claimed. The large Time Ascending change (24.106 to 7.945 ms) includes Present falling from 17.407 to 2.497 ms while writer spans fall only 0.959 to 0.568 ms. The data does not isolate why that presentation wait changed; it cannot establish a 16-ms CPU/comparison gain.

## Same-protocol, counterbalanced results

Milliseconds, median [minimum - maximum]. Each role has two five-repeat blocks and ten observations per request. Debug runs baseline/candidate/candidate/baseline; Release runs the reverse. Each role/profile has one exact executable receipt. The fixture, ten requests and product_paint_v2 timer boundaries are unchanged.

| Request | Baseline paint ms | Candidate paint ms | Baseline writer ms | Candidate writer ms |
| --- | ---: | ---: | ---: | ---: |
| 1. name ascending | 0.319 [0.121 - 1.224] | 0.290 [0.174 - 0.552] | 0.000 | 0.000 |
| 2. extension ascending | 8.983 [8.441 - 10.383] | 8.591 [6.642 - 13.739] | 1.591 | 0.883 |
| 3. time descending | 9.613 [8.942 - 10.839] | 8.739 [7.607 - 15.888] | 2.170 | 1.213 |
| 4. size descending | 8.652 [7.909 - 9.558] | 7.693 [6.976 - 10.983] | 1.889 | 1.074 |
| 5. none ascending | 6.878 [6.483 - 7.424] | 6.657 [5.673 - 9.195] | 1.402 | 0.801 |
| 6. name descending | 7.765 [7.014 - 9.165] | 7.979 [6.686 - 14.123] | 1.246 | 0.719 |
| 7. extension descending | 11.893 [3.377 - 24.371] | 10.799 [3.437 - 47.309] | 0.487 | 0.314 |
| 8. time ascending | 24.106 [7.716 - 46.912] | 7.944 [7.528 - 39.607] | 0.959 | 0.568 |
| 9. size ascending | 8.014 [7.606 - 19.957] | 7.954 [7.679 - 10.381] | 1.490 | 0.796 |
| 10. none ascending | 8.052 [5.939 - 8.383] | 8.201 [5.654 - 10.666] | 0.432 | 0.268 |

Every request and its writer-call, render, successful-Present and synthetic-warm counters remain in performance-comparison.json and selected original JSONL rows. A no-op cannot manufacture a frame. Changed sorts require a successful S_OK presentation; failed/occluded attempts do not qualify. Separate nested stage distributions are not additive, and recording cost inside drawing/layout scopes is not pure graphics CPU time. Full originals retain all per-item draw and frame families.

### Extension Ascending attribution

| Recorded scope | Baseline ms [range] | Candidate ms [range] |
| --- | ---: | ---: |
| FolderView.ApplyCurrentSort.OrderItems | 0.538 [0.449 - 0.875] | 0.648 [0.499 - 2.262] |
| render.layout_items_us | 2.333 [2.179 - 2.509] | 2.139 [1.973 - 4.006] |
| render.begin_to_enddraw_us | 3.418 [3.040 - 3.841] | 3.110 [2.696 - 4.605] |
| render.present_us | 0.272 [0.208 - 0.339] | 0.284 [0.071 - 0.862] |
| Request elapsed less counted writer spans | 7.470 [6.894 - 8.701] | 7.747 [5.897 - 11.109] |

The arithmetic remainder is computed for each original request before grouping. It includes other application/test work, graphics waits and per-call microsecond quantization; it is not pure CPU time or an estimate of capture-disabled latency. Nested scopes overlap and must not be added together.

Preparation remains separate from the request timer. The following retains every 5,000-item enumeration rank/sort-merge record, including counts, rather than selecting a fast initialization. These nested stage durations are not end-to-end folder-open latency.

| Preparation stage | Baseline ms [range], count | Candidate ms [range], count |
| --- | ---: | ---: |
| FolderView.ExecuteEnumeration.NaturalSortRanks | 0.639 [0.541 - 4.498], 20 | 0.628 [0.536 - 2.388], 20 |
| FolderView.ExecuteEnumeration.SortMerge | 4.143 [2.418 - 8.774], 20 | 3.958 [2.377 - 8.715], 20 |

## Runs and observation

| Role/block | Run ID | Receipt | Overlap polls | Max polling gap ms |
| --- | --- | --- | ---: | ---: |
| candidate-release-block1-01 | 20261020T224206Z-50348-7df13b97d0b54a4c847b7a5ce87204f8 | 63d844b82d44ad1a769dd81abd1e8696c13883c6f41f4d3a61c4b1ec7ce5513f | 0 | 744.086 |
| baseline-release-block1-01 | 20261020T224413Z-48608-0b678b49129b4f5d9454b8b4d2ecc243 | 9f7ad9ab6e85417d840e0a26efa5965414bf1228aa455dae165b056b19da5c3b | 0 | 853.258 |
| baseline-release-block2-02 | 20261020T224844Z-15504-18987b41568a4f3bb8052d5c944a1602 | 9f7ad9ab6e85417d840e0a26efa5965414bf1228aa455dae165b056b19da5c3b | 0 | 947.178 |
| candidate-release-block2-01 | 20261020T225102Z-97972-09c31189ef794186a17ceb99abec0395 | 63d844b82d44ad1a769dd81abd1e8696c13883c6f41f4d3a61c4b1ec7ce5513f | 0 | 852.882 |

All builds finish before timing. The shared observer requires a 30-second quiet gate and monitors the canonical runner window, nominally every 200 ms. It blocks compiler/build launchers, known DxUi test/benchmark executables, test drivers and RedPrism --self-test. Real already-running compiler and RedPrism self-test positive controls are retained; the generic probe label says synthetic, but neither control launched a test/compiler. Exact idle MSBuild identities require exited-parent and stable-CPU verification. Resumed CPU or a new tool blocks eligibility. Every rejected attempt is retained; acceptance never selects by latency.

This is snapshot-based absence of detected compilation or known native-test overlap, not a lossless process-start audit, proof of an idle host, or physical scanout timing. Poll gaps, unknown shells, process exclusions and observer CPU remain explicit in each observation. Two authorized chats were asked to hold new work after current operations finish; a third chat was not messaged at the user's direction. The observer waits for its compilation/tests to finish too; no independent process is interrupted.

## Correctness, source and qualification

Final test-enabled x64 Debug/Release builds and the final ten native cases per profile are separate from the superseded prototype. Native coverage includes complete JSONL schema/field order, controls and Unicode/malformed UTF-16, short-ASCII boundaries, uint64/HRESULT/counter semantics, immediate readback, concurrent rows, oversized scratch release, clear/reconfigure and same-path rotation. It also retains Windows ordering, current/selection, paint/no-op/failure/occlusion, rendering-error persistence and device-loss coverage. Source and tooling/format gates are recorded separately. No Full, ARM64 or ASan qualification is implied.

Baseline source is a6ced5f41 with only identical test-only writer observation added to the old producer. Candidate C++ matches the user-facing checkout. Source preservation strips exactly the baseline diagnostic blocks and compares the remaining producer with a6ced5f41; all protected product/compiler/pin inputs match in all three checkouts. Final identities bind each compiled source snapshot, receipt and current runtime closure.

A disk-full Z: volume caused preparation failures. The owned baseline and candidate checkouts were moved to E: and S:, then canonically restored and re-attested. Every final baseline/candidate run explicitly uses the same marked S:/RedSalamander.Perf fixture and capture root. Executable paths differ; loading/warm-up is outside measured requests. These fresh pairs must not be presented as a direct comparison with historical Z:-root absolute timings. Historical Z: sandbox findings remain untouched; the fresh S: audit has its own scope.

The unattested first build, shared-dependency refusal, disk-full compilation/restore failures, stale generated-path recovery and superseded prototype receipts remain in attempt-ledger.json with digest-bound full logs. The first Debug native pass succeeded but its wrapper assumed the former Z: artifact location; exact S: artifacts were recovered without changing/rerunning that result. Final native runs use the corrected explicit root and final nonblocking observer.

Completed attempts that overlap new compilation or exceed polling bounds are rejected even when native cases pass. Gates were also deliberately closed before any runner launched for bounded intervals granted to the already-authorized chat. Those interrupted gates have incomplete polling buffers and zero timing authority. Only first eligible completed blocks enter the comparison.

One excluded Release-baseline attempt was a conservative classifier false positive: an otherwise read-only tooling inspection referenced test.ps1 in its command text. Its original flagged poll and full command remain retained. It is not described as actual compilation, and no eligibility override or latency-based reselection was applied; the next eligible complete block is used.

After three accepted blocks, resumption preserves their exact fixed-order prefix instead of rerunning or reselecting samples. Fifteen guard tests plus actual verification cover source/receipt/profile/dependency identity, current compiled inputs/runtime closure under canonical profile locks, all required complete copied artifacts and the original polling rows/observer versions. Any absent ephemeral canonical source paths are explicitly listed in orchestrator-resumes.jsonl; the captured complete byte copies must still match the originally recorded hashes. No native code, timer, corpus or observer changes during this pause/resume. The original pre-resume orchestration source is retained too.

## Evidence and reproduction

artifact-bindings.json binds complete originals/extracts, byte counts, hashes and the exact selected metric families. All request samples are retained; full raw streams and all failed/rejected attempts remain under Z:/RedSalamander.Perf/evidence/debug-sort-product-cost-2026-10-20. Selected families retain every original row in order. Omitted per-item and other frame/preparation families remain intact externally and are included in analysis/metric counts. Receipt files are deduplicated by immutable ID. Archived orchestration scripts are evidence snapshots rather than supported repository commands.

Complete process observations, including every gate and measurement row, use lossless .jsonl.gz storage to bound the packet size after long waits. Bindings include original and compressed hashes/lengths. Independent verification hashes the decompressed original bytes before checking all measurement polls; no observation row is filtered or dropped.

```powershell
$env:REDSALAMANDER_TEST_ROOT='S:\RedSalamander.Perf'
Tools/Run-AllTests.ps1 -Suite Commands -Configuration Release -SkipBuild -CaseFilter folderView_perf_sort_toggle_stress -SelfTestRepeat 5
```

[Independent verification](independent-verification.json)
