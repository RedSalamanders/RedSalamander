# I45 Debug sort writer optimization - Debug

The unchanged 5,000-file product-paint scenario attributes this UI thread's synchronous JSONL writer cost before and after the optimization. Debug enables this diagnostic path by default. The old writer copied and converted constant run metadata and formatted timestamp/numeric fields repeatedly, including one record for every drawn item.

The candidate caches escaped metadata per configuration, proves short ASCII into bounded stack scratch, retains the strict Unicode fallback, appends unescaped runs in bulk, formats ASCII numbers/timestamps directly and reuses a row buffer with a 64-KiB retained ceiling. Configuration and append share one mutex. Appends remain synchronous and immediately readable; no metric is sampled, suppressed, deferred or removed. The shared partial-write helper preserves complete successful appends.

Windows natural comparison, listing-local ranks, full row reordering, current/selection, Time/Size numeric keys, layout, drawing, presentation, VSync, checked Debug iterators and graphics validation are unchanged. This is a captured/default-Debug path optimization; it does not establish the same improvement with capture disabled. Both versions use identical test-only per-thread writer observations. Counter reads never acquire the writer lock or hide contention outside the timer. Writer time is inclusive elapsed time, including contention, not CPU time.

The complete pairs localize the large Debug gain to recording: Extension Ascending paint median falls from 34.701 to 19.810 ms and counted writer spans from 17.312 to 2.479 ms. The per-request arithmetic remainder stays about 17 ms (17.467 to 17.205 ms). Debug ordering and drawing retain their ordinary unoptimized/checking costs: candidate OrderItems is 6.527 ms and begin-to-EndDraw 6.728 ms, versus Release 0.648/3.110 ms. These nested scopes include other diagnostic work and are not additive or pure CPU samples.

Release Extension Ascending is a smaller median change, 8.984 to 8.591 ms, with overlapping ranges. Name Descending and the final None request have small higher candidate paint medians, also with overlapping ranges; no uniform Release speedup is claimed. The large Time Ascending change (24.106 to 7.945 ms) includes Present falling from 17.407 to 2.497 ms while writer spans fall only 0.959 to 0.568 ms. The data does not isolate why that presentation wait changed; it cannot establish a 16-ms CPU/comparison gain.

## Same-protocol, counterbalanced results

Milliseconds, median [minimum - maximum]. Each role has two five-repeat blocks and ten observations per request. Debug runs baseline/candidate/candidate/baseline; Release runs the reverse. Each role/profile has one exact executable receipt. The fixture, ten requests and product_paint_v2 timer boundaries are unchanged.

| Request | Baseline paint ms | Candidate paint ms | Baseline writer ms | Candidate writer ms |
| --- | ---: | ---: | ---: | ---: |
| 1. name ascending | 0.498 [0.349 - 0.813] | 0.431 [0.317 - 0.670] | 0.000 | 0.000 |
| 2. extension ascending | 34.701 [32.175 - 37.510] | 19.809 [18.040 - 24.275] | 17.312 | 2.479 |
| 3. time descending | 45.193 [43.919 - 48.743] | 22.206 [20.276 - 22.452] | 26.290 | 3.715 |
| 4. size descending | 40.188 [37.934 - 43.879] | 19.049 [17.960 - 19.637] | 22.817 | 3.164 |
| 5. none ascending | 34.227 [32.384 - 36.150] | 17.221 [16.452 - 19.663] | 17.526 | 2.474 |
| 6. name descending | 32.203 [30.578 - 34.201] | 17.712 [17.039 - 21.948] | 15.707 | 2.231 |
| 7. extension descending | 17.201 [16.521 - 18.691] | 12.960 [11.952 - 13.903] | 5.336 | 0.779 |
| 8. time ascending | 25.767 [24.164 - 29.329] | 15.762 [15.237 - 16.825] | 10.588 | 1.540 |
| 9. size ascending | 31.790 [30.590 - 35.639] | 16.770 [15.738 - 18.785] | 16.782 | 2.371 |
| 10. none ascending | 17.312 [16.422 - 19.670] | 14.357 [12.653 - 15.547] | 4.226 | 0.683 |

Every request and its writer-call, render, successful-Present and synthetic-warm counters remain in performance-comparison.json and selected original JSONL rows. A no-op cannot manufacture a frame. Changed sorts require a successful S_OK presentation; failed/occluded attempts do not qualify. Separate nested stage distributions are not additive, and recording cost inside drawing/layout scopes is not pure graphics CPU time. Full originals retain all per-item draw and frame families.

### Extension Ascending attribution

| Recorded scope | Baseline ms [range] | Candidate ms [range] |
| --- | ---: | ---: |
| FolderView.ApplyCurrentSort.OrderItems | 6.250 [5.900 - 6.946] | 6.527 [6.109 - 6.889] |
| render.layout_items_us | 9.424 [9.095 - 11.016] | 3.022 [2.922 - 3.240] |
| render.begin_to_enddraw_us | 14.598 [12.821 - 16.341] | 6.728 [5.810 - 7.867] |
| render.present_us | 0.268 [0.231 - 0.373] | 0.268 [0.226 - 0.429] |
| Request elapsed less counted writer spans | 17.467 [16.464 - 19.979] | 17.205 [15.795 - 22.020] |

The arithmetic remainder is computed for each original request before grouping. It includes other application/test work, graphics waits and per-call microsecond quantization; it is not pure CPU time or an estimate of capture-disabled latency. Nested scopes overlap and must not be added together.

Preparation remains separate from the request timer. The following retains every 5,000-item enumeration rank/sort-merge record, including counts, rather than selecting a fast initialization. These nested stage durations are not end-to-end folder-open latency.

| Preparation stage | Baseline ms [range], count | Candidate ms [range], count |
| --- | ---: | ---: |
| FolderView.ExecuteEnumeration.NaturalSortRanks | 4.777 [4.624 - 5.307], 18 | 4.907 [4.602 - 5.268], 18 |
| FolderView.ExecuteEnumeration.SortMerge | 37.795 [23.650 - 45.149], 20 | 38.733 [25.459 - 43.611], 20 |

## Runs and observation

| Role/block | Run ID | Receipt | Overlap polls | Max polling gap ms |
| --- | --- | --- | ---: | ---: |
| baseline-debug-block1-04 | 20261020T205910Z-68428-7079b1d368264389878cc0a7e024c4f4 | 132aabce2ec14a1ba0e4931ad9c9ed0a86c4b77b77c61b145860b3899815f612 | 0 | 573.343 |
| candidate-debug-block1-01 | 20261020T210123Z-81532-a74f9aac91564ce2a0f032b9a4270a9a | c6963562258a28ed57e797d11423b84500da7368eb0c918cd8f3bf42ce7c7508 | 0 | 527.188 |
| candidate-debug-block2-01 | 20261020T210334Z-68324-b35976d9990f4eaba41b067fa00a50d9 | c6963562258a28ed57e797d11423b84500da7368eb0c918cd8f3bf42ce7c7508 | 0 | 466.111 |
| baseline-debug-block2-04 | 20261020T223958Z-26100-655c9e0bd9d64f51b719aa9a386dff28 | 132aabce2ec14a1ba0e4931ad9c9ed0a86c4b77b77c61b145860b3899815f612 | 0 | 472.036 |

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
Tools/Run-AllTests.ps1 -Suite Commands -Configuration Debug -SkipBuild -CaseFilter folderView_perf_sort_toggle_stress -SelfTestRepeat 5
```

[Independent verification](independent-verification.json)
