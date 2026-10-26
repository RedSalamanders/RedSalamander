# Original P0 versus current File Operations presentation

Ten fresh x64 Release processes pass (five per revision), in the predeclared ABBAABBAAB order, with zero disk-audit findings. All 270 copied runner files were independently rehashed. The test fixture is byte-identical, SHA256 `9c817b6cfe7570a3fea9c2d0b5a95da9379a5368df7612debd80628f095224fd`. Every run uses French/dark/reduced-motion/default-density, eight running Copy/Move cards and 128 selected Issues rows, equal workload digest `14201940575619722830`, and 1140 x 720 client pixels at 144 DPI for both windows. Power, hardware, font, process, executable and source identities are retained.

A is original P0 `e61e349c3e8d133e9d52a4c63ba4c49266bc4db5` with DxUi `f5c7fc7403e352425b8e82b7a4da1c042d942ede`. Only test registration, its private helper include, existing-render-count accessor and common fixture are overlaid. Its first build failed on the missing test-only include; that failed log is retained. Corrected full-solution build receipt is `f510718fd9d42c265bcc6317b12012088f266f7d4776a0e2930b4827bdeca7ed`.

B is the frozen V5 consumer on isolated base `ccebd5af68eb45f2e9deed64a7063780e71eb1ce`, 36 overlays matching primary `a86dd111`, pin `78b3de389a189c7f86f611787e0489fb6d474218`. Full-solution receipt is `e231b251d2a6a43e2e209b45aae9b2959b8b969727caeda6e693040f4d34f2da`. The current grid/menu branches are not adopted. This comparison includes the whole intervening consumer/library redesign; it cannot attribute a difference to one control or optimization.

| Synchronous work (eight redraws/updates per sample) | P0 median / p95 (ms) | Current median / p95 (ms) | Samples per revision |
| --- | ---: | ---: | ---: |
| Running popup redraw | 66.711 / 85.529 | 66.652 / 66.844 | 25 |
| Eight-task progress churn | 66.556 / 68.502 | 66.661 / 66.820 | 160 |
| Selected Issues redraw | 66.649 / 150.113 | 66.684 / 150.062 | 25 |

P95 uses pooled nearest-rank samples; the raw report also retains each process median. These are synchronous wall times, not presented FPS or frame intervals. The median differences are small; this fixture does not show a material redraw slowdown. Sparse tail samples and display synchronization limit stronger timing claims. No threshold was invented or relaxed.

Median process-private growth from each process's pre-popup sample is 44.445 MiB for P0 versus 8.227 MiB current when the popup first settles. After the full popup/Issues lifetime, it is 29.945 MiB versus -1.660 MiB. A negative net delta means sampled process-private bytes were lower than the cold sample; these snapshots do not identify what changed, so this is not evidence of negative UI allocation. Actual current private memory is lower at every post-cold stage: first-open 263.66–269.37 MiB versus P0 299.53–304.04 MiB. Full per-stage absolute and relative values are retained. Process CPU through teardown has median 1,421,875 us P0 versus 1,046,875 us current; the Windows process/thread accounting resolution is coarse.

All ten runs verify zero hosted rendering while hidden, restored selected row identity, removed diagnostic history and return of host/attachment counts after each destruction. Remaining process handles/GDI/USER counts are recorded rather than assumed to be zero. This is bounded lifetime evidence, not a long-duration leak proof.

The original-P0 common-workload comparison is complete. It does not qualify surface bytes, final combined pin, newer selected-detail/graph-disclosure parity, native monitor movement, keyboard/UIA/AT, or the complete P5 resource matrix. Equal-feature comparisons and those required gates remain open. The accepted grid memory tradeoff is not applied to this consumer measurement.

Full runner trees and complete metric streams remain in `C:/RedSalamander.Perf/evidence/i26-ui/common-p0-current-crossover-20260927`. This packet retains all outcomes, source overlays, build receipts and exact raw common-metric lines with source hashes/line numbers. Empty process snapshots produced no file through PowerShell's empty pipeline; the file map explicitly records those omissions. Before/after snapshots cannot exclude activity occurring only between them. Lightweight agent file checks, validators and Git operations continued during this sequence; it was not an exclusively idle machine. Timing values remain observational and do not establish a hard non-regression budget. SHA256SUMS binds every payload. Baseline smoke receipt and original failed setup stay in their separate external directories; the ten-run comparison starts only after both fixture smokes pass.

A separate Luna audit verified the 169 payload hashes and all reported values. Root checked its findings and corrected the initial draft's unsupported attribution of the negative process delta; no measurements or outcomes changed.
