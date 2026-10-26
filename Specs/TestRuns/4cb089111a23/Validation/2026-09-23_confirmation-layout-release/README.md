# Readable, scrollable confirmations — Release qualification

Copy/Move/Delete confirmations now present through the pinned composition host. This fixes the legacy render target returning success while the captured confirmation is black. Adaptive layout retains every option and action, wraps French labels at words, and scrolls body/options in short windows while keeping actions visible. Current typed facts and safe Cancel/Escape defaults are preserved.

- Source: isolated checkout `ccebd5af68eb45f2e9deed64a7063780e71eb1ce` plus the fifteen files in `raw/candidate-inputs.json` and `raw/candidate.patch`.
- Consumer DxUi pin remains `78b3de389a189c7f86f611787e0489fb6d474218`.
- Test-enabled x64 Release solution: **0 warnings, 0 errors**, 5:27; receipt `5f7d7f452b9107dd6954b73029de087ad62246ad1c138c03c03acf68eb82a6ee`.
- Nonactivating gallery run `20260923T133713Z-54328-bf367e3d1f664137af3639b4f8204a5f`: **1 passed, 0 failed, 0 skipped**, 109.3 seconds, disk audit 0.
- **376 original PNGs**, including **76 confirmations**: 20 wide, 20 narrow, 20 short top and 16 short scrolled. All 76 confirmation images were visually reviewed; no actionable visual defects remained. Copy/Move options become fully visible when scrolled; footer actions remain unobscured. Both Delete fixtures fit in the short host.
- Compared with the preceding 340-frame passing gallery, 334 shared PNGs are byte-identical. Only six narrow French option-wrap images change; 36 short/top/scrolled captures are added. This comparison is visual provenance, not a performance/resource waiver.
- Fresh Release PR run `20260923T134003Z-16796-8c5f46205009416994c57af78fb62d14`: **13 executed/promoted entries, 824 reported passes, 0 failures/skips**, 494,496 ms, disk audit 0. No reused or quarantined result. These are runner case counts; internal Release-only capability omissions remain in the original logs.

The semantic close-glyph witness requires four pixels within tolerance 64/channel. The retained old black frame matches zero for all three severity colors. Captures also require an attached composition host and prepared frame. Safe cancellation checks hidden/released presentation and backdrop resources for each prompt.

The native geometry probe reproduces the previous short-host label-fit failure and passes after sharing the actual painted close-control inset. Candidate checks cover 320/480/920-DIP widths, 320×80 fallback, all-option reachability, action hit isolation, resize/scroll clamping and 32 warm layout reuses. Native input/UIA behavior is not inferred from model or message-dispatch checks. The final read-only lifetime review found no concrete high-impact bug, but does not qualify concurrency or device loss.

Changed alert sources pass targeted clang-format. The central registry contains pre-existing formatting issues; only the three new declaration alignments were corrected after the build, with unchanged non-whitespace bytes and line count recorded in `raw/post-build-format-only.json`. The build and test receipts bind the pre-alignment hash explicitly.

[Actual screenshots and review](../../../../Plans/WIP/FileOperationsUiReview_2026-09-16/implementation/2026-09-23-confirmation-layout-release/README.md).
`SHA256SUMS.txt` binds the byte-preserved raw receipts. Preceding failed attempts and the intermediate passing attempt remain under `raw/prior-attempts`; no failed run is promoted. Full original runs remain in `C:/RedSalamander.Perf/evidence/i26-ui/alert-adaptive-hosted-20260923/`.

This is a bounded confirmation presentation/layout slice. Required Debug/Full, English/high-contrast/192-DPI, native keyboard/hover/drag/UIA/AT, device-loss/backdrop recovery, paired resource evidence and final combined DxUi pin qualification remain open. The foreground timing question is still pending; no desktop input was attempted. I26 remains ACTIVE.
