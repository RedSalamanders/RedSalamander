# File Operations interaction review: input

Result: **1 passed, 0 failed, 0 skipped**. This opt-in gallery case validates capture execution, not redesigned UI acceptance or Full repository coverage. No application source changes are included in this review.

The input lane uses the real application renderer with synthetic presentation fixtures and genuine French resources. Native display capture reports 144 and 96 DPI. The input lane uses the existing warning, desktop lease, owned-window checks and restoration; it captures Tab/Shift+Tab, hover, pressed and same-HWND return to the first display. It does not execute the complete decision or navigation matrix.

Source snapshot: `6174d1ab2428470c9d3992b649ce3c20aac43490d573d34a790733f071937e67`. Build receipt: `2c2fea151678b1a708c6c1f19794fe79d615e15b32c9bd22f1c244bb784f57bb`. Raw evidence: `S:/RedSalamander.Perf/evidence/i26-interaction-review-input-20260920`. Reviewed images and capture hashes are retained in the [interaction review](../../../../Plans/WIP/FileOperationsUiReview_2026-09-16/interaction-review-2026-09-20/README.md). No performance non-regression claim is made.
