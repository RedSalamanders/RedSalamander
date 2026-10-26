# French speed selector — measured caption width

The selector now measures its actual hosted Body-font caption and reserves both
chrome lanes plus padding. A cache bound to current snapshot slots, exact caption
and retained font identity avoids rebuilding text layouts on unchanged paints.
Running, queued and discovery layouts use the same width. No library pin changed.

The identical test fixture records the original failure before the correction:
438 captures, 16 of 136 visible-label checks failing and 8 of 12 same-window value
changes failing. Release and Debug candidates pass all 136 label checks and all
12 value changes at 480/760 DIP on the native 96/144-DPI displays. Each gallery
passes 1/0/0 with zero sandbox-audit findings; both full-solution builds report
zero warnings/errors. Exact runner IDs, source overlays and build receipts are
retained in this packet. The first candidate's C2065 compile failure is retained
separately and was corrected before these runs.

All 438 baseline/Release pairs were compared: 313 images are identical and 125
change. Root reviewed key originals and checked Sol High's independent digest,
dimension and visual audit. No new overlap was found in inspected action rows.
The 438 Release/Debug pairs have identical dimensions and body pixels; differences
are confined to the configuration-specific title-bar icon. Flyout captures show
their contents, but do not establish screen-space anchoring or native input/UIA.

Five rounds of eight synchronous redraws per DPI/width give Release median
changes of +1.59%, +0.41%, -1.21% and +6.93%; overlapping raw samples are retained.
These are synchronous observations, not presented FPS, allocation, memory, idle
CPU or final P5 acceptance. No memory tradeoff is accepted by this result.

This is bounded gallery qualification, not fresh PR/Full validation of the speed
change. Maximum uint64 speed, smaller widths, completed-group/terminal text,
native input/AT and the final combined pin/resource gates remain for subsequent
slices. Complete runs remain under C:/RedSalamander.Perf/evidence/i26-ui/.

| Profile | Run | Duration (ms) |
|---|---|---:|
| baseline | `20260923T161514Z-52956-93abaa6c45d14cdf91ccf114e43c4e1c` | 164598 |
| release | `20260923T163056Z-94096-d2b913874d4a4c5e8380abb33ab6c82e` | 132867 |
| debug | `20260923T163937Z-70116-89f9fe10f00841e699d5f60c7ef80c8b` | 144077 |

[Gallery](gallery.html) · [Baseline/Release comparison](baseline-release-comparison.json) · [Release/Debug comparison](release-debug-comparison.json).
