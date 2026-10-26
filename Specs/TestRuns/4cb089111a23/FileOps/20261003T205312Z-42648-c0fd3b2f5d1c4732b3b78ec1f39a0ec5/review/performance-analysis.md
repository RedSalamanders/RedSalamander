# Current-source F6 suspension counterfactual — analysis only

All 108 intervals retained: 54 OFF and 54 ON (18 states x 3 fresh processes). Receipt 256d18be002b07877b2e00cd18181d35fc30a3f3a702e6c436b8f8fc1ff35de7; fixture f348e82ed5a4b7cf7b5e2cc921642bfbd27ed02b3edf335b3ab4066fa99bbb03. Historical failed Debug baseline remains separate.

No optimized nonregression or I27 completion is asserted. Positive deltas below are adverse directions requiring review; CPU uses normalized actual interval duration, retaining raw coarse Windows figures in JSON/CSV.

| State | UI us/s OFF → ON | Process us/s OFF → ON | Private after bytes Δ | Working after bytes Δ | Handles after Δ | GDI/USER after Δ |
|---|---:|---:|---:|---:|---:|---:|
| initial-idle-visible | 0,0 → 0,0 | 23 425,8 → 0,0 | -241 664 | -729 088 | 0 | 0/-1 |
| initial-idle-hidden | 0,0 → 0,0 | 22 977,9 → 0,0 | -237 568 | -684 032 | 0 | 0/-1 |
| initial-idle-minimized | 0,0 → 0,0 | 0,0 → 0,0 | 4 399 104 | -679 936 | 0 | 0/-1 |
| running-visible | 69 137,2 → 23 964,7 | 69 137,2 → 23 964,7 | 4 927 488 | -36 864 | 2 | 0/0 |
| running-hidden | 0,0 → 0,0 | 0,0 → 0,0 | 4 767 744 | -233 472 | 0 | 0/0 |
| running-minimized | 0,0 → 0,0 | 0,0 → 0,0 | 4 898 816 | -131 072 | 0 | 0/0 |
| completing-visible | 0,0 → 23 461,0 | 23 286,1 → 23 461,0 | 1 708 032 | 65 536 | -2 | 0/0 |
| completing-hidden | 0,0 → 0,0 | 0,0 → 0,0 | 1 306 624 | -126 976 | 0 | 0/0 |
| completing-minimized | 0,0 → 0,0 | 0,0 → 0,0 | 1 241 088 | -229 376 | 0 | 0/0 |
| completed-idle-visible | 23 390,7 → 0,0 | 23 390,7 → 23 216,9 | -155 648 | -577 536 | 0 | 0/-1 |
| completed-idle-hidden | 23 113,9 → 0,0 | 23 531,6 → 0,0 | 110 592 | -663 552 | 0 | 0/-1 |
| completed-idle-minimized | 0,0 → 0,0 | 47 782,9 → 0,0 | -200 704 | -360 448 | 0 | 0/-1 |
| informational-running-visible | 143 348,6 → 94 553,7 | 238 914,4 → 141 830,6 | 327 680 | -180 224 | 0 | 0/0 |
| informational-running-hidden | 0,0 → 0,0 | 0,0 → 0,0 | -475 136 | -45 056 | 0 | 0/0 |
| informational-running-minimized | 0,0 → 0,0 | 69 137,2 → 0,0 | 8 192 | -69 632 | 0 | 0/0 |
| informational-completed-idle-visible | 0,0 → 23 496,2 | 0,0 → 23 855,0 | -131 072 | -606 208 | 0 | 0/-1 |
| informational-completed-idle-hidden | 0,0 → 0,0 | 0,0 → 0,0 | 253 952 | -290 816 | 0 | 0/-1 |
| informational-completed-idle-minimized | 0,0 → 0,0 | 0,0 → 0,0 | -413 696 | -569 344 | 0 | 0/-1 |

Adverse-direction observations retained: 50. JSON includes every state median/min/max/raw triplet, all 54 chronological pair comparisons, counter families, resource before/after/deltas, last taskbar state, and motion settlement durations.

Global C/D audit records remain truthful, including historical is_clean=false. Qualification is exact selected run closure plus unchanged external category/path/reason entry set; no historical-tree byte or global-clean claim.

Foreground protection is native boundary verification; no continuous HWND history is emitted. Resource retention is first-to-last observation inside F6, not process-exit or post-cleanup leak qualification. Root must assess adverse directions and measurement resolution before deciding the gate.