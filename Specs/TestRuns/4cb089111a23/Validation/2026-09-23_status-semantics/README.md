# Status semantics: matched Release gallery

The retained baseline and candidate use the same 17-source witness fixture, isolated build Git HEAD, dependency identity, x64 Release profile, French scenarios, and two native displays (144/96 DPI). Only Popup.cpp differs. The baseline gallery fails its four semantic families as designed: status chips 12/12, Done headings 8/8, completed-group badges 4/4, and hosted progress 6/18. The candidate gallery passes those counts and the previously passing terminal, group layout, and speed-label checks. The test-run disk audit is clean.

Baseline run `20260923T175208Z-8708-449dd3bbbbab45f6834b89d7270df21d` (failed 1); candidate run `20260923T180145Z-96520-f7e68d8f7ee34d96823bac9292d3247b` (passed 1). All 508 PNG pairs are hash/dimension/pixel-bound checked; 40 images differ. The 24 raw trace-bound redraw sample pairs are recomputed in `layout-comparison.json`.

The 24 paired in-process redraw medians range from -11.27% to +1.20% in time and -8.20 to +15.35 MiB in private bytes. These observations are not presented FPS, heap attribution, GPU bytes, or a final performance acceptance.

The gallery below contains 36 original selected captures; all 508 capture hashes remain in the comparison JSON. Scene 139 is synthetic text-fit stress, not a reachable stable engine case. This packet is Release-only: Debug, PR, Full, UIA, native-user interaction, and final consumer/performance acceptance remain open. Primary source HEAD is `9f2e1e9e`; isolated build HEAD is `ccebd5af` with the explicit hashed overlays. The governed runner validated the build receipt and executable before execution; the curator independently checks local archived provenance and comparisons, not a standalone cryptographic attestation.

[Focused actual captures](../../../../Plans/WIP/FileOperationsUiReview_2026-09-16/implementation/2026-09-23-status-semantics/README.md).
