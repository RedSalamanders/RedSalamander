# Terminal facts: matched Release evidence

Baseline projection `20260923T181530Z-28896-f1f5e65546bb4303ace94b01bb5df0d5` fails its count projection (2 passed, 1 failed) by design; its separate gallery `20260923T181633Z-93268-40bb71ba5e4c4785813e497d96ea2367` passes with 508 captures. Candidate projection `20260923T184219Z-66304-d5f4be7e474243b0b17eb5bab6ba8310` passes 3/3, and gallery `20260923T184329Z-63196-8765865d660f45bcbca73dce387148bc` passes with 516 captures. Both test runs per build have clean disk audits and receipt-bound logs. The baseline build had two C5258 test-witness capture warnings and no errors; candidate diagnostics must be zero.

The comparator checks 508 common PNG pairs by SHA-256, dimensions, and RGBA difference bounds (34 differ). Eight scene-144 PNGs exist only in the candidate. The retained 24 shared redraw fixtures have five trace-bound rounds each; their paired median redraw-time changes range -3.02% to +17.28%, and private-byte changes range -14.83 to +18.50 MiB. These are in-process observations, not presented FPS, heap attribution, GPU memory, or final performance acceptance.

The intermediate v1 candidate passed projection and gallery, but its 8/9 localization-contract result and missing settings-cache invalidation prevent it from being treated as final. Its build and run receipts, manifest, localization result, and initial production patch are retained concisely; the v2 settings follow-up patch and final receipts are retained separately. V2 adds twelve same-HWND transition stages per display, including a settings notification that rebuilds once and reuses the retained paragraph on the repeated paint.

The original baseline manifest has 18 rows for 17 distinct sources because SelfTest.cpp is repeated with the same hash; the corrected unique view and original metadata are both retained. The final source delta is four changed existing files and eight added localization/test inputs; fixture bytes are therefore not identical as a whole. Both primary and isolated build HEADs and dependency digests are explicitly compared by the curator.

The focused gallery has 12 shared actual captures (Move source kept, source unknown, and the completed group at both native displays) plus all eight new scene-144 captures. This packet covers Release only; Debug, PR, Full, UIA, native-user interaction, and final consumer acceptance remain open.

[Focused actual captures](../../../../Plans/WIP/FileOperationsUiReview_2026-09-16/implementation/2026-09-23-terminal-facts/README.md).
