# Pending-conflict inspection controls: final

Result: **15 passed, 3 failed, 0 skipped**, x64 Debug, six Commands cases repeated three times. Source/build receipt identities match.

Retained pre-change baseline: [navigation final](../2026-09-20_162858-i26-inspection-navigation-final/README.md), 18/18. Common renderer median/p95 66496/66865 us; differing 66547/66755 us. Same machine, 480/640/760 DIP French fixture, five batches of eight redraws at each width, three repeats and 45 ordered samples per presentation. Measurements are focused Debug redraw cost, not final Release FPS/resource acceptance.

The source and destination now have full-width hosted inspection buttons. Both use the published immutable prompt, owner navigation and exact selection; complete localized outcomes are measured, painted, accessible and announced for the current prompt. One request/result is retained; replacement prompts do not inherit old outcomes. Weak popup lifetime invalidation rejects late callbacks after destruction. Guarded foreground/hide/return remains a separate unfinished component.

Candidate is 15/18 with 12 compiler warning/context lines: the test incorrectly included the pinned UIA Button::Invoke(host,true) focus establishment in its asynchronous no-steal assertion. Final is 15/18 with zero warnings/errors: its unresolved-binding byte check incorrectly reused a post-completion ifstream reader. Verified corrects that existing fixture helper to a bounded read sharing READ/WRITE/DELETE, retains the byte assertion, and separates UIA invocation focus from subsequent navigation focus. It also covers the non-activating harness route and stale painted prompt refusal. These failures are retained, not waived.

French renderer checks use real button typography plus padding at all three widths. UIA invokes both roles in opposite panes, verifies exact selected identity and unchanged focus after invocation, missing-item outcome, replacement-result removal and unchanged destination bytes. Verified captures selected/missing result surfaces in French through the harness. Pointer, directed foreground/return, remote refusal/rollback, native mixed-DPI/AT and broad final acceptance remain open.

Full raw captures/metrics/results: `S:/RedSalamander.Perf/evidence/i26-inspection-controls-final-20260920`. Compacted metrics preserve raw hash and original row numbers. Original runner `C:/RedSalamander.Perf/evidence/i26-ui/inspection-controls-final.log`, SHA256 `765f332fd39d90fea3ebc0383dddab6c687aaf6bab09a77c3ac358caea424214`; normalization only strips trailing whitespace and normalizes line endings.

## Qualification limit

The accepted correctness matrix passes 18/18, but its differing-name redraw p95 is 88,361 us versus 66,755 us baseline. One justified unchanged-source confirmation passes 3/3 and records common/differing p95 79,436/81,740 us. Medians remain about 66.5 ms. Tail-latency non-regression is **not established**; do not dismiss these samples as noise or silently rebaseline. Profile presentation wait versus CPU/layout and retain a matched baseline before final resource acceptance. Earlier verified frames also expose a result/scope overlap; accepted frames and text-height assertions verify its repair. Overall R2/R3 and final qualification remain open.
