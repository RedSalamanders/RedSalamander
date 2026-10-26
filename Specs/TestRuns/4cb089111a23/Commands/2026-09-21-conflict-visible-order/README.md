# Visible conflict ordering: harness correction

Base `843eef39` plus the bound dirty snapshot; x64 Debug; unchanged qualified DxUi
pin `78b3de389a189c7f86f611787e0489fb6d474218`. Receipt:
`adde4b7fef9085c460c8ab3762b3a60886cfaa92d20fd8dc10d58d6a627dda41`.
The full test-enabled Debug build has zero warnings/errors.

The copy-merge case reported `promptBeforeActions=false` in broad and exact runs
(203 and 156 ms). The snapshot compared off-screen comparison-content coordinates
with fixed decisions. Rendering already clips those rows to the nested viewport.
The corrected snapshot applies that same viewport intersection before checking
ordering and records content bottom, visible bottom and action top separately.
Fixed subject/feedback bounds remain checked without the comparison clip.
No production drawing, action policy or viewport behavior changed in this slice.

| Coverage | Passed / failed / skipped |
|---|---|
| Exact failing baseline, stopped on first repetition | 2 / 1 / 6 |
| Copy-merge candidate, three repetitions | 9 / 0 / 0 |
| French comparison Commands fixture | 1 / 0 / 0 |
| Fairstream predecessor family, including inline F2 | 57 / 0 / 0 |

Counts include setup/cleanup. The candidate copy-merge cases complete in
234/250/234 ms; these correctness runs establish no general performance claim.
The 57-case run executes F2 after its predecessors (1359 ms). Phase5 10/10 and
R4A19 5/5 evidence remains in the linked Queue exemption archive. The final Full
and Release/configuration/resource gates remain open; targeted success is not a
replacement for them. Disk-audit findings (260) are retained without a clean claim.

Full logs, native archives and metrics are retained under
`C:/RedSalamander.Perf/evidence/i26-ui/conflict-visible-order-*20260921` and
`C:/RedSalamander.Perf/evidence/i26-ui/inline-queue-fairstream-order-20260921`.
The runner deletes inactive run roots on its next invocation; these copies were
made first. Source paths in raw JSON remain original provenance, not live links.

The deterministic French fixture saves twelve unmodified application PNGs at the
native 144-DPI display. Narrow, differing-name, provider fallback, wheel and
keyboard views were visually inspected. The before/after suffixes mean scrolling
within the same candidate build, not a UI redesign comparison. This run does not
claim native 96-DPI or complete assistive-technology qualification.
The fallback image still repeats its unsplit destination path in the question;
that presentation refinement remains explicit in I26. Do not guess a provider
filename to remove the repetition; keep the exact role path and use bucket wording.