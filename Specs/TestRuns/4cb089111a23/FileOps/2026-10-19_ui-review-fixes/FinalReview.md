# Final scoped review

No remaining actionable defect was found in the reviewed File Operations UI changes after the
final build and focused runtime checks. This conclusion is bounded by the qualification limits
in [README.md](README.md); it is not a Full or cross-platform verdict.

| Finding | Final behavior and evidence |
|---|---|
| Obsolete combined throughput-summary assertions | All five obsolete Copy/Move predicates now require separate speed and ETA regions and absence of the legacy summary. Actual localized Name/HelpText are checked in both plot states. The existing graph counts, placement, persistence, default omission, reload, restart and idempotence checks remain. All three Commands cases and three restart-child journeys pass. French discovery checks now require exact current copy wording and actual painted/accessible rows. |
| Live counts compared two copies of one formatted string | Expected denominators come from the independently known fixture tree. Changing numerators come from numeric inputs recorded at that actual paint generation. Painted text, descriptor text and retained accessible Name/HelpText must all match. Wrong denominators and missing accessible text are rejected. All five real copy routes pass; the complete live journey has 69 passing assertions. Completion counters retain their actual 64-bit width. |
| Percentage checks permitted missing required content | The matrix explicitly requires positive geometry and exact localized accessible percentage for determinate scenes, and absence of both region and provider for the other scenes. Forty-eight determinate checks and all 124 card-fact scenarios pass, including the new unknown-concurrency scene. |
| Delete described its workers as parallel transfers | The note requires unfinished Copy/Move, automatic concurrency and a positive effective budget. Delete, manual mode, unknown budget, collapsed and finished cards expose no such note. Both the visible capture and retained accessible-region assertions verify absence. |

Foreground testing also found a real regression in the focused source tooltip: finishing another
stream changed HelpText without changing the first path or its bounds, so the old path-only cache
left obsolete text visible. The cache now includes the tooltip content. The same assertion that
failed at 170 versus 71 code units passes with exact current text, while source focus remains stable.
Unchanged content still uses the cache; no new tooltip/layout operation is added to unchanged paints.

Two test-fixture assumptions were corrected without weakening functional assertions. Escape's
tooltip-dismissal check remains; the next reveal test first leaves the row to create an actual new
focus transition. Reload/restart fixtures reject active operations and use the canonical test helper
to retire idle popup presentation before owning its lifecycle. Native focus acknowledgement,
Unicode clipboard, path replacement/removal, long-path detail and action isolation checks still pass.

The final source review checked the UI-thread ownership of the new paint facts, count certainty
branches, percentage presence/absence, note eligibility, tooltip cache invalidation and retained
accessibility. The normative UI contract and user guide match the resulting behavior. Eight source,
specification, guide and dependency byte checks report zero mismatches. The unrelated dependency
pin remains unstaged. No product or test source changed after successful final qualification.
