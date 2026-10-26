# ASan qualification failure baseline

These are failed baseline receipts from source
`d195fe31b9b1948865fbfa06315d626a25f1a961`, x64 ASan Debug, build receipt
`ce927559c79beb5e8a65d0be3709ce956c28042d4961f6fbb4a3d07151978872`.
The Full run `20261025T140118Z-34484-fb6cb3f2682c46128609df962f2ea7e2`
completed all 18 groups: 15 passed and three failed. Compare recorded 252/3/21,
Commands 952/4/6 and FileOps 296/1/25 (passed/failed/skipped). All native verdicts
and failure reasons in those groups are retained byte for byte. The prior Debug
and Release Full passes do not clear this ASan failure.

The three search-service cases failed unchanged; controls changing only the
synthetic-root traversal seed passed. That is a diagnostic comparison, not a
production fix or startup performance claim. Code inspection found an optional
status overlay taking the same blocking volume mutex as cold index seeding;
a separately gated native regression must establish that dependency before repair.

The four UI cases repeated three times with unchanged binaries and ordinary
sanitizer settings: eight passed and four failed. Theme retention exceeded its
original private-byte guard on the first burst and passed its next two bursts.
Breadcrumb actions lacked Dismiss first; later repetitions reused an already
consumed synthetic clipboard sequence and failed admission. The native focus and
close cases passed all three isolated replays; their broad failures remain open.

The equivalent-selection replay recorded five passing and five failing target
repetitions, plus 20 passing setup/cleanup cases, with no skips. Its sampled
concurrency failure requires an actual transfer witness, not just configured
limits. Original byte equality, counting, source retention and concurrency bounds
remain protected. Ordered selected trace rows are bound to the complete original
streams by digest, byte count and line number.

`manifest.json` binds the complete curated payloads. All executable probes were
serial and source/binary hashes were checked. Existing failed receipts remain
failed. Diagnostic instrumentation, production fixes and final-source Full
qualification are pending; passing controls and replays are not repair evidence.
