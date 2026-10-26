# Current V47 qualification addendum

This supplements, rather than rewrites, the historical pre-refresh V46 source review. A78 now contains the exact approved integrated fixture `90f7301cc91e96d50985bf61451d6dcb1799254a5de6212928978137757ce740` and its Release solution rebuild succeeded with receipt `905bccdbb1698cef960bcc1064bb0ec1f9da73568d8007711624a78d0081ea1b`. The refresh archived the original inputs and changed only the fixture among the 46 baseline manifest rows. Evidence is in `a78-scan-v46-refresh`.

V47 changes only test diagnostics and the checkpoint from V46: native Tab counts, focus-change counts and the last two retained focus identities, plus forward-Tab pump diagnostics. The focus callback copies bounded owned strings when tests are enabled, so any test-enabled timing comparison still includes that instrumentation. No product focus repair or isolated DxUi attribution is claimed. The existing V46 source-delta review otherwise applies to the measured product paths.

Root reviewed the separate V46-named driver and comparator drafts for use with the explicitly supplied current V47 source manifest and post-Full Release receipt. Both enforce the approved fixture SHA before measurement and in archived replay. The comparator now validates all 275 integrated metric rows, four scan kinds, exact shared digest/geometry, five clean and five changing redraw samples per kind, scan surface/resource phases and per-process scan CPU deltas. The scan digest is parsed and checked inside Read-IntegratedRun before returning its data, correcting the earlier draft scope defect. No fabricated native fixture was used; the agent's offline schema check used retained real V45 metrics.

Reviewed external scripts:
- `run-integrated-a78-b49-pair-v46-draft.ps1`: SHA-256 `64afab5d28d0b5c05bb383639c66e1d4a10780b9336103020445c2b673ff8473`.
- `compare-integrated-a78-b49-v46-draft.ps1`: SHA-256 `d0f308f7148c00f259d7a050a97b9cf3377111419cacd6bebe22c6ebe83bfcd03`.

Execution is pending the final Release Full gate. This review is not performance acceptance. Use V:/RedSalamander.Perf on both sides, freshly reviewed resolved dependency identities and the post-Full B receipt. Retain all eight original runs, cross-pairs, failures and attribution limits; CPU ticks and five-sample tail maxima are descriptive.
