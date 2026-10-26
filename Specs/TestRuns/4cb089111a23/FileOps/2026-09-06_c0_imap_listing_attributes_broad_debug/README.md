# IMAP structural attributes: broad C0 Debug

Result: **18 passed, zero failed/skipped**, native duration 259.5 seconds.
The trace verifies IMAP 208/208, Copy 142/142, lookup 294/294 and Move-preflight
247/247, alongside the unchanged Delete/size/committed-error guards.
The two wide Copy scenarios retain the full 4,097-file corpus and existing deadline.

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild
-CaseFilter C0_ -TestRoot C:/RedSalamander.Perf
-RunId c0-imap-attributes-broad-debug-20260906`.

This uses the [focused Debug candidate's](../2026-09-06_c0_imap_listing_attributes_candidate_debug/README.md)
exact tested patch, source hashes, full clean build and source snapshot
`57bc7d9f6a1ddc2e80539eec07b74433272555dedc3d9cc4d842d78a580c756c`.
Receipt: `a3b75f8ef4e4590e2f4fb84014dd9839ba368322e37cdd2cb071b35485a2d319`.
No source/spec/index edits occurred during qualification. Raw metrics are selected
in original order with a digest-bound selection manifest; full raw evidence remains
under the governed run named in provenance.

The runner retains 206 disk-audit findings. No retry, quarantine or timeout change
was used. This is the C0 selection, not Fresh Full or live-provider qualification.
