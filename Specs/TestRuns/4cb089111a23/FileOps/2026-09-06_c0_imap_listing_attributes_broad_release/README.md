# IMAP structural attributes: broad C0 Release

Result: **18 passed, zero failed/skipped**, native duration 101,313 ms.
The trace verifies IMAP 208/208, Copy 142/142, lookup 294/294 and Move-preflight
247/247, alongside the unchanged Delete/size/committed-error guards.
Both wide Copy scenarios complete their full 4,097-file corpus.

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild
-CaseFilter C0_ -TestRoot C:/RedSalamander.Perf
-RunId c0-imap-attributes-broad-release-20260906`.

This uses the [focused candidate's](../2026-09-06_c0_imap_listing_attributes_candidate_release/README.md)
zero-warning/error full test-enabled build. Exact tested patch and working-byte
hashes are retained in the linked Debug candidate archive. Both configurations
share snapshot `57bc7d9f6a1ddc2e80539eec07b74433272555dedc3d9cc4d842d78a580c756c`.
Release receipt: `c38aa434016840c80c9f10fdc99067005f64010deb8e33b3883156bbf243e373`.
No source/spec/index changes occurred during qualification. Raw metrics are
selected in original order with their digest-bound selection manifest.

No retry, quarantine or increased deadline was used. The runner retains 209
disk-audit findings; existing runtime data was not cleaned. This C0 selection
does not establish live-provider, whole-operation memory or final Fresh Full
qualification. Post-test specification/evidence annotations do not claim another
native runtime result.
