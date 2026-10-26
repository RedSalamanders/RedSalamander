# IMAP structural attributes: failing Debug baseline

Starting commit: `f5858fbe0a56c99fd825958ae6434fda0e37e394`.
Only fixture/assertion and plan/performance notes changed before this run;
production behavior is unchanged. `tested.patch` is the LF-normalized Git diff
against that commit; `tested-source-hashes.json` binds the exact working bytes.
The full build receipt and final workspace manifest bind the tested snapshot.

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild
-CaseFilter C0_CurlImapListingTruth -TestRoot C:/RedSalamander.Perf
-RunId c0-imap-attributes-baseline-debug-20260906`.

Result: **110/134 assertions pass; 24 fail**. The native case fails cleanly;
the two setup cases pass. No retry or classification rerun was performed.
The test-enabled full Debug build completed in 187,751 ms with zero warnings/errors.
Receipt: `d7ed1766deae1596153c685b1985a307bdc72fc2b630a5330c34790d3f9c726d`.
Snapshot: `39cac13589c038ea41b93e2be271f3dcd77f360faa4e2934af117c11497af5d9`.

The existing 70 assertions are followed by thirteen four-check attribute-shape
scenarios and three four-check incomplete-summary scenarios. The fixture drives
the production parser/repair owner synchronously without network or credentials.
The selected metrics preserve their original ordering and digest-bound raw source;
request counts and map peaks are not whole-operation memory or live TLS proof.
The raw run remains under the exact governed root named in file provenance.
Its 202 disk-audit findings are retained, not silently repaired or discarded.

This expected-red archive is evidence for repair, not a passed C0/Full gate.
