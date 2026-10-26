# Curl Move-mode repair — Release

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -TestRoot C:\RedSalamander.Perf -CaseFilter C0_CurlMovePreflightTruth -RunId c0-curl-movemode-candidate-release-20260906`.

Result: **247/247 native checks**, 3/0/0 suite, no skips. Native case: 16,734 ms.
Both mixed-bulk cases pass with independent native-success/no-commit-refusal
receipts, partial aggregate status, no relay and preserved refused source.
Concurrency 1/4 takes 4,043 / 4,472 us, two mutation commands and 21 total commands
each. The 48 mode controls retain zero invalid-option I/O, zero cross-endpoint
native-only mutations, and successful same-endpoint file/directory renames.

Full test-enabled Release receipt:
`270090970d28ed12d238f8cab7f5a6b2da92fbed4b20933d36a10cbfdfc38472`.
Snapshot `049d8724b6ba200a243b870ce24bd6edc7c1b1c32925b225e823518c415afa5f`
matches Debug exactly. The eight-file tested patch/raw source hashes are in
`../2026-09-06_c0_curl_movemode_candidate_debug/`.

Build: 249,285 ms, **one warning, zero errors**. MSB3061 in the Google Drive
project's shared-output cleanup could not delete `Plugins/z.dll`; later staging
copied the dependency and the build attested successfully. Preserve this as a
further C6 shared-DLL Clean-ownership witness, not a zero-warning qualification.
The complete log and raw/normalized digests are retained. No manual deletion,
warning suppression or unrelated process termination was used.

The red Debug baseline is retained separately; cross-configuration timing is not
an equal-work throughput comparison. This is focused loopback evidence, not a live
endpoint or Fresh Full result. C0 memory/route qualification and C1–C6 remain open.
Exact-copy and ordered-metric provenance accompany the archive.
