# Initial bounded cursor-release witness

Diagnostic predecessor, not final qualification. The first Release build after
`96c5c5f8` succeeded in 249,265 ms with two warnings: C4457 for the new test's
`failed` cursor shadowing its output parameter, and the existing MSB3061 shared
`yyjson.dll` Clean ownership issue. The subsequent source correction renames
that cursor; this archive preserves the first tested seven-file patch and hashes.

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild
-TestRoot C:\RedSalamander.Perf -CaseFilter C0_CurlEntryLookupTruth
-RunId c0-curl-cursor-release-initial-lookup-release-20260906`.

Result: 3 passed, 0 failed, 0 skipped; native lookup 284/284, 24.4 seconds.
All ten `FileOps.Curl.CursorLifetime` observations pass, including one USER/two
LIST requests while parent and child objects coexist, intact unread parent rows,
4,097 streamed rows, empty/canceled/abandoned/failed cursors and healthy reborrow.
The late-426 fixture sends all listing bytes before failing its final reply;
neither the small cursor nor exact lookup exposes a successful partial result.
Source/resource policy: 197/197, zero skips, 21.2174620 seconds.

`build-provenance.json` binds the receipt and source snapshot. Copied results,
traces, policy XML and entry fingerprints are byte-exact with
`file-provenance.json`. Ordered selected metrics bind the raw digest, size,
row count and original line positions in `metric-selection.json`; the raw
stream remains in the owned run. The build log removes only trailing horizontal
whitespace, documented with both digests in `build-log-provenance.json`.

This focused result does not qualify broad Copy/Move, real server quotas, opaque
transport memory, the outstanding Debug wide-copy timeout, or C1-C6. No live
profile mutation, credentials, TLS policy or machine socket setting was changed.
