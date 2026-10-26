# Iterative native Copy candidate — wide cases still red

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_CurlCopyTraversalTruth -FailFast -RunId c0-curl-copy-traversal-candidate-debug-20260905`.

Based on regression checkpoint `07dfabca`. Build: 02:57, zero warnings/errors.
Receipt: `79a2d0c66af49402545410f8c3052944fb17e8accf4bba8ae0972d1053c9d51c`.
Workspace: `2d8e0ab053fb908e6580526cc73f650658015426fc14bf0c969ee69e44db2196`.
Tested `CopyMove.cpp` raw SHA-256:
`b489a7a001b52ee0e599a2cb382dcd24455155057ae316b5d7cd7202831897a8`.

The gate completes: **135/139 assertions; 44/46 scenarios pass**. Wrapper:
2 passed / 1 failed / 0 skipped, suite 148.6 s. All expected depth-128 Copy/Move
routes succeed, all depth-129 Copy routes refuse cleanly, and all strict-listing,
literal-name, early/late refusal and cancellation scenarios pass. The baseline
terminated before completing its first depth-128 scenario; this is behavioral
recovery evidence, not a proven diagnosis of its missing crash stack.

The same bounded cursor now feeds one iterative ancestry loop with serial or
scheduler-owned batches. It no longer recursively collects directories or runs a
recursive parallel producer. Batches drain before child discovery. Move membership
and original source-size commitments, staged publication, and partial-tree
preservation remain intact. Initial invalid discovery creates no children.

LIST cancellation: 54,745 / 55,087 / 54,790 / 54,730 us for first-list waits;
62,044 / 56,489 / 62,673 / 58,773 us for late waits. Earlier published bytes remain
present in every late case. Baseline cancellation was about one second, and both
parallel late cases reached later discovery without that earlier publication.
Maximum admitted Copy reservations: 129 frames, 128 queued files, 226,472 path
bytes and 6,670,544 metadata bytes. These are frame/batch reservations, not total
process memory, buffered publication lookups, or the separate Move commitment map.

Both wide cases remain failed:

- Serial: 120,005,799 us, 3,005/4,097 publications, ERROR_TIMEOUT during promotion;
  source endpoint also records ERROR_SEM_TIMEOUT from an idle control connection.
- Parallel: 248,970 us, 78/4,097 publications, E_FAIL in source proof/download.

The fixture's `truthful` flag compares Copy completion with the **expected fixture
status**; false for these expected-success cases does not prove a callback lied
about the actual failed HRESULT. Preserve exact source/destination checks and
compare actual status consistency separately when improving diagnostics.

Full raw perf retained: 197,826 bytes / 532 rows; SHA-256
`d85878c9be273a5ccd9501410789f874fe4668b89bc807ae29c30d45a09a2ef2`.
The archive includes all per-case reasons and traces. Source/localization Pester:
197/197, 14.27 s; owning spec inventory: zero blockers. Disk audit: 67 issues,
none deleted. The complete gate, Release qualification and C0 closeout remain open.
