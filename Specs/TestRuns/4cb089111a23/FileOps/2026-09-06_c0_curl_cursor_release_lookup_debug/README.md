# Bounded cursor release: Debug lookup

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild
-TestRoot C:\RedSalamander.Perf -CaseFilter C0_CurlEntryLookupTruth
-RunId c0-curl-cursor-release-lookup-debug-20260906`.

Result: 3 passed, zero failed/skipped, 26.2 seconds; native lookup 284/284.
Full test-enabled Debug build: 186,837 ms, zero warnings/errors. It shares source
snapshot `87965dd7ee6ceb0b341a8d6016ead438d846e6e6c216a05abeba6d45b8ad0829`
with Release. The sibling lookup Release archive preserves the exact seven-file
tested patch. All ten lifetime metric observations pass, including parent/child
single-session reuse and every empty/wide/abandon/cancel/final-failure case.

Copied results, traces and entry fingerprints are exact; provenance binds the
ordered metric selection to the raw stream. Build-log normalization and raw hash
are explicit. Broad Debug separately fails its serial wide-copy deadline; this
focused success cannot replace that run or qualify the whole C0 gate.
