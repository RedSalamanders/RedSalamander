# C0 Release — preserved malformed-list fixture failure

Command: test-enabled `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Release -CaseFilter C0_ -FailFast -RunId c0-curl-size-iterative-final-release-retry1-20260905`.

An earlier invocation stopped before build/test because the sandbox denied the
runner's residual-process ownership check. This approved retry completed the
governed build; it did not bypass process checks or kill independent processes.

Build: 213,036 ms, zero warnings/errors. Receipt
`bfe0471b6390efebf37870120da6ed4b648c6a582d30886f20c165a262b2259c`;
source `0620c05e5bd9dfc926d9014e8b7e2e7001b40caa0576c78e04f24dac7da1abbb`.
Host: 12 passed, one failed, one skipped; 16,672 ms. The failed bounds export has
126/127 assertions: malformed LIST, bulk Delete, concurrency four. Size qualification
was skipped by fail-fast, not passed. Aggregate classification remains unclassified.

The provider traversal row reports the expected `ERROR_INVALID_DATA`, but the old
fixture metric did not distinguish contents, receipt and server-health predicates.
The subsequent diagnostic archive supplies that witness; do not rewrite this run
as green. Its five raw files are unchanged, SHA-256-verified copies from the named
run under `C:\RedSalamander.Perf`, including the complete 1,814,720-byte perf stream.
The runner retained 54 disk-audit issues; no old evidence or user data was removed.
