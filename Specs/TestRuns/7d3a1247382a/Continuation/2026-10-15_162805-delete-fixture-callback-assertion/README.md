# Permanent Delete fixture callback assertion

The first focused `FileOps_PermanentDeleteCompletionProgress` attempt opened a CRT assertion at
`Task::FileSystemShouldCancel`, source line9132, requiring an active provider callback scope.
The direct synchronous test invocation initialized the Task operation-control callbacks but omitted
the executor's callback-scope lifetime. The first operation-control check asserted before deletion.

The run was interrupted and its owned process stopped. This is **not a passed case** and no completed
aggregate was produced. The provided assertion text, exact command, build receipt and partial trace
are retained here; the complete raw artifacts remain at the sandbox evidence path in the JSON.
No fresh dump was produced; the newest existing application crash artifact was from October9.

The fixture correction must use the existing selftest active-scope pattern around each synchronous
provider call, keeping the runtime assertion and actual host-driven test paths intact. A new verified
Debug build and a clean focused rerun are required before claiming the Delete fix validated.

## Follow-up on the next verified Debug receipt

The fixture now uses the existing callback-scope RAII pattern around each of its six synchronous
provider calls. On receipt `6b42b3d9d25bb8f351668a577f0ef2fcd4ee2279c5aff47505ef71e196ff27e7`,
the subsequent [focused Delete attempt](../../FileOps/2026-10-15_163647-FileOps_PermanentDeleteCompletionProgress/README.md)
completed without this assertion. Its actual partial/cancel/refusal accounting and host-driven
normal-parent/read-only-child checks passed before a later sequential bridge fixture guard failed.
That guard incorrectly required a parallel-only Ensure counter on the sequential route.

This resolves the observed callback-lifetime assertion in the fixture. The subsequent Delete case
still has a failed result (2 records passed, 1 failed); the complete case is not yet qualified.
The raw assertion evidence, interrupted status, old receipt and trace here remain unchanged.
Related completed runs retain 3 legacy sandbox-audit findings; no cleanup or clean-audit claim is made.
There is no repository-wide green verdict, and Release remains unqualified.

## Final focused Delete pass

The [final focused Delete case](../../FileOps/2026-10-15_164631-FileOps_PermanentDeleteCompletionProgress/README.md)
passed cleanly on test-enabled x64 Debug receipt
`047f98907061f915d00f1f490b967e89fab62bf15c3f8aa6169c531d0240b5af`: Setup, the selected case,
and Cleanup all passed, with 0 failed / 0 skipped. It exercised partial/canceled/refused deletion,
normal-parent/read-only-child consent, separate selected read-only file/empty-folder consents,
and both parallel and verified sequential bridge folder accounting without this assertion.

That later result supersedes the fixture qualification failures; it does not rewrite this
interrupted run or the intermediate failed attempt. Their raw evidence and receipts are retained.
The final focused aggregate still reports 3 legacy sandbox findings. No Full-repository or Release
qualification, clean-sandbox result, or before/after throughput improvement is inferred.
