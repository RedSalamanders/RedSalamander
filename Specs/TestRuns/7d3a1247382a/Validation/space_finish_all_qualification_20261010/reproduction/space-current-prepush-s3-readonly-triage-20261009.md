# Read-only triage: Debug PrePush PluginContract S3 failures (2026-10-09)

## Scope and current evidence

This corrected source review is against candidate checkout `C:\Users\PVB\.codex\worktrees\viewer-space-explorer-analyst\RedSalamander`, HEAD `993c3ff14ce60eb7331a0f36bd54475e227859e9`, and recorded run `20261009T201507Z-45224-aa600fc03f9e4418b3b3e7152722bb0f`. The outer log is `D:\RedSalamander.Perf\evidence\viewer-space-finish-all-20261009\prepush-debug-queued-path.log`; the run-state snapshot remains `status: running`, with `standalone.plugin-contract: failed` (the log also reports `PluginContractTests FAILED.`). No rerun or process interaction was done.

The recorded native output says:

- `[S3] debug selftest failed: the recursive Delete must list the prefix on the fixture`
- `[S3] debug selftest failed: a request whose body is still arriving must return after Cancel`
- `[S3] streaming request returned 15000 ms after Cancel`
- `[S3] debug selftest failed: Cancel must return a streaming request within a few body callbacks`

It separately reports a `FileSystem.dll` debug-selftest failure creating a relative file symlink and explicitly says `SeCreateSymbolicLinkPrivilege` is not held. Treat that as a distinct fixture/environment failure, not evidence for the S3 result.

## What the test establishes

`Plugins/FileSystemS3/FileSystemS3.Tests.FakeS3.cpp:1451-1535` is `RunS3StalledRequestCancelSelfTests`' first “bytes flowing” phase. It enables a 96-byte/40-ms drip listing, starts recursive delete for `/r0f/folder`, and waits up to 10 seconds for any fixture `GET` (`:1504-1515`). It then requests cancel, waits 15 seconds for the delete worker, and captures elapsed time before stopping the endpoint (`:1516-1523`). The 15,000 ms output is the expired worker wait; it is not a measured sub-three-second cancellation.

The fixture records a request after it parses the complete HTTP request (`:450-464`, `:822-840`); `RequestCount("GET")` only distinguishes method, not probe versus recursive-list phase. The failed listing assertion means no parsed GET was recorded in that wait. The body-flow assertion text describes the intended condition, but the fixture had not observed a GET to establish that condition. The two assertions plus duration all come from this one phase; after it the selftest returns if the worker missed the wait (`:1533-1535`), so this output does not show the later silent-server or ranged-reader phases running.

A missing GET establishes no more than absence of a parsed GET in the fake endpoint during the polling window. It does not tell whether the call stalled before connection, during request setup, or in another layer. If the worker did not return, the test stops the endpoint after the wait; that later action is not what caused the 15,000 ms value because the duration was captured first.

## Source path and transport limits

`ProbeS3Path` calls `TryGetPrefixExists` (`Plugins/FileSystemS3/FileSystemS3.Directory.cpp:1542-1575`). That function creates `Aws::S3Crt::Model::ListObjectsV2Request`, sets bucket/prefix/max-keys=1, arms operation control, and calls `client->ListObjectsV2` (`:1505-1539`). Recursive listing creates the same request model, arms control, and calls the same client method (`:1589-1634`). `DeleteItems` establishes `S3OperationOptionsScope` (`:6158-6172`); `ArmS3RequestControl` attaches the continue handler (`Plugins/FileSystemS3/FileSystemS3.Internal.h:265-276`). `MakeS3Client` constructs an `Aws::S3Crt::S3CrtClient` (`Plugins/FileSystemS3/FileSystemS3.Shared.cpp:800-826`).

This source review establishes that the plugin constructs and calls the AWS SDK S3 CRT client API. It does not inspect the SDK implementation linked into this build. The project comment in `FileSystemS3.Internal.h` says the handler runs on CRT event-loop threads and refers to a meta request, but that comment alone does not prove that this generic `ListObjectsV2` call uses CRT meta-request callbacks. I found no project-local binding showing a Windows WinHTTP transport either. Therefore this report does not attribute the stall to either WinHTTP or CRT callback behavior; the generic operation's internal transport and callback path remain unverified from the available candidate source.

The expected request-control calls are present in the production source, but their presence is not proof that the operation reached the endpoint or that the relevant SDK callback ran. The source evidence does not establish whether the failure is request admission, endpoint transport, SDK operation, control callback, or scheduling/timing.

## Diff scope

Candidate worktree changes are ViewerSpace-related. `git diff --name-only HEAD -- Plugins/FileSystemS3 Tests/PluginContractTests` returned no paths. Comparing those same S3/PluginContract paths against D-checkout HEAD `1277ffdec9cf78b5452d29f695aae7806fffe207` also returned no paths. The current PluginContract S3 failure is outside the candidate's ViewerSpace diff; that fact does not establish whether it is environmental or pre-existing.

## Named follow-up

After the canonical PrePush run reaches its final state, handle this separately as **FileSystemS3 R0f streaming-cancel fixture triage**. Only after the run owner authorizes a focused, receipt-verified PluginContract debug-selftest execution, capture the fake endpoint's parsed request method/path and distinguish the prefix-probe request from recursive listing. Preserve the existing drip rate, initial wait, cancel timing, 15-second worker wait, 3-second assertion, provider watchdog, clocks, test inventory, and coverage. First establish which stage reached the endpoint and whether the callback/control boundary was entered; only then propose a source change. Do not relax the assertions or infer successful cancellation from an unobserved GET.

This triage changed no source, test, normative documentation, index, build, test evidence, or process state. Its only output is this ignored proposal in the candidate worktree.
