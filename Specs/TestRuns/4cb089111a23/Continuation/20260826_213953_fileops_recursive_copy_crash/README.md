# File Operations recursive-copy crash continuation

## Purpose

This compact continuation record preserves the 2026-08-26 Fresh Full File
Operations crash before any focused rerun. The full dump and failed run remain
in their original local locations; they are intentionally not copied into the
repository.

## Failed run

- Command: `$env:REDSALAMANDER_TEST_ROOT='C:\RSPerf'; .\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -Configuration Debug -Platform x64 -BuildNumber 1`
- Run ID: `20260826T184510Z-115104-3722188961da4695b056fc727af14c60`
- Evidence root: `.build\ValidationEvidence\runs\20260826T184510Z-115104-3722188961da4695b056fc727af14c60`
- Result: 1,893 total; 1,860 passed; 1 failed; 32 skipped.
- Passing coverage before the crash included Commands 866/866 executed, source
  contracts 698/698, Compare 228/228 executed, and all Terminal contracts.
- Failing suite: File Operations returned exit code 1 after 2m53.8s. Its last
  durable family result was the preceding 53-case family, so the aggregate
  runner correctly reported the later coverage as missing rather than as an
  assertion failure.
- File Operations trace:
  `C:\RSPerf\runs\20260826T184510Z-115104-3722188961da4695b056fc727af14c60\artifacts\selftest\last_run\fileops\trace.txt`
- Last entered case: `Phase7_CopyItemsSingleFolderRecursiveParallelism`, after
  the intentionally throttled `Phase11_BridgeMultiFolderParallelCopyInFlightLines`.

## Crash identity

- WER exception: `0xC0000005` access-violation write.
- WER report ID: `e6baff53-1191-4353-ab9f-c9e685978573`.
- WER report:
  `C:\ProgramData\Microsoft\Windows\WER\ReportArchive\AppCrash_RedSalamander.ex_33d887c7d7fb8c76c59162fdf12b97d2381bd_5d135df1_e6baff53-1191-4353-ab9f-c9e685978573\Report.wer`
- Dump: `C:\Users\eric\AppData\Local\CrashDumps\RedSalamander.exe.106888.dmp`
- Dump size: 129,420,267 bytes.
- Dump SHA-256:
  `7CDDE52F7677208C1D8E2F460814FDF5A437F3B364D8FDDE86E92C9E6BE34D57`.
- WER fault offset: `RedSalamander.exe+0x1c8115e`.
- No matching current-run sidecar was written under
  `%LOCALAPPDATA%\RedSalamander\Crashes`; the WER dump is authoritative.

## Symbolicated stack

The crashing thread was symbolicated against the exact Debug binaries with
LLDB. Relevant frames, omitting standard-library formatting internals, are:

```text
#41 MakeCopyTempSiblingPath                         FileSystem.FileOps.cpp:3196
#42 PublishLocalCopyDirectory                      FileSystem.FileOps.cpp:4593
#43 EnsureLocalCopyDestinationDirectory            FileSystem.FileOps.cpp:4715
#44 CopyDirectoryChildrenParallel::processDirectory FileSystem.FileOps.cpp:5454
#48 CopyDirectoryChildrenParallel::runDirectory     FileSystem.FileOps.cpp:5640
#52 CopyDirectoryChildrenParallel::processDirectory FileSystem.FileOps.cpp:5574
#56 CopyDirectoryChildrenParallel::runDirectory     FileSystem.FileOps.cpp:5640
```

The immediate `std::format` input was an invalid `std::wstring_view` derived
from `destination.display`. Debug fill patterns and an unreadable source buffer
were visible in the dump. The path prefix reached only `deep_06_segment` of the
case's 96-directory input, which rules out the explicit traversal-depth guard
as the direct failure and points to traversal-owned path lifetime or stack
corruption.

## Continuation contract

1. Preserve the failed evidence root, `C:\RSPerf` run, WER report, and dump
   until a focused fix and final Fresh Full have passed.
2. Inspect recursive directory work-item ownership and recursion-frame state;
   do not raise the timeout or reduce the 96-directory stress input as a fix.
3. Rebuild the affected Debug x64 product and replay
   `Phase7_CopyItemsSingleFolderRecursiveParallelism` first.
4. Then run the complete File Operations suite before another Fresh Full.
5. Record the fixed commit, focused result, broad result, and final Fresh Full
   receipt in the active Ghostty upgrade plan before P11 closeout.

## Repair validation

- Root cause: `CopyDirectoryChildrenParallel` retained path objects, find data,
  directory authority, and the selected work item in every recursive producer
  frame. The 96-level case exhausted the effective Debug worker stack before the
  repository's explicit traversal ceiling and corrupted the deepest staged-path
  formatting call.
- Repair: move traversal-owned directory state and the retry-selected work item
  into RAII-owned heap frames while retaining the existing depth, metadata,
  conflict, cancellation, and queue contracts.
- Focused Debug x64 build: zero warnings and zero errors; attestation
  `3f21b9b502f980b18be58434135cd28b89ad5f42be80134601b641a42a06a7f8`.
- Exact replay run ID:
  `20260826T200200Z-107868-35596a52f33a45b397598d0ac2b99545`; 3 passed,
  0 failed, 0 skipped; wall time 8.2 seconds; disk audit zero.
- Complete File Operations run ID:
  `20260826T200236Z-90056-04e05f6cca8c4a59b215954726cf90cf`; 122 passed,
  0 failed, 20 intentional remote/opt-in skips; wall time 15m00.4s; disk audit
  zero.
- Final Fresh Full remains required before P11 closes.
