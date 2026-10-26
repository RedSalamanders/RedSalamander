# R4-A02 relation-receipt stack-overflow continuation

- Repro: `.\Tools\Run-AllTests.ps1 -Suite FileOps -ValidationMode Fresh -CaseFilter 'R4A02_'`
- Run: `20260904T223400Z-110428-3e9586bc35c44822a42b245c0a89e5ca`
- Failure: `0xC00000FD` in `TickProviderMatrix` before the first case result.
- Cause: the first R4-A02-2 draft embedded 64 task IDs in every one of the 64 deferred-consent receipt slots. Provider-matrix tests instantiate many `Task` objects in switch-case locals, so the enlarged `Task` type pushed the generated tick function over the UI-thread stack limit.
- Correction: keep each overlap relation as one heap-owned fixed-capacity set referenced by the receipt. Non-overlap receipts remain pointer-sized and the relation remains bounded to 64 IDs.
- Preserved checked-in evidence: the symbolized sidecar, governed run trace, and FileOps trace are beside this file. The raw minidump remains local and ignored under the repository's bounded-archive policy.
