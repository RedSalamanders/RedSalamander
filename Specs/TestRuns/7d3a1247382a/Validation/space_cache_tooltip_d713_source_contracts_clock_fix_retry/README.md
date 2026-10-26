# ViewerSpace focused source-contract retry

The first invocation transcript (`source-contracts-invocation-error.log`) records a PowerShell parameter transformation error before Pester started; it executed no tests and is retained as a failed invocation only.

The subsequent Pester 3.4.0 run (`source-contracts-retry.log`) passed 4 tests with 0 failures and 0 skips across the plugin lifetime consolidation and production `GetCursorPos()` source-contract checks. This is a focused source-inspection result, not a runtime or repository-wide pass.

Source snapshot: `846115f86b923381c3d9311025957793349fecaddbccab6fa1748db82982e131`. The two exact test files used by the successful invocation are included as `.source.ps1` payloads; `manifest.txt` records source paths, byte lengths, and SHA-256 values for both transcripts and both test files. The first invocation is not counted as a test run.
