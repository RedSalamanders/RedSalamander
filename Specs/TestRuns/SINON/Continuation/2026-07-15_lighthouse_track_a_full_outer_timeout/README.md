# Lighthouse Track A — interrupted Full-suite evidence

- Date: 2026-07-15
- Branch: `codex/lighthouse-loader-hardening`
- Command: `./Tools/Run-AllTests.ps1 -Suite Full -SkipBuild`
- Outer executor limit: 1,200 seconds
- Classification: harness/executor timeout, not a reproduced product hang or crash

The outer command limit terminated the run while File Operations was still making forward progress. The
last trace entry reached `Phase11_BridgeMultiFolderParallelCopyInFlightLines` after completing the preceding
Fairstream family. The last completed family artifact reports 45 passed, 0 failed, and 0 skipped cases. No
RedSalamander process survived termination and no crash artifact was produced.

This interrupted run is not closeout evidence. Before retrying, recover the abandoned artifact-operation
state with the repository-required serialized rebuild, then rerun Full with a bound long enough for all 18
File Operations families and the remaining Full-only executables.

Preserved files:

- `trace.txt`: application-level trace through interruption.
- `fileops-trace.txt`: File Operations family/case progress through interruption.
- `fileops-last-completed-family-results.json`: last atomically completed File Operations family result.
