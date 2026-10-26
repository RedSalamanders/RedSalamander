# C0 Curl partial-tree failure — pre-fix Debug

- Source: `58f22d0e` plus the recorded C0 working changes and new partial-tree test;
  production completion still inferred no-commit from the final HRESULT.
- Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_CurlPartialTreeFailure -FailFast -RunId c0-curl-partial-tree-red-debug-20260905`.
- Accepted build receipt: `10e4f690aff3cc0c007abadcfc20f752435eebbc7462189bc3d5eca4072baaf0`;
  build 2:58, zero warnings/errors.
- Harness: 2 passed, 1 failed, 0 skipped; native checks 11 passed, 2 failed.

Both serial paths (`DeleteItem`, `DeleteItems`) deleted the earlier file and its
directory before CWD into the later directory returned access denied. Backend
state assertions passed: two mutation commands, retained later child and sibling.
Only completion truth failed: the callback incorrectly supplied safe-Retry proof.
The concurrency-4 discovery-first controls issued zero mutation commands and
correctly retained their no-commit classification.

Durations (single serial / single parallel / bulk serial / bulk parallel):
8,297 / 4,277 / 5,080 / 4,889 us. These bounded loopback timings are not general
throughput or real-server qualification. Additional reporter/collision controls
and the host partial-tree witness were added after this baseline; they are not
covered by this failed run.

Raw `results.json`, `trace.txt`, aggregate results and the complete perf stream
were copied intact from the run-owned `artifacts/selftest/last_run` directory
beneath `C:\RedSalamander.Perf\runs\c0-curl-partial-tree-red-debug-20260905`.
The runner reported 37 retained-run disk-audit issues; nothing was removed.
