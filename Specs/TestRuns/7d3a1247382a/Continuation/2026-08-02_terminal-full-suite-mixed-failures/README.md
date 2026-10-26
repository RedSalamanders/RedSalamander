# Embedded Terminal full-suite failure progression

This bounded continuation record preserves the failures encountered while
closing the first embedded-terminal implementation. The large raw artifacts
from the first mixed run
`20260802T054453Z-4744-707e4439b5214c69b03348c87cc9039f` remain locally
retained under the ignored TestRuns tree; this versioned summary is the durable
diagnostic record and is not a green result.

Two later one-failure runs narrowed the remaining harness issues:

| Run | Aggregate | Failure | Disposition |
|---|---|---|---|
| `20260802T085118Z-46680-5973d5b4c9c449a3b1495e648a298dbe` | 1,179 pass, 1 fail, 53 skip | Monitor ETW selftest was launched after a normal product build had replaced the monitor binary with a tests-disabled build. | Rebuilt `RedSalamanderMonitor` with `RSBuildEnableTests=true`; exact ETW and next Full monitor gate passed. |
| `20260802T175839Z-38872-6de436cacc63493aaf5c529a7c48d64e` | 1,179 pass, 1 fail, 53 skip | Hidden-file removal was asserted after `WaitForPaneItems`, which proves inclusion but can accept the prior superset. | Moved visibility assertions after the exact-count navigation-shell barrier; the exact case passed 50 repeats and the full navigation prefix passed. |

The authoritative successor run is
`20260802T190401Z-4616-084da5bd053c43418460a8bf7a4ee84a`: 1,180 passed,
0 failed, 53 justified skips, all 17 suites passed, zero blocking
classification/quarantine, and zero TestSandbox disk-audit issues. The focused
real terminal lifecycle then passed 10/10 in
`20260802T204347Z-2000-ced799a18ae64b8ea52d28e82e290865` after the final x64
dependency-state restoration and zero-warning Terminal rebuild.
