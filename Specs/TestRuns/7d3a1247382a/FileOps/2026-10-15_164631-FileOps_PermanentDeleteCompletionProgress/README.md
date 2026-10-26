# Final focused Debug: deletion and completed folder counts

`FileOps_PermanentDeleteCompletionProgress` **passed**. The runner recorded 3 passed,
0 failed, 0 skipped (Setup, the selected case, Cleanup), exit code 0.

- Run: `20261015T144600Z-48204-814d6beb207a4d3eb5d01046f049e98d`.
- [Test-enabled x64 Debug receipt](build-receipt.json): `047f98907061f915d00f1f490b967e89fab62bf15c3f8aa6169c531d0240b5af`; the owning build passed with 0 warnings/errors.
- [Trace](fileops_trace.txt) / [raw aggregate](run-all-tests-results.json): actual provider and host-driven checks below, with foreground preservation where exercised.

| Witness | Observed result |
|---|---|
| Partial deletion before read-only consent | 1 file / 17 B removed; read-only child retained |
| Explicit read-only subtree confirmation | Cumulative 2 files / 1 folder / 46 B removed |
| Cancellation after a committed child | 1 file / 7 B retained in completed-work counters; root remains |
| Refused nonempty directory removal | 0 completed objects; read-only attributes and survivor retained |
| Host-driven parent/read-only-child consent | Separate initial permanent and read-only confirmations; final 2 files / 1 folder / 46 B |
| Selected read-only file and empty folder | Each requires its own explicit answer; no grant crosses roots; final 1 file / 1 folder / 17 B |
| Parallel and verified sequential bridge Copy | Each retains 2 files / 3 folders, exact payload bytes and completed summary counts |

This clean focused result supersedes the earlier fixture qualification failures: the
[callback assertion](../../Continuation/2026-10-15_162805-delete-fixture-callback-assertion/README.md)
and the [sequential bridge guard failure](../2026-10-15_163647-FileOps_PermanentDeleteCompletionProgress/README.md).
Their original interrupted/failed evidence remains intact.

The same receipt backs the passed [36-layout/18-capture presentation case](../2026-10-15_164700-FileOps_SkipCountingPresentation/README.md)
and [repeated-choice case](../2026-10-15_164725-FileOps_RepeatDecisionContracts/README.md).
[Source checks](final-source-contracts-summary.json) record 247 passed / 0 failed / 0 skipped.
The [question inventory](question-template-inventory.json) is a static resource-ID count, not
runtime coverage of every possible sentence or provider. [Review](REVIEW.md) records the wider source assessment.

Qualification is limited to these named x64 Debug scenarios. No Full-repository, Release, ARM64,
remote-provider or foreground-suite pass is claimed. The aggregate preserves `is_clean:false`
and 3 legacy sandbox findings (`archive-skip-check-7756799779d4`, `dxui-ci-first-37835055279`,
`dxui-ci-retry-37835055279-attempt2`). No raw statuses were rewritten or cleanup performed;
candidate measurements do not establish a before/after throughput improvement.
