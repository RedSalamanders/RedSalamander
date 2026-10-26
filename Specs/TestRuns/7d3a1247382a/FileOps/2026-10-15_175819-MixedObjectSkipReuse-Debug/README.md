# Mixed-object Skip reuse: focused x64 Debug

The reported sequence is an Access denied question for a file followed by the same question
for a folder, despite checked Apply to all and Skip. The cache accepted the choice but compared
source, destination and link kinds for every action. The older callback fixture covered only
two files and could not expose this gap.

The action-specific repeat key now ignores object shape only for non-mutating Skip in the
AccessDenied, SharingViolation, PathTooLong and NetworkOffline families. Task, error class,
source/destination profiles, destination namespace root and repeat eligibility still match.
Actual observed kinds remain in diagnostics. Replacement, Keep Both, metadata consent and
read-only deletion retain their existing authority boundaries. The capacity/store/removal/lookup
paths share this rule, including a matching grant in a full 64-entry cache.

## Native evidence

- Run: `20261015T155242Z-17292-57117ab2082c4fccb01056045e107257`.
- Case: `FileOps_RepeatDecisionContracts`; 3 passed, 0 failed, 0 skipped including Setup/Cleanup.
- [Debug build receipt](build-receipt.json):
  `a462ade886f9fed1b143a93c0a71636c63cca55df66669b6d9b1bf66bae29a68`.
- The test-enabled x64 Debug application/bundled-runtime build completed with zero warnings/errors.
  [Compiler output](debug-build-log.txt) is preserved without editing.
- [Native trace](fileops_trace.txt), [native results](fileops_results.json),
  [raw aggregate](run-all-tests-results.json), [workspace snapshot](workspace-final.json) and
  [raw performance records](perf/perf_metrics.jsonl) are retained byte for byte.

Each row below is a fresh task. Five same-scope metadata shapes use real Local metadata:
file/file, folder/folder, file/missing destination, folder/missing destination and both paths
missing. The issue status is injected deterministically; this does not recreate protected
Program Files ACLs or establish remote-provider behavior.

| First item | Checked | Same-scope prompts | Cached applications | Total prompts |
|---|---|---:|---:|---:|
| File | Yes | 1 | 4 | 5 |
| File | No | 5 | 0 | 9 |
| Folder | Yes | 1 | 4 | 5 |
| Folder | No | 5 | 0 | 9 |

Each task additionally requires fresh questions for a different source profile, destination
profile, destination root and SharingViolation. Four cached applications share one retained
aggregate row whose observed kinds describe the first cached item. Callback active scopes
return to zero. `FileOps.SelfTest.AccessDeniedSkipScope` value0 is the same-scope prompt count;
value1 is application count, not aggregate row count. Zero duration is a work-count witness.

The existing nested Keep Both fixture still reports one prompt/one cached application and exact
original/sibling bytes. Checked metadata consent still reports two prompts/two cached uses;
unchecked consent reports four prompts/no cached use with exact bytes. The in-memory matrix
covers all four failure classes, file/folder/link/unknown shapes, negative mutation/consent
actions, other conflict classes, eligibility and replacement at cache capacity.

## Runner and static validation

The exact case now joins the existing reviewed FileOps-only no-activation admission. Its
fixtures use owned task/window hooks and direct decisions, without desktop-global input or
assertions. Exact casing, no proof injection and the FileOps-only selection remain required.
Broad selections and near-misses retain their desktop lease. The actual run reported
`PENDING_INTERACTIVE (0)` and launched with `--selftest-no-activate`.

The four source/governance/inventory suites passed all 237 tests. In the combined five-suite
attempt, 304 of 305 tests passed and the native-inventory planner test required an explicit
validation profile; [that original result](source-contracts-initial-summary.json) is preserved.
With x64/Debug explicitly supplied, the [68 planner tests](planner-contracts-retry-summary.json)
all passed. This supplies 305 distinct passing checks across the recorded attempts, not a
single uninterrupted 305-test pass. Spec inventory reported zero blocking findings.

An earlier Debug compilation had two aggregate-initialization warnings in the new matrix;
explicit array braces fixed them before the final build. Its noninteractive invocation selected
zero native tests under the prior conservative desktop classification. Neither attempt is a
native regression pass or a performance baseline.

## Qualification limits

[The Release build attempt](release-build-attempt.json) was blocked before compilation by the
independently running Release application, PID 51848. It was not stopped or replaced. Close it
before rebuilding that profile. No Release runtime, Full-repository, ARM64, foreground,
protected-directory ACL, remote-provider, throughput improvement or latency qualification is
claimed. The raw aggregate retains the same three unrelated legacy sandbox-directory findings;
this is not a clean sandbox-audit verdict. No baseline/candidate timing comparison was made.
