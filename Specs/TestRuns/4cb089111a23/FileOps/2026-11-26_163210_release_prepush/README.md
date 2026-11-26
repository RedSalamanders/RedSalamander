# I66 x64 Release non-interactive PrePush

Fresh run `20261126T144231Z-76004-1019cc96a5804abbbe89eb344337d076` completed in 46 minutes and 12 seconds. It exercised 14 scheduled entries on frozen source `2d41ebc552455a91b5a22df48b0d934957f81200` with x64 Release binaries. The build had zero compiler warnings and errors. The runner verified the existing artifacts under the final receipt `b826f24318753439412044c21bf8e99cee93e57eefa6084bf787826930dd7e74` and snapshot `67a53a21a3c79a2becb5b1852d036cdc0b83aa28d98877168f187e0d378c1818`.

| Result | Recorded value |
| --- | --- |
| Change verdict | `PASSED` |
| Repository verdict | `NOT_EVALUATED` |
| Runner exit code | 2 |
| Scheduled entries | 14 passed, all promoted from this run |
| Reported cases | 1,801 passed, zero failed, 29 skipped |
| Tooling Pester | 1,500 passed, including the R0e admission and archive checks |
| Compare Directories | 255 passed and 29 environment skips |

The run used no desktop lease. Commands, FileOperations, ViewerPETests.Interactive and ViewerSqliteTests.Interactive remain outside this run's coverage. The environment skips include unconfigured remote profiles and unavailable filesystem and test-hook capabilities. Native ARM64 and later main commits have separate qualification.

The runner archived the results and all case reasons under `C:/RedSalamander.Perf/evidence/completed-runs/20261126T144231Z-76004-1019cc96a5804abbbe89eb344337d076/artifacts/selftest/last_run`. Its disk audit separately reported three unexpected root files, nine unexpected run directories and one stale run directory. Those paths were preserved.

| Archived file | SHA-256 |
| --- | --- |
| `run-all-tests-results.json` | `5dbd82daead91a7fb9503a506d94b27ba69f9d278e9ecd9298590ba23f9304a8` |
| `run-all-tests-results-v1.json` | `4a6a5e10e9c30f8b5fa0a6e409ce17c79117762bea1b5c6794ed888d86c5a860` |
| `run-all-tests-case-history.jsonl` | `a2879c2cc17389b028094d4cccb8599b4f351cc14f06156624d19472cd7e685e` |

## Hosted follow-up

The x64 job for source `d68962598e7b9b8f14f30d3ec8365fb8d6770f3c` finished with 3,028 passed, one failed and 69 skipped. Its sole failed case was the Change Case documentation gallery. That source predates the gallery repair that passed and merged in PR #325. It passed `cmd_pane_fileops_conflict_two_pending_task_identity` and `Phase5_DiscoveryCancelReleasesSlot` on the original witness. The pointer candidate still needs its own result. [Hosted run](https://github.com/RedSalamanders/RedSalamander/actions/runs/45811712240)

The preserved x64 artifact is in local evidence collection `20261126T153611Z-57200-7fc4efb2c63d4b028aefd3fd84921e06`. Nested hosted run `20261126T135315Z-3236-b4eed0f90c864a9cab53f7dbb0db7c6b` has result SHA-256 `94ed298a12c01186ef6064db6a7d6df2e7ba661b214601f72cc87cbef82de664`. Its case history records the pointer and Start now passes on rows 2453 and 2986 and the gallery failure on rows 2597 and 2757.

The ARM64 Commands execution on the same source reported 964 passed, two failed and six skipped. Change Case failed its gallery capture. The pending-decision case changed task B's disclosure and preserved both unresolved choices. It failed because the UI Automation route lost actual foreground and focused-control verification. The isolated repeat of that case reported one pass, zero failures and zero skips. The outer runner then failed the foreign-console retirement check. The pointer draft addresses the earlier x64 viewport clipping. ARM64 foreground ownership still requires separate evidence.

The archived ARM64 collection `20261126T154026Z-72976-49900dfed70c4954abb0a817803b8912` preserves original and retry Commands results in nested run `20261126T141346Z-6076-2859d2900c3d4c0f9c6292d5f40d33ef`.

| Archived Commands file | SHA-256 |
| --- | --- |
| `previous_run/commands/results.json` | `e7502a5dcb7e7dda9c56fac7bc5a4264b812388fee0ac20a92dbfa2d60613256` |
| `last_run/commands/results.json` | `1a4dbcde22f542e4ec0726e04f3c2e2fd1095f825dbc939e3680e18f333d443e` |
