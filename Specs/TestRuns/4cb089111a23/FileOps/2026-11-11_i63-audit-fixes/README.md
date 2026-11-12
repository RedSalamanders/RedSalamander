# I63 audit fixes: Fresh Full run on `dfb4cff5b`

This folder records the second Full run of the I63 branch. It was made after the renumbering, the merge of
`origin/main` (`67f987ca5`) and the fixes for the confirmed findings of an external audit of 2026-11-11.
The repository verdict of the run is **FAILED** (2 failed cases of 2,863). The File Operations, Commands
and CompareDirectories suites have no failure. One failure also exists on `main`. The other came from this
branch and was fixed after the run in a test file; the run was not repeated in full. Nothing in this folder
claims a passing Full gate.

## Run

| Fact | Value |
| --- | --- |
| Run id | `20261111T162023Z-49172-62af6ff7467b4bde979a2de8f4d62dd1` |
| Commit | `dfb4cff5b` (clean working tree) |
| Command | `.\Test-Changes.ps1 -Mode Full -Interactive -AllowAlternateVolumeTestRoot -TestRoot C:\RedSalamander.Perf` |
| Mode | Fresh, Debug x64, interactive entries included, second-volume cases enabled |
| Build | 0 warnings, 0 errors |
| Total | 2,863 cases: 2,802 passed, 2 failed, 59 skipped; 127 min 1 s |

## Result per entry

| Entry | Passed | Failed | Skipped |
| --- | ---: | ---: | ---: |
| FileOperations | 298 | 0 | 25 |
| Commands | 966 | 0 | 5 |
| CompareDirectories | 257 | 0 | 27 |
| ToolsPesterTests | 1,238 | 2 | 2 |
| ToolsPesterBuildToolchain | 30 | 0 | 0 |
| The other 13 entries (both ViewerPETests entries, both ViewerSqliteTests entries, ProductUiTests, LocalizationTests, SettingsSchemaTests, RedConfigureTests, PluginContractTests, CrashHandlingTests, FileSystemCurlTests, PerformanceTests2, VcpkgMergeSynthetic) | 1 each | 0 | 0 |

Cases named because earlier runs or the audit fixes concern them:

- `DiscoveryScope_MultipleFoldersLiveProgress` passed. Before this run it passed 36 of 36 in a focused run
  of 12 repeats on the fixed code. Its one earlier failure is attributed to the stale in-flight entry that
  the audit fix clears. The attribution is by mechanism: the failing field of the original run was not
  recorded.
- `FileOps_ReviewRemediationContracts` passed. It holds the new regression cases: compaction of the
  stream entries, a location that does not enforce sharing, Cancel on the no-Recycle-Bin question, and the
  keep-awake retry. These cases were not run against the code before the fixes, so no failing run of them
  exists.
- `R4A19_DiscoveryIndependentVolumes` passed.
- The ViewerPETests close case that took 643 ms in the first Full run passed here. Its cause was not
  investigated; one pass does not show that it is stable.

## The two failures

| Entry | Case | Classification |
| --- | --- | --- |
| ToolsPesterTests | `keeps manual ASan on both native architectures with seeded detector proof` | Not from this branch. The test, the workflows and the engine pin are identical on `origin/main`; see the first run's folder `2026-11-11_i62-closeout`. Not fixed here. |
| ToolsPesterTests | `keeps IronLedger FolderView data-integrity guards wired` | From this branch. The audit fix replaced `*hit` with one `hit.value()` read in `FolderView::OnLButtonDown`; the source contract still pinned the old spelling. The product code is as intended. |

## After the run

The source contract was updated in commit `954eb77b8` (one test file, no product code). It pins the new
spelling and refuses a `*hit` dereference in that function. On `954eb77b8`:

- `.\Tools\Run-AllTests.ps1 -Suite PR -TestRoot C:\RedSalamander.Perf -AllowAlternateVolumeTestRoot`, run id
  `20261111T190633Z-51960-824878de26b44acdb258ebca51594193`: build with 0 warnings and 0 errors; 1,253
  cases, 1,250 passed, 1 failed, 2 skipped; 21 min 29 s. ToolsPesterTests: 1,239 passed, 1 failed, 2
  skipped. The one failure is the ASan contract above. The PR suite holds the standalone entries and the
  tooling tests; it does not run the File Operations, Commands or CompareDirectories suites.
- Between the two runs only `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` changed, so the three
  product suites of the Full run above were run on the same product code as `954eb77b8`.

No Full run was made on `954eb77b8` or on any later commit. Later commits of this branch change only this
archive and documents.

## Skipped cases

The skips and their reasons are the same as in the first run's folder: missing connection profiles and
SMB or MTP resources, opt-in ViewerSpace benchmarks, two legacy 7z qualifications that are not
implemented, a journal cursor that is not available and a sandbox volume that is not ReFS. Each reason is
in the archived files.

## Files

| File | Content |
| --- | --- |
| `run-all-tests-results.json` | Aggregate v2: plan, decision and snapshot digests, the build receipt id and the verdict of each of the 18 entries. |
| `run-all-tests-results-v1.json` | Every case of every entry with its status and skip reason. |
| `fileops_results.json` | The File Operations suite result: 323 cases with status, reason and duration. |

Traces and the performance metric stream stay in local evidence storage under the run id; they exceed the
archive size limits. The PR suite run after the fix is recorded by its run id only.
