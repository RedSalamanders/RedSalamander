# I62 closeout: Fresh Full run on `ded63e43b`

This folder records the verification of the I62 closeout commit. The repository verdict of the run is
**FAILED** (5 failed cases of 2,833). The File Operations, Commands and CompareDirectories suites have no
failure. The five failures are outside the code this branch changes; each is classified below. This run is
therefore not a passing Full gate, and nothing in this folder claims one.

## Run

| Fact | Value |
| --- | --- |
| Run id | `20261111T121453Z-18008-65721582e53d407b815eea8e81d114e0` |
| Commit | `ded63e43b` (clean working tree) |
| Command | `.\Test-Changes.ps1 -Mode Full -Interactive -AllowAlternateVolumeTestRoot -TestRoot C:\RedSalamander.Perf` |
| Mode | Fresh, Debug x64, interactive entries included, second-volume cases enabled |
| Build | 0 warnings, 0 errors |
| Total | 2,833 cases: 2,769 passed, 5 failed, 59 skipped; 150 min 32 s |

## Result per entry

| Entry | Passed | Failed | Skipped |
| --- | ---: | ---: | ---: |
| FileOperations | 298 | 0 | 25 |
| Commands | 966 | 0 | 5 |
| CompareDirectories | 257 | 0 | 27 |
| ToolsPesterTests | 1,206 | 4 | 2 |
| ToolsPesterBuildToolchain | 30 | 0 | 0 |
| ViewerPETests.Noninteractive | 0 | 1 | 0 |
| The other 12 entries (ViewerPETests.Interactive, both ViewerSqliteTests entries, ProductUiTests, LocalizationTests, SettingsSchemaTests, RedConfigureTests, PluginContractTests, CrashHandlingTests, FileSystemCurlTests, PerformanceTests2, VcpkgMergeSynthetic) | 1 each | 0 | 0 |

Two File Operations cases are named because earlier I62 runs reported them:

- `R4A19_DiscoveryIndependentVolumes` passed. Its earlier failure was a run started without
  `-AllowAlternateVolumeTestRoot`, not a limit of the machine.
- `DiscoveryScope_MultipleFoldersLiveProgress` passed. Its one earlier failure is not explained and stays
  a watch item of I63 (`RW1`).

## The five failures

| Entry | Case | Classification |
| --- | --- | --- |
| ToolsPesterTests | `keeps native ARM64 regression behind main admission and PRs compiler-free` | Environment. The test reads the ignored `.ci-engine` checkout, which this worktree did not have. |
| ToolsPesterTests | `Release workflow source contracts` (Describe setup) | Environment, same cause. |
| ToolsPesterTests | `Winget shared publication route` (Describe setup) | Environment, same cause. |
| ToolsPesterTests | `keeps manual ASan on both native architectures with seeded detector proof` | In this run: the same missing checkout. With the checkout present the case still fails; see below. |
| ViewerPETests.Noninteractive | `Close returns while the export writer is held and module unload stays blocked` | Timing. The close succeeded and unload stayed blocked, but it took 643 ms against a 500 ms bound. Not rerun. |

Checks made after the run, on the same commit:

- The CI engine was cloned into `.ci-engine` at the revision that `.github/workflows/ci.yml` pins
  (`904a8cb6824da010e2ddcd787d2037c610b0df9c`). The three affected Describe blocks were then run with
  `Invoke-Pester`: 36 passed, 1 failed.
- The remaining failure is `keeps manual ASan on both native architectures with seeded detector proof`.
  The test expects the literal text `platform: x64` in the engine's `qualification.yml`; the pinned engine
  selects its platforms from an input and does not contain that text. The test file, the workflows and the
  engine pin are identical on this branch and on `origin/main` (`99624c538`), so this failure does not come
  from this branch. It is not fixed here.
- This branch changes no file under `Plugins/ViewerSpace`, `Plugins/ViewerPE` or `Tests/ViewerPETests`.
  The cause of the 643 ms close was not investigated.
- An exact Resume of this run was attempted and refused (`CORRUPT EVIDENCE: Resume origin contains an
  invalid promotion or evidence record`, exit 5) after its build produced a new artifact identity. It ran no
  test and contributes no result.

## Skipped cases

Every skip keeps its reason in the archived files. For File Operations (25): no SMB path is configured
(5), no connection profile exists for S3 (3), OneDrive Personal (3), OneDrive Business (2), SharePoint (2),
FTP (2), SFTP (2), SCP (2) and IMAP (2), and two legacy 7z qualifications are not implemented. For Commands
(5): opt-in ViewerSpace benchmarks. For CompareDirectories (27): the same connection profiles, SMB and MTP
resources, a journal cursor that is not available, and a sandbox volume that is not ReFS.

The run's sandbox audit lists nine stray items in the test root that predate this run (three text files
from October 2026 and six old run folders). They were left in place.

## Files

| File | Content |
| --- | --- |
| `run-all-tests-results.json` | Aggregate v2: plan, decision and snapshot digests, the build receipt id and the verdict of each of the 18 entries. |
| `run-all-tests-results-v1.json` | Every case of every entry with its status and skip reason. |
| `fileops_results.json` | The File Operations suite result: 323 cases with status, reason and duration. |

Traces and the performance metric stream stay in local evidence storage under the run id; they exceed the
archive size limits.
