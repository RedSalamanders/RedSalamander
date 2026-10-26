# Last-three-days remediation performance evidence — 2026-07-11

## Identity and scope

- Machine profile: `4cb089111a23` (`SINON`), Debug x64, same machine for baseline and candidate.
- Repository commit embedded in both runs: `eaf640798eb0fa21f2b5b6e3883891ac2442a255`; the candidate also contains the uncommitted remediation worktree.
- Candidate binary identity at measurement time:
  - `RedSalamander.exe` SHA-256 `672BABE20FE2AF05C4606C6E9EB92EAE7E2900279F1840765116FFA51B2A3E8A`
  - `Common.dll` SHA-256 `EC18053925AC8EE83B2BC05CC9A71898EB131769000A29F14301472DD088FBD1`
- Final validation identity after the settings/Folder View follow-up:
  - `RedSalamander.exe` SHA-256 `E40E9A192F2AF48C1B53597E0AA40751177E87DE3E6E896DDD01543625B91C23`
  - `Common.dll` SHA-256 `DFE57B551AA44922725F918E6906AAEF7BFC19E82FA722F368B8586F45F61C27`
- Raw telemetry remains local and ignored by Git under the run folders named below. Under the TestRuns
  archive policy this checked-in README is the accepted compact correctness/performance artifact; its
  commands, counts, identities, and classifications supersede ignored raw paths.

## Commands and correctness

Baseline raw run: `Specs/TestRuns/4cb089111a23/Commands/2026-07-11_111538`

```powershell
.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-timeout-multiplier=3 --selftest-case=cmd_pane_fileops_conflict_prompt_compacts_actions --selftest-repeat=100
```

Result: 100 passed, 0 failed; suite duration 247,608 ms.

Candidate raw run: `.build/ArchivedRawTestRuns/2026-07-11/4cb089111a23/Commands/2026-07-11_130136`

```powershell
.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-timeout-multiplier=3 --selftest-case=cmd_pane_fileops_conflict_prompt_metadata_and_actions,cmd_pane_fileops_completed_group_and_navigation,cmd_pane_fileops_popup_presentation_settings_and_taskbar --selftest-repeat=100
```

Result: 300 passed, 0 failed; suite duration 291,170 ms. The three candidate cases are the split successors of the baseline monolith. Their combined duration is not a product-performance comparison because each case now owns independent setup and teardown.

Warmup: neither run used a separate discarded warmup pass; the first repetition is included in each
reported distribution. Decisions therefore use the documented p95 sample gates and same-machine
baseline/candidate comparison, not isolated cold-start medians.

## File Operations metrics

Repository analyzer quality gates are 200 samples for p95 and 1,000 for p99.

| Metric | Run | Count | P50 us | P95 us | P99 us | Max us | p95 / p99 quality |
|---|---|---:|---:|---:|---:|---:|---|
| `FileOps.Conflict.MetadataPromptUs` | baseline | 200 | 103 | 55,704 | 60,657 | 107,778 | pass / fail |
|  | candidate | 400 | 39,594 | 50,599 | 55,702 | 136,410 | pass / fail |
| `FileOps.Conflict.MetadataUs` | baseline | 300 | 50 | 54,515 | 60,079 | 107,778 | pass / fail |
|  | candidate | 500 | 36,471 | 49,682 | 55,090 | 136,410 | pass / fail |
| `FileOps.Popup.Render.TotalUs` | baseline | 10,370 | 5,311 | 9,436 | 11,269 | 131,861 | pass / pass |
|  | candidate | 10,588 | 4,960 | 9,157 | 10,996 | 122,727 | pass / pass |
| `FileOps.Popup.Render.BuildSnapshotUs` | baseline | 10,370 | 234 | 423 | 649 | 128,323 | pass / pass |
|  | candidate | 10,588 | 187 | 366 | 583 | 119,673 | pass / pass |
| `FileOps.Settings.EnqueueUs` | candidate only | 800 | 1,353 | 1,785 | 2,212 | 2,694 | pass / fail |
| `FileOps.Settings.SaveUs` | candidate only | 401 | 7,766 | 82,676 | 208,254 | 472,192 | pass / fail |

Analyzer p95 decisions:

- Conflict prompt metadata: `-9.2%`, noise (no regression).
- Total conflict metadata: `-8.9%`, noise (no regression).
- Popup render total: `-3.0%`, noise (no regression).
- Popup snapshot construction: `-13.5%`, improvement.
- Settings enqueue/save have no baseline because the asynchronous instrumentation is new. The 401 background saves from 800 enqueue samples demonstrate coalescing; persistence latency is no longer caller-thread latency.

The baseline monolith and split candidate emit different fast-path sample mixes, so their metadata medians are not comparable. The p95 distributions meet sample quality and are the keep/revert signal. P99 is claimed only for the two render metrics.

Analyzer command pattern:

```powershell
$baseline = 'Specs/TestRuns/4cb089111a23/Commands/2026-07-11_111538'
$candidate = '.build/ArchivedRawTestRuns/2026-07-11/4cb089111a23/Commands/2026-07-11_130136'
.\Tools\Show-PerfRuns.ps1 -CompareRun $baseline,$candidate -Metric 'FileOps.Conflict.MetadataPromptUs' -MinimumSamplesForP95 200 -MinimumSamplesForP99 1000
.\Tools\Show-PerfRuns.ps1 -CompareRun $baseline,$candidate -Metric 'FileOps.Conflict.MetadataUs' -MinimumSamplesForP95 200 -MinimumSamplesForP99 1000
.\Tools\Show-PerfRuns.ps1 -CompareRun $baseline,$candidate -Metric 'FileOps.Popup.Render.TotalUs' -MinimumSamplesForP95 200 -MinimumSamplesForP99 1000
.\Tools\Show-PerfRuns.ps1 -CompareRun $baseline,$candidate -Metric 'FileOps.Popup.Render.BuildSnapshotUs' -MinimumSamplesForP95 200 -MinimumSamplesForP99 1000
```

## Search service responsiveness

The foreground lifetime regression is also visible in focused same-machine case duration. Baseline source: `.build/ArchivedRawTestRuns/2026-07-11/Specs/TestRuns/SINON/Continuation/2026-07-10_141000_compare_foreground_pipe_only_readiness_output_contracts/compare.results.json`. Candidate source: `.build/SearchRemediationFinal/runs/search-remediation-prefix-postfix-20260711/artifacts/selftest/last_run/results.json`.

| Case | Baseline ms | Candidate ms | Delta | Candidate result |
|---|---:|---:|---:|---|
| `search_service_foreground_logs_request_status` | 10,252 | 290 | -97.2% | passed |
| `search_service_filters_cached_descendants_denied_to_client` | 10,150 | 333 | -96.7% | passed |

The final Search family passed 35 cases with 0 failures and 4 capability skips in 60,525 ms. The exact five readiness/cutover cases passed 5/5 in 11,060 ms.

## Abnormal-run classification and harness guard

The first completed-group repeat stopped after 76 passes because a concurrent `build.ps1` invocation force-terminated the exact-path self-test process 235 ms after its last metric. There was no dump, WER event, assertion, or crash-handler artifact. Evidence is preserved locally at `Specs/TestRuns/4cb089111a23/Continuation/2026-07-11_124246_completed_group_repeat_abnormal_exit` and `.build/CrashEvidence/2026-07-11_124246_completed_group_repeat_abnormal_exit`.

`build.ps1` now aborts with PID/path/command-line diagnostics instead of killing an exact-path process whose command line contains a self-test flag (or whose command line cannot be inspected). Four focused process-selection tests plus five inventory tests passed under both PowerShell 7 and Windows PowerShell.

## Folder View deletion regression

Deleting a file with successful-operation auto-dismiss enabled caused the popup teardown to enqueue
an asynchronous settings save. The directory watcher could observe that write before the save worker
recorded its final file stamp, classify the app's own write as external, and refresh running plugins.
Provider release cleared both pane item models; provider restoration retained the folder paths but did
not re-enumerate them, leaving both panes empty.

The atomic settings writer now returns the identity of the finalized temporary file it moved into
place. Hot reload publishes that exact stamp and ends its internal-save epoch before schema I/O, so
an external replacement during post-write work cannot be mislabeled as the app's write. Epoch changes
retry/repost instead of consuming a notification, process shutdown uses one bounded final-save
request, and late worker state is process-lifetime fenced. Provider rediscovery captures each pane's
raw provider path and instance context, qualifies it once, and explicitly re-enumerates both models.

Final focused commands:

```powershell
.\Tools\Run-AllTests.ps1 -SkipBuild -Suite Commands -CaseFilter 'settings_store_no_recovery_and_file_stamp,settings_hot_reload_self_save_suppression,settings_save_queue_serializes_coalesces_and_flushes,cmd_app_rereadAssociations_reloads_actions_and_refreshes_panes'
.\Tools\Run-AllTests.ps1 -SkipBuild -Suite Commands -CaseFilter 'settings_store_no_recovery_and_file_stamp,settings_hot_reload_self_save_suppression,settings_save_queue_serializes_coalesces_and_flushes,cmd_app_rereadAssociations_reloads_actions_and_refreshes_panes' -SelfTestRepeat 5 -SelfTestShuffleSeed '20260711'
```

Results: 4/4 at run `20260711T150508Z-71928-619bc6c6235f47ceb40d9d8528d23c01`;
20/20 across five shuffled repetitions at run
`20260711T150534Z-93132-88f048b874144584b34962cddf32c806`. Coverage includes exact finalized
stamp equality, external replacement during self-save, failed-save epoch retry, bounded production
final-save timeout behind a 300-second injected stall, local-pane re-enumeration, and a 7z pane with a
nonempty archive instance context. Both runs reported 200 pre-existing standalone DxUi/ViewerPE/
ViewerSqlite directories in the shared TestSandbox audit; none belonged to either runner-owned run ID.

## Final focused validation and Full-suite classification

- Debug x64 `Common` build: `.build/logs/msbuild-20260711_170110_735.log`, zero warnings / zero errors.
- Isolated Debug x64 application build with already-built dependencies:
  `.build/logs/msbuild-redsalamander-isolated-20260711_170148_410.log`, zero warnings / zero errors.
- Exact tooling commands were `Invoke-Pester -Path <file> -PassThru` for
  `TestInventory.Tests.ps1`, `TestHarnessSourceContracts.Tests.ps1`,
  `DocumentationDriftContracts.Tests.ps1`, `TestRunArchive.Tests.ps1`,
  `BuildOutputProcess.Tests.ps1`, and `ResourceLocalizationContracts.Tests.ps1`, plus
  `.\Tools\Test-TestRunArchive.ps1`.
- Tooling results: inventory 5/5, source contracts 132/132, documentation drift 9/9, TestRuns archive
  5/5, build-output collision guards 4/4, resource localization 4/4, and archive lint green for the
  one curated changed artifact.
- Continuation reconciled three concurrently added viewer contracts and the ViewerText checked-seek
  expectation; the current source-contract suite is green at 132/132.
- `.\Tools\Run-AllTests.ps1 -Suite Full` was attempted. It stopped during the solution build before
  tests at `.build/logs/msbuild-20260711_161736_710.log`: concurrent uncommitted ViewerSqlite test
  work did not link `Common.dll`'s perf writer or define its executable TraceLogging provider.
- The shared worktree changed again during final validation; a later dependency build stopped at
  `.build/logs/msbuild-20260711_164723_672.log` because concurrent uncommitted ViewerWeb work omitted
  an `environment` lambda capture. Neither blocker is in the settings, Folder View, FileOps, Search,
  or harness remediation. The isolated owning application target and all focused regressions are green;
  the solution-wide Full gate is classified as externally blocked, not green.

## Decision

Keep the remediation. Correctness/stress cases are green, both metadata p95 distributions are below baseline, popup rendering is stable or improved, settings disk persistence is off the input thread, and Search foreground shutdown no longer pays the former request-lifetime timeout. No p95 analyzer regression remains.
