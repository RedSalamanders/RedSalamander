# Lighthouse Track A — Full-suite closeout attempt

- Date: 2026-07-15
- Branch: `codex/lighthouse-loader-hardening`
- Command: `./Tools/Run-AllTests.ps1 -Suite Full -SkipBuild`
- Run id: `20260715T063912Z-68848-0c7538cb385a4701ad263d91ce76ee93`
- Duration: 2,206,178 ms (36m 46s)
- Result: 1,138 passed, 10 failed, 53 skipped

The loader/action-specific focused gates were green before this run: five
`settings_store_file_actions_v16_` cases and two `file_action_external_launch_` cases passed. The Full run
also completed all File Operations families and the first-party plugin/viewer executables under the safe DLL
search policy; no launcher, external-action, settings-file-action, plugin-load, or DLL-resolution failure was
reported.

The ten Full failures are outside Track A's changed behavior:

- Commands UI isolation/settling: six Preferences, Find, Compare Options, and pane-filter interaction cases.
- FolderView performance interaction: `folderView_perf_huge_folder_scale` and
  `folderView_perf_relayout_churn_while_scrolled`.
- `RedSalamanderMonitorEtwLatency`: executable exited 1 without captured output.
- `ToolsPesterTests`: Pester adapter exited 1.

The aggregate artifact classifies the three failing suites as `UNCLASSIFIED_FAILURE`; none was quarantined or
converted to a retry-only pass. These remain external Full-suite blockers and are not accepted as a green
Lighthouse closeout gate.

Preserved files:

- `run-all-tests-results.json`: complete aggregate result, case list, classifications, and disk audit.
- `fileops-last-family-results.json`: last completed File Operations family result.
