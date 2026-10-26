# Unfocused embedded-terminal dimming evidence

## Result

The x64 Debug lifecycle case
`cmd_pane_embedded_terminal_plugin_lifecycle` passed three consecutive repeats:
3 passed, 0 failed, and 0 skipped. The strengthened case explicitly focuses the
terminal, forces a repaint, returns focus to the source FolderView, and forces an
inactive repaint before continuing its existing ConPTY reuse, path-follow, input,
and teardown coverage.

The candidate was built from clean implementation commit `9e133d3059407cd12b5a788591f614a084a3eb71`.
The full-solution Debug rebuild completed with 0 warnings and 0 errors in
`msbuild-20260809_122002_657.log`; the post-commit host refresh also completed
with 0 warnings and 0 errors in `msbuild-20260809_122555_626.log`.

## Render diagnostics

The retained focus-tagged `terminal.render.frame_us` rows contain 31 samples at
116 columns by 56 rows:

- focused: 21 samples, median 3,166 us;
- unfocused: 10 samples, median 3,975 us.

The existing same-machine Debug archive at
`Specs/TestRuns/4cb089111a23/Commands/2026-08-09_105758/` recorded 14 untagged
samples with a 4,025 us median. A pre-change run of the current lifecycle case
passed 3/3 but emitted no render samples because it did not force a terminal
paint; the candidate selftest closes that instrumentation gap. These small Debug
sets show that the new single full-client overlay executes without an obvious
order-of-magnitude regression. They are diagnostic evidence, not a Release
percentile or performance-improvement claim.

The metric rows have an empty embedded `commit` field because the in-product Git
reader resolves this linked worktree's branch name but not its loose ref in the
repository's common Git directory. The run began from the clean commit identified
above, and no source changed between that build and this capture.

## Reproduction

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild `
  -CaseFilter 'cmd_pane_embedded_terminal_plugin_lifecycle' `
  -SelfTestRepeat 3 -TimeoutMultiplier 3
```
