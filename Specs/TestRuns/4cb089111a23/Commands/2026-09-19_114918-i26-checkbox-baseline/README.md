# Repeat scope — retained baseline failure

Unchanged product at `9d3318382`, with the new focused test/instrumentation. x64 Debug build
completed with zero warnings/errors. The presentation case fails as expected because the repeat-scope
strip exposes an ordinary Button instead of UIA Toggle. See results, trace and build receipt here.

Runner: `20260919T114541Z-34560-c5b8877489894d7685c0844cab11410e`.
Full local log: `C:/RedSalamander.Perf/evidence/i26-ui/i26-checkbox-baseline.log`.
The five eight-redraw batches after four warm-ups are 66431, 66516, 66773, 66437 and 66619 us
(median 66516 us). The full metric stream is retained. Compare the
[candidate](../2026-09-19_120441-i26-checkbox-candidate/README.md).

Command: `Run-AllTests.ps1 -Suite Commands -CaseFilter
cmd_pane_fileops_popup_presentation_settings_and_taskbar -Platform x64 -Configuration Debug
-TestRoot C:/RedSalamander.Perf -TimeoutMultiplier 3`, with `RSBuildEnableTests=true`.
The harness used its notified interactive desktop lease. The separate sandbox audit retains 225
unrelated pre-existing run-directory findings. This expected failure is not Full acceptance.
