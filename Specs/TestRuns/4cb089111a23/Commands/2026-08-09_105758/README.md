# Embedded terminal rendering and scrollbar evidence

## Result

The focused x64 Debug Commands run passed
`cmd_preferences_dialog_scroll_host_preserves_retained_page_state` and
`cmd_pane_embedded_terminal_plugin_lifecycle` three times each: 6 passed,
0 failed, and 0 skipped. Runner id:
`20260809T085729Z-29544-5a045f10789a4c3abc2e3877220a7718`.

The Preferences case proves the right page host retains one DxUi scrollbar
owner without a native `WS_VSCROLL`. The terminal lifecycle case opens a real
ConPTY-backed `Terminal.dll` child and exercises visible rendering and teardown.

## Diagnostic render samples

The run emitted 14 `terminal.render.frame_us` samples at 116 columns by 56
rows. The median was 4,025 us. This Debug sample set establishes that the new
per-frame instrumentation is live; it is below percentile-quality thresholds
and is not a Release performance claim. There is no same-machine pre-change
baseline for this newly introduced metric.

## Reproduction

```powershell
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild `
  -CaseFilter 'cmd_preferences_dialog_scroll_host_preserves_retained_page_state,cmd_pane_embedded_terminal_plugin_lifecycle' `
  -SelfTestRepeat 3 -TimeoutMultiplier 3
```

The compact correctness artifacts and relevant `terminal.*` metric rows are
retained beside this file.
