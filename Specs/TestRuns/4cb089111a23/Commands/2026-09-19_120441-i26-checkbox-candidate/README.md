# Repeat scope — native Checkbox candidate

The presentation case passes 1/1, zero skips, in x64 Debug. Build: zero warnings/errors.
It checks real Checkbox type/visible state, UIA Toggle, repeated model acknowledgement without
another toggle, Space, pointer activation, release outside the target and keyboard/pointer parity.
The existing direct Skip flow then preserves the destination. The task pause shortcut now lets the
focused repeat-scope Checkbox own Space. The consumer DxUi pin is unchanged.

Runner: `20260919T120102Z-8544-2fc261a3e951486db425764aae75e22f`.
Receipt: `dc1f85f74d5c4b30f94d5fe7f2919088ff64723515d5ff88c50ea245faea65e6`.
Source snapshot: `6aec08a5318f18fac9dda52c03c5fc747464a2af63aa44fb9b3e5a41cf301f0c`.
HEAD: `9d3318382525a0b725418576b06ff5a948dc5909`, with dirty implementation/test inputs.
Full local log: `C:/RedSalamander.Perf/evidence/i26-ui/i26-checkbox-candidate-v2.log`.

The five eight-redraw batches after four warm-ups are 65321, 67311, 66993, 66308 and 67118 us
(median 66993 us), versus [baseline](../2026-09-19_114918-i26-checkbox-baseline/README.md)
66516 us: +0.72%. The measured block and prompt are identical. Candidate adds later pointer checks
outside the measured block; fixture formatting also changed after baseline. This synchronous Debug
diagnostic does not establish Release frame pacing/resource acceptance. Full metric streams and
build receipts preserve the exact input/output identities.

The first candidate exposed a real Space/pause interception defect; its failed trace and metrics
remain under `C:/RedSalamander.Perf/evidence/i26-ui/checkbox-candidate-first`. Its behavior is not
counted as passing. Application captures, disabled-state geometry, French wrapping, fresh mixed-DPI
and complete product qualification remain open.

Command: `Run-AllTests.ps1 -Suite Commands -CaseFilter
cmd_pane_fileops_popup_presentation_settings_and_taskbar -Platform x64 -Configuration Debug
-TestRoot C:/RedSalamander.Perf -TimeoutMultiplier 3`, with `RSBuildEnableTests=true`.
The harness used its notified interactive desktop lease and restoration. The separate sandbox audit
still reports 225 preserved unrelated run directories; this is not clean Full closeout evidence.
