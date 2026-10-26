# I26 conflict placement — native popup integration

`cmd_pane_fileops_popup_presentation_settings_and_taskbar` passed 1/1 in x64 Debug, zero skips.
Build: zero warnings/errors. The existing notified interactive desktop lease was used under the
user's explicit authorization. This case is a focused native integration witness, not a claim that
all locale/DPI/layout and accessibility acceptance is complete.

- Runner: `20260919T105917Z-4256-dffbfaa4e1594f9e85c786b0a0036da8`.
- Receipt: `a1b48c470e586dcbfff44c57306014e1c852a192f12dccc15796b2c242936422`.
- Source snapshot: `2a7136c4614ba746c09d426ba72792c9212510c290ccf6d2c4539667603e67f9`.
- HEAD: `a212edec433d658748ae1273cba68633c86de8b2`, with dirty implementation/test inputs.

The test checks exact KeepBoth/Overwrite/Skip/Cancel primary order, SkipAll-only overflow, immutable
loading/default/Escape state, keyboard/UIA behavior, and stale overflow Skip rejection. Direct Skip
then resolves the real conflict with the existing destination payload preserved. Existing taskbar,
density/settings, reduced-motion, high-contrast and footer-only recreation assertions also pass.
Full results, trace, metrics and build receipt are retained alongside this document.

Command: `Run-AllTests.ps1 -Suite Commands -CaseFilter
cmd_pane_fileops_popup_presentation_settings_and_taskbar -Platform x64 -Configuration Debug
-TestRoot C:/RedSalamander.Perf -TimeoutMultiplier 3`, with `RSBuildEnableTests=true`.
Full log: `C:/RedSalamander.Perf/evidence/i26-ui/i26-policy-commands.log`.
The separate sandbox audit still reports 225 preserved pre-existing run directories. This is not
clean Full closeout evidence. See the [policy comparison](../../FileOps/2026-09-19_105819-i26-policy-candidate/README.md)
for the matched construction-cost diagnostic and its limits.
