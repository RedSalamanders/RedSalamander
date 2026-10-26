# I12 terminal repair and closeout qualification

Source: `3dcf7ab167211451882f2c50e202aec68a5a4318` plus the I12 repair worktree.
Status: I12 qualification completed in the [final Fresh Full archive](../2026-09-08_200340_i12_final_full/README.md): **2,116 passed / 0 failed / 52 declared skips**, all 35 entries promoted. The attempts retained here keep their original verdicts.

## First Full and foreground-test warning

`full-first-attempt/` preserves the failed Fresh Full
`20260908T164307Z-104260-ba28d93d5bb041e5900189f2db0b2445`: 2,101 passed,
14 failed, 52 declared skips, with 34 of 35 entries promoted. All failures
are in Commands; File Operations passed 180/0/20. The owner confirmed
interacting with the test windows. This establishes interference, but does not
prove the cause of every failure. All case statuses and skip reasons are retained.

The initial PowerShell/WinForms notice was rejected after the owner pointed out
the existing `DirectedSelfTestInputWarning` and requested a larger, centered
surface. Its historical probe/contract results remain under `qualification/`;
they are superseded, not qualification of the final implementation.

The final implementation reuses that native helper across Commands, DxUi
Menu/NativeTextInput and interactive ViewerPE/ViewerSqlite cases. It uses one
large, DPI-scaled, centered, non-activating, click-through surface, names keyboard
and mouse, and provides a three-second lead. Nested input guards borrow it.
Commands starts the warning before startup activation; no-activation lanes and
case listing remain silent. The original desktop lease is unchanged.

`desktop-replay-first/` retains run
`20260908T181539Z-8780-44014c9a25e444f7b11082f75ca7cf9d`: 136/2/0.
The 46-case predecessor-inclusive filter covered all 14 original failures three
times. Find result shortcuts failed only on repetitions two and three because
the pane reused its prior deletion/move enumeration despite fixture files being
recreated on disk. The shared Find setup now explicitly refreshes before waiting
for the new contents; neither product caching nor assertion deadlines changed.

`desktop-replay-native/` retains the corrected repeat:
`20260908T182930Z-97660-586ba5c9e75145f5b8839707c4bb9b6c`, **138/0/0**.
Debug rebuilt in 2m44s with zero warnings/errors. All 14 original failures and
their predecessor groups passed three times. The [final Fresh Full](../2026-09-08_200340_i12_final_full/README.md) subsequently passed on the final inputs,
including the adjustment that starts the warning before app startup activation.

## Repair and required witnesses

The floating opener previously required focus inside an existing terminal, so
Ctrl+Alt+T and Command Shell Window could not create the first session from a
file pane. Its application-level state now uses the same supported active launch
source as dispatch. `cmd_terminal_session_menu_floating_dispatch` checks shortcut
dispatch, main-menu dispatch, and the existing embedded-session menu route.

The blank embedded tab was a ConPTY launch defect. `CREATE_NO_WINDOW` produced
no ConPTY output with either a stock PowerShell prompt or the owner's profile.
Removing that flag alone allowed redirected host handles to receive shell
output. Explicit `STARTF_USESTDHANDLES` with null standard handles, without
`CREATE_NO_WINDOW`, delivered the marker and prompt through ConPTY. The same
configuration worked without the old startup resize workaround. Empty layouts
now preserve the current grid; the first nonzero host layout sets its size.
Separate isolated probes also delivered the marker and prompt through ConPTY
from Windows PowerShell 5.1 and cmd.exe with the corrected launch configuration.

The isolated probe ran only owned child processes in `Z:\RedSalamander.Perf`.
`qualification/conpty-launch-probe.json` retains all observations and binds each
original local result by SHA-256 while omitting prompt contents. The handle
inheritance behavior and explicit-null remedy are also described by the Windows
Terminal maintainer in [Microsoft's ConPTY discussion](https://github.com/microsoft/terminal/discussions/15814).
The integrated lifecycle case now requires initial formatted VT screen text in
addition to Running and authenticated idle state. A test-only hook reuses the
production bounded formatter and accepts customized prompt formats. Its first
UIA-based implementation was rejected after a broker-read timeout/fail-fast;
the failed attempt is preserved separately, with no Full pass claimed for it.

## Focused evidence

- `release/`: nine exact terminal/focus/FileOps cases passed, zero failed/skipped;
  run `20260908T163756Z-28496-52f8387c7b26486cb0c3fd2d3b69e1cb`, build receipt
  `81328c5a7cc2c556b29815ac7b6d08c983bfd0e8ace3c485f1da122aa8eb0a90`.
  Build completed in 3m44s with zero warnings/errors. The full lifecycle case
  passed in 6,167 ms, including initial screen text and confirmed folder follow.
- `qualification/live-release-entrypoints.json`: actual Alt+7, Ctrl+Alt+T from
  a file pane, and the Command Shell Window menu displayed customized prompt
  output with the production repair. Some private-use prompt glyphs displayed
  as boxes in the selected font; no glyph-rendering improvement is claimed.
- `release-first-attempt/`: the original ten-case attempt remains failed
  (8/2/0): Debug-only preview text instrumentation and pending fixture input
  before folder follow. The preview witness remains required in Debug Full.
- `release-uia-timeout/`: the first nine-case attempt remains incomplete/failed
  because the UIA-based prompt probe timed out and failed fast. The replacement
  probe reads the production bounded formatter without a UIA worker.
- `debug/`: six exact host/focus/FileOps cases passed, zero failed/skipped;
  run `20260908T160151Z-96976-68fa7bc472ff4deeb072ad767c29a12b`, build receipt
  `61222954624165f7a65c0e56e771a9d5d6e1e9dc9fc7aef3e32880cced76b270`.
  The wildcard-looking embedded-terminal filter did not select those cases;
  this run does not claim their coverage. The nine-case Release run and final Full provide it.
- `qualification/focused-contract-results.json`: all 298 linked Pester cases
  passed, zero failed/skipped. The exact I12 Done-path admission accepts only
  the reviewed filename and rejects an unreviewed sibling. Spec and tool
  inventories separately report zero blocking findings.

## Ghostty decision

Retain admitted commit `9c3ec931d64561a8407dde7ac984ce156ae91539` for this repair.
The canonical online status at `2026-09-08T15:25:38.842Z` observed upstream
`b0c421fcd2e290629d4285c181b52fe2f2095f06`, 211 commits ahead, with no advisories
returned by the successful GitHub/OSV checks. `qualification/status.json`
preserves that observation. This is a dated status check, not a guarantee about
future advisories or an admission of the newer runtime.

The reproduced defects are in host command state and process launch; the existing
runtime displays correctly with the launch repair. The reviewed upstream changes
do not establish a need to update it for these defects. A separate update still
needs the repository's paired runtime/performance and native ARM64 qualification;
I15's outstanding fixture/evidence work is not satisfied by this host repair.

## Performance scope

Use the existing `terminal.open_to_visible_us`, `terminal.render.frame_us`, and
`folder.removal_focus.begin_us` metrics. This work makes no general throughput or
latency improvement claim. Current Release measurements remain in this archive;
[final Full metrics](../2026-09-08_200340_i12_final_full/README.md) retain their tested source identity.
