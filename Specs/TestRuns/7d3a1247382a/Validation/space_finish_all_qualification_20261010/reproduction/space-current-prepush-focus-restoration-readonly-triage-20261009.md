# Current PrePush interactive focus-restoration triage (read-only)

Run: `20261009T201507Z-45224-aa600fc03f9e4418b3b3e7152722bb0f` at candidate worktree `C:\Users\PVB\.codex\worktrees\viewer-space-explorer-analyst\RedSalamander`.

## Recorded witness

The preserved `D:\RedSalamander.Perf\runs\20261009T201507Z-45224-aa600fc03f9e4418b3b3e7152722bb0f\artifacts\selftest\last_run\ViewerPETests.Interactive.output.log` ends at lines 114–115 with:

- `[FAIL] TestViewerWebUsesDxUiComboHostWithoutVisibleLegacyCombo parent observes restored foreground and keyboard focus`
- `[FAIL] Viewer suite stops before another child after a containment or desktop-restoration failure`

For this Web case, the same output records successful child launch/resume/exit and child pass, no remaining owned job processes, saved desktop identity/input ownership permitting restoration, and desktop ownership remaining intact immediately before restoration. The PE combo case immediately before it records the parent observing the exact saved foreground and keyboard focus without restoration writes. The output contains no ViewerSpace-named case and ends `ViewerPETests failed.`

## What the restore code establishes—and does not

At candidate `Tests/ViewerPETests/ViewerPE.Tests.Runner.cpp:864-880`, each activation case saves `GetForegroundWindow()`, its process/thread identity, `GetGUIThreadInfo` focus, and `GetLastInputInfo`. The `scope_exit` cleanup first joins/terminates only the owned child/job and verifies it is quiet (around lines 914–1007), then at lines 1015–1048 validates that the saved HWND/process/thread/focus identity is still valid, last-input time is unchanged, and the current foreground is either empty, the saved one, the child, or the harness. If the saved state is not already restored, lines 1079–1100 first observe it, recheck ownership just before restoration, call `SetForegroundWindow(previousForeground)`, conditionally call `SetFocus` only when the saved thread is the current thread, and finally require exact foreground and saved GUI-thread focus equality. Line 1103 emits the failing parent witness.

Thus this log establishes that the saved identity/input preconditions and immediate pre-restoration ownership check passed, yet the final restore predicate was false. It does not record whether `SetForegroundWindow` returned false, whether foreground remained different, whether `GetGUIThreadInfo` failed, or whether the focus HWND differed. The code's detailed `[INFO] restoration refused...` output is only emitted when the earlier permission predicate fails; no such line occurs here. There are no final foreground/focus HWND values or API return values in the preserved output, so the specific failed final conjunct and its cause remain unknown. Nothing in these artifacts establishes external/user input or another environmental cause.

The loop in `RunFullSuiteInFreshProcesses` at lines 15555–15572 sets `desktopBoundaryIntact` from the restoration result and breaks before the next child if false. The following generic stop assertion consequently fails as intended. This is a containment/desktop-boundary stop, not evidence that the unlaunched next case failed.

## Cases not reached

The registered activation-required sequence is at lines 15514–15550. The log proves the PE combo and Web combo children were reached; PE passed its restoration check, and Web failed it. The suite then breaks. Later activation cases not run in this ViewerPE interactive entry are:

- `TestViewerImgRawUsesDxUiComboHostWithoutVisibleLegacyCombo`
- `TestViewerTextUsesDxUiComboHostWithoutVisibleLegacyCombo`
- `viewer_text_encoding_picker_open_and_filter_full_catalog`
- `TestViewerSpaceWindowOpensWithoutVisibleChildFallbackAndEscapeCloses`
- `TestViewerVlcWindowTabTransfersFocusToHudAndClosesCleanly`
- `TestViewerShellComboHostsLongRunOpenCloseStayStable`
- Debug-only `TestViewerTextAsyncOpenAndUtf8HexTerminalContracts`

`TestViewerSpaceWindowOpensWithoutVisibleChildFallbackAndEscapeCloses` is the ViewerSpace activation-required case in this list. The ViewerSpace validation, lifecycle, bounded-scan, directory-summary, and other noninteractive cases are registered separately as noninteractive at lines 15529–15535, so they are not part of this interactive group. Aggregate `run-all-tests-results.json` records `viewer-pe.interactive` incomplete; it separately records `viewer-sqlite.interactive` promoted after it, so the PE child-suite stop did not prevent the later SQLite interactive plan entry from running.

## Follow-up boundary

No rerun, build, process, desktop/input operation, source edit, or index change was performed. If further diagnosis is authorized, preserve the existing restoration predicates and desktop lease; capture read-only final values for `activated`, `GetForegroundWindow()`, `GetGUIThreadInfo` success, and observed vs saved focus around the existing final check. Do not weaken the focus assertion or replace the desktop lease based on this incomplete witness.
