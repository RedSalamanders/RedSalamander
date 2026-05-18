# SelfTest Specification

## Overview

RedSalamander ships three debug self-test suites:
- `--compare-selftest`
- `--commands-selftest`
- `--fileops-selftest`

This document defines the normative result contract shared by those suites.

Related documents:
- `Specs/Testing/Testing_TestCoverage.md` — comprehensive per-suite test case inventory
- `Specs/Testing/Testing_SelfTestRemoteCredentials.md`
- `Specs/Testing/Testing_PerformanceValidation.md`
- `Specs/TestRuns/README.md`
- `Tests/README.md` — central test infrastructure index
- `Tools/Run-AllTests.ps1` — unified test runner with summary reporting

## Result Contract

### Case coverage

Every declared self-test case must produce exactly one case result in the suite `results.json`.

Allowed statuses are:
- `passed`
- `failed`
- `skipped`

There must be no declared case with a missing status.

### Skip semantics

- `skipped` is the correct result when a declared case cannot run because a required precondition is absent.
- Every skipped case must carry a human-readable reason.
- Skipping is part of normal suite behavior for conditional coverage and must not make the suite fail by itself.

### Setup failure behavior

If a suite encounters a fatal setup failure before all declared cases can run:
- the setup failure itself must be recorded explicitly,
- unreached declared cases must still be emitted as `skipped`,
- each backfilled skipped case must explain that it was not executed because of suite setup failure.

## Conditional Coverage Rules

Conditional coverage must remain part of the suite membership. Preconditions change execution status, not case existence.

Examples:
- remote smoke tests skip when required connection profiles, secrets, or sandbox roots are absent,
- plugin-dependent tests skip when the plugin is not available,
- machine-dependent filesystem coverage such as ReFS skips when the required volume is not present.

A case must not disappear from the suite just because the current machine lacks its prerequisites.

Environment variables may select alternate test inputs such as profile names, but they must not be used to remove a declared case from suite membership.

## Suite and Aggregated Results

- Suite `results.json` files must preserve per-case status and reason.
- Aggregated self-test results must count `passed`, `failed`, and `skipped` consistently with the suite artifacts.
- In-product self-test suites must use the shared result-emission helper for suite-level case insertion, status counts, and first-failure propagation. Suites may keep distinct execution models, such as `SelfTest::RunCase` or FileOperations phase-state execution, but summary emission must stay centralized.
- The unified runner must compare each executed in-product suite's result names with runner-native `--selftest-list-cases` output for the same suite/filter and fail on duplicate expected names, duplicate actual names, missing declared results, or unexpected extra results. CompareDirectories may emit an extra `setup` result for explicit setup failure reporting.
- When `Tools/Run-AllTests.ps1` is invoked with a non-empty `-CaseFilter`, runner-native case listing must return at least one expected case for each executed self-test suite. A zero-expected, zero-actual filtered run is invalid evidence and must fail as result coverage drift.
- Runner-injected coverage failures must contribute to the effective suite status, displayed failure counts, runner aggregate `exit_code`, and process exit code even when the native self-test process exits successfully.
- Checked-in archived runs may contain skipped cases; that is valid when the skip reason documents the missing precondition.
- `Tools/Run-AllTests.ps1` must tolerate both direct suite JSON (`commands_results.json`, `compare_results.json`, `fileops_results.json`) and aggregated run JSON (`selftest_run_results.json`) when summarizing archived results.
- `Tools/Run-AllTests.ps1` must also write a runner-owned `run-all-tests-results.json` artifact for every invocation. Multi-suite runs must preserve each suite's status counts, case names, failure reasons, skipped reasons, and wrapper exit code in that aggregate file so later suites cannot overwrite earlier evidence.
- `Tools/Run-AllTests.ps1` must read artifacts from `REDSALAMANDER_SELFTEST_ROOT\last_run` when `REDSALAMANDER_SELFTEST_ROOT` is set; otherwise it must use `%LOCALAPPDATA%\RedSalamander\SelfTest\last_run`. This keeps isolated worktrees and concurrent local runs from reading another checkout's shared `last_run` evidence.
- If a self-test process exits early, crashes, or writes only partial artifacts, `Tools/Run-AllTests.ps1` must report the runner failure from the available JSON/trace data instead of failing its own parser.

## Command-Line Safety

- `--selftest-timeout-multiplier=N` accepts only finite numeric values.
- Invalid or non-finite multiplier values are command-line errors.
- Finite values outside `[0.1, 100.0]` are clamped to that range with a diagnostic.
- Scaled nonzero timeouts must remain finite and at least 1 ms.
- PowerShell harnesses that need the `RedSalamander.exe` self-test exit code
  must use `Start-Process -Wait -PassThru` or `System.Diagnostics.Process`
  rather than relying on `$LASTEXITCODE` from direct invocation of the
  GUI-subsystem executable.
- PowerShell/Pester tests that launch build-toolchain child processes must use
  a finite timeout and must capture stdout/stderr to named log files. A hung
  child build must fail the test with those log paths instead of blocking
  `Run-AllTests.ps1 -Suite Full` indefinitely.

## Self-Test Roots and UI Navigation

- `REDSALAMANDER_SELFTEST_ROOT` may be supplied as a relative path by local
  runner scripts, but the in-product self-test helpers must normalize it to an
  absolute path before exposing it to test cases.
- Self-test fixture directory and file names should keep on-disk path segments
  concise. Worktree-local absolute roots can already be long; tests should use
  variables, comments, and assertion text for readability rather than embedding
  long descriptive names into temporary filesystem paths.
- Adversarial filename fixtures that claim near-maximum filename coverage must
  budget for the full absolute self-test temp root plus every generated
  subdirectory before choosing the file-name component length. A component that
  is individually legal can still break the test when the total path exceeds the
  active Windows path budget.
- Column, grid, or viewport fixtures that depend on row counts must derive those
  counts from the same visible UI mode they assert. Horizontal scrollbars,
  display mode, DPI, and density can change the pane's visible height; a
  no-scroll probe row count is not enough evidence for a scrolled fixture.
- Commands UI tests that call `FolderWindow::SetFolderPath(...)` must wait for
  the requested pane path to settle before dispatching commands that consume
  the current pane roots.
- Commands UI tests that clear pane visibility/filter state and then navigate
  to a fixture folder must not let the reset refresh an old pane path race the
  fixture enumeration. Clear state without refreshing, or perform the reset
  after the fixture path is active and wait for the resulting enumeration/debug
  state before asserting empty, filtered, or watermark UI.
- Empty-folder and filter-watermark self-tests must also clear pane operation
  alerts, explicit empty-state messages, and background watermarks before
  asserting the built-in empty-folder placeholder. These surfaces intentionally
  outlive some folder navigation paths, so visibility/filter reset alone is not
  enough isolation evidence.
- Commands GUI self-tests that validate pointer clicks, keyboard focus, or
  live DxUi traversal must run in a foreground-capable desktop session. Hidden
  or background launches are not valid evidence for those cases because they
  can change focus routing, pane activation, and pointer targeting.
- Commands GUI self-tests that validate product routes based on
  `WindowFromPoint` or screen-coordinate hit testing must first bring the
  target dialog/window to the foreground/top of the z-order and verify the
  target point belongs to that window's root before sending synthetic wheel or
  pointer messages.
- Commands GUI self-tests that mutate persisted or runtime settings, such as
  Connection Manager profiles, must snapshot and restore the touched settings
  before the case returns. Full-order keyboard or focus traversal coverage must
  seed deterministic data and must not depend on rows, profiles, or selection
  state left behind by earlier cases.
- Commands GUI self-tests that continue after activating a focused file must
  account for any transient top-level viewer, prompt, or tool window opened by
  that activation. The test must wait for or close that transient UI and
  re-establish the intended pane focus before validating the next keyboard mode
  or focus-sensitive command.
- Commands GUI self-tests that drive modal prompts from a worker thread must
  keep UI Automation reads bounded and must always leave a path that closes the
  prompt. A blocked UIA provider call must degrade to a failed assertion with a
  trace message, not leave the owner window disabled and the suite waiting
  forever.
- Commands self-tests that validate Win32 clipboard formats must treat the
  clipboard as desktop-global state: do not run clipboard GUI cases in
  parallel, use bounded `OpenClipboard` retries for reads, report the open owner
  and registered-format availability on failure, and allocate DROPFILES read
  buffers with room for the trailing null terminator.
- Commands GUI self-tests that dispatch pane commands through `WM_COMMAND` and
  depend on the focused pane must wait until the intended folder-view HWND is
  the actual focused folder view after transient UI cleanup. A single
  `SetActivePane`/`SetFocus` call is not durable evidence when queued focus
  restore or focus-loss messages may still run.
- Commands GUI self-tests that reactivate focus-exit modes such as integrated
  Quick Search must also reassert stable folder-view focus after the command
  activates and before sending synthetic `WM_CHAR` input. `WM_KILLFOCUS` is a
  valid product exit path for those modes, so active debug state alone is not
  enough evidence that subsequent typed input is isolated from queued focus
  changes.
- Commands GUI self-tests that call Win32 `SetFocus` must assert the resulting
  focused HWND, such as by polling `GetFocus()`, rather than comparing
  `SetFocus`'s return value with the target. `SetFocus` returns the previous
  focus HWND on success, so direct equality with the target is valid only when
  the target was already focused and is order-dependent test evidence.
- Self-tests that validate an external process by reading a marker file must
  wait for the expected marker content, not only for file existence. Process
  launch and shell redirection can create the file before the first line is
  flushed, especially in full-suite order.
- Self-tests that validate file-action macro expansion through an external
  marker must match the documented macro context. Macros expanded inside an
  `arguments` template are Windows command-line arguments and therefore include
  the quoting/escaping needed for safe argument boundaries.
- Preferences page-specific self-tests should navigate to their setup page by
  named category selection or the documented visible root-row helper. They must
  not encode `End` plus a fixed number of `Up` keys as a shortcut for a page,
  because expanded child nodes and visible root-order changes make that setup
  order-dependent. Dedicated category-tree tests remain responsible for
  validating real keyboard Home/End/Up/Down behavior.
- UI self-tests that assert no unrelated host repaint churn during a stimulus
  must first wait for the unrelated host's debug render count to settle after
  navigation and page setup. Pending renders from setup must not be charged to
  the action under test, but the stimulus window may still require zero extra
  unrelated renders, resize churn, or resize failures after that settled
  baseline. The settle wait must require a meaningful idle window with multiple
  unchanged samples; a single unchanged poll is not enough evidence that
  pending setup paints have drained.
- Repaint-churn baselines must be captured after foregrounding, hit-test
  preparation, `UpdateWindow`, and other setup that can pump messages or
  invalidate UI hosts. Those setup paints must not be measured as part of the
  stimulus being validated.
- UI self-tests that assert render-count budgets for the stimulated host itself
  must also measure from a settled pre-action baseline to a settled post-action
  snapshot. Render-count deltas captured around a single message-pump pass are
  order-dependent evidence because queued setup, hover, or paint work can be
  charged to the wrong action.
- UI self-tests that drive a control through debug-only helpers and then assert
  repaint evidence must make the stimulus deterministic: establish a scrollable
  starting position, verify the behavioral state change such as scroll offset,
  and force/update the target window's paint when the helper invalidates without
  going through the real message dispatch path.
- UI self-tests that assert hover-dependent state after pointer clicks must keep
  the real cursor position and synthetic mouse messages aligned. Sending
  `WM_MOUSEMOVE` without moving the cursor can be overwritten by normal message
  pumping and turn hover assertions into current-desktop artifacts.
- UI self-tests that wait on debug snapshots must format failure diagnostics
  from the post-wait snapshot, not from a snapshot captured before the wait.
  This is required for focus/scroll assertions where the observed state can
  change while the wait helper is polling.
- UI self-test debug snapshots that expose retained focus targets must report
  `None` when no DxUi control currently owns retained focus. Family-specific or
  optional controls must be pointer-guarded before comparison so a null retained
  focus cannot be misreported as a control that is absent on the active page.
- UI self-tests that expose both native focus ownership and retained DxUi focus
  targets must assert those states independently. Native focus on another child
  HWND, such as a dialog category tree, does not require a DxUi host's retained
  focus target to become `None` when the retained control still belongs to that
  host.
- UI self-tests that send native edit messages to a DxUi text-input target must
  wait for both the retained DxUi text-field focus and a valid focused Win32
  input target after any deferred rebuild, search refresh, or list rebind. A
  matching retained focus target alone is not enough evidence that native
  `WM_CHAR` / edit-message routing is ready to receive input; tests that probe
  compatibility `EM_SETSEL` / `EM_REPLACESEL` paths must drive the native DxUi
  host target and retained text session directly.
- UI self-test debug snapshots that expose a single `focusTarget` field and use
  it to drive keyboard input must report the active native-focus owner, not a
  retained-only DxUi control from a host that no longer owns native focus.
- GUI Tab-traversal self-tests must include retained focus, native focus
  ownership, and modifier-state evidence in failure diagnostics when those
  states can affect routing. Ordered traversal scripts must stop at the first
  failed step instead of continuing into cascaded focus traces that obscure the
  original failure.
- The shared Commands message pump must be cooperative and bounded. It must not
  drain the queue indefinitely because active DxUi hover, cursor, timer, or
  paint traffic can otherwise turn a bounded wait into a long-running queue
  drain and hide the real test step duration.

## Artifact Contract

Self-test artifacts must preserve enough detail to explain why a run passed, failed, or skipped:
- `results.json` records final case status and reason,
- `trace.txt` records supporting diagnostic context,
- `perf_metrics.jsonl` must be preserved when the scenario emits performance metrics,
- archived copies under `Specs/TestRuns/` must keep those files intact.

Archive-to-repo discovery must only accept a candidate repository root when it
contains all of:
- `RedSalamander.sln`,
- `Specs/TestRuns`,
- a `.git` directory or `.git` worktree file.

The parent walk from the executable directory must remain explicitly bounded.

## DxUi Popup Validation

For command selftests that validate migrated app chrome using DxUi popup menus:
- owned `DxUi_ContextMenu` popup windows are an authoritative “menu opened” signal,
- popup dismissal may be validated by observing that owned popup window close, not only by `GUI_INMENUMODE`,
- tests must not rely on `GUI_INMENUMODE` alone once the validated surface routes through the shared DxUi popup path instead of a native Win32 menu loop.

## DxUi Pointer Validation

For command selftests that validate real pointer interaction on DxUi controls:
- a target inside a scrollable DxUi host must be scrolled into the viewport
  before the selftest sends mouse messages,
- debug host/client rectangles consumed by pointer tests must be in the target
  HWND's client coordinate space after any `ScrollPanel` offset has been
  applied,
- a control's semantic `IsVisible()` state is not enough to prove it is
  onscreen and clickable when the control lives in scrollable content.

## NavigationView DxUi Text-Host Validation

For command selftests that validate NavigationView address-bar edit mode or full-path popup edit mode:
- `NavigationViewDebugSnapshot` is the authoritative contract for edit visibility, focus target, current edit text, selection range, `currentEditUsesNativeTextInput`, `currentEditHostHwnd`, the active backend-neutral `currentEditInputHwnd`, caret screen rectangle validity/coordinates, and active composition start/end state,
- tests must not enumerate descendant native `Edit` / `RICHEDIT50W` windows to prove edit mode, because NavigationView no longer exposes a visible native edit child and no longer installs NavigationView policy on a hidden child edit surface.
- path-region keyboard activation selftests must verify active native IME composition owns Enter/Escape/Tab before NavigationView submit/cancel/tab policy runs.
- edit-suggest keyboard-routing selftests must verify active native IME composition owns Up/Down before NavigationView suggestion-selection policy runs, so candidate/navigation keys do not change edit-suggest selection while composing.
- address-bar edit clipboard coverage must verify the focused DxUi host remains active and pane-level Select All, Copy, and Paste commands mutate/copy the edit text instead of falling through to FolderView command handling.
- invalid-path validation coverage must assert edit mode remains active and `NavigationViewDebugSnapshot::currentEditHelpText` exposes the rejected path text through the retained DxUi `TextField` HelpText contract.
- NavigationView pointer and region-keyboard selftests must not inherit the
  navigation-bar visibility left by earlier commands. Before using hit-test
  rectangles or `DebugFocusNavigationViewRegion(...)`, they must snapshot the
  previous navigation-bar visibility, reveal the target pane's NavigationView,
  wait until its child HWND is visible, and restore the original visibility
  before returning.

## Preview Pane and ViewerVLC Validation

For command selftests that validate embedded preview behavior:
- `FolderWindow::PreviewPaneDebugSnapshot` is the authoritative contract for preview activity, source/host panes, selected Folder/Preview tab, hosted plugin ID, tab-strip HWND and tab hit rectangles, Preview close visibility, tab-strip visible/pending tooltip text, content HWND, embedded viewer HWND, and embedded viewer instance identity.
- Preview tests must verify that embedded viewers do not take keyboard focus from the source pane, that focus changes reuse the current embedded viewer instance and HWND when the resolved plugin is unchanged, that stale content is cleared before the new file is reported as rendered, and that preview close/replacement persists changed plugin configuration.
- Preview source-contract tests must cover menu-bearing embedded viewers so Preview-appropriate menu options remain reachable from right-click context menus without requiring a visible embedded menubar or standalone filename/header chrome, while standalone-only actions and shortcut labels stay out of the embedded menu.
- Preview responsiveness coverage must verify that rapid same-plugin and cross-plugin preview switches do not wait for slow media/player teardown on the UI thread.
- Preview source-contract tests must cover embedded media audio-only paths so VLC visualizers cannot create top-level player or visualization windows outside the Preview host.
- When saved viewer associations are empty or resolve only the default text viewer, preview tests must cover the built-in embedded viewer defaults before `builtin/viewer-text` fallback.

For command selftests that validate `builtin/viewer-vlc`:
- `WndMsg::ViewerVlcDebugGetSnapshot` is the authoritative contract for HUD state, volume/mute state, snapshot dimensions, Fluent icon glyph usage, and filled-button HUD styling.
- Wheel-seek coverage must exercise both normal viewer surfaces and libVLC-owned child video windows so embedded preview and standalone playback keep the same wheel behavior.
- Slow-stop coverage must use `WndMsg::kViewerVlcDebugSetStopDelay` to ensure media-to-media preview switches, including video-to-audio transitions, avoid slow player retirement, media-to-image preview switches stay responsive while VLC player cleanup continues in the background, and VLC debug snapshots must assert that the embedded video surface remains a child of the preview viewer after same-plugin media navigation.

## Search-Specific Coverage

Search coverage follows the same contract:
- local, fallback, indexed, and service search cases stay declared,
- ReFS validation stays declared even on machines without ReFS,
- a machine without a fixed ReFS volume records `skipped` with a reason instead of silently omitting the case.
- Direct SQLite cutover, prefilter, and injected SQLite failure cases require a
  readable live NTFS journal cursor for the requested root. When that cursor is
  unavailable, the case must remain declared and record `skipped` with that
  precondition as the reason.
- SQLite service no-wait query tests must distinguish a healthy service that
  answers by live scan from host fallback. `DEGRADED_NO_INDEX` means the service
  stayed healthy but could not use a ready/current database; only
  `SERVICE_UNAVAILABLE` means the host switched away from the service
  transport.
- Service tests whose assertion is not about the default ProgramData database
  must use isolated foreground-service storage paths so pre-existing local
  SQLite state cannot change the result.

## Source Organization

Project-owned self-test implementation must live in `.cpp` source files, with optional `.h` declarations when cross-translation-unit declarations are needed.

Project code must not introduce `.inl` self-test implementation files.

The `--commands-selftest` suite may continue to include family source files from `RedSalamander/SelfTest/Commands/Commands.SelfTest.cpp`, but those included family files must still be `.cpp` files rather than `.inl` fragments.

`Commands.SelfTest.Preferences.cpp` may include Preferences chunk files for
navigation, but each `Commands.SelfTest.Preferences.*.cpp` chunk must own a
complete anonymous namespace wrapper. A chunk must not close a namespace opened
by the coordinator or leave a namespace open for the next family include.
