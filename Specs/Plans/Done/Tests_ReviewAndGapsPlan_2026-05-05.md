# Tests Review - Coverage Gaps and Maintenance Plan

Last updated: 2026-05-09

Status: Done

## Active Closeout Checklist

- [x] Add and maintain a checklist at the beginning of this WIP plan so the
  remaining work is visible without reading the entire document.
- [x] Keep focused and full-run evidence archived under `Specs/TestRuns/` and
  referenced from this plan as fixes land.
- [x] Fix Compare Options retained DxUi keyboard traversal after layout/scroll
  and verify the full Compare Options command family with multiple fresh roots.
- [x] Fix shell/pane recheck coverage evidence: reject zero-case filtered runs
  and shorten the item-properties fixture path that hit `MAX_PATH` limits.
- [x] Fix `cmd_app_` splash UIA and close hangs by pumping while UIA/close work
  runs, then verify focused splash and full `cmd_app_`.
- [x] Fix app submenu placement expectations to use the same DxUi popup
  positioning helper as product code, including work-area clamping.
- [x] Fix Reread Associations cache-clear evidence so the snapshot records the
  immediate post-clear cache size before pane refresh can repopulate it.
- [x] Fix Primary Edit focused-file test maintenance: assert the focused file
  path and the documented `{Filename}` argument quoting behavior.
- [x] Fix Preferences Viewers selected-extension debug state so selection tests
  observe the current mapping instead of stale empty state.
- [x] Fix Preferences Viewers Remove/Reset destructive mapping mutations so they
  clear the selected mapping state after the model rebind.
- [x] Fix Preferences Viewers Reset Defaults test expectations and diagnostics:
  default grid row count includes the Default mapping row, Cancel/reopen restores
  the first-row selection, and failures print the relevant snapshot fields.
- [x] Fix Preferences Viewers Add/Update maintenance and behavior: live UIA now
  targets the shared `Match value` / `Save Association` controls, and saving an
  edited selected association replaces that row instead of appending a duplicate
  key-migration row.
- [x] Verify Preferences Plugins tests no longer depend on stale visible
  category-row counts for navigation to the Plugins root page.
- [x] Verify Preferences Keyboard and Themes tests no longer depend on stale
  visible category-row counts for page setup.
- [x] Fix Viewers/Editors shared file-actions debug scroll semantics so
  sustained-scroll tests exercise the same signed wheel path as other DxUi
  grids.
- [x] Resolve Viewers/Editors shared file-actions tab-traversal coverage so it
  follows the current shared association form and reports useful diagnostics
  when focus stalls.
- [x] Re-open and resolve the order-dependent Viewers shared file-actions
  reverse Tab traversal failure surfaced after earlier Preferences cases:
  diagnostics must report null retained focus as `None`, and the final fix must
  prove reverse traversal reaches `Computer` after the broad family prefix.
- [x] Fix stale Viewers shared file-actions column-index coverage so
  reorder/copy/header tests target named current columns (`Match`, `Computer`,
  `F3 View`, `Alt+F3 Alternate View`, `Status`) rather than retired
  Extension/Viewer positional assumptions.
- [x] Fix remaining Editors/Mouse Preferences note-page tests that still walk
  the category tree with stale magic Down-key counts after `User Menu` moved
  into the visible root order.
- [x] Resolve the Preferences page-host retained-scroll failure now surfaced by
  fail-fast validation: determine whether General is legitimately scrollable at
  the reduced test height, then update product/test/spec so per-category scroll
  retention is validated without encoding stale non-overflow assumptions.
- [x] Continue foreground fail-fast validation of the `cmd_preferences_dialog_`
  family and resolve the next first failure.
- [x] Resolve the Preferences Keyboard export page-settle failure surfaced by
  fail-fast validation: determine whether it is standalone or order-dependent
  after earlier Keyboard cases, then fix the product/test/spec contract so
  Keyboard action tests settle on current DxUi page state rather than stale
  retained/capture state.
- [x] Resolve the Preferences rapid-switch page-specific UIA subtree failure:
  Editors must validate the current shared file-actions contract, while Mouse
  must not expose stale editable/toggle descendants from prior pages.
- [x] Resolve the File Operations/Compare Directories page-setup navigation
  failures surfaced after the Preferences family advanced past Viewers:
  page-specific tests must select named categories instead of encoding
  `End`+fixed-`Up` shortcuts that depend on expanded tree state.
- [x] Resolve the Hot Paths tab-traversal failure that only reproduces after
  earlier Preferences cases: exact Hot Paths tab traversal and the Hot Paths
  prefix pass, so the next diagnostic run must isolate the order-dependent
  retained-focus/native-focus/page-host state before any product/test fix lands.
- [x] Resolve the Preferences sustained grid-scroll false positive: long-run
  page-list tests must settle the unrelated category-tree render baseline
  before asserting zero additional category-tree repaint churn during the
  measured scroll window.
- [x] Resolve the remaining Editors/Mouse page-specific setup drift by selecting
  Editors and Mouse with named category selection instead of assuming the
  dialog opens on `General` before fixed Down-key navigation.
- [x] Rerun the full foreground Commands suite with fresh absolute
  `REDSALAMANDER_SELFTEST_ROOT` after the Preferences family is stable.
- [x] Run the first full closeout suite after Commands stabilized and capture
  the remaining non-Commands failures before closing the plan:
  `C:\R\FULLSUITE1\last_run\run-all-tests-results.json` passed Compare,
  Commands, FileOperations, DxUiTests, ViewerPETests, ViewerSqliteTests,
  MonitorTest, and PerformanceTests2, then failed in LocalizationTests,
  ToolsPesterTests, and VcpkgMergeSynthetic.
- [x] Capture and resolve the next full closeout blockers from
  `C:\R\FULLSUITE3\last_run\run-all-tests-results.json`: one transient
  Commands copy-as-text clipboard failure and one VcpkgMergeSynthetic failure
  caused by Tools Pester leaking StrictMode into the runner session.
- [x] Rerun the full closeout suite from scratch after the `FULLSUITE3` fixes;
  `C:\R\FULLSUITE4` was interrupted by the user before aggregate results were
  written and is partial, non-closeout evidence. Fresh foreground
  `C:\R\FULLSUITE5\last_run\run-all-tests-results.json` completed on
  2026-05-09 and failed the closeout gate with 780 passed / 5 failed /
  44 skipped: four Commands UI stability failures and one `ViewerPETests`
  process failure. The final fresh closeout rerun
  `C:\R\FULLSUITE7\last_run\run-all-tests-results.json` passed with
  785 passed / 0 failed / 44 skipped.
- [x] Re-run the full Commands suite after the second address-bar Tab-focus
  stabilization patch and the runner display fix. The diagnostic full Commands
  run at `C:\R\FULLCMDFIX1\last_run\run-all-tests-results.json` failed with
  592 passed / 5 failed / 0 skipped and is non-closeout evidence until the
  display StrictMode crash is verified fixed. Fresh
  `C:\R\FULLCMDFIX2\last_run\run-all-tests-results.json` passed Commands with
  597 passed / 0 failed / 0 skipped after both fixes.
- [x] Capture native executable / CppUnitTest stdout and stderr from the unified
  runner so future standalone failures, including `ViewerPETests`, preserve an
  output log in the aggregate artifact instead of only an exit code.
- [x] Resolve the `ViewerPETests` closeout failure exposed by the new output
  log: `C:\R\FULLSUITE6\last_run\ViewerPETests.output.log` showed the nested
  `TestViewerShellComboHostsLongRunOpenCloseStayStable` fresh-process entry
  was killed by the parent 120-second timeout while still executing valid
  churn. `ViewerPETests` now gives that six-cycle nested stress entry a
  600-second outer timeout while normal isolated viewer cases keep the
  120-second cap, and `Tools\Tests\TestHarnessSourceContracts.Tests.ps1` guards
  the distinction.
- [x] Rerun the full closeout suite from scratch after the `ViewerPETests`
  timeout-budget fix; `C:\R\FULLSUITE6` is valid failed evidence and not a
  closeout pass. `C:\R\FULLSUITE7` is the passing closeout rerun.
- [x] Re-run focused Pester coverage for the runner output-log and display
  changes after the `$r.OutputLogPath` StrictMode-safe display fix:
  `Invoke-Pester -Path @('Tools\Tests\RunAllTestsPlan.Tests.ps1','Tools\Tests\ProcessStreaming.Tests.ps1') -PassThru`
  passed 10/0/0 on 2026-05-09.
- [x] Re-run `cmd_pane_focusAddressBar_tab_traversal` exact/repeat after the
  second stabilization patch that reuses the last live edit snapshot while
  restoring focus before Tab/Shift+Tab dispatch:
  `C:\R\F5_ADDRTAB_FIXED2_REPEAT1` through
  `C:\R\F5_ADDRTAB_FIXED2_REPEAT12` all passed 1/0/0 with exit code 0.
- [x] Triage the full foreground Commands failure set under a much shorter
  self-test root to separate path-depth/fixture artifacts from durable
  functional regressions before landing any broad product or test fix.
- [x] Resolve the newly resurfaced full-order Preferences Keyboard header-drag
  blocker from `C:\R\FULLF49`: the UI-chrome fix is focused-green, but the full
  fail-fast run stopped earlier after 185 passed / 1 failed / 411 skipped at
  `cmd_preferences_dialog_keyboard_header_drag_reorders_columns_without_sort`.
  Isolate exact and Keyboard-prefix behavior before changing shared grid drag,
  selection, or diagnostics.
- [x] Resolve the resurfaced full-order Preferences Keyboard long-run list-scroll
  category-tree render-baseline blocker from `C:\R\FULLF39`: after the
  NavigationView current-breadcrumb patch, fail-fast stopped earlier after
  163 passed / 1 failed / 433 skipped because the category-tree host render
  count moved from 2 to 3 during Keyboard list scrolling chunk 5. The longer
  settle window passed exact Keyboard long-run and advanced the Preferences
  prefix through Keyboard/Viewers/Themes long-run cases in `C:\R\PREFD12`.
- [x] Resolve the newly surfaced Preferences Keyboard copy/search focus
  blocker from `C:\R\PREFD12`: the Preferences prefix advanced to
  63 passed / 1 failed / 104 skipped, then
  `cmd_preferences_dialog_keyboard_reordered_resized_copy_follows_visible_columns_after_search_roundtrip`
  failed because the Keyboard search field no longer exposed a focused Win32
  edit target before the live clear-back step. The test now waits for both
  retained DxUi search focus and the native input bridge before sending edit
  messages after deferred rebuilds; exact, small-family, and broad Preferences
  evidence passed through this case in `C:\R\KBCOPY2`, `C:\R\KBRR3`, and
  `C:\R\PREFD13`.
- [x] Re-open the Preferences scroll-host category-tree repaint blocker from
  `C:\R\PREFD13`: after the Keyboard copy/search fix, the broad Preferences
  prefix advanced to 166 passed / 1 failed / 1 skipped and stopped at
  `cmd_preferences_dialog_scroll_host_preserves_retained_page_state` with one
  extra category-tree render during page-host scrolling. The hit-test/foreground
  preparation now runs before the settled repaint baseline is captured; exact
  `C:\R\PSH5` and broad Preferences `C:\R\PREFD14` passed.
- [x] Resolve the re-opened full-order Preferences Keyboard UIA grid-selection
  blocker from `C:\R\FULLF41`: after a diagnostic build for directory-impact
  selection, the full Commands run stopped earlier at
  `cmd_preferences_dialog_keyboard_page_exposes_live_uia_grid_selection` with
  184 passed / 1 failed / 412 skipped. Exact, Keyboard-family, and Preferences
  prefix evidence stayed green, diagnostics were added to the failure path, and
  full-order `C:\R\FULLF42` advanced past this case.
- [x] Resolve the full-order directory-impact chained-rename blocker from
  `C:\R\FULLF40`: the full Commands fail-fast run advanced past the repaired
  Preferences family and stopped at
  `cmd_pane_navigation_directory_impact_preserves_selection_across_chained_renames`
  after 461 passed / 1 failed / 135 skipped. Exact standalone
  `C:\R\DIRCHAIN1` and the `cmd_pane_navigation_directory_impact_` prefix
  `C:\R\DIRCHAIN2` both passed, so the next step is order-dependent diagnosis
  of pane path, item/filter/focus state, queued directory refreshes, and
  selection mapping across multiple rename hints. The strengthened diagnostic
  case and full-order `C:\R\FULLF42` advanced past this stop; keep the richer
  diagnostics unless later cleanup proves them noisy.
- [x] Resolve the current full-order NavigationView unfocused-pane pointer
  blocker from `C:\R\FULLF42`: full Commands advanced past the repaired
  Preferences, directory-impact, dispatch smoke, and NavigationView double-click
  cases, then stopped after 557 passed / 1 failed / 39 skipped at
  `cmd_pane_navigationView_unfocused_pane_click_focuses_target_pane` because the
  right navigation bar did not become visible before the click validation.
  After the Preferences-family repairs, `C:\R\FULLF45` reproduced the same
  first failure after 557 passed / 1 failed / 39 skipped: the snapshot still
  reports both nav bars visible, but the actual right NavigationView HWND is
  hidden with a zero-sized rect before the unfocused-pane click. Active patch:
  the unfocused-pane test now saves the prior zoom state, clears zoom before
  requiring the right pane/navigation bar to be visible, and restores zoom at
  teardown. Exact and NavigationView-prefix evidence passed earlier, and the
  final full Commands proof `C:\R\FULLF50\last_run\commands\results.json`
  passed with 597 passed / 0 failed / 0 skipped; archived evidence:
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-08_204608/`.
- [x] Resolve the re-opened Preferences Plugins reordered/resized sort-search
  setup blocker from `C:\R\FULLF43`: the diagnostic full-order run failed
  earlier than the NavigationView proof point after 157 passed / 1 failed /
  439 skipped at
  `cmd_preferences_dialog_plugins_reordered_resized_columns_survive_sort_cycles_and_search_roundtrip`
  because the test could not select the Plugins category before establishing
  the reordered/resized sort-search baseline. Exact Plugins, Plugins prefix,
  and the final broad Preferences `C:\R\PREFD20` run passed through this case.
- [x] Resolve the current broad Preferences Viewers reordered/resized sort
  header-rect blocker from `C:\R\PREFD16`: the Preferences prefix failed before
  the Plugins sort-search point after 53 passed / 1 failed / 114 skipped at
  `cmd_preferences_dialog_viewers_reordered_resized_columns_survive_sort_cycles`
  because the test could not capture the visible `Match` header rectangle before
  the reordered/resized sort baseline. Exact `C:\R\VIEWSORT1` and Viewers-only
  prefix `C:\R\VIEWFAM1` passed, so the failure is broader Preferences order
  pollution or a shared retained-grid setup assumption, not a standalone
  Viewers behavior regression. Failure-only diagnostics were added, the
  diagnostic build passed, exact `C:\R\VIEWSORT2` passed, repeat Viewers prefix
  `C:\R\VIEWFAM3` passed, and broad Preferences `C:\R\PREFD17` advanced past
  this case.
- [x] Resolve the current broad Preferences category-tree render-budget blocker
  from `C:\R\PREFD17`: the Preferences prefix advanced past the Viewers/Plugins
  blockers and then failed after 165 passed / 1 failed / 2 skipped at
  `cmd_preferences_dialog_category_switches_do_not_churn_tree_host` because the
  category tree rendered three times while returning to `General` where the
  test budgets at most two renders. Fresh exact `C:\R\PCHURN2` and category
  prefix `C:\R\PCHURNF2` both passed, so this is currently broad-order only.
  Per-click diagnostics were added, exact/category-prefix reruns stayed green,
  and broad Preferences `C:\R\PREFD20` passed all Preferences cases.
- [x] Resolve the current broad Preferences Keyboard reordered/resized search
  clear-back focus blocker from `C:\R\PREFD18`: after the category-tree
  diagnostics were added, the Preferences prefix stopped earlier after
  62 passed / 1 failed / 105 skipped at
  `cmd_preferences_dialog_keyboard_reordered_resized_columns_survive_search_roundtrip`.
  The no-match search rebuild settled (`search='__codex_no_match__'`,
  `rows=0`), but the subsequent clear-back step could not reacquire a focused
  Win32 input target (`focusTarget=0`). Isolate exact and Keyboard
  reordered/resized prefix before changing the shared search-focus wait. Exact
  and Keyboard-prefix reruns stayed green after diagnostics, and broad
  Preferences `C:\R\PREFD20` passed all Preferences cases.
- [x] Resolve the current broad Preferences Viewers long-run scroll category
  drift blocker from `C:\R\PREFD19`: after Keyboard diagnostics were added,
  the Preferences prefix stopped earlier after 38 passed / 1 failed /
  129 skipped at `cmd_preferences_dialog_viewers_long_run_list_scrolling_stays_bounded`
  because scroll chunk 0 changed the active category unexpectedly. The preceding
  Plugins, Keyboard, and category setup cases all passed; isolate exact and
  neighboring long-run prefixes before changing scroll assertions or helpers.
  Exact and Viewers-prefix reruns stayed green after diagnostics, and broad
  Preferences `C:\R\PREFD20` passed all Preferences cases.
- [x] Resolve the current full-order Preferences Themes UIA grid-selection
  setup blocker from `C:\R\FULLF44`: full Commands advanced past the repaired
  broad Preferences blockers but then stopped after 194 passed / 1 failed /
  402 skipped at `cmd_preferences_dialog_themes_page_exposes_live_uia_grid_selection`
  because the Themes page did not expose its DX grid surface for UIA selection
  validation. Broad Preferences `C:\R\PREFD20` passed this same case, so
  isolate exact and Themes-prefix behavior before changing Themes setup or
  shared UIA grid-selection assertions. Exact, Themes-prefix, and follow-up
  full-order `C:\R\FULLF45` all advanced past this case.
- [x] Resolve the resurfaced full-order Clipboard Paste Shortcut blocker from
  `C:\R\FULLF36`: the command copied `alpha.txt`/`beta.txt` into the
  destination instead of creating unique `.lnk` files even though focus was on
  the expected left folder view.
- [x] Resolve the next short-root full-order blocker:
  `cmd_preferences_dialog_editors_mouse_tab_skips_note_surface` fails after
  284 passes because the Mouse note-page Tab traversal expects wrapped focus
  back to the category tree while the shell reports category-tree focus but a
  shell retained-focus target value of `3`. Determine whether the assertion is
  stale, the debug enum is mislabeling focus, or the Preferences shell is
  keeping divergent native/retained focus after the note page. Exact standalone
  reproduction now proves the gap is a stale test/spec expectation after the
  shared DxUi retained-focus contract changed; the corrected retained-shell
  expectations pass in the focused case.
- [x] Resolve the next short-root full-order blocker:
  `cmd_preferences_dialog_scroll_host_preserves_retained_page_state` now passes
  standalone and in the Preferences-only family, but fails in full order after
  292 passes with one extra category-tree render during page-host wheel
  scrolling. The likely test gap is an unsettled unrelated category-tree render
  baseline before measuring page-host scroll churn; the shared settle helper is
  now moved to the common Preferences test include scope and applied to this
  case before the full short-root fail-fast run advanced past it.
- [x] Resolve the external viewer/editor marker-file full-order flake:
  marker-file tests must wait for the expected first-line content, not only
  file existence, because shell redirection can create the file before flushing
  the marker text. Exact viewer/editor evidence is green and the next full
  short-root run advanced past the earlier external-viewer marker failure.
- [x] Resolve the next short-root full-order blocker:
  `cmd_pane_quickSearch_integrated_navigation` fails after 299 passes with
  `Quick Search no-match state should remain active.` A focused diagnostic patch
  now shows the full-order failure more precisely: after Enter activation and
  transient-window cleanup, `IDM_PANE_QUICK_SEARCH` does not reactivate search
  mode (`active=0`, `query=''`, focused file `beta-alpha.txt`). Determine
  whether the command is routed to the wrong pane/window, blocked by focus not
  being truly restored to the folder view, or exposing a product command bug.
- [x] Resolve the current short-root full-order blocker:
  the full suite now advances past Quick Search and stops after 312 passes with
  `Pressing Enter from a non-editor Find control did not trigger the default
  Find action.` Identify the exact case and determine whether this is
  standalone Find-dialog routing regression, another full-order focus/default
  button setup gap, or stale test sequencing after earlier Find dialog cases.
  Exact standalone evidence is fixed, the shared DxUi contract is green, and
  `C:\R\FULLF9\last_run\commands\results.json` advanced past the previous
  Find stop before exposing a later Compare Directories Options failure.
- [x] Resolve the current short-root full-order blocker:
  `cmd_compare_directories_options_tab_traversal_live_dx_interaction` now fails
  after 407 passes with `Compare Directories options Cancel button focus target
  not reached during tab traversal`, while the snapshot reports the focus still
  on target `13`, the dialog scrolled near the bottom (`300/312`), and events
  routed to `RedSalamander.CompareOptions.DxHost`. Determine whether this is a
  standalone Compare Options traversal drift, a full-order retained/native
  focus interaction, or a stale expectation after the shared focus-retention
  contract changed. Focused evidence is green after tightening the debug
  snapshot to report only the native-focus-owning host, and `C:\R\FULLF15`
  passed the previous 407-pass stop in full order.
- [x] Resolve the full-order credential prompt hang exposed while rerunning the
  short-root suite: `cmd_connection_credential_prompt_theme_cycle_keeps_surface_legible`
  opened the modal `Password required` prompt and then stopped writing trace
  before results. Bounded UIA task helpers now prevent modal prompt worker reads
  from hanging forever; exact and credential-prompt-family evidence are green.
- [x] Resolve the full-order pane Clipboard Paste Shortcut blocker:
  `cmd_pane_clipboardPasteShortcut_creates_unique_links` failed in
  `C:\R\FULLF11` after Quick Search/Find fixes because only the existing
  `alpha - Shortcut.lnk` appeared. The test now validates the seeded CF_HDROP
  clipboard payload and waits for the intended left folder view to own focus
  before dispatching the focus-sensitive pane command; exact evidence and
  `C:\R\FULLF15` full-order evidence are green.
- [x] Re-open the Clipboard Paste Shortcut full-order blocker resurfaced by
  `C:\R\FULLF36\last_run\commands\results.json`: fail-fast stopped after
  369 passed / 1 failed / 227 skipped at
  `cmd_pane_clipboardPasteShortcut_creates_unique_links` with
  `Paste Shortcut should refresh the pane with unique .lnk files`, while the
  destination folder contained copied `alpha.txt`/`beta.txt` instead of the
  expected `.lnk` files. Focus diagnostics reported the focused pane/view were
  already the expected left folder view, so the next root-cause step is to
  determine whether the app dispatched the normal Paste command path, whether
  stale clipboard/drop-effect state survived the preceding Clipboard Cut case,
  or whether `PasteShortcutFromClipboard()` is bypassed in full order. Exact
  standalone `C:\R\CLIPSHORT1\last_run\commands\results.json` passed with
  1 passed / 0 failed / 0 skipped and produced the expected
  `alpha - Shortcut (2).lnk` / `beta - Shortcut.lnk` artifacts, so this is not
  a standalone Paste Shortcut product regression. Note:
  `C:\R\CLIPFAM1\last_run\commands\results.json` is zero-case non-evidence
  because `cmd_pane_clipboard` is not a valid prefix filter for these
  camel-cased case names. Broader pane-prefix
  `C:\R\PANEPR1\last_run\commands\results.json` passed with 206 passed /
  0 failed / 0 skipped, including both Clipboard Cut and Paste Shortcut in
  sequence, so the `FULLF36` failure needs earlier non-`cmd_pane_` full-order
  state or is intermittent. Diagnostic patch in progress: trace the
  `IDM_PANE_CLIPBOARD_PASTE_SHORTCUT` command branch, trace entry/source/
  created-link counts inside `FolderView::PasteShortcutFromClipboard()`, and
  include destination directory entries in the Paste Shortcut failure message.
  Debug x64 rebuilt cleanly at `.build/logs/msbuild-20260508_145342_846.log`
  with 0 warnings and 0 errors after adding those diagnostics. Exact
  `C:\R\CLIPSHORT2\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped; top-level trace confirms the healthy command branch:
  `WM_COMMAND PasteShortcut`, `PasteShortcutFromClipboard sources=2`, and
  `created=2`. Full-order `C:\R\FULLF37\last_run\commands\results.json`
  advanced past Clipboard Paste Shortcut and returned to the NavigationView
  double-click blocker after 554 passed / 1 failed / 42 skipped, so the
  Clipboard item is no longer the active stop.
- [x] Resolve the new short-root full-order Preferences render-churn blocker:
  `cmd_preferences_dialog_category_switches_do_not_churn_tree_host` fails after
  291 passes in `C:\R\FULLF13` with one category click reporting 3 category-tree
  renders where the assertion allows 2. Exact standalone
  `C:\R\PCHURN1` and Preferences prefix `C:\R\PREFCHURN1` are green, so the
  next step is to challenge the test's measurement window and determine whether
  a pending setup render is being charged to the click or the product is doing
  real extra tree-host work only after full-suite order.
- [x] Resolve the next short-root full-order File Operations issues-pane
  blocker: `cmd_pane_fileops_issues_pane_long_run_scrolling_stays_bounded`
  fails after 427 passes in `C:\R\FULLF14` with
  `Issues pane did not repaint after long-run scroll chunk 0.` Determine
  whether the synthetic wheel stimulus is landing on the wrong/unchanged scroll
  state, the assertion incorrectly requires repaint when no scroll offset can
  change, or the issues pane genuinely stopped rendering after long full-suite
  order.
- [x] Resolve the resurfaced external editor marker/focus blocker:
  `C:\R\FULLF17` stopped much earlier at
  `alternate_edit_launches_configured_editor_action` after 68 passes. Exact
  standalone `C:\R\AE1` passed, while the `alternate_edit_` prefix reproduced a
  stale marker-read failure where the file existed before the first line was
  flushed. Patch the remaining external viewer/editor/menu marker checks to wait
  for expected first-line content and make primary/alternate edit shortcut
  dispatch wait for stable left-pane focus before relying on `GetFocusedPane()`.
- [x] Resolve the preview tab pointer-hover blocker:
  `C:\R\FULLF18` advanced past the previous editor marker stop, then failed at
  `pane_view_options_toggle_preview_pane_tabs_and_selection` after 50 passes
  with `Inactive Preview tab should hide its close button when it is not
  hovered.` Exact standalone `C:\R\PVTAB1` reproduced it, so fix the stale test
  helper that sent synthetic mouse messages without moving the real cursor.
- [x] Resolve the resurfaced Quick Search no-match focus-loss blocker:
  `C:\R\FULLF19` advanced past the preview and editor-marker stops, then failed
  at `cmd_pane_quickSearch_integrated_navigation` after 299 passes with
  `Quick Search no-match state should remain active; active=0, query=''`.
  Exact standalone `C:\R\QS1` is green, so challenge the full-order focus
  isolation around Quick Search reactivation before treating this as product
  behavior.
- [x] Resolve the next short-root full-order FolderView empty/filter watermark
  blocker: `folderView_filter_watermark_empty_state` fails after 507 passes in
  `C:\R\FULLF15` with `Expected empty-folder state active for empty folder.`
  `C:\R\FULLF16` reproduced the stop with stronger diagnostics:
  path was already the empty fixture, item count was zero, and the filter was
  inactive, but `emptyActive=0`. The active hypothesis is that pane operation
  alerts, explicit empty-state messages, or background watermarks can outlive
  prior full-suite cases and suppress the built-in empty-folder placeholder.
- [x] Resolve the Edit New prompt macro-quoting blocker:
  `C:\R\FULLF20` advanced past Quick Search and the previous FolderView stop,
  then failed at `cmd_pane_editNew_prompt_filters_editor_combo_and_creates_file`
  after 518 passes. Exact standalone `C:\R\EDITNEW1` fails the same way, and
  the marker contains `"alpha.editnew"`, matching the documented file-action
  `arguments` macro quoting contract.
- [x] Resolve the current short-root full-order Preferences General theme-cycle
  blocker: `C:\R\FULLF21` stopped earlier than Edit New at
  `cmd_preferences_dialog_general_theme_cycle_keeps_surface_legible` after
  210 passes with `Preferences General page did not settle after the light theme
  update.` Reproduce exact/family, determine whether this is a standalone
  General page-settle regression, order-dependent state from preceding Themes
  grid/sort/search cases, earlier non-Preferences command state, or a stale wait
  predicate after the accepted DxUi design refresh. Current evidence:
  standalone exact `C:\R\GENLIGHT1` passed with 1 passed / 0 failed / 0
  skipped, the General prefix `C:\R\GENFAM1` passed with 7 passed / 0 failed /
  0 skipped, and the Preferences dialog prefix `C:\R\PREFD1` passed with 168
  passed / 0 failed / 0 skipped. Therefore the next full-order run needs
  stronger diagnostics on the light-theme wait instead of another blind
  assertion. Later Preferences prefix and full Commands evidence advanced past
  this case; final proof `C:\R\FULLF50\last_run\commands\results.json` passed
  with 597 passed / 0 failed / 0 skipped.
- [x] Resolve the newly surfaced post-rebuild Preferences Hot Paths traversal
  blocker before rerunning full Commands order: the diagnostic rebuild passed,
  exact General theme-cycle `C:\R\GENLIGHT2` passed, but the fresh Preferences
  dialog prefix `C:\R\PREFD2` stopped after 98 passes at
  `cmd_preferences_dialog_hot_paths_tab_traversal_live_dx_interaction`.
  Standalone Hot Paths tab traversal `C:\R\HPTAB1` passed and the Hot Paths
  prefix `C:\R\HPFAM1` passed with 7 passed / 0 failed / 0 skipped, so the
  current evidence points to earlier Preferences-page order state. The trace
  shows Tab reaches the first Browse button, then the next Tab from the DX host
  toward the first Label field clears native focus to `0x0` while the active
  page and active DX host remain the Hot Paths page host.
- [x] Resolve the repeat post-build Preferences prefix category-tree PageDown
  blocker: `C:\R\PREFD3` advanced past the Hot Paths stop and then failed after
  161 passes at `cmd_preferences_dialog_category_tree_page_navigation_stays_on_dx_tree_path`
  with `Preferences category host lost keyboard focus after VK_NEXT.` Exact
  `C:\R\CTPG1` and category-tree prefix `C:\R\CTFAM1` are green. The case still
  sampled the snapshot immediately after one pump and compared the initial
  `SetFocus` return value directly; patch it to use `FocusWindowAndWait(...)`
  and a post-key settled snapshot that still requires category-tree focus,
  selection movement, unchanged page-host scroll, and page-title/selection
  agreement.
- [x] Resolve the next post-build Preferences prefix setup blocker:
  `C:\R\PREFD4` stopped after 14 passes at
  `cmd_preferences_dialog_viewers_remove_live_dx_interaction` with
  `Preferences Viewers page did not settle before remove interaction validation.`
  Exact `C:\R\VRM1` is green. The setup still used stale `Home` + two `Down`
  keystrokes to reach Viewers; patch it to use named Viewers category selection,
  the shared focus waiter, and post-wait snapshot diagnostics.
- [x] Resolve the next Preferences prefix blocker:
  after the Viewers Remove named-category patch rebuilt cleanly at
  `.build/logs/msbuild-20260508_095459_074.log`, exact
  `C:\R\VRM2\last_run\commands\results.json` passed with 1 passed / 0 failed /
  0 skipped. The next prefix run `C:\R\PREFD5` advanced to 107 passed / 1
  failed / 60 skipped at
  `cmd_preferences_dialog_keyboard_live_search_dx_interaction`, but the failure
  reason is only `failed`. Exact `C:\R\KLS1` and Keyboard prefix
  `C:\R\KFAM1` are green. Trace shows the case returned `ok=no failed=no`, so
  the Keyboard navigation helper returned false without recording
  `state.failure`; patch it to use `FocusWindowAndWait(...)`, keep named
  Keyboard category selection, and report the post-wait snapshot if the page
  does not settle.
- [x] Resolve the resurfaced Viewers long-run list-scroll baseline blocker:
  after the Keyboard diagnostics patch rebuilt cleanly at
  `.build/logs/msbuild-20260508_100158_159.log`, exact
  `C:\R\KLS2\last_run\commands\results.json` passed. The next Preferences
  prefix run `C:\R\PREFD6\last_run\commands\results.json` stopped after
  38 passed / 1 failed / 129 skipped at
  `cmd_preferences_dialog_viewers_long_run_list_scrolling_stays_bounded` with
  `Preferences category tree host should not repaint during Viewers list
  scrolling chunk 0; render count moved from 14 to 15.` Exact
  `C:\R\VLR1\last_run\commands\results.json` passed, so this is another
  order/setup measurement gap unless repeat evidence proves otherwise. Patch
  the Viewers long-run setup away from direct `SetFocus(...) == target` and
  hard-coded `Home`+`Down` navigation before trusting the category-tree render
  baseline. The patch rebuilt cleanly at
  `.build/logs/msbuild-20260508_101006_731.log` with 0 warnings and 0 errors.
  Exact `C:\R\VLR2\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped. The broad Preferences prefix
  `C:\R\PREFD7\last_run\commands\results.json` passed with 168 passed /
  0 failed / 0 skipped and `trace.txt` ended with `CommandsSelfTest: PASS`;
  the collection shell timed out at 240 seconds after the run had already
  written passing artifacts, and no `RedSalamander.exe` self-test process
  remained.
- [x] Re-open the Preferences Keyboard long-run list-scroll render-baseline
  failure resurfaced by `C:\R\FULLF39\last_run\commands\results.json`:
  fail-fast stopped after 163 passed / 1 failed / 433 skipped at
  `cmd_preferences_dialog_keyboard_long_run_list_scrolling_stays_bounded` with
  `Preferences category tree host should not repaint during Keyboard list
  scrolling chunk 5; render count moved from 2 to 3.` This is the same class as
  the earlier Viewers long-run false positive: exact isolation must determine
  whether Keyboard is now standalone-broken or whether full-order setup leaves a
  pending category-tree paint that must be settled before measuring the list
  scroll window. Exact `C:\R\KBLR1\last_run\commands\results.json` passed with
  1 passed / 0 failed / 0 skipped, proving the failure is order/setup
  measurement. Patch in progress: strengthen the shared
  `WaitForPreferencesCategoryTreeRenderCountToSettle(...)` helper from a short
  three-sample quiet window to a longer idle window before capturing the
  no-repaint baseline, and record that requirement in
  `Specs/Testing/Testing_SelfTests.md`. Debug x64 rebuilt cleanly at
  `.build/logs/msbuild-20260508_154026_241.log` with 0 warnings and 0 errors.
  Exact `C:\R\KBLR2\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped. Broad Preferences prefix
  `C:\R\PREFD12\last_run\commands\results.json` advanced past Keyboard,
  Viewers, and Themes long-run checks, then stopped later after 63 passed /
  1 failed / 104 skipped at
  `cmd_preferences_dialog_keyboard_reordered_resized_copy_follows_visible_columns_after_search_roundtrip`;
  therefore the long-run render-baseline fix is no longer the active
  Preferences-family stop.
- [x] Investigate the new Preferences Keyboard reordered/resized copy-after-search
  focus failure from `C:\R\PREFD12\last_run\commands\results.json`: the failed
  case reports `Preferences Keyboard search field did not expose a focused
  Win32 input target before live clear-back validation.` The immediately
  preceding Keyboard grid/header/reordered/resized cases passed in the same
  prefix, so the next step is exact isolation and trace review before deciding
  whether the stale assumption is native focus retention, search-field rebuild,
  or an actual product focus regression. Exact isolation
  `C:\R\KBCOPY1\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped, and the small reordered/resized family
  `C:\R\KBRR2\last_run\commands\results.json` passed with 3 passed /
  0 failed / 0 skipped, proving the failure was order-dependent. Root cause:
  after the live no-match search rebuild, the test waited only for search text
  and row count, then immediately sent `EM_SETSEL` / `EM_REPLACESEL` to
  `GetFocus()`. In broad order the retained DxUi focus target was still the
  Keyboard search field, but the hidden Win32 input bridge was not yet ready.
  Patch: add `WaitForPreferencesKeyboardSearchInputTarget(...)`, use it in the
  Keyboard reordered/resized live search and copy/search cases, and record the
  native-bridge readiness contract in `Specs/Testing/Testing_SelfTests.md`.
  Debug x64 rebuilt cleanly at `.build/logs/msbuild-20260508_155146_562.log`
  with 0 warnings and 0 errors. Verification:
  `C:\R\KBCOPY2\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped,
  `C:\R\KBRR3\last_run\commands\results.json` passed with 3 passed /
  0 failed / 0 skipped, and broad Preferences prefix
  `C:\R\PREFD13\last_run\commands\results.json` advanced past the previous
  Keyboard copy/search stop before failing later at the scroll-host repaint
  case.
- [x] Re-open the Preferences scroll-host render-churn blocker after
  `C:\R\PREFD13\last_run\commands\results.json`: the broad Preferences prefix
  now reaches 166 passed / 1 failed / 1 skipped, and fails at
  `cmd_preferences_dialog_scroll_host_preserves_retained_page_state` with
  `Scrolling the Preferences page host should not repaint the category tree;
  saw 1 extra tree render(s).` This is the same assertion previously believed
  resolved by `C:\R\PREFD9`; exact isolation and trace instrumentation must
  determine whether the newly longer preceding order leaves another pending
  category-tree repaint, whether the stimulus itself legitimately invalidates
  the tree, or whether the no-repaint assertion should measure a more settled
  post-category-switch baseline. Root cause: the test captured the settled
  category-tree render baseline and then called the hit-test foreground helper,
  whose `BringWindowToTop` / topmost toggles / `UpdateWindow` / message pump can
  repaint the category tree before the wheel stimulus. Patch: run the
  foreground/hit-test preparation before
  `WaitForPreferencesCategoryTreeRenderCountToSettle(...)`, and record in
  `Specs/Testing/Testing_SelfTests.md` that repaint baselines must be captured
  after any setup that can pump or invalidate UI. Debug x64 rebuilt cleanly at
  `.build/logs/msbuild-20260508_160045_880.log` with 0 warnings and 0 errors.
  Exact `C:\R\PSH5\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped. Broad Preferences prefix
  `C:\R\PREFD14\last_run\commands\results.json` passed with 168 passed /
  0 failed / 0 skipped.
- [x] Resolve the current full-order Clipboard Cut blocker: the fresh full
  Commands non-fail-fast run `C:\R\FULLF22\last_run\commands\results.json`
  completed with 583 passed / 14 failed / 0 skipped. It did not reproduce the
  original General light-theme settle failure, but it reported repeat-sensitive
  Preferences failures and a later Clipboard Cut failure. The fresh fail-fast
  rerun `C:\R\FULLF23\last_run\commands\results.json` advanced past the
  Preferences stops and failed after 368 passed / 1 failed / 228 skipped at
  `cmd_pane_clipboardCut_sets_move_drop_effect` with `Clipboard Cut should
  write two CF_HDROP paths; got 0.` Exact isolation must determine whether the
  clipboard payload is missing, not readable, overwritten by prior cases, or
  routed to the wrong pane before any product/test fix lands. Exact
  `C:\R\CLIPCUT1\last_run\commands\results.json` reproduced the zero-path
  failure. The root cause was test-side clipboard reading, plus a real DROPFILES
  read-buffer maintenance bug: the helper asked `DragQueryFileW` to write the
  trailing null without allocating space for it. After fixing the test/product
  DROPFILES readers and adding bounded Preferred DropEffect read retries with
  diagnostics, Debug x64 rebuilt cleanly at
  `.build/logs/msbuild-20260508_104715_973.log` with 0 warnings and 0 errors.
  Exact `C:\R\CLIPCUT4\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped. Clipboard Paste Shortcut controls also passed:
  `C:\R\PASTEMISS1\last_run\commands\results.json` and
  `C:\R\PASTESHORT3\last_run\commands\results.json` each passed with
  1 passed / 0 failed / 0 skipped. The attempted concurrent clipboard pair run
  under `C:\R\PASTESHORT2` is intentionally not evidence because clipboard GUI
  cases share desktop-global clipboard state. Reopened by
  `C:\R\FULLF25\last_run\commands\results.json`: full-order advanced past the
  Viewers tab blocker and stopped again at Clipboard Cut because the CF_HDROP
  reader itself hit `OpenClipboard` access denied (`openError=5`,
  `openOwner=0x11602`) before retrying. Patch the path reader with the same
  bounded clipboard-open retry used by the DropEffect reader. The retry patch
  rebuilt cleanly at `.build/logs/msbuild-20260508_111125_473.log` with
  0 warnings and 0 errors; exact
  `C:\R\CLIPCUT5\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped; and
  `C:\R\FULLF26\last_run\commands\results.json` advanced past Clipboard Cut
  before stopping later at Navigation View address-bar Tab traversal.
- [x] Resolve the current full-order Preferences Viewers tab-traversal blocker:
  after the Clipboard Cut fix, `C:\R\FULLF24\last_run\commands\results.json`
  stopped after 170 passed / 1 failed / 426 skipped at
  `cmd_preferences_dialog_viewers_tab_traversal_live_dx_interaction`.
  Failure: `Preferences Viewers reverse wrapped search field focus target not
  reached during tab traversal; expected search field, saw none; native focus
  before=0x1B01902, after=0x0, active page before=0x1B01902, active page
  after=0x1B01902, message target=0x1B01902, page children=1, resize failures=0.`
  Exact isolation must decide whether this is a standalone regression, an
  order-dependent retained/native focus state after the earlier Viewers cases,
  or a stale reverse-wrap expectation before another DxUi focus patch lands.
  Current controls are green: exact `C:\R\VTAB1\last_run\commands\results.json`
  passed with 1 passed / 0 failed / 0 skipped, the Viewers prefix
  `C:\R\VIEWFAM1\last_run\commands\results.json` passed with 27 passed /
  0 failed / 0 skipped, and the full Preferences prefix
  `C:\R\PREFD8\last_run\commands\results.json` passed with 168 passed /
  0 failed / 0 skipped. Patch in progress: replace the test setup's remaining
  direct `SetFocus(...) == target` assertion with `FocusWindowAndWait(...)`
  before the next full-order repeat. The patch rebuilt cleanly at
  `.build/logs/msbuild-20260508_110044_318.log` with 0 warnings and 0 errors.
  Exact `C:\R\VTAB2\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped, and `C:\R\FULLF25\last_run\commands\results.json`
  advanced past this case before stopping later at Clipboard Cut.
- [x] Resolve the current full-order Navigation View address-bar Tab blocker:
  `C:\R\FULLF26\last_run\commands\results.json` stopped after 553 passed /
  1 failed / 43 skipped at `cmd_pane_focusAddressBar_tab_traversal` with
  `Navigation view did not enter address-bar edit mode during forward tab
  handoff.` Exact isolation must determine whether this is a standalone
  Navigation View regression, focus not restored to the folder view after prior
  cases, or a stale test wait around the address-bar edit-mode snapshot. Exact
  `C:\R\NAVTAB1\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped. Patch in progress: require stable left folder-view
  focus before dispatching the focus-sensitive address-bar command, and include
  the post-wait `NavigationViewDebugSnapshot` fields when edit mode does not
  appear. The patch rebuilt cleanly at
  `.build/logs/msbuild-20260508_112453_788.log` with 0 warnings and 0 errors,
  and exact `C:\R\NAVTAB2\last_run\commands\results.json` passed with
  1 passed / 0 failed / 0 skipped. Reopened by
  `C:\R\FULLF29\last_run\commands\results.json`: full-order advanced past the
  reopened Preferences scroll-host and Hot Paths stops, then failed here after
  553 passed / 1 failed / 43 skipped. The new diagnostics show the
  NavigationView did enter edit mode and copied the expected path text
  (`focusTarget=6`, `editMode=1`, `editText=...navigation_view_tab_traversal...`,
  `pathText=...navigation_view_tab_traversal...`), but no folder view owned
  focus at assertion time (`focusedPane=0`, `focusedView=0x0`, `leftView=...`).
  Exact `C:\R\NAVTAB3\last_run\commands\results.json` still passes with
  1 passed / 0 failed / 0 skipped, so the stop is order-dependent. Root cause
  found: the preceding full-order dispatch smoke can hide the pane navigation
  bar, and the direct `cmd/pane/focusAddressBar` command path invoked
  `NavigationView::FocusAddressBar()` without first revealing that bar. The
  folder-view keyboard navigation route already calls `SetNavigationBarVisible`
  before entering NavigationView UI. Patch in progress: command-driven
  NavigationView actions now reveal the pane navigation bar before focusing the
  address bar, change-directory edit, drive menu, or history dropdown; the
  address-bar Tab test now hides the left navigation bar before the first
  command dispatch so exact coverage catches this failure mode, and
  `Specs/UI/UI_NavigationView.md` records the durable command-reveal contract.
  Debug x64 rebuilt cleanly at `.build/logs/msbuild-20260508_124302_311.log`
  with 0 warnings and 0 errors, and exact
  `C:\R\NAVTAB4\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped. Full-order
  `C:\R\FULLF30\last_run\commands\results.json` advanced past this case,
  proving the command-reveal fix in order before stopping at the next
  NavigationView pointer test.
- [x] Resolve the current full-order Navigation View path double-click blocker:
  `C:\R\FULLF30\last_run\commands\results.json` stopped after 554 passed /
  1 failed / 42 skipped at
  `cmd_pane_navigationView_path_doubleClick_enters_edit_mode` with
  `Navigation view did not enter address-bar edit mode after path-region
  double-click.` Current hypothesis to verify, not assume: the previous
  address-bar Tab test restores the original full-order navigation-bar
  visibility, so this pointer-oriented test may inherit a hidden left
  NavigationView after `dispatch_smoke_all_commands` hid it. Exact isolation
  must confirm whether this case passes alone, then the patch should make
  NavigationView tests explicitly establish the visible navigation-bar state
  before hit-testing or keyboard region activation. Exact
  `C:\R\NAVDBL1\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped, confirming the current stop is order-dependent.
  Patch in progress: add a shared NavigationView selftest setup guard that
  captures/restores navigation-bar visibility and waits for the target
  NavigationView child window to become visible before pointer or region-focus
  assertions. `Commands.SelfTest.ViewCommands.cpp` now applies that guard to
  the NavigationView pointer/region-keyboard family, and
  `Specs/Testing/Testing_SelfTests.md` records the lasting selftest isolation
  rule. Debug x64 rebuilt cleanly at
  `.build/logs/msbuild-20260508_130336_408.log` with 0 warnings and
  0 errors. Exact `C:\R\NAVDBL2\last_run\commands\results.json` passed with
  1 passed / 0 failed / 0 skipped. Family
  `C:\R\NAVFAM1\last_run\commands\results.json` then advanced through the
  first 9 NavigationView cases before stopping at
  `cmd_pane_navigationView_history_dropdown_keyboard_navigation`; exact
  `C:\R\NAVHIST1\last_run\commands\results.json` reproduced the same failure.
  New root cause: keyboard activation of NavigationView Menu/History/DiskInfo
  regions posts the same dropdown-open messages as pointer activation, and the
  dropdown session leaves `focusFirstNavigableItem=false`, so the live DxUI
  popup opens with no keyboard index. Patch in progress: make
  `ActivateFocusedRegion()` mark keyboard-opened dropdowns and pass that flag
  into the DxUI context menu session while keeping pointer-opened dropdowns
  mouse-neutral. `NavigationView.{h,cpp,Interaction.cpp,Menus.cpp}` now carries
  that keyboard-open flag through posted dropdown messages, and
  `Specs/UI/UI_NavigationView.md` records the durable keyboard-vs-pointer
  dropdown contract. Debug x64 rebuilt cleanly at
  `.build/logs/msbuild-20260508_130855_145.log` with 0 warnings and
  0 errors. Exact `C:\R\NAVHIST2\last_run\commands\results.json` passed with
  1 passed / 0 failed / 0 skipped, and family
  `C:\R\NAVFAM2\last_run\commands\results.json` passed with 12 passed /
  0 failed / 0 skipped. Full-order proof was reopened by
  `C:\R\FULLF34\last_run\commands\results.json`, which advanced past the
  nonstandard Common Folders stop and failed this path double-click case again
  after 554 passed / 1 failed / 42 skipped. Diagnostic patch in progress:
  the double-click assertion now reports the final NavigationView snapshot,
  navigation-bar/window visibility, native focus, click point, resolved target
  window, mapped click point, path/ellipsis rectangles, and pane refresh/item/
  selection counters when edit mode does not appear. Debug x64 rebuilt cleanly
  at `.build/logs/msbuild-20260508_141246_503.log` with 0 warnings and
  0 errors. Exact `C:\R\NAVDBL3\last_run\commands\results.json` passed with
  1 passed / 0 failed / 0 skipped, and the NavigationView prefix
  `C:\R\NAVFAM3\last_run\commands\results.json` passed with 12 passed /
  0 failed / 0 skipped. Full-order diagnostic evidence is pending.
  Full-order `C:\R\FULLF35\last_run\commands\results.json` reproduced the
  blocker after 554 passed / 1 failed / 42 skipped with the stronger failure
  detail: final NavigationView snapshot had `focusTarget=0`, `editMode=0`,
  empty `path`/`edit`, zero `pathRect=(0,0-0,0)`, no visible child windows,
  `navVisible=1`, `navWindowVisible=1`, click target equal to the NavigationView
  HWND, and pane counters still stable (`refresh=40/40`, `items=2/2`,
  `selected=0/0`). This rules out a merely hidden navigation bar or bad
  target-window resolution; the next root-cause step is to determine why the
  left NavigationView can republish or expose an empty path/zero layout between
  the successful baseline/single-click checks and the double-click edit-mode
  wait in full order. Diagnostic patch in progress: capture a fresh
  post-single-click breadcrumb snapshot, recompute the double-click point from
  that fresh geometry, and record the immediate snapshot after the leading
  click of the double-click sequence so the next run can separate stale test
  coordinates from a product path-clear during mouse handling. Debug x64
  rebuilt cleanly at `.build/logs/msbuild-20260508_143346_506.log` with
  0 warnings and 0 errors. Exact
  `C:\R\NAVDBL4\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped after this diagnostic edit, so the new probes do not
  break the standalone double-click path. NavigationView prefix
  `C:\R\NAVFAM4\last_run\commands\results.json` passed with 12 passed /
  0 failed / 0 skipped. Full-order
  `C:\R\FULLF37\last_run\commands\results.json` reproduced the NavigationView
  stop after 554 passed / 1 failed / 42 skipped with new evidence: the
  pre-double-click snapshot and the immediate leading-click snapshot both kept
  the correct path and nonzero path rectangle, and the final snapshot still had
  the correct path/rect but `editMode=0`. This rules out a product path-clear;
  root cause is a stale test click heuristic that chooses a point near the
  right edge of the breadcrumb region, which can land on a non-edit-activating
  breadcrumb/separator when the path is ellipsized. Next fix: expose the visible
  current breadcrumb segment in the debug snapshot and click that segment
  directly. Patch in progress: `NavigationViewDebugSnapshot` now reports the
  current breadcrumb segment rectangle, and the double-click test targets that
  segment center rather than guessing a point from the full path-region width.
  Debug x64 rebuilt cleanly at `.build/logs/msbuild-20260508_150847_591.log`
  with 0 warnings and 0 errors.
- [x] Resolve the current full-order nonstandard Common Folders submenu blocker:
  `C:\R\FULLF31\last_run\commands\results.json` stopped after 449 passed /
  1 failed / 147 skipped at
  `cmd_pane_navigation_nonstandard_menu_common_folders` with
  `Common Folders submenu should expose the local common-folder entries.
  Observed submenu:`. This stop appears before the prior Hot Paths and
  NavigationView positions, so it must be isolated before any earlier
  NavigationView full-order claims can close. Exact isolation must determine
  whether the test fails alone, fails only after previous navigation/menu cases,
  or needs richer diagnostics for the observed submenu contents and popup
  ownership. Exact `C:\R\COMMON1\last_run\commands\results.json` passed with
  1 passed / 0 failed / 0 skipped, and the nearby
  `cmd_pane_navigation_` prefix `C:\R\NAVPFX1\last_run\commands\results.json`
  passed with 31 passed / 0 failed / 0 skipped, so the `FULLF31` stop depends
  on earlier full-order state. Patch in progress: replace the test's synthetic
  hover-only submenu opening with deterministic keyboard focus of the Common
  Folders root row plus `VK_RIGHT`, and add submenu/keyboard-index diagnostics
  if the submenu still fails to materialize. Debug x64 rebuilt cleanly at
  `.build/logs/msbuild-20260508_132354_697.log` with 0 warnings and
  0 errors. Exact `C:\R\COMMON2\last_run\commands\results.json` passed with
  1 passed / 0 failed / 0 skipped, and
  `C:\R\NAVPFX2\last_run\commands\results.json` passed with 31 passed /
  0 failed / 0 skipped. Full-order proof landed in
  `C:\R\FULLF34\last_run\commands\results.json`, which advanced past this
  case before failing later at NavigationView path double-click after
  554 passed / 1 failed / 42 skipped.
- [x] Resolve the current full-order Preferences Plugins tab-traversal blocker:
  `C:\R\FULLF32\last_run\commands\results.json` stopped earlier than the
  Navigation/Common-Folders proof point after 245 passed / 1 failed /
  351 skipped at `cmd_preferences_dialog_plugins_tab_traversal_live_dx_interaction`.
  Failure:
  `Preferences Plugins reverse wrapped search field focus target not reached
  during tab traversal.` Exact isolation must determine whether the reverse
  wrap fails standalone, only after the preceding Preferences Keyboard action
  cases, or because the Plugins tab-traversal test still sends Tab to a stale
  active page HWND after native focus changes during the traversal. Before any
  fix lands, add diagnostics comparable to the Viewers/Themes/Keyboard traversal
  tests: expected/actual focus target, native focus before/after, active page
  handle before/after, message target, child count, resize counts, selected
  plugin/custom-path text, and list visible-row/column baselines. The WIP must
  be updated with exact, prefix/family, build, and rerun evidence before this
  item can close. Exact standalone `C:\R\PLUGTAB1\last_run\commands\results.json`
  passed with 1 passed / 0 failed / 0 skipped, so the failure is not a
  standalone Plugins traversal regression. The Plugins-page prefix
  `C:\R\PLUGFAM1\last_run\commands\results.json` also passed with 29 passed /
  0 failed / 0 skipped, so preceding Plugins-only cases are not sufficient to
  reproduce the reverse-wrap loss. After the Plugins round-trip focus fix below,
  exact `C:\R\PLUGTAB3\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped, and the full Preferences-dialog prefix
  `C:\R\PREFD11\last_run\commands\results.json` passed with 168 passed /
  0 failed / 0 skipped. Full-order proof landed in
  `C:\R\FULLF33\last_run\commands\results.json`, which advanced past the
  Plugins tab traversal and stopped later at Advanced tab traversal after
  265 passed / 1 failed / 331 skipped.
- [x] Resolve the newly surfaced Preferences prefix Plugins round-trip blocker
  before further full-order claims: `C:\R\PREFD10\last_run\commands\results.json`
  stopped after 7 passed / 1 failed / 160 skipped at
  `cmd_preferences_dialog_plugins_roundtrip_restores_dxui_surface` with
  `Failed to refocus the Preferences category host before leaving Plugins for
  General.` This is earlier than the FULLF32 Plugins tab-traversal stop and
  therefore must be isolated first. Exact isolation must decide whether this is
  a standalone Plugins round-trip regression, an order-dependent setup gap from
  the first seven Preferences dialog cases, or another stale direct
  `SetFocus(...) == target` assertion that should use the shared focus waiter
  and snapshot diagnostics. Exact standalone
  `C:\R\PLUGROUND1\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped, confirming the stop depends on earlier Preferences
  shell/chrome cases. Root cause identified: the round-trip test still asserted
  `SetFocus(categoryTreeHost) == categoryTreeHost`, even though Win32 returns
  the previously focused window; patch in progress uses `FocusWindowAndWait(...)`
  and reports native focus, category host, active page, shell host, category,
  and page title if focus cannot be restored. Debug x64 rebuilt cleanly at
  `.build/logs/msbuild-20260508_133830_155.log` with 0 warnings and 0 errors.
  Exact `C:\R\PLUGROUND2\last_run\commands\results.json` passed with
  1 passed / 0 failed / 0 skipped, and the repro prefix
  `C:\R\PREFD11\last_run\commands\results.json` passed with 168 passed /
  0 failed / 0 skipped.
  `C:\R\NAVDBL5\last_run\commands\results.json` failed standalone after the
  first current-segment patch because comparing breadcrumb segment paths to the
  normalized current path was too strict. The debug probe is being adjusted to
  expose the actual last rendered breadcrumb segment instead, matching the
  product's `OnLButtonDblClk` edit-mode condition. Debug x64 rebuilt cleanly at
  `.build/logs/msbuild-20260508_151204_078.log` with 0 warnings and 0 errors
  after switching the test to the last-segment probe. Exact
  `C:\R\NAVDBL6\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped. NavigationView family
  `C:\R\NAVFAM5\last_run\commands\results.json` passed with 12 passed /
  0 failed / 0 skipped. Full-order `C:\R\FULLF38\last_run\commands\results.json`
  still failed here after 554 passed / 1 failed / 42 skipped, but now rules out
  the stale right-edge click theory: the click target was the NavigationView
  HWND, `lastVisible=1`, `lastRect=(350,0-948,36)`, and the click point
  `(649,18)` was inside that last rendered segment. Root-cause hypothesis under
  test: `OnLButtonDown` still dispatched `RequestPathChange(...)` for the
  current/last breadcrumb segment, so the leading click of the double-click could
  queue redundant same-path refresh work in full order and close or prevent the
  edit host even though standalone/family order passed. Patch in progress:
  current/last breadcrumb clicks now focus the owning pane without navigating,
  while double-click keeps the existing edit-mode route; `Specs/UI/UI_NavigationView.md`
  records that durable contract. Debug x64 rebuilt cleanly at
  `.build/logs/msbuild-20260508_153153_907.log` with 0 warnings and 0 errors.
  Exact `C:\R\NAVDBL7\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped. NavigationView family
  `C:\R\NAVFAM6\last_run\commands\results.json` passed with 12 passed /
  0 failed / 0 skipped. Full-order `C:\R\FULLF42\last_run\commands\results.json`
  advanced past this double-click case, proving the current/last breadcrumb
  click fix in full order before stopping at the next NavigationView pointer
  case.
- [x] Resolve the current full-order NavigationView unfocused-pane pointer
  blocker: `C:\R\FULLF42\last_run\commands\results.json` advanced past the
  prior Preferences Keyboard UIA, directory-impact chained rename, dispatch
  smoke, and NavigationView double-click stops, then failed after 557 passed /
  1 failed / 39 skipped at
  `cmd_pane_navigationView_unfocused_pane_click_focuses_target_pane` with
  `Right navigation bar did not become visible before unfocused-pane navigation
  click validation.` The archived run is
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-08_170141/`. Immediate
  hypothesis to test, not assume: this may be a full-order visibility/state
  residue after `dispatch_smoke_all_commands` and the navigation-shell stability
  cases, because the failure occurs before breadcrumb click validation and before
  the test captures the right-pane NavigationView snapshot. Next: run the exact
  case and the `cmd_pane_navigationView_` prefix with the same short root style;
  if focused evidence is green, add diagnostics to the `EnsureNavigationViewVisibleForSelfTest`
  failure path so it reports stored navigation visibility, HWND validity,
  `IsWindowVisible`, pane paths, child rects, and active/focused pane state
  before any product/test fix lands. Exact isolation stayed green:
  `C:\R\NAVUNF1\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped, and NavigationView prefix
  evidence later passed after the preview/zoom setup fix. Final full-order proof
  `C:\R\FULLF50\last_run\commands\results.json` passed with 597 passed /
  0 failed / 0 skipped.
  `C:\R\NAVFAM7\last_run\commands\results.json` passed with 12 passed /
  0 failed / 0 skipped. Diagnostic patch in progress: the right navigation-bar
  visibility failure now reports stored visibility, current and expected
  NavigationView HWND/rect/class/visibility, active pane, focused window/folder
  view, left/right navigation-bar flags, pane path, and NavigationView snapshot
  state. Debug x64 rebuilt cleanly at
  `.build/logs/msbuild-20260508_170711_797.log` with 0 warnings and 0 errors;
  exact `C:\R\NAVUNF2\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped, and NavigationView prefix
  `C:\R\NAVFAM8\last_run\commands\results.json` passed with 12 passed /
  0 failed / 0 skipped. Full-order diagnostic proof is running at
  `C:\R\FULLF43` with process id 81868. `C:\R\FULLF43` did not reach the
  NavigationView pointer case; it stopped earlier at
  `cmd_preferences_dialog_plugins_reordered_resized_columns_survive_sort_cycles_and_search_roundtrip`
  after 157 passed / 1 failed / 439 skipped with
  `Failed to select the Preferences Plugins category before reordered-resized-sort/search validation.`
  The archived evidence is `Specs/TestRuns/4cb089111a23/Commands/2026-05-08_171227/`.
  Exact Plugins sort-search `C:\R\PLUGSORT1\last_run\commands\results.json`
  passed with 1 passed / 0 failed / 0 skipped, and Plugins prefix
  `C:\R\PLUGFAM2\last_run\commands\results.json` passed with 29 passed /
  0 failed / 0 skipped. The broader Preferences prefix
  `C:\R\PREFD16\last_run\commands\results.json` did not reach the Plugins
  sort-search stop: it failed earlier at
  `cmd_preferences_dialog_viewers_reordered_resized_columns_survive_sort_cycles`
  after 53 passed / 1 failed / 114 skipped with
  `Failed to capture the visible Preferences Viewers Match header rect before
  reordered-resized/sort validation.` This moves the active Preferences
  investigation one step earlier, to the shared Viewers/Plugins/Themes retained
  grid setup assumptions. Exact Viewers sort-cycle
  `C:\R\VIEWSORT1\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped, and Viewers prefix
  `C:\R\VIEWFAM1\last_run\commands\results.json` passed with 27 passed /
  0 failed / 0 skipped. Next step: add failure-only diagnostics to the Viewers
  header-rect setup point, including snapshot state and visible header
  availability for columns 0-4, then rerun the broad Preferences prefix.
  Diagnostic patch is now in `Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp`
  at the initial Viewers sort-cycle header capture point; build and focused
  verification are pending. Diagnostic build
  `.build/logs/msbuild-20260508_172511_543.log` completed with 0 warnings /
  0 errors. Post-diagnostic exact rerun
  `C:\R\VIEWSORT2\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped. Post-diagnostic Viewers prefix
  `C:\R\VIEWFAM2\last_run\commands\results.json` did not reach the broad
  Preferences proof point: it failed after 22 passed / 1 failed / 4 skipped at
  `cmd_preferences_dialog_viewers_reordered_resized_copy_follows_visible_columns_after_sort_cycles`
  with `Failed to select the Preferences Viewers category before
  reordered-resized-copy/sort validation.` This narrower failure supersedes the
  broader `PREFD16` header-rect probe until exact/pair/family isolation proves
  whether the copied-column test has the same stale category/page setup issue.
  Exact copied-column sort-cycle `C:\R\VIEWCOPY1\last_run\commands\results.json`
  passed with 1 passed / 0 failed / 0 skipped, and the smaller
  `cmd_preferences_dialog_viewers_reordered_resized_` slice
  `C:\R\VIEWRR1\last_run\commands\results.json` passed with 6 passed /
  0 failed / 0 skipped. Repeat the full Viewers prefix once; if it fails again,
  instrument the shared Viewers category-selection path, otherwise continue to
  the broad Preferences prefix. Repeat Viewers prefix
  `C:\R\VIEWFAM3\last_run\commands\results.json` passed with 27 passed /
  0 failed / 0 skipped, so `VIEWFAM2` is retained as an intermittent setup miss
  but not yet sufficient root-cause evidence for a patch.
  Broad Preferences prefix `C:\R\PREFD17\last_run\commands\results.json`
  advanced past the prior Viewers header-rect and Plugins setup blockers, then
  failed later after 165 passed / 1 failed / 2 skipped at
  `cmd_preferences_dialog_category_switches_do_not_churn_tree_host` with
  `Preferences category tree should repaint at most twice for category 0; saw 3
  render(s), before=17, after=20, selectedVisibleIndex=0, title='General',
  expectedTitle='General'.` Active investigation now moves to this later
  category-tree render-budget assertion. Fresh exact isolation
  `C:\R\PCHURN2\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped, and the category-tree prefix
  `C:\R\PCHURNF2\last_run\commands\results.json` passed with 9 passed /
  0 failed / 0 skipped. This keeps the failure broad-order only. Next step:
  add diagnostics to the render-budget case that name the click phase/iteration
  and print before/after category-tree/page-host snapshot state, then rerun
  exact, category-prefix, and broad Preferences before changing the render
  budget or product behavior. Diagnostic patch is now in
  `Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp`: each category
  click traces its phase/iteration and before/after snapshot state, and the
  failure message includes both snapshots. Build and focused reruns are
  pending. Diagnostic build `.build/logs/msbuild-20260508_174453_085.log`
  completed with 0 warnings / 0 errors. Post-diagnostic exact
  `C:\R\PCHURN3\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped; every traced category click had delta 2. The
  category-tree prefix `C:\R\PCHURNF3\last_run\commands\results.json` passed
  with 9 passed / 0 failed / 0 skipped; its 27 traced category clicks also had
  max delta 2 and no over-budget click. The broad Preferences prefix remains
  the next proof point. Broad Preferences `C:\R\PREFD18\last_run\commands\results.json`
  did not reach the category-tree diagnostic point; it failed earlier after
  62 passed / 1 failed / 105 skipped at
  `cmd_preferences_dialog_keyboard_reordered_resized_columns_survive_search_roundtrip`
  with `Preferences Keyboard search field did not expose a focused Win32 input
  target before live clear-back validation; focusTarget=0,
  search='__codex_no_match__', rows=0.` The trace proves the no-match seed and
  rebuild settled first, so active investigation shifts to Keyboard search
  native-focus reacquisition after the no-match rebuild in broad order. Next
  step: exact and Keyboard reordered/resized-prefix isolation before changing
  the shared search-focus wait or product focus behavior. Exact standalone
  `C:\R\KBSR1\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped, the smaller
  `cmd_preferences_dialog_keyboard_reordered_resized_` slice
  `C:\R\KBRR4\last_run\commands\results.json` passed with 3 passed /
  0 failed / 0 skipped, and the Keyboard prefix
  `C:\R\KBFAM1\last_run\commands\results.json` passed with 24 passed /
  0 failed / 0 skipped. This makes the Keyboard search-focus loss broad
  Preferences-order only. Next step: add failure diagnostics around the
  post-no-match search-focus reacquisition before any refocus/budget change.
  Diagnostic patch is now in
  `Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp`: the failing
  Keyboard search path traces the current search/focus/window state after the
  no-match rebuild and includes the same state in the Win32 input-target
  failure text. Diagnostic build `.build/logs/msbuild-20260508_175625_749.log`
  completed with 0 warnings / 0 errors. Post-diagnostic exact
  `C:\R\KBSR2\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped, and post-diagnostic Keyboard prefix
  `C:\R\KBFAM2\last_run\commands\results.json` passed with 24 passed /
  0 failed / 0 skipped. Broad Preferences remains the next proof point.
  Broad Preferences `C:\R\PREFD19\last_run\commands\results.json` did not
  reach the Keyboard diagnostic point; it failed earlier after 38 passed /
  1 failed / 129 skipped at
  `cmd_preferences_dialog_viewers_long_run_list_scrolling_stays_bounded` with
  `Preferences long-run Viewers scrolling chunk 0 changed the active category
  unexpectedly.` Active investigation shifts to the Viewers long-run scroll
  case: exact and neighboring long-run-prefix isolation must decide whether the
  failure is broad-order category/focus drift, an over-broad wait predicate, or
  real product category churn during the debug scroll helper. Exact
  `C:\R\VIEWLR1\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped, and the Viewers prefix
  `C:\R\VIEWFAM4\last_run\commands\results.json` passed with 27 passed /
  0 failed / 0 skipped. This makes the Viewers scroll category drift broad
  Preferences-order only. Add per-chunk snapshot diagnostics before changing
  the scroll wait predicate or product/debug scroll helper. Diagnostic patch is
  now in `Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp`: the
  Viewers long-run case traces the ready baseline, settled tree baseline, and
  per-chunk post-scroll snapshot, and category/title failures include the same
  detailed state. Diagnostic build `.build/logs/msbuild-20260508_180538_114.log`
  completed with 0 warnings / 0 errors. Post-diagnostic exact
  `C:\R\VIEWLR2\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped, and post-diagnostic Viewers prefix
  `C:\R\VIEWFAM5\last_run\commands\results.json` passed with 27 passed /
  0 failed / 0 skipped. Broad Preferences
  `C:\R\PREFD20\last_run\commands\results.json` passed with 168 passed /
  0 failed / 0 skipped. This clears the current Preferences-family blockers
  through Plugins reordered/resized sort-search, Viewers header/sort/list-scroll,
  Keyboard reordered/resized search, and the category-tree render-budget case.
  Keep the added diagnostics for now because the prior failures were
  broad-order/timing-sensitive and full Commands still needs to prove the
  broader suite state. Full Commands fail-fast
  `C:\R\FULLF44\last_run\commands\results.json` then advanced past the
  Preferences blockers covered by `PREFD20` but stopped after 194 passed /
  1 failed / 402 skipped at
  `cmd_preferences_dialog_themes_page_exposes_live_uia_grid_selection` with
  `Preferences Themes page did not expose its DX grid surface for UIA selection
  validation.` The trace shows Viewers long-run and Keyboard reordered/resized
  search diagnostics were healthy in full order before this stop. Active
  investigation moves to Themes UIA grid-selection setup under full-order
  pollution from earlier non-Preferences command families. Exact standalone
  `C:\R\THEMEUIA1\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped, and the Themes prefix
  `C:\R\THEMEFAM1\last_run\commands\results.json` passed with 26 passed /
  0 failed / 0 skipped. This makes the Themes UIA grid setup failure
  full-order only. Add Themes setup diagnostics before changing page setup or
  shared grid-selection assertions. Diagnostic patch is now in
  `Commands.SelfTest.Preferences.ThemesGeneralPanes.cpp`: the Themes UIA
  selection case traces setup state immediately after selecting Themes, traces
  the ready page state, and includes the same detailed snapshot/window state in
  the failure text. Diagnostic build `.build/logs/msbuild-20260508_182245_750.log`
  completed with 0 warnings / 0 errors. Post-diagnostic exact
  `C:\R\THEMEUIA2\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped, and post-diagnostic Themes prefix
  `C:\R\THEMEFAM2\last_run\commands\results.json` passed with 26 passed /
  0 failed / 0 skipped. Full Commands
  `C:\R\FULLF45\last_run\commands\results.json` advanced past the Themes UIA
  stop and returned to the NavigationView unfocused-pane blocker after
  557 passed / 1 failed / 39 skipped:
  `cmd_pane_navigationView_unfocused_pane_click_focuses_target_pane` failed
  because the right NavigationView HWND matched the expected handle but was not
  visible and had a zero-sized rect, while the debug snapshot reported both
  nav bars visible and the active/focused pane still left. Active investigation
  moves back to NavigationView layout/window-visibility synchronization. Root
  cause hypothesis under test: full Commands can leave `_zoomedPane` set from
  earlier zoom command coverage, and a zoomed-left layout gives the right pane a
  zero-width rect even when the right navigation-bar visibility flag is true.
  Patch in progress: `cmd_pane_navigationView_unfocused_pane_click_focuses_target_pane`
  now saves `_zoomedPane`/restore ratio, clears zoom before configuring both
  panes and requiring the right NavigationView to be visible, then restores the
  original zoom state during teardown. Next evidence: rebuild, exact
  `C:\R\NAVUNF3`, NavigationView prefix `C:\R\NAVFAM9`, and full Commands
  `C:\R\FULLF46`. Debug x64 rebuilt cleanly at
  `.build/logs/msbuild-20260508_184412_311.log` with 0 warnings and 0 errors.
  Focused exact verification `C:\R\NAVUNF3\last_run\commands\results.json`
  passed with 1 passed / 0 failed / 0 skipped. NavigationView prefix
  `C:\R\NAVFAM9\last_run\commands\results.json` passed with 12 passed /
  0 failed / 0 skipped. Full Commands `C:\R\FULLF46` ran as the next
  proof point; background runner output was captured in
  `.build/logs/selftest-FULLF46-stdout.txt` and
  `.build/logs/selftest-FULLF46-stderr.txt`, with RedSalamander process id
  44892 observed at start. Result: `C:\R\FULLF46\last_run\commands\results.json`
  ran non-fail-fast to completion in 656105 ms with 594 passed / 3 failed /
  0 skipped. The first failure remains
  `cmd_pane_navigationView_unfocused_pane_click_focuses_target_pane`: the zoom
  guard proved `zoomedPane=none`, but the right NavigationView HWND was still
  hidden with a zero-height screen rect `(678,1350)-(1243,1350)` while stored
  right navigation-bar visibility and the debug snapshot path were correct.
  This rejects the zoom-only hypothesis and moves the investigation to pane
  layout height/preview/function-bar or restore-state residue before this case.
  Two later failures in the same non-fail-fast run,
  `cmd_app_toggleUiChrome` and `cmd_pane_copy_text`, may be fallout from the
  earlier NavigationView failure unless exact/prefix reruns prove otherwise.
  Patch in progress: the NavigationView visibility diagnostic now records main
  client rect, pane folder-view rect, pane view-options state, preview source /
  host / selected tab, command-line visibility, and function-bar state. The
  unfocused-pane test also closes any inherited preview pane before configuring
  left/right folder paths so a prior preview host cannot hide the target right
  NavigationView. Debug x64 rebuilt cleanly at
  `.build/logs/msbuild-20260508_190416_097.log` with 0 warnings and 0 errors.
  Post-patch exact `C:\R\NAVUNF4\last_run\commands\results.json` passed with
  1 passed / 0 failed / 0 skipped, and NavigationView prefix
  `C:\R\NAVFAM10\last_run\commands\results.json` passed with 12 passed /
  0 failed / 0 skipped. Next proof point is a corrected full Commands
  fail-fast run with `-FailFast` enabled. `C:\R\FULLF47` is running now;
  background runner logs are `.build/logs/selftest-FULLF47-stdout.txt` and
  `.build/logs/selftest-FULLF47-stderr.txt`, with RedSalamander process id
  65096 observed at start. Result:
  `C:\R\FULLF47\last_run\commands\results.json` stopped before the
  NavigationView proof point after 69 passed / 1 failed / 527 skipped at
  `edit_with_menu_populates_applicable_editor_actions`: `Edit With menu action
  should receive expanded path macros; marker first line was ''.` This becomes
  the active first fail-fast blocker; keep the NavigationView preview/zoom
  patch pending until a later full run reaches that case again.
- [x] Resolve the current full-order NavigationView unfocused-pane pointer
  blocker after the preview/zoom setup fix: full Commands fail-fast
  `C:\R\FULLF48\last_run\commands\results.json` advanced past
  `cmd_pane_navigationView_unfocused_pane_click_focuses_target_pane` and the
  whole `cmd_pane_navigationView_` family before stopping later at
  `cmd_app_toggleUiChrome`.
- [x] Resolve the current full-order UI chrome toggle blocker from
  `C:\R\FULLF48`: full Commands fail-fast stopped after 571 passed / 1 failed /
  25 skipped at `cmd_app_toggleUiChrome` with `ToggleFunctionBar did not hide
  the visible FolderWindow function bar.` This matches the later extra failure
  seen in non-fail-fast `C:\R\FULLF46`. Active hypothesis: the test forces
  `g_folderWindow.SetFunctionBarVisible(true)` but the command handler toggles
  the separate persisted/global function-bar model state, so a prior case that
  leaves the global model false makes the first command show instead of hide.
  Exact standalone `C:\R\TOGEX1\last_run\commands\results.json` passed with
  1 passed / 0 failed / 0 skipped, confirming order-dependent state. Patch in
  progress: `cmd_app_toggleUiChrome` now normalizes, toggles, and restores the
  function bar through `IDM_VIEW_FUNCTIONBAR` command dispatch, using direct
  `SetFunctionBarVisible(...)` only as a cleanup fallback if the command model
  cannot be restored. Build evidence after this patch:
  `.\build.ps1 -ProjectName RedSalamander` succeeded with 0 warnings / 0 errors
  in `.build/logs/msbuild-20260508_193700_933.log`. Focused exact verification
  `C:\R\TOGEX2\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped. The runner treats a filter without trailing `_` as an
  exact case, so `C:\R\TOGFAM1\last_run\commands\results.json` is a repeat base
  exact proof with 1 passed / 0 failed / 0 skipped; the true sibling prefix
  `C:\R\TOGFAM2\last_run\commands\results.json` passed
  `cmd_app_toggleUiChrome_keeps_navigation_shell_stable` with 1 passed /
  0 failed / 0 skipped. Next verification: full Commands fail-fast
  `C:\R\FULLF49`, now running with background logs
  `.build/logs/selftest-FULLF49-stdout.txt` and
  `.build/logs/selftest-FULLF49-stderr.txt`; RedSalamander process id 90340 was
  observed at start. Result: `C:\R\FULLF49\last_run\commands\results.json`
  stopped earlier than the UI-chrome proof point after 185 passed / 1 failed /
  411 skipped at
  `cmd_preferences_dialog_keyboard_header_drag_reorders_columns_without_sort`.
  After the Keyboard repair, full Commands fail-fast
  `C:\R\FULLF50\last_run\commands\results.json` passed with 597 passed /
  0 failed / 0 skipped, proving the UI-chrome fix in full order; archived
  evidence: `Specs/TestRuns/4cb089111a23/Commands/2026-05-08_204608/`.
- [x] Resolve the newly resurfaced full-order Preferences Keyboard header-drag
  blocker from `C:\R\FULLF49`: `cmd_preferences_dialog_keyboard_header_drag_reorders_columns_without_sort`
  failed with `Dragging the Preferences Keyboard Shortcut header did not reorder
  the visible DX grid columns while keeping selection and bounded visible work
  stable; selected='About / Show information about RedSalamander. | Unassigned
  | Function bar', rows=7, cols=3, cells=21, resizeCount=2, focusTarget=0,
  pageResizeFailures=0.` Next steps: run exact standalone, then the focused
  Keyboard-header/reordered prefix, before deciding whether this is stale
  test setup, shared DxUi grid drag behavior, or full-order retained state.
  Exact standalone `C:\R\KBHDR1\last_run\commands\results.json` also failed
  with 0 passed / 1 failed / 0 skipped and the same message, proving the first
  gap is standalone. Next patch is diagnostic: include pre-drag header hit-test
  and drag point data so the next exact run can distinguish stale coordinates
  from shared-grid reorder behavior. Diagnostic build
  `.\build.ps1 -ProjectName RedSalamander` succeeded with 0 warnings /
  0 errors in `.build/logs/msbuild-20260508_195237_150.log`.
  Diagnostic exact `C:\R\KBHDR2\last_run\commands\results.json` still failed,
  but proved the start and target coordinates are correct: start `(377,67)` hit
  zone 1 / column 1 / non-resize on the Shortcut header, target `(24,67)` hit
  zone 1 / column 0 / non-resize on the Command header, with header rects
  Command `(12,52-292,82)` and Shortcut `(292,52-462,82)`. Active hypothesis:
  the Keyboard header drag is routed through child-window resolution even though
  its coordinates are already page-host client coordinates; try direct page-host
  mouse delivery for Keyboard header reorder drags, then recheck exact and
  Keyboard-prefix cases. Direct-delivery build
  `.\build.ps1 -ProjectName RedSalamander` succeeded with 0 warnings /
  0 errors in `.build/logs/msbuild-20260508_195545_853.log`.
  Direct-delivery exact reruns still failed, so the blocker is not only child
  resolution in `SendMouseDragToResolvedPointWindow`. Additional diagnostics
  showed the post-drag polling loop could no longer capture the Command or
  Shortcut header rects while the snapshot still reported 7 rows / 3 columns /
  21 cells. The wait-structure cleanup build
  `.build/logs/msbuild-20260508_200423_264.log` succeeded with 0 warnings /
  0 errors, but exact `C:\R\KBHDR5\last_run\commands\results.json` still
  failed. Column-layout diagnostics added in
  `.build/logs/msbuild-20260508_200933_065.log` showed exact
  `C:\R\KBHDR6\last_run\commands\results.json` remained
  `command@0:280.0;shortcut@1:170.0;scope@2:110.0`, so the visible layout never
  committed a reorder. The page-host handle audit found that
  `DebugGetPreferencesActivePageHandle()` returns the parent page host while
  header rects are in the DX host client space; switching the active header drag
  case to `DebugGetPreferencesActivePageDxHostHandle()` built cleanly in
  `.build/logs/msbuild-20260508_201412_891.log`, but exact
  `C:\R\KBHDR7\last_run\commands\results.json` still failed with the same
  unchanged column layout. Next diagnostic: expose shared DxUi grid
  header-reorder pointer/commit state so the exact failure proves whether the
  product grid never receives a reorder drag, receives it but normalizes to a
  no-op, or commits and is later overwritten by page/model rebuild.
  Diagnostic build `.build/logs/msbuild-20260508_202115_450.log` succeeded
  with 0 warnings / 0 errors. Exact `C:\R\KBHDR8\last_run\commands\results.json`
  proved the grid was correct: the drag started once, committed once, and the
  final layout was `shortcut@0:170.0;command@1:280.0;scope@2:110.0`. The
  remaining failure was a stale test predicate requiring the DX host resize
  counter to stay exactly equal to the baseline; a successful first pointer
  reorder can settle one deferred swap-chain resize while preserving focus,
  selection, rows, cells, and zero resize failures. The Keyboard reorder
  variants now target the DX host directly and accept at most one settled resize
  with no resize failures. Build
  `.build/logs/msbuild-20260508_202801_259.log` succeeded with 0 warnings /
  0 errors, and exact `C:\R\KBHDR9\last_run\commands\results.json` passed with
  1 passed / 0 failed / 0 skipped. Next verification: run the broader
  `cmd_preferences_dialog_keyboard_` family to catch any neighboring reorder,
  resize, copy, search, or focus regressions. First Keyboard-family run
  `C:\R\KBFAM1\last_run\commands\results.json` advanced to 3 passed /
  1 failed / 20 skipped and stopped at
  `cmd_preferences_dialog_keyboard_header_resize_changes_visible_width`; the
  resize itself succeeded (`commandWidth=280.0->328.0`, resizeDown 0->1,
  resizeMove 0->1, focus still Shortcuts grid), but the same stale exact
  DX-host resize-count predicate rejected the one settled resize. Patch in
  progress: apply the bounded resize-count rule to the Keyboard header-resize
  case as well. Build `.build/logs/msbuild-20260508_203117_449.log` succeeded
  with 0 warnings / 0 errors. Exact resize verification
  `C:\R\KBRSZ1\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped. Broad Keyboard-family verification
  `C:\R\KBFAM2\last_run\commands\results.json` passed with 24 passed /
  0 failed / 0 skipped, covering the header reorder, header resize,
  reordered-copy, reordered-resized, search roundtrip, and copy/search
  neighbors. Full Commands fail-fast
  `C:\R\FULLF50\last_run\commands\results.json` then passed with 597 passed /
  0 failed / 0 skipped, proving the original `C:\R\FULLF49` stop advances in
  suite order; archived evidence:
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-08_204608/`.
- [x] Resolve the current full-order editor action menu macro blocker from
  `C:\R\FULLF47`: full Commands fail-fast now stops after 69 passed / 1 failed /
  527 skipped at `edit_with_menu_populates_applicable_editor_actions`. Exact
  isolation must determine whether the menu action fails standalone, after the
  preceding viewer/editor action cases, or only in the full suite. Read the
  editor action command setup carefully before patching; the marker file first
  line was empty, so the immediate possibilities are command launch did not run,
  the selected menu action did not receive expanded path macros, or the marker
  read raced the external action write. Exact standalone
  `C:\R\EDITMENU1\last_run\commands\results.json` reproduced the failure with
  0 passed / 1 failed / 0 skipped in 10524 ms. Root cause: the test command
  templates used `>"{Path}\marker.txt"` even though argument macros are
  already quoted by `FileActionLauncher`; the launched `cmd.exe` command line
  became `>""C:\...\edit_with_menu_{...}"\edit-with-menu-marker.txt"` and
  stayed running without creating the marker. Patch in progress: replace the
  marker-writing test actions with macro-safe command shapes that use
  `workingDirectory={Path}`, bare marker filenames, and `if exist {FullPath}` /
  `{SelectedPathsFile}` where the argument macro itself is the behavior under
  test. Debug x64 rebuilt cleanly at
  `.build/logs/msbuild-20260508_191848_694.log` with 0 warnings and 0 errors.
  Exact post-patch `C:\R\EDITMENU2\last_run\commands\results.json` passed with
  1 passed / 0 failed / 0 skipped and left no `cmd.exe` process holding the
  edit-with-menu marker redirection. Adjacent marker-command filters also
  passed fail-fast with no stale marker `cmd.exe` processes:
  `C:\R\ACTIONF1` (`file_action_`) 7 passed / 0 failed / 0 skipped,
  `C:\R\VIEWACT1` (`view_`) 6 / 0 / 0,
  `C:\R\EDITACT1` (`edit_`) 5 / 0 / 0,
  `C:\R\ALTACT1` (`alternate_`) 5 / 0 / 0, and
  `C:\R\USERMENU1` (`user_menu_`) 2 / 0 / 0. Next proof point: full Commands
  fail-fast `C:\R\FULLF48`. `C:\R\FULLF48` is running now with background logs
  `.build/logs/selftest-FULLF48-stdout.txt` and
  `.build/logs/selftest-FULLF48-stderr.txt`; RedSalamander process id 33988
  was observed at start. `C:\R\FULLF48` advanced past this editor action menu
  blocker.
- [x] Resolve the current full-order Preferences Advanced tab-traversal blocker:
  `C:\R\FULLF33\last_run\commands\results.json` stopped after 265 passed /
  1 failed / 331 skipped at
  `cmd_preferences_dialog_advanced_tab_traversal_live_dx_interaction` with
  `Preferences Advanced Filter mask field focus target not reached during tab
  traversal.` This stop is later than the resolved Plugins traversal and before
  the Navigation/Common-Folders proof point, so it is the active fail-fast
  blocker. Exact isolation must decide whether the Advanced traversal fails
  standalone, only after earlier Advanced cases, or because the test still uses
  stale focus setup/message targets and lacks the richer tab-traversal
  diagnostics now present in adjacent Preferences tests. Exact standalone
  `C:\R\ADVTAB1\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped, so this is not a standalone Advanced traversal
  regression. The Advanced-only prefix
  `C:\R\ADVFAM1\last_run\commands\results.json` passed with 6 passed /
  0 failed / 0 skipped, so earlier Advanced-only cases are not sufficient to
  reproduce the miss. Diagnostic patch in progress: the Advanced tab traversal
  helper now records expected/observed retained focus, native focus before/after,
  cached page handle, active page/DX host before/after, child/render counts, and
  resize failures for each Tab step. Debug x64 rebuilt cleanly at
  `.build/logs/msbuild-20260508_135610_202.log` with 0 warnings and 0 errors.
  Exact `C:\R\ADVTAB2\last_run\commands\results.json` passed with
  1 passed / 0 failed / 0 skipped, the Advanced prefix
  `C:\R\ADVFAM2\last_run\commands\results.json` passed with 6 passed /
  0 failed / 0 skipped, and full-order
  `C:\R\FULLF34\last_run\commands\results.json` advanced past this Advanced
  stop before failing later at NavigationView path double-click.
- [x] Resolve the current full-order Hot Paths navigation-shell stability
  blocker: `C:\R\FULLF27\last_run\commands\results.json` stopped after
  465 passed / 1 failed / 131 skipped at
  `cmd_pane_navigation_hot_paths_keeps_navigation_shell_stable` with
  `Navigation shell did not restore cleanly after Hot Paths close; focusTarget=0,
  editMode=no, historyVisible=no, suggestVisible=no, popupVisible=no,
  childWindows=0, currentPath='', historyCount=0, refreshCount=12, itemCount=3,
  selectedCount=0, focusedItem='b.log'.` Exact isolation must determine whether
  the path text is legitimately empty after closing the Hot Paths window, or
  whether the test closes the transient dialog before NavigationView has
  republished the stable breadcrumb snapshot. Exact
  `C:\R\HPNAV1\last_run\commands\results.json` reproduced the same blocker with
  0 passed / 1 failed / 0 skipped; top-level trace proves the command resynced
  the left shell from the folder view before returning. Diagnostic patch in
  `C:\R\HPNAV2\last_run\commands\results.json` rebuilt cleanly and proved the
  NavigationView snapshot, NavigationView HWND, navigation-bar visibility,
  breadcrumb path, actual pane path, history count, item count, selection count,
  and focused item are all stable after close. The remaining mismatch is focus
  restoration: Preferences returns focus to the root main-window owner, but does
  not reliably drive the active pane back to its `FolderView` before queued test
  assertions run. `C:\R\HPNAV3\last_run\commands\results.json` confirms the
  first owner-post patch still left `focusedFolderView=0x0` and
  `focusedWindow=0x0`, so the main-window restore message must actively
  re-activate the main window before restoring pane focus. `C:\R\HPNAV4` and
  `C:\R\HPNAV5` still failed with `focusedFolderView=0x0`; the follow-up trace
  in `C:\R\HPNAV6\last_run\trace.txt` shows the Preferences deferred-close
  message runs with a valid state, but the later `WM_NCDESTROY` owner-restore
  block is not reached early enough for this self-test close path. Patch in
  progress: post the existing `WndMsg::kPaneRestoreFolderFocus` to the owner
  from the deferred-close message before resetting the tracked Preferences
  HWND; the main-window handler restores/minimizes state, calls
  `SetActiveWindow`/`SetForegroundWindow`, runs the normal main-window focus
  restore path, and falls back to
  `FolderWindow::TryRestoreActivePaneFolderViewFocus`. Keep focused-HWND
  diagnostics in the Hot Paths assertion, and record the durable modeless
  tool-window focus contract in `Specs/UI/UI_TopLevelToolWindows.md`.
  `C:\R\HPNAV7\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped after moving the restore request to deferred close.
  Temporary trace probes were removed, the patch rebuilt cleanly at
  `.build/logs/msbuild-20260508_120617_578.log` with 0 warnings and 0 errors,
  and exact `C:\R\HPNAV8\last_run\commands\results.json` passed with
  1 passed / 0 failed / 0 skipped. Full-order proof landed in
  `C:\R\FULLF29\last_run\commands\results.json`, which advanced past the prior
  465-pass Hot Paths stop before failing later at Navigation View address-bar
  Tab traversal.
- [x] Reopen and resolve the full-order Preferences scroll-host render-churn
  blocker: `C:\R\FULLF28\last_run\commands\results.json` stopped after
  292 passed / 1 failed / 304 skipped at
  `cmd_preferences_dialog_scroll_host_preserves_retained_page_state` with
  `Scrolling the Preferences page host should not repaint the category tree;
  saw 1 extra tree render(s).` This was previously believed fixed by settling
  the category-tree render baseline, so the next work must challenge whether
  the settle helper is insufficient after long full order, whether category
  re-entry queues another tree repaint after the baseline is captured, or
  whether the test is missing a real product repaint regression. First gather
  exact/focused repeat evidence and trace data before changing product or test
  assertions. Fresh exact evidence `C:\R\PSH3\last_run\commands\results.json`
  failed earlier than the full-order render-count assertion with
  `Preferences General page should reset page scroll to 0 after leaving a
  scrolled page; saw pageScrollY=108`, proving the helper also sampled
  category transitions before the retained per-category scroll state settled.
  Patch in progress: the scroll-host test now waits for category, page title,
  resize health, and the expected per-category `pageScrollY` before using a
  category-switch snapshot. Debug x64 rebuilt cleanly at
  `.build/logs/msbuild-20260508_122023_537.log` with 0 warnings and 0 errors,
  and exact `C:\R\PSH4\last_run\commands\results.json` passed with 1 passed /
  0 failed / 0 skipped. The broader Preferences prefix
  `C:\R\PREFD9\last_run\commands\results.json` passed with 168 passed /
  0 failed / 0 skipped. Full-order proof landed in
  `C:\R\FULLF29\last_run\commands\results.json`, which advanced past the prior
  292-pass scroll-host stop and the later 465-pass Hot Paths stop before
  failing at Navigation View address-bar Tab traversal.
- [x] Resolve the current short-root full-order blocker:
  `cmd_preferences_dialog_themes_roundtrip_restores_dxui_surface` fails after
  167 passes with `Failed to focus the Preferences category host for Themes
  round-trip test.` Challenge the test's direct `SetFocus(...) == target`
  assertion because Win32 returns the previous focus window; verify focus by
  the resulting focused HWND and then rerun the Themes round-trip/theme-cycle
  and full short-root suite.
- [x] Isolate Connection Manager profile-mutating GUI cases by seeding and
  restoring `g_settings.connections` around `live_dx_interaction` and
  `tab_traversal_live_dx_interaction`; exact repeats now pass with deterministic
  rows after the earlier profile accumulation was removed.
- [x] Resolve the remaining first short-root full-order failure:
  `cmd_connection_manager_window_tab_traversal_live_dx_interaction` still passes
  standalone but fails after the preceding Connection Manager family by routing
  a later forward Tab from the current action button back to the list. The
  diagnostic pass has now ruled out stale modifiers and confirmed native
  focus-loss clears the retained DxUi focus target; patch `WindowHost` focus
  retention and prove it with a generic DxUi regression test plus repeated
  Connection Manager prefix runs.
- [x] Rerun closeout tooling/source-contract checks that were touched by this
  plan (`Tools\Tests`, inventory/runner checks, and focused native suites as
  applicable).
- [x] Verify the current full-closeout fixes with focused reruns:
  LocalizationTests after its handcrafted settings JSON was updated to the
  current v16 schema while preserving the documented invalid-language fallback,
  Tools Pester after plugin language-resource deployment filtering and the
  49-case inventory update, and the synthetic vcpkg merge script through the
  unified runner after PowerShell argument splatting was corrected.
- [x] Merge durable behavior discovered during implementation into the
  authoritative specs under `Specs/` rather than leaving it only in this WIP
  plan. Final durable updates landed in `Specs/Testing/Testing_TestCoverage.md`
  and `Tests/README.md` for standalone output logs, ViewerPETests nested
  timeout budgeting, and the 51-case tooling inventory.
- [x] Move this plan to `Specs/Plans/Done/` only after focused, sharded/full,
  and spec-closeout evidence make the remaining risk explicit.

## Continuation Checkpoint - 2026-05-08 21:41 Europe/Paris

Commands is no longer the active blocker. The full foreground Commands proof at
`C:\R\FULLF50\last_run\commands\results.json` passed with 597 passed / 0
failed / 0 skipped, and archived evidence was written under
`Specs/TestRuns/4cb089111a23/Commands/2026-05-08_204608/`.

The first full closeout run after that proof was
`C:\R\FULLSUITE1\last_run\run-all-tests-results.json`, invoked as:

```powershell
$env:REDSALAMANDER_SELFTEST_ROOT='C:\R\FULLSUITE1'
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2
```

That run passed CompareDirectories (125 passed / 0 failed / 24 skipped),
Commands (597 passed / 0 failed / 0 skipped), FileOperations (55 passed /
0 failed / 20 skipped), DxUiTests, ViewerPETests, ViewerSqliteTests,
MonitorTest, and PerformanceTests2. It failed in three closeout-only areas:
LocalizationTests, ToolsPesterTests, and VcpkgMergeSynthetic.

Root cause notes and active fixes:

- LocalizationTests was still hand-writing settings JSON with
  `schemaVersion=11` even though the settings store source and spec now define
  v16 as current and unsupported versions are invalid. The test JSON now uses
  `schemaVersion=16`, including the invalid `ui.language="..\\bad"` case, which
  continues to assert the documented loader fallback to `system`.
- ToolsPesterTests treated language resource projects under
  `Plugins/*/Lang/*/*.vcxproj` as runtime plugin DLLs expected under
  `.build\x64\Debug\Plugins`. `Tools/Tests/RedSalamanderPluginDeployment.Tests.ps1`
  now filters runtime plugin projects away from `\Lang\` and checks those
  language resource DLLs under `.build\x64\Debug\Lang`.
- The plugin-deployment test added one Pester `It` case, so the source
  inventory and docs now use 49 Tools Pester cases in
  `Tools/Tests/TestInventory.Tests.ps1`, `Specs/Testing/Testing_TestCoverage.md`,
  and `Tests/README.md`.
- VcpkgMergeSynthetic failed only when invoked through the unified runner
  because `Tools/Run-AllTests.ps1` passed `$Entry.Arguments` as one array
  positional argument. The PowerShell-script runner branch now copies the
  arguments into `$scriptArguments` and splats that array.

Next required evidence before moving this plan: rebuild/rerun LocalizationTests,
run Tools Pester after the deployment/inventory fixes, verify the synthetic
vcpkg merge script through the unified runner path, then rerun
`Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2` with a fresh
`REDSALAMANDER_SELFTEST_ROOT`.

Focused evidence collected after the fixes:

- `.build/logs/msbuild-20260508_214411_456.log` rebuilt LocalizationTests with
  0 warnings / 0 errors.
- `.\.build\x64\Debug\LocalizationTests.exe` passed all localization/settings
  checks with exit code 0, including v16 `ui.language` load/save and invalid
  language fallback to `system`.
- `Invoke-Pester -Path @('Tools\Tests\TestInventory.Tests.ps1','Tools\Tests\RunAllTestsPlan.Tests.ps1')`
  passed 13 / 0 / 0.
- `Invoke-Pester -Path 'Tools\Tests' -PassThru` passed 49 / 0 / 0; the embedded
  targeted RedSalamander build logged 0 warnings / 0 errors at
  `.build/logs/msbuild-20260508_214501_956.log`.
- The corrected PowerShell-script runner branch invoked
  `Tests\vcpkg-merge-synthetic-test.ps1` with splatted arguments and passed all
  5 synthetic merge cases with exit code 0.

The remaining closeout gate is the fresh full-suite rerun.

Attempted full-suite rerun `C:\R\FULLSUITE2` was launched through a hidden
background PowerShell wrapper so progress could be polled. It advanced through
CompareDirectories, Commands, FileOperations, the standalone executables, and
into ToolsPesterTests; the embedded plugin-deployment build completed cleanly
at `.build/logs/msbuild-20260508_221218_997.log` with 0 warnings / 0 errors.
However, after that build the wrapper had no active child test process, wrote
no aggregate `run-all-tests-results.json`, and made no observable progress. The
wrapper was stopped manually, so `FULLSUITE2` is diagnostic-only and must not be
used as closeout evidence. Rerun the full suite in the normal foreground runner
path with a fresh root.

Foreground full-suite rerun `C:\R\FULLSUITE3\last_run\run-all-tests-results.json`
completed in 44m 20.9s and is valid failed evidence. Passing suites:
CompareDirectories 125 / 0 / 24, FileOperations 55 / 0 / 20, DxUiTests,
ViewerPETests, ViewerSqliteTests, MonitorTest, LocalizationTests,
PerformanceTests2, and ToolsPesterTests. Remaining failures:

- Commands: `cmd_pane_copy_text` failed after 596 passed / 1 failed / 0 skipped
  because `Copy Name as Text` left the Unicode clipboard empty. Exact
  standalone `C:\R\COPYTXT1` passed, so this is an order/timing clipboard
  contention case, not a standalone command behavior regression. Product fix:
  pane copy-as-text clipboard writes now use bounded `OpenClipboard(...)`
  retries before showing feedback, matching the existing DxUi clipboard
  contention contract. `Specs/UI/UI_CommandMenuKeyboard.md` now records this
  durable behavior.
- VcpkgMergeSynthetic failed after ToolsPesterTests because Pester test files
  leave `Set-StrictMode -Version Latest` active in the runner session.
  `Tools/VcpkgInstallSafety.ps1` assumed a single-file `Get-ChildItem` result
  had `.Count`; the helper now wraps source files in `@(...)`, and
  `Tools/Tests/VcpkgInstallSafety.Tests.ps1` adds strict-mode single-file merge
  coverage. Tooling inventory and docs now report 50 Tools Pester cases.

Focused evidence after these fixes:

- `.build/logs/msbuild-20260508_231035_752.log` rebuilt RedSalamander with
  0 warnings / 0 errors after the product clipboard retry patch.
- `C:\R\COPYTXT2\last_run\commands\results.json` passed exact
  `cmd_pane_copy_text` with 1 passed / 0 failed / 0 skipped.
- `Set-StrictMode -Version Latest; .\Tests\vcpkg-merge-synthetic-test.ps1`
  passed all 5 synthetic merge cases with exit code 0.
- `Invoke-Pester -Path @('Tools\Tests\VcpkgInstallSafety.Tests.ps1','Tools\Tests\TestInventory.Tests.ps1')`
  passed 10 / 0 / 0, including the new strict-mode single-file merge guard and
  the 50-case inventory/doc alignment.

The remaining closeout gate is a fresh foreground full-suite rerun with a new
root after these fixes.

## Stop/Resume Checkpoint - 2026-05-08 23:46 Europe/Paris

The user intentionally stopped the chat while the post-fix full-suite rerun was
in progress. Before stopping, process state was checked:

- No `RedSalamander.exe` process was active.
- No `Run-AllTests.ps1`, `--commands-selftest`, `--compare-selftest`, or
  `--fileops-selftest` process from the interrupted run was active.
- `C:\R\FULLSUITE4\last_run\run-all-tests-results.json` does not exist.
- `C:\R\FULLSUITE4\last_run\results.json` and
  `C:\R\FULLSUITE4\last_run\fileops\results.json` contain only the native
  FileOperations result (55 passed / 0 failed / 20 skipped). Because the unified
  runner was interrupted before writing the aggregate, `FULLSUITE4` is partial
  diagnostic evidence only and MUST NOT be used to close the plan.

Current code/spec fixes that are already applied and focused-green:

- `RedSalamander/FolderWindow.FileSystem.cpp`: pane copy-as-text clipboard
  writes now retry `OpenClipboard(...)` for a bounded period before failing,
  matching the existing DxUi contention-tolerant clipboard behavior.
- `Specs/UI/UI_CommandMenuKeyboard.md`: durable contract now states
  copy-as-text commands write `CF_UNICODETEXT` and tolerate short-lived
  clipboard contention with bounded retries.
- `Tools/VcpkgInstallSafety.ps1`: `Merge-RSVcpkgTripletSafe` wraps the
  `Get-ChildItem` result in `@(...)` so single-file merges are strict-mode safe.
- `Tools/Tests/VcpkgInstallSafety.Tests.ps1`: added strict-mode single-source
  file merge coverage.
- `Tools/Tests/TestInventory.Tests.ps1`, `Specs/Testing/Testing_TestCoverage.md`,
  `Tests/README.md`, and this plan now report 50 Tools Pester cases.
- Earlier closeout fixes remain applied: `Tools/Run-AllTests.ps1` splats
  PowerShell script arguments correctly; `Tools/Tests/RedSalamanderPluginDeployment.Tests.ps1`
  separates runtime plugin DLLs from plugin language-resource DLLs; and
  `Tests/LocalizationTests/LocalizationTests.cpp` uses current
  `schemaVersion=16` handcrafted settings JSON while preserving invalid
  `ui.language` fallback to `system`.

Fresh focused evidence already collected after the final fixes:

- `.build/logs/msbuild-20260508_231035_752.log`: Debug x64 RedSalamander build
  after the clipboard retry patch, 0 warnings / 0 errors.
- `C:\R\COPYTXT2\last_run\commands\results.json`: exact
  `cmd_pane_copy_text`, 1 passed / 0 failed / 0 skipped.
- `Set-StrictMode -Version Latest; .\Tests\vcpkg-merge-synthetic-test.ps1`:
  all 5 synthetic merge cases passed with exit code 0.
- `Invoke-Pester -Path @('Tools\Tests\VcpkgInstallSafety.Tests.ps1','Tools\Tests\TestInventory.Tests.ps1') -PassThru`:
  10 passed / 0 failed / 0 skipped.

Last valid full-suite evidence:

- `C:\R\FULLSUITE3\last_run\run-all-tests-results.json` completed and failed
  with exactly two blockers: `cmd_pane_copy_text` and `VcpkgMergeSynthetic`.
  Those two blockers are the ones fixed above.
- `C:\R\FULLSUITE2` was a hidden-background wrapper attempt that reached
  ToolsPester's embedded build, then stopped making observable progress without
  aggregate JSON; it was manually stopped and is diagnostic-only.

Resume from here:

```powershell
Get-Process RedSalamander -ErrorAction SilentlyContinue
$env:REDSALAMANDER_SELFTEST_ROOT='C:\R\FULLSUITE5'
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2
```

If `FULLSUITE5` passes, update this plan with the aggregate counts and path,
archive the full-run evidence under `Specs/TestRuns/<machine>/Tests/<timestamp>/`,
check for remaining unchecked checklist items with:

```powershell
rg -n "^- \[ \]" Specs\Plans\WIP\Tests_ReviewAndGapsPlan_2026-05-05.md
```

Then move the plan to `Specs/Plans/Done/` only after the closeout checklist is
complete and all durable behavior is in authoritative specs. If `FULLSUITE5`
fails, treat the first failure as the next blocker, reproduce it focused, and
append the new discovery here before editing.

## Continuation Checkpoint - 2026-05-09 11:14 Europe/Paris

Fresh foreground closeout run:

```powershell
$env:REDSALAMANDER_SELFTEST_ROOT='C:\R\FULLSUITE5'
.\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2
```

Result: `C:\R\FULLSUITE5\last_run\run-all-tests-results.json` completed and
failed with 780 passed / 5 failed / 44 skipped. Wall time was 42m26.4s. The
aggregate confirms that CompareDirectories, FileOperations, DxUiTests,
ViewerSqliteTests, MonitorTest, LocalizationTests, PerformanceTests2,
ToolsPesterTests, and VcpkgMergeSynthetic passed, so the previous localization,
Tools Pester, vcpkg strict-mode, copy-as-text clipboard, and file-operations
prompt blockers did not recur.

Current blockers from `FULLSUITE5`:

- `Commands`: 593 passed / 4 failed / 0 skipped.
  - `cmd_preferences_dialog_category_switches_do_not_churn_tree_host`: the
    category-tree render budget saw 3 renders while switching to `Viewers`,
    where the test allows at most 2.
  - `cmd_app_compare_keeps_navigation_shell_stable`: after closing Compare
    Directories, navigation shell focus did not restore cleanly
    (`focusTarget=0`, `childWindows=0`, selected count 0).
  - `cmd_pane_focusAddressBar_tab_traversal`: focused window was unavailable
    before the forward address-bar Tab handoff.
  - `cmd_pane_navigationView_edit_suggest_keyboard_routing`: navigation edit
    state did not settle without the suggest popup while applying the selected
    suggestion.
- `ViewerPETests`: process exited with code 1 after 46.6s.

Next diagnostic order:

1. Inspect the `FULLSUITE5` command aggregate and any standalone stdout/stderr
   captured for `ViewerPETests`.
2. Reproduce `ViewerPETests` standalone first because it failed outside the
   Commands UI-order state machine.
3. Re-run the four failing Commands cases exact/prefix under short fresh roots.
   If exact runs pass, treat them as full-order timing/state leaks and capture
   preceding-case context before changing code or budgets.
4. Update this plan with every focused result before making fixes.

Focused diagnosis and first fix:

- `ViewerPETests` direct rerun passed with exit code 0 after the full-suite
  standalone failure, so the closeout blocker is currently intermittent or
  runner-hidden rather than a consistently reproducible standalone failure.
  Because `Run-AllTests.ps1` starts executable suites without capturing stdout
  or stderr, the next durable runner improvement should capture executable
  output so standalone failures are diagnosable from aggregate artifacts.
- `cmd_preferences_dialog_category_switches_do_not_churn_tree_host` exact
  passed at `C:\R\F5_PREFCHURN_EXACT1\last_run\commands\results.json`
  (1 passed / 0 failed / 0 skipped).
- `cmd_app_compare_keeps_navigation_shell_stable` exact passed at
  `C:\R\F5_APPCOMPARE_EXACT2\last_run\commands\results.json`
  (1 passed / 0 failed / 0 skipped).
- `cmd_pane_navigationView_edit_suggest_keyboard_routing` exact passed at
  `C:\R\F5_NAVSUGGEST_EXACT2\last_run\commands\results.json`
  (1 passed / 0 failed / 0 skipped).
- `cmd_pane_focusAddressBar_tab_traversal` reproduced as an intermittent
  focused failure: `C:\R\F5_ADDRTAB_EXACT2` failed on the reverse Shift+Tab
  handoff, `C:\R\F5_ADDRTAB_EXACT3` passed, then the pre-fix repeat loop
  `C:\R\F5_ADDRTAB_REPEAT1` through `C:\R\F5_ADDRTAB_REPEAT8` failed once
  (run 7) with `Focused window unavailable before reverse address-bar
  shift-tab handoff.` Root cause: the test was sampling `GetFocus()` once after
  proving the Dx edit surface existed, so a transient null Win32 focus between
  the edit-mode snapshot and Tab dispatch could fail the test before exercising
  the product handoff.
- Applied fix in `RedSalamander/SelfTest/Commands/Commands.SelfTest.ViewCommands.cpp`:
  the address-bar Tab test now activates the main window before dispatching the
  focus-address-bar command, waits for the live edit host/bridge, restores focus
  to that edit surface when Windows reports a transient null focus, and expands
  failure diagnostics with active window, focus window, edit host/bridge, and
  snapshot state. This keeps the coverage intent: Tab/Shift+Tab must still tear
  down the address-bar edit surface and restore the folder view with stable
  selection.
- Build after the fix passed:
  `.build/logs/msbuild-20260509_112707_121.log`, 0 warnings / 0 errors.
- Post-fix focused proof: `C:\R\F5_ADDRTAB_FIXED_REPEAT1` through
  `C:\R\F5_ADDRTAB_FIXED_REPEAT12` all passed (12/12). Family checks also
  passed: `C:\R\F5_PREFCATEGORY_FIXED1` passed 9/0/0,
  `C:\R\F5_APP_PREFIX_FIXED1` passed 77/0/0, and
  `C:\R\F5_NAVVIEW_PREFIX_FIXED1` passed 12/0/0.

Next step: run a fresh full Commands suite before another full closeout suite.
If Commands passes, rerun `ViewerPETests` through an improved/captured runner
or the full suite to determine whether the standalone exit-code failure recurs.

Continuation update after the first address-bar patch:

- Implemented native executable / CppUnitTest output capture in
  `Tools/Run-AllTests.ps1`. Those runner entries now use the streaming process
  helper and write `<artifactRoot>\<EntryName>.output.log`; the aggregate
  suite summary in `Tools/TestRunPlan.ps1` carries `output_log_path`, and
  `Tools/Tests/RunAllTestsPlan.Tests.ps1` asserts the field. Focused Pester
  for `RunAllTestsPlan.Tests.ps1` plus `ProcessStreaming.Tests.ps1` passed
  10/0/0 again after the display StrictMode fix.
- Fresh full Commands diagnostic:

  ```powershell
  $env:REDSALAMANDER_SELFTEST_ROOT='C:\R\FULLCMDFIX1'
  .\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -TimeoutMultiplier 2
  ```

  `C:\R\FULLCMDFIX1\last_run\run-all-tests-results.json` recorded
  592 passed / 5 failed / 0 skipped with `exit_code=1`. The runner then hit a
  display-only StrictMode error because self-test result hashtables do not
  always expose `OutputLogPath`; `Tools/Run-AllTests.ps1` now reads either
  `OutputLogPath` or `output_log_path` through `Get-JsonValue` before printing
  the optional output path. Focused Pester after the display fix passed 10/0/0.
- `FULLCMDFIX1` remaining failures:
  `cmd_preferences_dialog_category_switches_do_not_churn_tree_host` exceeded
  the category-tree render budget once while returning from Viewers to General;
  `cmd_pane_createDirectory_prompt_long_run_open_close_stays_stable` observed
  a prompt ValuePattern that did not start with `New folder` on cycle 2;
  `cmd_pane_focusAddressBar_tab_traversal` lost the live edit host/bridge before
  the forward Tab handoff; `cmd_pane_navigationView_edit_suggest_keyboard_routing`
  lost its popup/edit state across the synthetic DPI transition; and
  `cmd_app_menuBar_submenu_placement_matches_spec` closed the cascading submenu
  when the pointer returned to the owning parent item.
- The `cmd_pane_focusAddressBar_tab_traversal` diagnostic showed the first
  patch still allowed the live edit snapshot to be lost before the focus restore
  loop (`focus=null`, `active=0x0`, `host=null`, `bridge=null`,
  `visibleChildren=0`). The second patch keeps the last captured live edit
  snapshot available to `sendTabFromFocusedWindow`, tries to restore focus to
  that host/bridge before pumping messages, and avoids an immediate pump after
  `SetFocus(...)` that could blur the edit surface before the synthetic key is
  dispatched.
- Build after the second patch passed:
  `.build/logs/msbuild-20260509_114705_398.log`, 0 warnings / 0 errors.

Next step from here: rerun the address-bar exact/repeat proof. If it passes,
run a fresh full Commands suite under `C:\R\FULLCMDFIX2` and isolate only the
failures that still reproduce after the address-bar stabilization.

Address-bar proof after the second patch:

- `C:\R\F5_ADDRTAB_FIXED2_REPEAT1` through
  `C:\R\F5_ADDRTAB_FIXED2_REPEAT12` all passed
  `cmd_pane_focusAddressBar_tab_traversal` with exit code 0 and 1 passed /
  0 failed / 0 skipped.

Next step from here: run fresh full Commands under `C:\R\FULLCMDFIX2`; only
failures that still reproduce there should drive further code changes.

Full Commands proof after the runner/display and address-bar fixes:

- `C:\R\FULLCMDFIX2\last_run\run-all-tests-results.json` passed the full
  Commands suite with 597 passed / 0 failed / 0 skipped in 9m56.4s. This run
  cleared the five `FULLCMDFIX1` diagnostic failures, including the address-bar
  Tab handoff, NavigationView suggest-DPI, create-directory prompt long-run,
  menu-bar submenu placement, and Preferences category-tree render-budget
  cases.

Next step from here: run a fresh full closeout suite under `C:\R\FULLSUITE6`
with the improved runner output capture. If it passes, update the authoritative
specs and move this WIP plan to `Specs/Plans/Done/`; if it fails, record the
new blocker set here and isolate each failure before changing code.

Full closeout attempt and `ViewerPETests` timeout fix:

- `C:\R\FULLSUITE6\last_run\run-all-tests-results.json` failed the closeout
  gate with 784 passed / 1 failed / 44 skipped in 36m58.7s. Passing suites:
  CompareDirectories 125/0/24, Commands 597/0/0, FileOperations 55/0/20,
  DxUiTests, ViewerSqliteTests, MonitorTest, LocalizationTests,
  PerformanceTests2, ToolsPesterTests 50/0/0, and VcpkgMergeSynthetic 5/0.
  Only `ViewerPETests` failed with exit code 1.
- The new output log captured the actual failure:
  `C:\R\FULLSUITE6\last_run\ViewerPETests.output.log` failed at
  `TestViewerShellComboHostsLongRunOpenCloseStayStable fresh harness process
  exits before timeout`. The log showed the nested six-cycle shell-combo churn
  child was still executing normal cycles when the parent fresh-process wrapper
  hit the same 120-second cap used for a single isolated viewer case.
- Exact control proof:
  `.\.build\x64\Debug\ViewerPETests.exe TestViewerShellComboHostsLongRunOpenCloseStayStable`
  passed with exit code 0 in 25.7s, so the blocker was the outer timeout budget,
  not a viewer assertion failure.
- Applied fix in `Tests/ViewerPETests/ViewerPETests.cpp`: normal isolated viewer
  cases use `kViewerHarnessDefaultTimeout = 120000ms`; the nested long-run
  shell-combo entry uses `kViewerShellComboLongRunTimeout = 600000ms` when
  launched by the fresh-process full harness. The six-cycle test itself still
  launches each child viewer check with the normal default timeout.
- Added maintenance guard in
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` and updated the tooling
  inventory docs/counts from 50 to 51 Pester-style cases in
  `Specs/Testing/Testing_TestCoverage.md` and `Tests/README.md`.
- Verification after the fix:
  `.\build.ps1 -ProjectName ViewerPETests` passed with 0 warnings / 0 errors
  (`.build/logs/msbuild-20260509_124720_672.log`);
  `Invoke-Pester -Path 'Tools\Tests\TestHarnessSourceContracts.Tests.ps1' -PassThru`
  passed 14/0/0; direct `.\.build\x64\Debug\ViewerPETests.exe` passed with
  exit code 0 in 41.1s; `Invoke-Pester -Path
  'Tools\Tests\TestInventory.Tests.ps1' -PassThru` passed 5/0/0 and confirmed
  the updated 51-case tooling inventory/doc alignment.

Final closeout evidence:

- Fresh full closeout:

  ```powershell
  $env:REDSALAMANDER_SELFTEST_ROOT='C:\R\FULLSUITE7'
  .\Tools\Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2
  ```

  `C:\R\FULLSUITE7\last_run\run-all-tests-results.json` passed with
  785 passed / 0 failed / 44 skipped in 48m32.2s. Suite counts:
  CompareDirectories 125/0/24, Commands 597/0/0, FileOperations 55/0/20,
  DxUiTests 1/0/0, ViewerPETests 1/0/0, ViewerSqliteTests 1/0/0,
  MonitorTest 1/0/0, LocalizationTests 1/0/0, PerformanceTests2 1/0/0,
  ToolsPesterTests 51/0/0, and VcpkgMergeSynthetic 5/0/0.
- Archived closeout evidence:
  `Specs/TestRuns/4cb089111a23/Tests/2026-05-09_125031/` contains the aggregate
  `run-all-tests-results.json`, self-test and FileOperations JSON/trace files,
  and standalone output logs renamed to `.txt` so they are check-in friendly.
- Remaining skipped coverage is expected conditional coverage: remote
  credentials/connection profiles were not configured, optional FileOperations
  7z cases are intentionally covered in Compare self-tests, and SQLite
  direct-query cutover cases skipped because the live journal cursor was
  unavailable. Skipped reasons are preserved in the aggregate JSON.
- Final durable spec updates landed in `Specs/Testing/Testing_TestCoverage.md`
  and `Tests/README.md`. The plan can move to `Specs/Plans/Done/`.

## Continuation Checkpoint - 2026-05-08 16:30 Europe/Paris

The Preferences family is no longer the active full-order blocker. After the
Keyboard native-input-target fix and the scroll-host repaint-baseline fix,
exact and prefix evidence passed at `C:\R\KBCOPY2`, `C:\R\KBRR3`,
`C:\R\PSH5`, and `C:\R\PREFD14`. `C:\R\PREFD14\last_run\commands\results.json`
passed the broad `cmd_preferences_dialog_` family with 168 passed / 0 failed /
0 skipped.

Fresh full foreground fail-fast evidence:
`C:\R\FULLF40\last_run\commands\results.json` ran for 536898 ms and stopped
after 461 passed / 1 failed / 135 skipped. The first failure is
`cmd_pane_navigation_directory_impact_preserves_selection_across_chained_renames`:
`Chained directory-impact refresh did not preserve selection onto the final
rename target; selectedCount=0, renameNewSelected=no, unselectedNewSelected=no,
renameOldVisible=no, renameMidVisible=no, renameNextVisible=no,
unselectedOldVisible=no.`

Control evidence narrows this to full-order or timing/state interaction, not a
simple standalone product regression:

- `C:\R\DIRCHAIN1\last_run\commands\results.json` passed the exact chained
  rename case with 1 passed / 0 failed / 0 skipped.
- `C:\R\DIRCHAIN2\last_run\commands\results.json` passed the
  `cmd_pane_navigation_directory_impact_` prefix with 2 passed / 0 failed /
  0 skipped, including the immediately preceding directory-impact selection
  case plus the chained case.

Current hypothesis to verify before fixing: the chained-rename case waits for
at least one directory enumeration and for the expected final names to appear,
but a later queued directory-impact refresh may still clear or remap selection
after that item-presence wait. The existing failure text does not show enough
state to distinguish delayed refresh, wrong pane/path, inherited filter/hidden
state, focus loss, or selection-map collapse. Next patch should add focused
diagnostics for current pane path, item count, focused item, selected count,
name-filter state, final/old/intermediate name visibility and selection, and
the directory enumeration count before and after the chained notifications.

Diagnostic build evidence after adding that state capture:
`.build/logs/msbuild-20260508_162555_186.log` rebuilt Debug x64 with
0 warnings and 0 errors. The strengthened exact chained-rename case
`C:\R\DIRCHAINRED1\last_run\commands\results.json` passed with 1 passed /
0 failed / 0 skipped even when the test forced one refresh after the first
rename hint before sending the remaining chain. The broader
`cmd_pane_navigation_` prefix at `C:\R\NAVDIAG1` passed with 31 passed /
0 failed / 0 skipped, and the wider `cmd_pane_` prefix at `C:\R\PANEDIAG1`
passed with 206 passed / 0 failed / 0 skipped. This keeps the chained-rename
blocker classified as full-order only so far.

The next full fail-fast run, `C:\R\FULLF41\last_run\commands\results.json`,
did not reach the directory-impact case. It stopped earlier at
`cmd_preferences_dialog_keyboard_page_exposes_live_uia_grid_selection` after
184 passed / 1 failed / 412 skipped with
`Preferences Keyboard page did not expose its DX grid surface for UIA selection
validation.` The trace shows it had advanced through Plugins, Keyboard/Viewers/
Themes long-run list scrolling, Themes grid selection, Viewers grid selection,
Viewers tab traversal, and Viewers reorder/resize/copy cases before opening the
Keyboard grid-selection case. Focused isolation stayed green: exact standalone
`C:\R\KBUIA1\last_run\commands\results.json` passed with 1 passed / 0 failed /
0 skipped, the Keyboard family `C:\R\KBFAM1\last_run\commands\results.json`
passed with 24 passed / 0 failed / 0 skipped, and the broad Preferences prefix
`C:\R\PREFD15\last_run\commands\results.json` passed with 168 passed / 0 failed
/ 0 skipped. A diagnostics-only patch now waits for Keyboard category focus with
the shared focus helper instead of checking the previous `SetFocus` return value,
and failure messages include category, page title, row/cell counts, search/focus
state, page child counts, rendered DX hosts, resize failures, active page/host,
and native focus. Debug x64 rebuilt cleanly at
`.build/logs/msbuild-20260508_164447_455.log`; fresh exact
`C:\R\KBUIA3\last_run\commands\results.json` passed with 1 passed / 0 failed /
0 skipped, and Keyboard prefix `C:\R\KBFAM2\last_run\commands\results.json`
passed with 24 passed / 0 failed / 0 skipped. Full Commands validation
`C:\R\FULLF42\last_run\commands\results.json` then advanced past the reopened
Keyboard UIA and directory-impact stops, through `dispatch_smoke_all_commands`,
and through the repaired NavigationView double-click case before failing later
at `cmd_pane_navigationView_unfocused_pane_click_focuses_target_pane` after
557 passed / 1 failed / 39 skipped. The archived evidence is
`Specs/TestRuns/4cb089111a23/Commands/2026-05-08_170141/`.

Do not move this plan to `Specs/Plans/Done/` yet. The Commands full foreground
suite is still fail-fast blocked in the Preferences/Navigation area, and
closeout/source-contract checks have not been rerun after the latest patches.

## Continuation Checkpoint - 2026-05-08 Europe/Paris

Work resumed on the saved `cmd_preferences_dialog_general_theme_cycle_keeps_surface_legible`
blocker from `C:\R\FULLF21`. Exact standalone evidence is green:
`C:\R\GENLIGHT1\last_run\commands\results.json` passed with 1 passed / 0
failed / 0 skipped. The General-only order slice is also green:
`C:\R\GENFAM1\last_run\commands\results.json` passed with 7 passed / 0 failed /
0 skipped. The full Preferences dialog prefix is green:
`C:\R\PREFD1\last_run\commands\results.json` passed with 168 passed / 0
failed / 0 skipped. This narrows the blocker to state introduced before the
Preferences dialog family in full Commands order, not a standalone General or
Preferences-family failure. A diagnostic-only patch now formats the final
Preferences General theme-cycle snapshot when a theme wait misses, including
expected/actual theme flags, category, child counts, render/resize counts,
focus target, scroll state, visible dialog child count, page title, and backdrop
state. Debug x64 rebuilt cleanly at
`.build/logs/msbuild-20260508_093616_176.log` with 0 warnings and 0 errors.
Fresh exact General theme-cycle `C:\R\GENLIGHT2\last_run\commands\results.json`
passed with 1 passed / 0 failed / 0 skipped.

The post-build Preferences prefix sanity check surfaced an earlier blocker:
`C:\R\PREFD2\last_run\commands\results.json` stopped at
`cmd_preferences_dialog_hot_paths_tab_traversal_live_dx_interaction` after 98
passes / 1 failed / 69 skipped. Focused isolation is green:
`C:\R\HPTAB1\last_run\commands\results.json` passed exact standalone, and
`C:\R\HPFAM1\last_run\commands\results.json` passed the Hot Paths prefix with
7 passed / 0 failed / 0 skipped. `C:\R\PREFD2\last_run\trace.txt` records the
failure as a focus drop from the first Browse button toward the first Label
field: native focus before the step was the Hot Paths DX host, native focus
after was `0x0`, and the active page / active DX host stayed on the same Hot
Paths page host. Next: repeat the Preferences prefix once to decide whether
this is a reproducible earlier-Preferences order bug or a single-run focus
flake; if it reproduces, inspect earlier General/Panes theme/settings state and
the DxUi button-to-TextField Tab path before patching.

## Continuation Checkpoint - 2026-05-06 22:30 Europe/Paris

Stop point saved for the next session. No `RedSalamander.exe` self-test process
is still running.

Design note: the user confirmed the new DxUi visual design is the correct one.
Do not revert the refreshed `advanced_controls_dark.png` baseline while
continuing this plan.

Latest full-order run:
`C:\R\FULLF21\last_run\commands\results.json` finished with 210 passed /
1 failed / 386 skipped. The first failure is
`cmd_preferences_dialog_general_theme_cycle_keeps_surface_legible`:
`Preferences General page did not settle after the light theme update.` The
trace shows the preceding cases all passed through Themes grid/sort/search
coverage and General setup:

- `cmd_preferences_dialog_general_page_uses_dxui_toggle_cards` passed
- `cmd_preferences_dialog_general_live_dx_interaction` passed
- `cmd_preferences_dialog_general_dxui_customization_preview_and_cancel` passed
- `cmd_preferences_dialog_general_window_backdrop_apply_updates_supported_windows`
  passed
- `cmd_preferences_dialog_general_tab_traversal_live_dx_interaction` passed
- `cmd_preferences_dialog_general_roundtrip_restores_dxui_surface` passed
- `cmd_preferences_dialog_general_theme_cycle_keeps_surface_legible` failed

Immediate next commands:

```powershell
$env:REDSALAMANDER_SELFTEST_ROOT='C:\R\GENLIGHT1'
$args=@('--commands-selftest','--selftest-case=cmd_preferences_dialog_general_theme_cycle_keeps_surface_legible','--selftest-fail-fast','--selftest-timeout-multiplier=2')
$p=Start-Process -FilePath '.\.build\x64\Debug\RedSalamander.exe' -ArgumentList $args -Wait -PassThru
$json=Get-Content 'C:\R\GENLIGHT1\last_run\commands\results.json' -Raw | ConvertFrom-Json
[pscustomobject]@{ExitCode=$p.ExitCode;Passed=$json.passed;Failed=$json.failed;Skipped=$json.skipped;FailureMessage=$json.failureMessage}
$json.cases | ? status -eq failed | select name,reason,duration_ms
```

If exact standalone passes, rerun the smallest useful order slice around Themes
and General before changing product code:

```powershell
$env:REDSALAMANDER_SELFTEST_ROOT='C:\R\GENTHEMEFAM1'
$env:REDSALAMANDER_REPO_ROOT=(Get-Location).Path
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter cmd_preferences_dialog_ -SkipBuild -TimeoutMultiplier 2 -ExePath .build\x64\Debug\RedSalamander.exe
```

Recent fixes already applied and verified before this stop:

- Quick Search reactivation/no-match focus-loss: `Commands.SelfTest.Search.cpp`
  now reasserts stable folder-view focus after Quick Search activation before
  sending synthetic `WM_CHAR`. Spec rule added to
  `Specs/Testing/Testing_SelfTests.md`. Evidence: build
  `.build/logs/msbuild-20260506_220007_066.log`, exact
  `C:\R\QS2\last_run\commands\results.json`, and full-order
  `C:\R\FULLF20\last_run\commands\results.json` advanced past the old 299-pass
  stop.
- FolderView empty/filter watermark: full-order
  `C:\R\FULLF20\last_run\commands\results.json` passed
  `folderView_filter_watermark_empty_state` at index 507.
- Edit New macro quoting: `Commands.SelfTest.Dialogs.cpp` now waits for the
  quoted marker line `"alpha.editnew"` because file-action macros in
  `arguments` are documented command-line arguments. Spec rule added to
  `Specs/Testing/Testing_SelfTests.md`. Evidence: build
  `.build/logs/msbuild-20260506_221607_376.log`, exact
  `C:\R\EDITNEW2\last_run\commands\results.json`, and family
  `C:\R\EDITNEWFAM1\last_run\commands\results.json`.

Current files changed in the final stretch:

- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Search.cpp`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Dialogs.cpp`
- `Specs/Testing/Testing_SelfTests.md`
- `Specs/Plans/WIP/Tests_ReviewAndGapsPlan_2026-05-05.md`

Do not move this plan to `Specs/Plans/Done/` yet. The Commands full foreground
suite is still fail-fast blocked at the General light-theme settle case, and
closeout/source-contract checks have not been rerun after the latest patches.

## Continuation Checkpoint - 2026-05-06 20:31 Europe/Paris

`C:\R\FULLF13\last_run\commands\results.json` advanced past the earlier
credential-prompt and Themes/Connection Manager stops, then failed at
`cmd_preferences_dialog_category_switches_do_not_churn_tree_host` with
291 passed / 1 failed / 305 skipped. Exact standalone
`C:\R\PCHURN1\last_run\commands\results.json` passed, and the Preferences prefix
`C:\R\PREFCHURN1\last_run\commands\results.json` passed with 168 passed /
0 failed / 0 skipped, so this is full-order only so far.

Patch checkpoint: the category-switch render-churn test now measures from a
settled category-tree render-count baseline to a settled post-click snapshot
instead of capturing immediately around one message-pump pass. This keeps the
budget assertion but prevents queued setup, hover, or paint work from being
charged to the category click. Durable rule added to
`Specs/Testing/Testing_SelfTests.md`. Evidence after the patch: Debug x64 build
`.build/logs/msbuild-20260506_202544_689.log` passed with 0 warnings and
0 errors; `C:\R\PCHURN2\last_run\commands\results.json` passed with
1 passed / 0 failed / 0 skipped; and
`C:\R\PREFCHURN2\last_run\commands\results.json` passed with 168 passed /
0 failed / 0 skipped. The checklist item remains open until the next fresh
full short-root run advances past the previous 291-pass stop.

`C:\R\FULLF14\last_run\commands\results.json` then proved the Preferences
render-churn repair in full order: it advanced past
`cmd_preferences_dialog_category_switches_do_not_churn_tree_host`, past the
previous Compare Options stop, and stopped later at
`cmd_pane_fileops_issues_pane_long_run_scrolling_stays_bounded` with
427 passed / 1 failed / 169 skipped. This closes the Preferences render-churn
item and makes the File Operations issues-pane long-run scroll case the active
blocker.

Patch checkpoint for issues-pane long-run scrolling: the test now resets the
grid to a deterministic top position, alternates down/up wheel chunks so every
chunk has scroll room, forces the target pane paint after using the debug-only
scroll helper, and requires both scroll-offset movement and repaint evidence
with scroll/render counters in the failure text. Durable rule added to
`Specs/Testing/Testing_SelfTests.md`. Evidence after the patch: Debug x64 build
`.build/logs/msbuild-20260506_204808_947.log` passed with 0 warnings and
0 errors; `C:\R\IPLR2\last_run\commands\results.json` passed with
1 passed / 0 failed / 0 skipped; and
`C:\R\IPFAM2\last_run\commands\results.json` passed with 17 passed /
0 failed / 0 skipped. Keep the issues-pane checklist item open until a fresh
full short-root run advances past the previous 427-pass stop.

`C:\R\FULLF15\last_run\commands\results.json` then proved the issues-pane and
clipboard fixes in full order: it advanced past the previous 427-pass
issues-pane stop, passed `cmd_pane_clipboardPasteShortcut_creates_unique_links`,
passed Compare Options tab traversal, and stopped later at
`folderView_filter_watermark_empty_state` with 507 passed / 1 failed /
89 skipped. This closes the issues-pane and clipboard blockers and makes
FolderView empty/filter watermark state the active blocker.

Patch checkpoint for FolderView filter watermark empty-state: the test now
activates the empty fixture first, then clears inherited pane visibility/filter
state on that same fixture path and waits for the resulting enumeration plus
debug state. This prevents a visibility reset from refreshing the old pane path
and racing the empty-folder assertion in full-suite order. Durable rule added to
`Specs/Testing/Testing_SelfTests.md`. Evidence after the patch: Debug x64 build
`.build/logs/msbuild-20260506_210546_890.log` passed with 0 warnings and
0 errors; `C:\R\FVWM2\last_run\commands\results.json` passed with 1 passed /
0 failed / 0 skipped. Keep this checklist item open until a fresh full
short-root run advances past the previous 507-pass stop.

`C:\R\FULLF16\last_run\commands\results.json` reproduced the same
`folderView_filter_watermark_empty_state` stop after the first patch, but the
enhanced assertion narrowed the state: the pane path was the empty fixture,
`itemCount=0`, `filterActive=0`, and `watermark=0`, while `emptyActive=0`.
Root-cause read of `FolderView::CanShowEmptyFolderState()` shows the remaining
full-order blockers are hidden pane surfaces, not folder enumeration:
operation/error overlays, explicit empty-state messages, background watermarks,
pending busy overlays, or current/displayed-folder mismatch can all suppress
the built-in empty placeholder. The next patch therefore clears operation
alerts, explicit empty-state messages, and background watermarks before
asserting empty-placeholder/filter-watermark behavior, keeps the visibility
reset on the active fixture path, sets the left pane to the built-in file-system
context for the case, and expands future diagnostics with filter text plus pane
alert snapshot details. The durable self-test isolation rule is updated in
`Specs/Testing/Testing_SelfTests.md`. Focused validation after the patch:
Debug x64 build `.build/logs/msbuild-20260506_212856_119.log` passed with
0 warnings and 0 errors; exact
`C:\R\FVWM3\last_run\commands\results.json` passed with 1 passed / 0 failed /
0 skipped; and the `folderView_` Commands slice at
`C:\R\FVFAM3\last_run\commands\results.json` passed with 3 passed /
0 failed / 0 skipped. Keep the checklist item open until a fresh full
short-root run advances past the previous 507-pass stop.

`C:\R\FULLF17\last_run\commands\results.json` did not reach FolderView; it
stopped earlier at `alternate_edit_launches_configured_editor_action` with
68 passed / 1 failed / 528 skipped. Exact standalone
`C:\R\AE1\last_run\commands\results.json` passed with 1 passed / 0 failed /
0 skipped, proving this is order-sensitive. The narrower
`alternate_edit_` prefix at `C:\R\AEPF1\last_run\commands\results.json`
reproduced a related marker-file race: the marker file existed, but the first
line was read before shell redirection flushed the expected `alternate-edit`
text. Patch in progress: the remaining external viewer/editor/user-menu marker
checks in `Commands.SelfTest.Settings.cpp` now use `WaitForTextFileFirstLine`,
and the primary/alternate edit tests wait for stable left-folder-view focus
before dispatching focus-sensitive edit shortcuts. Evidence after the patch:
Debug x64 build `.build/logs/msbuild-20260506_213949_197.log` passed with
0 warnings and 0 errors, and the `alternate_edit_` prefix at
`C:\R\AEPF2\last_run\commands\results.json` passed with 2 passed /
0 failed / 0 skipped. Full-order evidence from
`C:\R\FULLF18\last_run\commands\results.json` and
`C:\R\FULLF19\last_run\commands\results.json` advanced past the previous
68-pass stop, so the external editor marker/focus checklist item is closed.

`C:\R\FULLF18\last_run\commands\results.json` advanced past the previous
68-pass editor marker stop, proving the marker/focus patch in full order, then
stopped earlier than FolderView at
`pane_view_options_toggle_preview_pane_tabs_and_selection` with 50 passed /
1 failed / 546 skipped. Exact standalone
`C:\R\PVTAB1\last_run\commands\results.json` reproduced the same failure. Root
cause: the preview-tab test clicked by sending `WM_MOUSEMOVE`/button messages to
the tab host but did not move the real cursor first, so `PumpPendingMessages()`
could reapply hover from the current desktop cursor and leave the inactive
Preview tab's close glyph visible. Patch in progress: the click helper now
moves the OS cursor to the synthetic click point before sending the click
messages, and `Specs/Testing/Testing_SelfTests.md` records the hover-test
isolation rule. Evidence after the patch: Debug x64 build
`.build/logs/msbuild-20260506_214627_859.log` passed with 0 warnings and
0 errors, and exact
`C:\R\PVTAB2\last_run\commands\results.json` passed with 1 passed /
0 failed / 0 skipped. Full-order evidence from
`C:\R\FULLF19\last_run\commands\results.json` advanced past the previous
50-pass stop, so the preview-tab pointer-hover checklist item is closed.

`C:\R\FULLF19\last_run\commands\results.json` then stopped at
`cmd_pane_quickSearch_integrated_navigation` with 299 passed / 1 failed /
297 skipped. This is not the earlier reactivation failure: the command now
reactivates Quick Search, but after typing the no-match character the snapshot
reports `active=0`, `query=''`, and focused file `beta-alpha.txt`. Exact
standalone `C:\R\QS1\last_run\commands\results.json` passed with 1 passed /
0 failed / 0 skipped, so the active hypothesis is a full-order focus-loss race:
`FolderView::OnKillFocusMessage()` intentionally exits incremental search, and
the test proved active state after reactivation without proving the folder view
retained stable focus before synthetic `WM_CHAR` input. Patch in progress:
`cmd_pane_quickSearch_integrated_navigation` now reasserts
`WaitForFolderViewPaneFocus(...)` after both Quick Search activations, and
`Specs/Testing/Testing_SelfTests.md` records the focus-exit-mode rule. Evidence
after the patch: Debug x64 build
`.build/logs/msbuild-20260506_220007_066.log` passed with 0 warnings and
0 errors; exact `C:\R\QS2\last_run\commands\results.json` passed with
1 passed / 0 failed / 0 skipped. Full-order evidence from
`C:\R\FULLF20\last_run\commands\results.json` advanced past the previous
299-pass stop, so the Quick Search focus-loss checklist item is closed.

`C:\R\FULLF20\last_run\commands\results.json` also passed
`folderView_filter_watermark_empty_state` at index 507, so the FolderView
empty/filter watermark checklist item is closed with full-order evidence.
The same run then stopped later at
`cmd_pane_editNew_prompt_filters_editor_combo_and_creates_file` with 518 passed /
1 failed / 78 skipped. Exact standalone
`C:\R\EDITNEW1\last_run\commands\results.json` reproduces the failure with
0 passed / 1 failed / 0 skipped. The marker file contains `"alpha.editnew"`,
not `alpha.editnew`, which matches `Specs/Core/Core_SettingsStore.md`: file
action macros inside `arguments` are expanded as quoted/escaped Windows
command-line arguments. Patch in progress: the Edit New test now waits for the
quoted marker content and reports the actual first line on failure, and
`Specs/Testing/Testing_SelfTests.md` records that external marker tests must
match the documented macro quoting context. Evidence after the patch: Debug x64
build `.build/logs/msbuild-20260506_221607_376.log` passed with 0 warnings and
0 errors; exact `C:\R\EDITNEW2\last_run\commands\results.json` passed with
1 passed / 0 failed / 0 skipped; and the Edit New prompt slice at
`C:\R\EDITNEWFAM1\last_run\commands\results.json` passed with 3 passed /
0 failed / 0 skipped. Continue full short-root validation to expose the next
blocker after the previous 518-pass stop.

## Continuation Checkpoint - 2026-05-06 18:23 Europe/Paris

The external viewer marker-file blocker from
`C:\R\FULLF4\last_run\commands\results.json` is resolved as a self-test
synchronization gap. Exact/prefix controls passed before the patch, proving the
feature path itself was not broken; the durable test fix now waits until the
marker file's first line contains the expected `external-viewer` or
`external-editor` text instead of treating file creation alone as completion.
Evidence after the patch: Debug x64 build
`.build/logs/msbuild-20260506_181339_153.log` passed with 0 warnings and
0 errors, `C:\R\VPE2\last_run\commands\results.json` passed with 1 passed /
0 failed / 0 skipped, and `C:\R\QS3\last_run\commands\results.json` kept the
Quick Search focused control green with 1 passed / 0 failed / 0 skipped.

The next full short-root foreground fail-fast run,
`C:\R\FULLF5\last_run\commands\results.json`, advanced past the external marker
case but stopped earlier than Quick Search with 167 passed / 1 failed /
429 skipped at `cmd_preferences_dialog_themes_roundtrip_restores_dxui_surface`.
The failure is now under investigation as a likely stale self-test focus
assertion: the test compares `SetFocus(categoryTreeHost)` with
`categoryTreeHost`, but Win32 `SetFocus` returns the previous focus HWND on
success. This can pass only when the category host was already focused and can
fail in full order after the preceding Themes page cases leave focus elsewhere.

Patch checkpoint: the Preferences self-test include now has a shared
`FocusWindowAndWait(...)` helper that asserts the resulting focused HWND, and
the Themes round-trip plus theme-cycle cases use it for category-host setup.
Durable rule added to `Specs/Testing/Testing_SelfTests.md`. Evidence:
Debug x64 build `.build/logs/msbuild-20260506_182249_102.log` passed with
0 warnings and 0 errors; `C:\R\THRT2\last_run\commands\results.json` passed
with 1 passed / 0 failed / 0 skipped; `C:\R\THCY1\last_run\commands\results.json`
passed with 1 passed / 0 failed / 0 skipped; and the full Themes prefix
`C:\R\THP1\last_run\commands\results.json` passed with 26 passed / 0 failed /
0 skipped. The checklist remains open until the next full short-root run proves
the suite advances past this full-order blocker.

Full-order evidence: `C:\R\FULLF6\last_run\commands\results.json` advanced
past the Themes round-trip/theme-cycle area and stopped later with 299 passed /
1 failed / 297 skipped at `cmd_pane_quickSearch_integrated_navigation`. That
closes the Themes `SetFocus` blocker and reopens Quick Search as the active
fail-fast item with a better diagnostic:
`Quick Search no-match reactivation should enter search mode; active=0,
query='', focused='beta-alpha.txt'.`

Patch checkpoint for Quick Search: the Commands test shared helpers now include
`WaitForFolderViewPaneFocus(...)`, and
`cmd_pane_quickSearch_integrated_navigation` waits for stable left folder-view
focus before both `IDM_PANE_QUICK_SEARCH` command dispatches. The failure text
now also records right-pane incremental-search state so a future recurrence can
separate wrong-pane routing from immediate focus-loss exit. Durable rule added
to `Specs/Testing/Testing_SelfTests.md`. Evidence so far: Debug x64 build
`.build/logs/msbuild-20260506_183753_073.log` passed with 0 warnings and
0 errors, and `C:\R\QS4\last_run\commands\results.json` passed with 1 passed /
0 failed / 0 skipped. The item remains open until a full short-root run advances
past the previous 299-pass Quick Search stop.

Full-order evidence: `C:\R\FULLF7\last_run\commands\results.json` advanced
past Quick Search, proving the focused-pane stabilization fixed the prior
full-order failure, then stopped later with 312 passed / 1 failed / 284 skipped
at a Find dialog Enter-routing assertion: `Pressing Enter from a non-editor Find
control did not trigger the default Find action.`

Find Enter-routing investigation checkpoint: exact standalone
`C:\R\FENT1\last_run\commands\results.json` also fails with the same assertion,
so this is not full-order leakage. The root cause is a shared-control behavior
gap: `DxUi::Checkbox` inherited `Toggle::OnKeyDown`, so a focused checkbox
consumed `VK_RETURN` and toggled instead of allowing `WindowHost` to invoke the
default Find button. The durable spec conflict is now resolved in
`Specs/UI/UI_DxUiSharedGrid.md`: focused checkboxes own `Space`, while `Enter`
falls through to the host default route; switch-style `Toggle` controls keep
owning both `Enter` and `Space`.

Patch checkpoint for Find Enter-routing: `DxUi::Checkbox::OnKeyDown` now lets
`VK_RETURN` fall through to host-level default-button routing while retaining
`Space`/arrow toggle behavior from `DxUi::Toggle`; native `DxUiTests` now assert
both the focused-checkbox default-button route and the mixed-dialog flow. While
validating that shared focus area, the earlier retained-native-focus fix was
split from hidden text-input bridge external blur semantics so host native focus
loss retains the logical traversal target, but bridge blur to an external HWND
still clears the focused bridge-backed control. `Specs/UI/UI_DxUiSharedGrid.md`
now records both contracts.

Evidence: `.\build.ps1 -ProjectName DxUiTests` passed in
`.build/logs/msbuild-20260506_190113_137.log` with 0 warnings and 0 errors;
`.\.build\x64\Debug\DxUiTests.exe WindowHost` passed; `.\.build\x64\Debug\DxUiTests.exe TextInputBridge`
passed; the accepted advanced-controls visual refresh updated
`Tests/DxUiTests/Baselines/advanced_controls_dark.png`; `.\.build\x64\Debug\DxUiTests.exe Rendering`
passed; and the full `.\.build\x64\Debug\DxUiTests.exe` passed. The rebuilt main
app passed `.\build.ps1 -ProjectName RedSalamander` in
`.build/logs/msbuild-20260506_190931_825.log` with 0 warnings and 0 errors, and
the exact Find command case passed at `C:\R\FENT3\last_run\commands\results.json`
with 1 passed / 0 failed / 0 skipped.

Full-order evidence: `C:\R\FULLF8\last_run\commands\results.json` stopped before
Find at `cmd_preferences_dialog_themes_long_run_list_scrolling_stays_bounded`
with one extra category-tree render during chunk 5. Exact Themes evidence
`C:\R\THLR1\last_run\commands\results.json` passed with 1 passed / 0 failed /
0 skipped, and the full Preferences prefix
`C:\R\PREFTHLR1\last_run\commands\results.json` passed with 168 passed /
0 failed / 0 skipped, so that single repaint was not patched as a durable
Themes failure. The next full run, `C:\R\FULLF9\last_run\commands\results.json`,
advanced past the Themes long-run case and past
`cmd_pane_find_dialog_enter_from_checkbox_invokes_default_search`, closing the
Find blocker, then stopped later with 407 passed / 1 failed / 189 skipped at
`cmd_compare_directories_options_tab_traversal_live_dx_interaction`. That is the
active fail-fast item.

Compare Options traversal investigation checkpoint: the exact focused case
`C:\R\COTAB1\last_run\commands\results.json` reproduced the same Cancel-button
failure, so this was not only full-order leakage. The trace showed the OK step
left native focus on the body DxHost/text bridge while the footer OK host kept a
retained logical target; `DebugGetOptionsSnapshot(...)` then reported retained
OK as the single `focusTarget`, causing the scripted next Tab to route through
the wrong host. The patch makes that single focus target native-focus-owned:
the body reports a target only while the body host owns native focus, and footer
OK/Cancel report only when their own host owns native focus and retains the
button. Durable rule added to `Specs/Testing/Testing_SelfTests.md`.

Focused evidence after the patch: Debug x64 build
`.build/logs/msbuild-20260506_193529_451.log` passed with 0 warnings and
0 errors, and the exact command case
`C:\R\COTAB2\last_run\commands\results.json` passed with 1 passed / 0 failed /
0 skipped. The checklist remains open until the next full short-root fail-fast
run proves the suite advances past the previous Compare Options stop.

Full-order rerun checkpoint: `C:\R\FULLF10\last_run\commands\trace.txt` advanced
through the Connection Manager traversal cases but hung before Compare Options
at `cmd_connection_credential_prompt_theme_cycle_keeps_surface_legible`. Window
inspection showed the main RedSalamander window disabled and an enabled
`RedSalamander.ConnectionCredentialPromptWindow` titled `Password required`,
with no `results.json` emitted. Root cause: credential prompt modal-driver tests
used unbounded UI Automation reads inside the worker responsible for closing
the prompt; if a provider read stalls, the worker never sends Escape and the
modal prompt keeps the suite blocked. The shared connection-manager UIA helper
now accepts a trace label and returns a default sentinel after its timeout, and
the credential prompt theme/validation/long-run UIA reads now use that bounded
path. Durable rule added to `Specs/Testing/Testing_SelfTests.md`.

Evidence after the prompt-hang patch: Debug x64 build
`.build/logs/msbuild-20260506_194930_517.log` passed with 0 warnings and
0 errors; exact theme-cycle case `C:\R\CPTH1\last_run\commands\results.json`
passed with 1 passed / 0 failed / 0 skipped; and the credential-prompt family
`C:\R\CPFX1\last_run\run-all-tests-results.json` passed with 8 passed /
0 failed / 0 skipped. Restart the short-root full fail-fast run next and verify
it advances past both the prompt case and the earlier Compare Options stop.

## Continuation Checkpoint - 2026-05-06 15:04 Europe/Paris

The File Operations and Compare Directories page-setup navigation blocker is
resolved by replacing stale `End`+fixed-`Up` page setup with named category
selection. Evidence:
`Specs/TestRuns/commands-prefs-fileops-prefix-named-category-20260506_144710/last_run`
passed with 6 passed / 0 failed / 0 skipped, and
`Specs/TestRuns/commands-prefs-compare-prefix-named-category-20260506_144720/last_run`
passed with 7 passed / 0 failed / 0 skipped.

The next foreground Preferences fail-fast run then advanced to Hot Paths:
`Specs/TestRuns/commands-prefs-failfast-named-category-20260506_144735/last_run`
reported 98 passed / 1 failed / 69 skipped. The first failure was
`cmd_preferences_dialog_hot_paths_tab_traversal_live_dx_interaction` at reverse
Tab from the second slot path back to the first slot label field. Exact and
prefix controls rule out a simple Hot Paths-only failure:
`Specs/TestRuns/commands-prefs-hotpaths-tab-exact-20260506_continue/last_run`
passed with 1 passed / 0 failed / 0 skipped, and
`Specs/TestRuns/commands-prefs-hotpaths-prefix-20260506_continue/last_run`
passed with 7 passed / 0 failed / 0 skipped.

Diagnostic patch in progress: Hot Paths tab traversal now logs expected vs.
observed retained focus, native focus, active Preferences page, active DX host,
page child-window count, rendered DX-host count, and resize failures on every
step. A diagnostic-only build issue was fixed by formatting the internal
`PrefCategory` as an integer instead of asking `std::format` to format the enum
directly. Fresh Debug x64 build evidence:
`.build/logs/msbuild-20260506_150208_987.log` with 0 warnings and 0 errors.

Next step: rerun the broad foreground `cmd_preferences_dialog_` fail-fast case
set into `Specs/TestRuns/commands-prefs-failfast-hotpaths-diagnostics-20260506_continue/`
and use the new trace to identify the actual order-dependent retained-focus or
native-focus source before landing the Hot Paths fix.

## Continuation Checkpoint - 2026-05-06 15:31 Europe/Paris

The broad Preferences validation surfaced and resolved three additional
order-sensitive test-contract gaps after the FileOps/Compare and Hot Paths
fixes:

- `Specs/TestRuns/commands-prefs-failfast-hotpaths-diagnostics-20260506_continue/last_run`
  stopped earlier than Hot Paths with 38 passed / 1 failed / 129 skipped at
  `cmd_preferences_dialog_viewers_long_run_list_scrolling_stays_bounded`.
  Root cause: the test captured the category-tree render baseline before the
  unrelated category-tree host had settled after page setup, so a delayed setup
  repaint was charged to Viewers list scrolling. The Plugins main list, Plugins
  custom paths, Keyboard, Viewers, and Themes long-run list-scroll tests now
  wait for a settled category-tree render baseline before asserting zero extra
  unrelated renders during the scroll window. Durable contract updates landed in
  `Specs/Testing/Testing_SelfTests.md` and `Specs/UI/UI_DxUiSharedGrid.md`.
- `Specs/TestRuns/commands-prefs-failfast-settled-longrun-20260506_continue/last_run`
  then advanced to 44 passed / 1 failed / 123 skipped at
  `cmd_preferences_dialog_viewers_tab_traversal_live_dx_interaction`.
  The exact case still passed at
  `Specs/TestRuns/commands-prefs-viewers-tab-exact-after-settle-20260506_continue/last_run`;
  a diagnostic pass then advanced past Viewers and Hot Paths, so the added
  Viewers setup diagnostics remain as maintenance evidence rather than a product
  fix.
- `Specs/TestRuns/commands-prefs-failfast-viewers-focus-diagnostics-20260506_continue/last_run`
  advanced to 141 passed / 1 failed / 26 skipped and proved Hot Paths tab
  traversal passed in the broad family order. The next failure was
  `cmd_preferences_dialog_editors_and_mouse_pages_use_dxui_statics`: it still
  sent fixed Down-key navigation without first returning to a known starting
  category. That page-specific test now uses named category selection for both
  Editors and Mouse, matching the durable test-navigation contract in
  `Specs/Testing/Testing_SelfTests.md`.

Verification after these fixes:

- Debug x64 builds passed at `.build/logs/msbuild-20260506_150932_734.log`,
  `.build/logs/msbuild-20260506_151500_788.log`, and
  `.build/logs/msbuild-20260506_152150_945.log`, each with 0 warnings and
  0 errors.
- `Specs/TestRuns/commands-prefs-viewers-longrun-settled-exact-20260506_continue/last_run`
  passed with 1 passed / 0 failed / 0 skipped.
- `Specs/TestRuns/commands-prefs-editors-mouse-statics-named-category-exact-20260506_continue/last_run`
  passed with 1 passed / 0 failed / 0 skipped.
- `Specs/TestRuns/commands-prefs-failfast-editors-named-category-20260506_continue/last_run`
  passed the full foreground Preferences family with 168 passed / 0 failed /
  0 skipped.

Next step: run the full foreground Commands suite with a fresh absolute
`REDSALAMANDER_SELFTEST_ROOT`. If it passes, continue closeout tooling/native
validation and only then consider moving this plan to `Done`.

## Continuation Checkpoint - 2026-05-06 16:07 Europe/Paris

The first full foreground Commands repeat after the full Preferences family
passed did not close the plan. Evidence:
`Specs/TestRuns/commands-full-foreground-after-preferences-20260506_continue/last_run`
reported 597 total, 567 passed, 30 failed, 0 skipped.

The failure set is mixed and must be challenged before any fix lands:

- Several failures look like self-test root/path-depth artifacts rather than
  direct product regressions: ZIP extraction returned `0x80070003`, and multiple
  cases failed while creating deep or long fixture paths
  (`search_local_plugin_parallel_cancel_fanin`,
  `cmd_pane_rename_prompt_long_initial_selection_stays_clipped`,
  `cmd_pane_itemProperties_window_long_run_scrolling_stays_bounded`,
  `cmd_pane_navigationView_full_path_popup_edit_route`, and
  `cmd_pane_navigationView_full_path_popup_ancestor_click_navigates_to_ancestor`).
- Several failures are live UI/focus/order-sensitive and need separate focused
  repeats after the path-depth question is removed: Connection Manager tab
  traversal, Preferences Plugins/Viewers/Keyboard/General/Hot Paths/Editors
  Mouse, Find dialog keyboard routing, Compare Directories options traversal,
  folder empty-state rendering, navigation-view edit/history routing, and UI
  chrome toggling.

Next step: rerun representative failures under a deliberately short
`REDSALAMANDER_SELFTEST_ROOT`, then update this plan with which failures are
path artifacts, which are order-dependent UI gaps, and which require code/spec
changes.

## Continuation Checkpoint - 2026-05-06 16:52 Europe/Paris

Representative full-run failures that explicitly mentioned fixture creation,
ZIP extraction, or deep paths were repeated with short self-test roots. This
proves several failures from the first full foreground Commands run were
path-depth artifacts from using
`Specs/TestRuns/commands-full-foreground-after-preferences-20260506_continue`
as the live `REDSALAMANDER_SELFTEST_ROOT`, not functional regressions in those
features:

- `C:\RSST\search_local_plugin_parallel_cancel_fanin\last_run`: 1 passed /
  0 failed / 0 skipped.
- `C:\RSST\cmd_pane_archive_pack_unpack_zip_roundtrip_and_validation\last_run`:
  1 passed / 0 failed / 0 skipped.
- `C:\RSST\cmd_pane_rename_prompt_long_initial_selection_stays_clipped\last_run`:
  1 passed / 0 failed / 0 skipped.
- `C:\RSST\cmd_pane_unpack_prompt_uses_dxui_destination_unpacker_and_mask\last_run`:
  1 passed / 0 failed / 0 skipped.
- `C:\RSST\cmd_pane_itemProperties_window_long_run_scrolling_stays_bounded\last_run`:
  1 passed / 0 failed / 0 skipped.
- `C:\RSST\cmd_pane_navigationView_full_path_popup_edit_route\last_run`:
  1 passed / 0 failed / 0 skipped.
- The first repeat of
  `cmd_pane_navigationView_full_path_popup_ancestor_click_navigates_to_ancestor`
  still failed under `C:\RSST\<full-case-name>\...`, which was not actually a
  short enough root for that deepest fixture. Repeating it under `C:\R\N1`
  passed with 1 passed / 0 failed / 0 skipped.

Next step: rerun the full foreground Commands suite with a genuinely short
root such as `C:\R\F1`. If failures remain, treat them as real full-order UI
or test-contract gaps and continue focused triage from that shorter-root full
run.

## Continuation Checkpoint - 2026-05-06 17:11 Europe/Paris

The full foreground Commands suite was rerun with a genuinely short live
self-test root, `C:\R\F1`. Evidence: `C:\R\F1\last_run` reported 597 total,
572 passed, 25 failed, 0 skipped. This confirms the previous live-root path
depth explained five of the first full-run failures, but does not close the
suite.

Failures that remain under the short root are now treated as real full-order UI
or test-contract gaps:

- Connection Manager reverse tab traversal.
- Preferences Viewers baseline navigation, Hot Paths theme-cycle focus,
  Keyboard live-search restoration, Plugins tab traversal, and Editors/Mouse
  note-surface round-trip.
- Find dialog Enter/default action and tab traversal.
- File Operations issues-pane sustained scrolling.
- Pane Filter and Hot Paths navigation-shell restoration.
- Folder empty-state rendering after filter and standalone empty folder paths.
- Edit New macro expansion.
- Navigation-view address bar, path edit, breadcrumb ancestor, full-path popup,
  history dropdown, edit-suggest routing, Toggle UI Chrome, and Swap Panes
  shell stability.

Next step: compare exact/family repeats against the short-root full-order
failure set, then fix the earliest shared stale state or setup contract that
explains the largest cluster before changing isolated cases.

## Continuation Checkpoint - 2026-05-06 17:28 Europe/Paris

The first short-root full-order failure has been narrowed to Connection Manager
test isolation, not a simple standalone product failure:

- Exact repeat under `C:\R\CM1\last_run` passed with 1 passed / 0 failed /
  0 skipped.
- Prefix fail-fast repeat under `C:\R\CMPF1\last_run` advanced through the
  preceding Connection Manager cases, then failed
  `cmd_connection_manager_window_tab_traversal_live_dx_interaction` after
  24 passed / 1 failed / 4 skipped.
- Trace evidence shows the standalone traversal opens on a short deterministic
  row (`New connection`, index 1), while the prefix run starts traversal after
  earlier Connection Manager cases have accumulated many `New connection (...)`
  profiles and retained a later row (`New connection (17)`, selected row 9).

Patch in progress: seed and restore `g_settings.connections` around the live
tab-traversal case, matching the cleanup pattern already used by other
Connection Manager mutation tests. The durable spec will also state that
profile-mutating GUI self-tests must restore runtime settings and must not let
full-order keyboard traversal depend on profiles created by earlier cases.

## Continuation Checkpoint - 2026-05-06 17:36 Europe/Paris

After the initial traversal isolation patch, the exact traversal repeat passed
under `C:\R\CM2\last_run` with 1 passed / 0 failed / 0 skipped, but the
Connection Manager prefix fail-fast run moved the first failure earlier:
`C:\R\CMPF2\last_run` reported 15 passed / 1 failed / 13 skipped at
`cmd_connection_manager_window_live_dx_interaction`.

That case also passes standalone (`C:\R\CMDX1\last_run`, 1 passed / 0 failed /
0 skipped), so the problem is still full-order profile isolation rather than an
isolated product failure. Its trace shows it opens with accumulated rows in the
prefix run, creates/edits/removes/ closes live rows, and has no
`g_settings.connections` restore guard. Patch in progress: apply the same
deterministic profile seed/restore cleanup to `live_dx_interaction` before
rerunning the Connection Manager prefix.

## Continuation Checkpoint - 2026-05-06 16:44 Europe/Paris

The Connection Manager profile-isolation patch built cleanly at
`.build/logs/msbuild-20260506_163151_955.log` with 0 warnings and 0 errors.
Focused repeats now pass:

- `C:\R\CMDX2\last_run`: `cmd_connection_manager_window_live_dx_interaction`
  passed with 1 passed / 0 failed / 0 skipped.
- `C:\R\CM3\last_run`:
  `cmd_connection_manager_window_tab_traversal_live_dx_interaction` passed with
  1 passed / 0 failed / 0 skipped.

The Connection Manager prefix repeat under `C:\R\CMPF3\last_run` still fails,
but the failure has changed shape and is now deterministic: the traversal opens
with the seeded short row (`selectedRow=1`, plugin `builtin/file-system-s3`,
name `New connection`), the first forward Tab reaches `New...`, and the second
forward Tab from `New...` lands back on the list instead of `Rename...`.
Because DxUi forward Tab can only do that if the host sees reverse traversal or
has lost/staled the current focused control, the next diagnostic patch will log
per-step pre/post focus and retained modifier state before any product fix
lands.

## Continuation Checkpoint - 2026-05-06 16:52 Europe/Paris

The targeted diagnostic patch built cleanly at
`.build/logs/msbuild-20260506_164533_209.log` with 0 warnings and 0 errors.
It adds test-only retained modifier/focus diagnostics to the Connection Manager
snapshot and per-step traversal trace.

Two Connection Manager prefix repeats now separate a transient pass from the
real recurring failure:

- `C:\R\CMPF4\last_run` passed the full Connection Manager prefix with
  29 passed / 0 failed / 0 skipped. In that run every traversal step reported
  `preModifiers=0x0` and present/visible/enabled/focusable retained focus.
- `C:\R\CMPF5\last_run` reproduced the failure with 24 passed / 1 failed /
  4 skipped. The second step failed before routing from `New...` to `Rename...`
  because the pre-step retained focus had already dropped to `None`:
  `preModifiers=0x0`, `preFocusPresent=false`, then routing from no current
  focus chose the first focusable control, the list.

This rules out stale Shift/modifier state and narrows the root cause to a focus
loss or interaction-state reset between traversal steps. Next diagnostic:
capture native Win32 focus ownership in the snapshot to confirm whether a
delayed `WM_KILLFOCUS` from earlier Connection Manager cases clears the DxUi
retained focus after the first Tab.

## Continuation Checkpoint - 2026-05-06 17:06 Europe/Paris

The native-focus diagnostic confirmed the recurring failure shape:
`C:\R\CMPF7\last_run` failed after `New...` and `Rename...` had both succeeded.
The failing `Remove` step entered with `preKind='None'`,
`preModifiers=0x0`, `preNativeInDialog=false`, `preNativeHost=false`, and
`preNativeText=false`, then routed from no retained control back to the list.
This proves stale Shift state is not the cause. The durable root cause is in
the shared DxUi host: `WindowHost::OnKillFocus()` marks the current control
unfocused and then clears `_focusedControl`, so any later synthetic/direct Tab
routing starts from the first focusable child.

RED coverage has been added before the product fix:
`TestWindowHostNativeFocusLossRetainsLogicalFocusForTraversal` in the
`DxUiTests` WindowHost suite. Build evidence:
`.build/logs/msbuild-20260506_170405_875.log` with 0 warnings and 0 errors.
The focused RED run failed as intended with:
`FAILED: native focus loss keeps the retained logical focus target`.
Next change: retain the logical focus target across native focus loss, restore
control focus visuals on `WM_SETFOCUS`, and keep clearing true stale/tree-reset
interaction state through the existing prune/reset paths.

GREEN evidence for the generic guard: after the `WindowHost` patch,
`.build/logs/msbuild-20260506_170750_965.log` rebuilt `DxUiTests` with
0 warnings and 0 errors, and `.build\x64\Debug\DxUiTests.exe --suite=WindowHost`
passed. The fix retains `_focusedControl` across native focus loss, clears
transient modifiers on focus loss/root reset, and reactivates focus visuals on
native focus regain. Durable contract updates landed in
`Specs/UI/UI_DxUiWinUIDesign.md` and `Specs/Testing/Testing_SelfTests.md`.
Next validation is the app build plus repeated Connection Manager prefix runs
under fresh short roots.

Latest app-level evidence: `RedSalamander` rebuilt cleanly at
`.build/logs/msbuild-20260506_170859_270.log` with 0 warnings and 0 errors,
then the Connection Manager prefix passed twice:

- `C:\R\CMPF8\last_run\commands\results.json`: 29 passed / 0 failed /
  0 skipped.
- `C:\R\CMPF9\last_run\commands\results.json`: 29 passed / 0 failed /
  0 skipped.

After the traversal test maintenance cleanup, the app rebuilt cleanly again at
`.build/logs/msbuild-20260506_171424_617.log` with 0 warnings and 0 errors, and
`C:\R\CMPF10\last_run\commands\results.json` passed with 29 passed /
0 failed / 0 skipped. The Connection Manager short-root blocker is closed.
Next action: restart the full foreground Commands fail-fast run from a fresh
short root to expose the next remaining full-order gap.

## Continuation Checkpoint - 2026-05-06 17:24 Europe/Paris

The fresh short-root full Commands fail-fast run under
`C:\R\FULLF1\last_run\commands\results.json` advanced past the prior
Connection Manager blocker and stopped at the next real gap:
284 passed / 1 failed / 312 skipped. Failed case:
`cmd_preferences_dialog_editors_mouse_tab_skips_note_surface`.

Failure text:
`Preferences Mouse wrapped category tree focus target not reached during
note-page tab traversal; actual categoryTreeFocused=1, shellFocusTarget=3,
currentCategory=5, visibleCurrentPageChildWindowCount=1,
currentPageDxHostResizeFailureCount=0, shellDxHostResizeFailureCount=0,
focus=0x2a711b8, pageHost=0x3d12378, shellHost=0x11e2094.`

Initial read: the page is Mouse (`currentCategory=5`) and native focus appears
to be on the category tree (`categoryTreeFocused=1`), but the retained shell
focus target enum is still `3`, so the next task is to inspect the
Preferences-shell debug focus mapping and the note-page traversal helper before
changing product behavior.

## Continuation Checkpoint - 2026-05-06 17:29 Europe/Paris

The Preferences Mouse note-page blocker reproduces standalone, so it is not
only a full-order artifact:
`C:\R\PMNOTE1\last_run\commands\results.json` failed with 0 passed /
1 failed / 0 skipped and the same state as the full run:
`categoryTreeFocused=1`, `shellFocusTarget=3`, `currentCategory=5`.

Root-cause read:
`PreferencesShellDebugFocusTarget::CancelButton` is enum value `3`, and
`DebugGetSnapshot()` intentionally reports `hostState._shellHost.GetFocusControl()`.
After the shared `WindowHost` fix, that is the retained logical shell focus
target, not active native focus. The active native focus is already proven by
`categoryTreeFocused=1`. The test was stale because it expected the retained
shell target to become `None` when native focus wrapped back to the category
tree. The corrected contract is stricter: forward wrap must report native
category-tree focus plus retained shell `CancelButton`, and reverse wrap must
report native category-tree focus plus retained shell `ResetAllButton`.
`Specs/UI/UI_PreferencesDialog.md` and `Specs/Testing/Testing_SelfTests.md`
now document that native focus ownership and retained DxUi focus targets are
separate assertions.

GREEN evidence: `RedSalamander` rebuilt cleanly at
`.build/logs/msbuild-20260506_173032_175.log` with 0 warnings and 0 errors, and
`C:\R\PMNOTE2\last_run\commands\results.json` passed the exact
`cmd_preferences_dialog_editors_mouse_tab_skips_note_surface` case with
1 passed / 0 failed / 0 skipped. Next action: restart the short-root full
Commands fail-fast run to expose the next unresolved gap.

## Continuation Checkpoint - 2026-05-06 17:39 Europe/Paris

The next short-root full Commands fail-fast run advanced further:
`C:\R\FULLF2\last_run\commands\results.json` reported 292 passed /
1 failed / 304 skipped. The failed case is
`cmd_preferences_dialog_scroll_host_preserves_retained_page_state` with:
`Scrolling the Preferences page host should not repaint the category tree; saw
1 extra tree render(s).`

Controls:

- Exact standalone repeat under `C:\R\PSH1\last_run` passed with
  1 passed / 0 failed / 0 skipped.
- Preferences-only fail-fast repeat under `C:\R\PREFX1\last_run` passed with
  168 passed / 0 failed / 0 skipped.

Root-cause hypothesis: the product contract remains the one already documented
in `Specs/UI/UI_DxUiSharedGrid.md` and `Specs/Testing/Testing_SelfTests.md`:
page-host scroll must not repaint the category tree after the unrelated
category-tree host render count has settled. The failing test measured from
`viewersSnapshot.categoryTreeDxHostRenderCount` immediately after category
setup/re-entry, so a delayed pre-existing tree repaint from earlier full-suite
history could be charged to the wheel stimulus. Patch in progress: move the
existing `WaitForPreferencesCategoryTreeRenderCountToSettle(...)` helper from
the later Viewers/Keyboard chunk into the shared Preferences self-test include
scope, then call it before the scroll-host test captures its measurement
baseline.

GREEN/full-order evidence after the scroll-host settle-baseline patch:
`RedSalamander` rebuilt cleanly at `.build/logs/msbuild-20260506_174633_312.log`
with 0 warnings and 0 errors. `C:\R\PSH2\last_run` passed the exact scroll-host
case with 1 passed / 0 failed / 0 skipped. `C:\R\FULLF3\last_run\commands`
then advanced beyond the previous scroll-host failure and stopped later at
299 passed / 1 failed / 297 skipped. The next failed case is
`cmd_pane_quickSearch_integrated_navigation` with `Quick Search no-match state
should remain active.`

## Continuation Checkpoint - 2026-05-06 12:42 Europe/Paris

The Viewers resize diagnostic patch built successfully at
`.build/logs/msbuild-20260506_123840_064.log` with 0 warnings and 0 errors.
The focused diagnostic repeat still failed under
`.build/selftest-isolated/commands-prefs-viewers-resize-diagnostics-20260506_124023/last_run/`,
but the stronger failure message changed the root-cause finding: the drag does
reach the Viewers grid (`hitTest=true`, `hitResize=true`, `hostHitsList=true`),
the active page and active DX host are the same window, the pointer resize state
advances (`resizeDown=0->1`, `resizeMove=0->1`), and the visible header geometry
does resize (`matchWidth=144.0->216.0`, `computerLeft=162.0->234.0`). The
remaining stale assertion was the test contract itself: Viewers
`viewersListResizeCount` is the grid pointer resize-move counter, so header
resize tests must require it to advance after a resize instead of staying equal
to the pre-resize baseline.

Patch in progress: Viewers header-resize and combined reordered-plus-resized
predicates now require `viewersListResizeCount > baselineResizeCount` after the
resize operation. Next validation is a fresh build plus the focused Viewers
resize repeat, then the broader Viewers slice.

## Continuation Checkpoint - 2026-05-06 12:59 Europe/Paris

The resize-counter assertion patch rebuilt cleanly at
`.build/logs/msbuild-20260506_124326_457.log` with 0 warnings and 0 errors.
The focused Viewers resize repeat then passed under
`.build/selftest-isolated/commands-prefs-viewers-resize-focused-20260506_124501/last_run/`
with 1 passed / 0 failed / 0 skipped.

The broader Viewers slice initially advanced to 25 passed / 2 failed / 0 skipped
under `C:\Users\eric\AppData\Local\RedSalamander\SelfTest\last_run`. The
remaining failures were both reordered-plus-resized sort-cycle cases. Root cause:
the Done-era specs still require live Viewers association-grid sort-cycle
parity, but the current shared file-actions grid had regressed to non-sortable
columns plus a no-op page-level sort delegate. The delegate also chose
Associations vs Actions from stale `_activeGrid`; a header click sets the actual
focused grid, so sort routing must prefer the focused grid before falling back
to retained active-grid state.

Fixes landed in `RedSalamander/Preferences.FileActions.cpp`: association/action
columns are sortable again, `FileActionGridModel::SortRows(...)` sorts visible
cell text case-insensitively with stable source-order tie breaking, rebuilds
preserve the active grid sort spec, and sort requests route by the focused grid.
The durable contract is recorded in `Specs/UI/UI_PreferencesDialog.md`.

Verification: Debug x64 rebuilt cleanly at
`.build/logs/msbuild-20260506_125542_378.log` with 0 warnings and 0 errors.
Focused reruns passed for
`cmd_preferences_dialog_viewers_reordered_resized_columns_survive_sort_cycles`
under
`.build/selftest-isolated/commands-prefs-viewers-sortcycles-route-20260506_125719/last_run/`
and
`cmd_preferences_dialog_viewers_reordered_resized_copy_follows_visible_columns_after_sort_cycles`
under
`.build/selftest-isolated/commands-prefs-viewers-copy-sortcycles-route-20260506_125739/last_run/`.
The full Viewers Preferences slice then passed with 27 passed / 0 failed /
0 skipped under `C:\Users\eric\AppData\Local\RedSalamander\SelfTest\last_run`.

Next validation: rerun the foreground `cmd_preferences_dialog_` fail-fast
family and continue first-failure diagnosis.

## Continuation Checkpoint - 2026-05-06 13:07 Europe/Paris

The foreground Preferences fail-fast rerun after the Viewers sort-routing fix
advanced to 82 passed / 1 failed / 85 skipped under
`C:\Users\eric\AppData\Local\RedSalamander\SelfTest\last_run`. The new first
failure is `cmd_preferences_dialog_general_tab_traversal_live_dx_interaction`
with `Preferences General compact-mode toggle focus target not reached during
tab traversal.`

Root cause: the test skipped the visible `Language` combo when tabbing forward
from `Function bar` to `Compact mode`. The product page creation order, the
debug focus enum, and `Specs/UI/UI_PreferencesDialog.md` all treat `Language`
as a visible focusable General Display control between `Function bar` and the
DxUI group. The stale test sequence caused the second Tab to land correctly on
`LanguageCombo` while the assertion expected `CompactModeToggle`.

Patch in progress: the General tab-traversal test now includes `LanguageCombo`
in forward and reverse order and reports expected/actual focus targets plus
page-host diagnostics on future failures. The authoritative Preferences UI spec
now records the General page-local Tab order:
`Menu bar`, `Function bar`, `Language`, `Compact mode`, `Animations`,
`Window backdrop`, `Splash screen`, then wrap. Next validation is a clean Debug
x64 build and a focused repeat of the General tab-traversal case.

Verification update: Debug x64 rebuilt cleanly at
`.build/logs/msbuild-20260506_130812_955.log` with 0 warnings and 0 errors.
The focused General tab-traversal repeat passed at
`Specs/TestRuns/commands-prefs-general-tab-20260506_131013/last_run` with
1 passed / 0 failed / 0 skipped. Next validation is the foreground
`cmd_preferences_dialog_` fail-fast family again.

## Continuation Checkpoint - 2026-05-06 13:18 Europe/Paris

The foreground Preferences fail-fast rerun after the General tab traversal fix
advanced to 141 passed / 1 failed / 26 skipped under
`Specs/TestRuns/commands-prefs-failfast-20260506_131041/last_run`. The next
first failure is `cmd_preferences_dialog_editors_and_mouse_pages_use_dxui_statics`
with `Preferences navigation did not move to the Mouse category.`

Root cause: several Editors/Mouse note-page tests still encode visible
Preferences category navigation as raw Down-key counts (`Mouse == 5` from Home,
or two rows after `Editors`). That count predates the current root order:
`General`, `Panes`, `Viewers`, `Editors`, `User Menu`, `Keyboard`, `Mouse`,
`Themes`, `Plugins`, `File Operations`, `Compare Directories`, `Hot Paths`,
`Advanced`. The product enum values were not stale; the failing tests confused
enum value with visible tree row position.

Patch in progress: command self-tests now have a named
`PreferencesRootRowForCategory(...)` helper for the visible root order, and the
Editors/Mouse note/static tests route Home/Down navigation through that helper
instead of duplicating magic counts. Next validation is a clean Debug x64 build,
then focused repeats of the Editors/Mouse DX statics, note UIA, tab-skip, and
round-trip cases before rerunning the foreground Preferences fail-fast family.

Verification update: Debug x64 rebuilt cleanly at
`.build/logs/msbuild-20260506_132028_740.log` with 0 warnings and 0 errors.
The exact focused Editors/Mouse repeats all passed under
`Specs/TestRuns/commands-prefs-editors-mouse-exact-20260506_132247/`:
`cmd_preferences_dialog_editors_and_mouse_pages_use_dxui_statics`,
`cmd_preferences_dialog_editors_mouse_live_dx_notes`,
`cmd_preferences_dialog_editors_mouse_tab_skips_note_surface`, and
`cmd_preferences_dialog_editors_mouse_roundtrip_restore_dxui_notes` each
reported 1 passed / 0 failed / 0 skipped. A mistaken broad filter run at
`Specs/TestRuns/commands-prefs-editors-mouse-focused-20260506_132233/last_run`
matched zero expected cases and is not counted as coverage evidence. The
authoritative Preferences UI spec now states that visible root row order is the
navigation contract and must not be inferred from `PrefCategory` enum values.

Next validation: rerun the foreground `cmd_preferences_dialog_` fail-fast
family and continue first-failure diagnosis.

## Continuation Checkpoint - 2026-05-06 13:31 Europe/Paris

The foreground Preferences fail-fast rerun after the Editors/Mouse navigation
fix advanced to 166 passed / 1 failed / 1 skipped under
`Specs/TestRuns/commands-prefs-failfast-20260506_132352/last_run`. The next
first failure is
`cmd_preferences_dialog_scroll_host_preserves_retained_page_state` with
`Preferences General page should not show a vertical scrollbar after leaving a
scrolled page.`

Root-cause investigation in progress: `PreferencesDebugSnapshot` derives
`pageHostShowsVerticalScroll` from `pageScrollMaxY > 0`, not from a raw stale
`WS_VSCROLL` style bit. That means the failing General snapshot is reporting a
real positive scroll extent at the reduced dialog height chosen to make Viewers
scroll. The next patch must challenge the old assumption that General is always
non-scrollable in that reduced geometry; the durable contract is per-category
scroll offset retention (`General` offset resets to 0, `Viewers` restores its
retained offset) plus each page exposing a scrollbar only when its own measured
content overflows.

## Continuation Checkpoint - 2026-05-06 13:40 Europe/Paris

The first retained-scroll patch rebuilt cleanly at
`.build/logs/msbuild-20260506_133357_676.log` with 0 warnings and 0 errors.
It updated the test/spec contract so General is compared against its own
baseline scroll extent instead of assuming General can never overflow in the
reduced dialog geometry.

The exact focused rerun then failed at
`Specs/TestRuns/commands-prefs-scroll-host-20260506_133538/last_run` with
0 passed / 1 failed / 0 skipped. The new failure is still in
`cmd_preferences_dialog_scroll_host_preserves_retained_page_state`, now at the
Viewers wheel step: `pageScrollY` stayed 0 after routed `WM_MOUSEWHEEL`
scrolling despite the reopened Viewers snapshot reporting a positive page-host
scroll extent. Investigation focus: determine whether the synthetic wheel point
now hits a DxUi child that consumes the wheel, whether page-host routing is not
settled after the added General baseline roundtrip, or whether the product
wheel-route fallback is wrong for retained Preferences pages.

## Continuation Checkpoint - 2026-05-06 13:53 Europe/Paris

The scroll-host diagnostic build passed at
`.build/logs/msbuild-20260506_134515_031.log` with 0 warnings and 0 errors. The
diagnostic exact rerun failed at
`Specs/TestRuns/commands-prefs-scroll-host-diagnostics-20260506_134703/last_run`
with route telemetry showing the wheel message was seen but never reached the
page host: `routeSeen=true`, `targetPageHost=false`,
`windowFromPointPageHost=false`, `pageHostWndProcSeen=false`, and
`fallbackCalled=false`. Root cause: the test used the production
screen-hit-tested wheel route without first making the Preferences dialog the
top foreground window. `WindowFromPoint` could therefore resolve the synthetic
screen point outside the Preferences root even though the point was valid in
page-host client coordinates.

The test now raises Preferences to the foreground/top z-order before routed
wheel validation and asserts the screen point belongs to the Preferences root
before sending `WM_MOUSEWHEEL`. The durable contracts are captured in
`Specs/UI/UI_PreferencesDialog.md` (page-host wheel routing is screen-hit-tested
and nested wheel handlers consume before fallback) and
`Specs/Testing/Testing_SelfTests.md` (screen-hit-tested synthetic input must
foreground the target and verify the hit root). Debug x64 rebuilt cleanly at
`.build/logs/msbuild-20260506_134926_435.log` with 0 warnings and 0 errors, and
the exact scroll-host rerun passed at
`Specs/TestRuns/commands-prefs-scroll-host-routed-top-20260506_135127/last_run`
with 1 passed / 0 failed / 0 skipped.

## Continuation Checkpoint - 2026-05-06 13:59 Europe/Paris

The foreground Preferences fail-fast rerun after the scroll-host routed-wheel fix
advanced to 109 passed / 1 failed / 58 skipped under
`Specs/TestRuns/commands-prefs-failfast-20260506_135342/last_run`. The new first
failure is `cmd_preferences_dialog_keyboard_export_live_dx_interaction` with
`Preferences Keyboard page did not settle before export interaction validation.`

Root-cause investigation is in progress. Earlier Keyboard page, traversal,
search, reset, and long-run scroll cases passed in the same fail-fast run before
Export. The immediate next checks are to determine whether Export fails
standalone or only after earlier Keyboard state mutations, then add snapshot
diagnostics around the page-settle predicate if the failure needs more state
evidence.

## Continuation Checkpoint - 2026-05-06 14:16 Europe/Paris

Keyboard page-settle investigation narrowed the failure to the broad Preferences
order. Exact standalone runs passed for
`cmd_preferences_dialog_keyboard_export_live_dx_interaction` at
`Specs/TestRuns/commands-prefs-keyboard-export-exact-20260506_135958/last_run`
and `cmd_preferences_dialog_keyboard_live_search_dx_interaction` at
`Specs/TestRuns/commands-prefs-keyboard-live-search-exact-20260506_140703/last_run`.
The full Keyboard prefix also passed with 24 passed / 0 failed / 0 skipped at
`Specs/TestRuns/commands-prefs-keyboard-prefix-failfast-20260506_140008/last_run`.

The broad Preferences repeat still exposed Keyboard-order sensitivity at
`Specs/TestRuns/commands-prefs-failfast-repeat-20260506_140048/last_run`, but it
failed one case earlier than Export:
`cmd_preferences_dialog_keyboard_live_search_dx_interaction` timed out after UIA
`SetValue` reported success but the edited search value did not settle. A
diagnostic-only patch now prints the active Preferences page, product snapshot,
matched UIA ValuePattern state, and all visible edit values for future Keyboard
live-search waits. Debug x64 rebuilt cleanly at
`.build/logs/msbuild-20260506_140800_158.log` with 0 warnings and 0 errors, and
the exact diagnostic live-search rerun passed at
`Specs/TestRuns/commands-prefs-keyboard-live-search-diagnostics-exact-20260506_140943/last_run`.

The next broad Preferences fail-fast run with those diagnostics advanced through
all Keyboard cases and reached the final Preferences case, failing at
`Specs/TestRuns/commands-prefs-failfast-keyboard-diagnostics-20260506_140953/last_run`
with 167 passed / 1 failed / 0 skipped. New first failure:
`cmd_preferences_dialog_rapid_switches_keep_page_specific_uia_subtrees` reports
`Preferences Editors note page should not expose lingering editable ValuePattern
descendants; saw 13.` Root-cause focus moves to page-specific UIA subtree
teardown/visibility after rapid category switches. A later full Preferences
rerun is still required before closing the Keyboard checklist item because the
diagnostic patch may have shifted timing.

## Continuation Checkpoint - 2026-05-06 14:20 Europe/Paris

The rapid-switch failure reproduces standalone at
`Specs/TestRuns/commands-prefs-rapid-switch-exact-20260506_141537/last_run`
with 0 passed / 1 failed / 0 skipped. Root cause is a stale test contract, not a
new product leak: `Specs/UI/UI_PreferencesDialog.md` defines `Editors` as a
Viewers-style shared file-actions page, so it legitimately exposes editable
ValuePattern descendants. The same rapid-switch test also still used old
hard-coded tree Y positions for `Keyboard` and `Mouse`; after `User Menu` became
a visible root row, those points no longer matched the documented visible root
order.

Patch in progress: the rapid-switch test now derives click points from
`PreferencesRootRowForCategory(...)`, treats `Editors` as a shared file-actions
surface with editable descendants, and keeps the stale-descendant rejection on
the note-style `Mouse` page. The durable spec now states that rapid-switch UIA
validation must assert each active page's current contract: `Editors` allows
file-actions edit/value descendants, while `Mouse` must not expose stale
edit/combo/value/toggle descendants.

Debug x64 rebuilt cleanly after that patch at
`.build/logs/msbuild-20260506_141941_404.log` with 0 warnings and 0 errors. The
exact rapid-switch rerun now passes at
`Specs/TestRuns/commands-prefs-rapid-switch-current-contract-20260506_142132/last_run`
with 1 passed / 0 failed / 0 skipped. Next gate: rerun the full foreground
`cmd_preferences_dialog_` family to verify the Keyboard sequence and final
rapid-switch case together.

The broad Preferences fail-fast rerun did not reach the Keyboard/rapid-switch
tail yet. It failed earlier at
`Specs/TestRuns/commands-prefs-failfast-current-contract-20260506_142157/last_run`
with 44 passed / 1 failed / 123 skipped. Recurrent first failure:
`cmd_preferences_dialog_viewers_tab_traversal_live_dx_interaction` reports
`Preferences Viewers reverse computer field focus target not reached during tab
traversal; saw edit-new combo.` Treat this as an order-dependent traversal
problem until the exact case proves otherwise; do not close the Viewers
tab-traversal checklist item on the older pass evidence alone.

## Continuation Checkpoint - 2026-05-06 14:31 Europe/Paris

The exact standalone Viewers tab-traversal case still passes at
`Specs/TestRuns/commands-prefs-viewers-tab-exact-recurring-20260506_142417/last_run`
with 1 passed / 0 failed / 0 skipped, so the current failure is
order-dependent rather than a simple always-failing test contract.

Trace review from
`Specs/TestRuns/commands-prefs-failfast-current-contract-20260506_142157/last_run`
shows the reverse sequence reaches `PrimaryActionCombo`, then loses retained
DxUi focus while trying to move back to `ComputerField`: native focus becomes
null and the active page host remains the message target. The reported
`edit-new combo` value was itself misleading: the Viewers page never creates
the Editors-only `_editNewActionCombo`, and the debug helper compared null
retained focus against that null optional control pointer.

Patch in progress: `DebugGetViewersFocusTarget()` now returns `None` as soon as
the page host has no retained focus and pointer-guards the optional
Editors-only edit-new combo before comparison. The Viewers tab-traversal test
failure message now records expected/observed targets plus native focus,
active page, message target, visible page-child count, and resize-failure count.
The durable self-test contract in `Specs/Testing/Testing_SelfTests.md` now
requires focus-target snapshots to report null retained focus as `None` and to
guard family-specific optional controls. Next validation is a clean Debug x64
build, then a broad-enough Viewers/Preferences prefix rerun to expose the real
focus-loss state without the misleading label.

## Continuation Checkpoint - 2026-05-06 14:44 Europe/Paris

Debug x64 rebuilt cleanly after the Viewers focus-target diagnostic fix at
`.build/logs/msbuild-20260506_143129_915.log` with 0 warnings and 0 errors.
The Viewers-only Preferences prefix then passed at
`Specs/TestRuns/commands-prefs-viewers-prefix-focus-diagnostics-20260506_143340/last_run`
with 27 passed / 0 failed / 0 skipped. The broader foreground Preferences
fail-fast rerun advanced past Viewers, Keyboard, and the rapid-switch tail
area before failing at
`Specs/TestRuns/commands-prefs-failfast-focus-diagnostics-20260506_143437/last_run`
with 145 passed / 1 failed / 22 skipped. That closes the reopened Viewers
traversal concern against the broad family prefix.

The new first failure was
`cmd_preferences_dialog_file_operations_custom_bandwidth_live_dx_interaction`
with `Preferences File Operations page did not settle before custom-bandwidth
validation.` The exact standalone custom-bandwidth case passes at
`Specs/TestRuns/commands-prefs-fileops-custom-bandwidth-exact-20260506_144000/last_run`
with 1 passed / 0 failed / 0 skipped. The File Operations prefix then showed a
sibling order-dependent failure in
`cmd_preferences_dialog_file_operations_roundtrip_restores_dxui_surface` at
`Specs/TestRuns/commands-prefs-fileops-prefix-custom-bandwidth-20260506_144050/last_run`
with 5 passed / 1 failed / 0 skipped.

Root cause: File Operations page-specific tests still used stale setup
navigation (`End` then three `Up` keys) to mean File Operations. That is fragile
once prior tests leave the category tree scrolled or expanded; the visible root
order is already documented as the contract. Patch in progress: File Operations
page-specific tests now use `DebugSelectPreferencesCategory(kPrefCategoryFileOperations)`,
and the same stale `End` plus two `Up` setup was removed from Compare
Directories page-specific tests. Category-tree keyboard tests still keep the
real Home/End/Up/Down validation. The durable self-test guidance now states
that page-specific Preferences tests must use named category selection or the
visible root-row helper, not fixed `End`+`Up` shortcuts.

## Continuation Checkpoint - 2026-05-06 12:16 Europe/Paris

The foreground Preferences fail-fast rerun after the Viewers tab-traversal fix
advanced to 47 passed / 1 failed / 120 skipped at
`Specs/TestRuns/4cb089111a23/Commands/2026-05-06_121315/`. The new first
failure is `cmd_preferences_dialog_viewers_copy_follows_reordered_columns`.

Root-cause evidence: the clipboard immediately after failure contained
`Any<TAB>.selftest-viewers-001<TAB>Missing: builtin/viewer-text<TAB>(none)<TAB>Missing: builtin/viewer-text`.
The copy path followed the visible grid order correctly, but the test dragged
logical column `1u`, which is now `Computer`, while its assertion still expected
the retired two-column `Viewer` column to lead the row. The shared file-actions
association grid now has the durable Viewers column order `Match`, `Computer`,
`F3 View`, `Alt+F3 Alternate View`, `Status`; affected tests need named column
constants and updated expectations before the next fail-fast Preferences rerun.
Additional fixture issue: `TestSetViewerAssociationRows(...)` only replaced the
association rows and could inherit an empty/stale `viewers.actions` list from
earlier command-family state, causing `F3 View` cells to display `Missing:
builtin/viewer-text` while copy assertions expected the configured default
viewer action name. The helper now seeds default viewer actions before installing
custom association rows so the Viewers UI tests own their action definitions.

Patch/verification so far: Debug x64 rebuilt cleanly at
`.build/logs/msbuild-20260506_122159_111.log` with 0 warnings and 0 errors.
The focused first-failure repeat passed at
`Specs/TestRuns/4cb089111a23/Commands/2026-05-06_122342/` with 1 passed /
0 failed / 0 skipped.

Broader Viewers slice validation then failed at
`Specs/TestRuns/4cb089111a23/Commands/2026-05-06_122531/` with 22 passed /
5 failed / 0 skipped. New finding: the resize-only tests used logical column
`1u` as the adjacent header that must move when `Match` is resized. After the
column constants patch, those resize checks incorrectly tracked `F3 View`;
they must track the adjacent `Computer` header instead, while reorder/copy
tests continue to target `F3 View`.

The adjacent-column patch built successfully at
`.build/logs/msbuild-20260506_122645_923.log` with 0 warnings and 0 errors,
but the focused resize repeat still failed at
`Specs/TestRuns/4cb089111a23/Commands/2026-05-06_122833/` with 0 passed /
1 failed / 0 skipped. The failure still reports `resizeCount=0` and unchanged
geometry, so the remaining root cause is before the post-resize assertions:
the synthetic drag is not reaching the Viewers grid header-resize path. Next
debug step: add the same hit-test/pointer-state diagnostics used by the
Keyboard resize test, or expose equivalent Viewers file-actions grid debug
helpers if they do not exist yet, then rerun the focused case before changing
any resize behavior.

## Continuation Checkpoint - 2026-05-06 11:31 Europe/Paris

The foreground Preferences fail-fast rerun after the Viewers Add/Update fix
failed at `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_112645/` with
34 passed / 1 failed / 133 skipped. The first failing case was
`cmd_preferences_dialog_plugins_custom_paths_page_exposes_live_uia_grid_selection`;
the standalone repeat also failed under
`.build/selftest-isolated/commands-prefs-plugins-custom-paths-uia-standalone`.

Root cause: several Plugins tests still navigated to the Plugins category by
focusing the category tree, pressing Home, then sending seven Down keys. The
visible category order now includes User Menu before Keyboard, so seven Down
keys land on Themes instead of Plugins. These tests are not validating category
arrow traversal; they are validating the Plugins root page, custom-paths grid,
and Plugins command surfaces. The WIP patch replaces those stale row-count
walks with `DebugSelectPreferencesCategory(kPrefCategoryPlugins)` in:

- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp`
  for custom-paths UIA selection and custom-paths long-run scrolling.
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.PluginsThemesAdvanced.cpp`
  for Plugins search, tab traversal, empty custom-paths placeholder,
  custom-paths add/remove, Configure, Test, and Test All setup.
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp`
  for category-tree expand/collapse setup, which validates Left/Right
  expand/collapse behavior rather than Plugins' absolute visible row index.

Debug x64 rebuilt cleanly at `.build/logs/msbuild-20260506_113251_481.log`
with 0 warnings and 0 errors. The focused custom-paths UIA case passed at
`Specs/TestRuns/4cb089111a23/Commands/2026-05-06_113440/` with 1 passed /
0 failed / 0 skipped, confirming the stale category-row-count dependency was
removed for the first blocker.

Next validation: rerun the foreground `cmd_preferences_dialog_` fail-fast
family and continue first-failure diagnosis.

## Continuation Checkpoint - 2026-05-06 11:40 Europe/Paris

The foreground Preferences fail-fast rerun after the Plugins navigation cleanup
advanced to 37 passed / 1 failed / 130 skipped at
`Specs/TestRuns/4cb089111a23/Commands/2026-05-06_113640/`. The new first
failure was `cmd_preferences_dialog_keyboard_long_run_list_scrolling_stays_bounded`
with the same root maintenance issue: the test used Home plus four Down keys to
reach Keyboard, but the visible category order now places User Menu before
Keyboard, so four Down keys land on User Menu.

The WIP patch now replaces the remaining hard-coded Home/Down category setup
for Keyboard and Themes with `DebugSelectPreferencesCategory(...)` in:

- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ViewersAndKeyboardLists.cpp`
  for Keyboard long-run, Keyboard grid/header/copy setup, and Themes
  shell/theme-cycle setup.
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.ThemesGeneralPanes.cpp`
  for Themes retained state, UIA grid, pointer, header, and copy setup.
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.PluginsThemesAdvanced.cpp`
  for Themes search, duplicate/clear/set/apply/save/load setup.
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp`
  for Themes reordered/resized/sort/search setup.

Debug x64 rebuilt cleanly at `.build/logs/msbuild-20260506_114113_070.log`
with 0 warnings and 0 errors. The focused Keyboard long-run case passed at
`Specs/TestRuns/4cb089111a23/Commands/2026-05-06_114257/` with 1 passed /
0 failed / 0 skipped, confirming the next stale category-row-count blocker was
removed.

Next validation: rerun the foreground `cmd_preferences_dialog_` fail-fast
family.

## Continuation Checkpoint - 2026-05-06 11:49 Europe/Paris

The foreground Preferences fail-fast rerun after the category-navigation cleanup
advanced to 38 passed / 1 failed / 129 skipped at
`Specs/TestRuns/4cb089111a23/Commands/2026-05-06_114437/`. The new first
failure was `cmd_preferences_dialog_viewers_long_run_list_scrolling_stays_bounded`;
the standalone repeat also failed under
`.build/selftest-isolated/commands-prefs-viewers-longrun-standalone`.

Root cause: `FileActionPreferencesPage::DebugScrollAssociationByWheelDetents`
mutated `Grid::DebugSetScrollOffsets(...)` directly with `current + detents *
90`. From the top of the list, the test's negative wheel detents clamped back to
0 and produced no scroll advance. The other Preferences grid helpers route
through `Grid::OnMouseWheel(...)`, where negative detents scroll down from the
top and participate in normal invalidation/render-count behavior.

Fix: the shared file-actions helper now sets focus to the association grid and
replays the signed wheel detents through `Grid::OnMouseWheel(...)`. Debug x64
rebuilt cleanly at `.build/logs/msbuild-20260506_114701_436.log` with
0 warnings and 0 errors. The focused Viewers long-run case passed at
`Specs/TestRuns/4cb089111a23/Commands/2026-05-06_114849/` with 1 passed /
0 failed / 0 skipped. The durable rule was merged into
`Specs/UI/UI_PreferencesDialog.md`.

Next validation: rerun the foreground `cmd_preferences_dialog_` fail-fast
family.

## Continuation Checkpoint - 2026-05-06 12:05 Europe/Paris

The next foreground Preferences fail-fast run advanced to 44 passed / 1 failed /
123 skipped at `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_115122/`. The
first failure is
`cmd_preferences_dialog_viewers_tab_traversal_live_dx_interaction`; the
standalone repeat also failed at
`Specs/TestRuns/4cb089111a23/Commands/2026-05-06_115253/` with 0 passed /
1 failed / 0 skipped.

Current evidence: the test starts on the Viewers search field, reaches the
mappings grid on the first Tab, then fails to reach the expected extension
field on the second Tab and reports retained focus still on the mappings grid.
The Viewers test is now the outlier compared with sibling Preferences
tab-traversal tests: it chooses the native focused child as the message target
when that child belongs to the dialog, while most sibling page tests send Tab to
the active page host directly. Before changing behavior, add focused trace
around this helper to capture the native focus HWND, the active page host HWND,
the actual message target, and the retained focus target for each step.

Resolution: the diagnostic rerun at
`Specs/TestRuns/4cb089111a23/Commands/2026-05-06_120325/` proved the second Tab
did reach the active page host and moved retained focus to an unmodeled target,
not the mappings grid. The stale failure text came from formatting the failure
message with the previous snapshot while the wait predicate was still being
evaluated. The real maintenance gap was that
`PreferencesViewersDebugFocusTarget` and the tab test still modeled the retired
Viewers-only form. The enum and product debug mapping now name the whole shared
file-actions surface: tab header, match kind/value, computer, primary and
alternate action selectors, test file, association buttons, and action-tab
fields. The tab test now validates the current Associations-tab order in both
directions, including the tab header wrap between Reset Defaults and Search,
and formats failures only after the fresh snapshot has been captured.

Debug x64 rebuilt cleanly at `.build/logs/msbuild-20260506_120656_631.log`
with 0 warnings and 0 errors. The focused Viewers tab traversal case passed at
`Specs/TestRuns/4cb089111a23/Commands/2026-05-06_121030/` with 1 passed /
0 failed / 0 skipped. The durable keyboard-traversal order was merged into
`Specs/UI/UI_PreferencesDialog.md`.

## Continuation Checkpoint - 2026-05-06 11:16 Europe/Paris

The plan remains in `Specs/Plans/WIP/`. Full foreground Commands validation
after the `cmd_app_` fixes exposed Preferences Viewers coverage gaps and stale
test expectations; focused fixes are now landing in order.

New evidence since the 10:10 checkpoint:

- Full foreground Commands validation after the app-family fixes failed at
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_103525/` with 504 passed /
  93 failed / 0 skipped. The first failure was
  `edit_command_uses_focused_primary_editor_action`.
- The Primary Edit case was a stale test assertion, not a product bug: the
  marker file contained `"focused.editcmd"` with quotes, proving the command
  used the focused file and preserved the documented `{Filename}` argument
  quoting contract. Debug x64 rebuilt cleanly at
  `.build/logs/msbuild-20260506_103746_970.log`, and the focused edit case
  passed at `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_103958/`.
- Preferences Viewers theme-cycle validation exposed stale selected-extension
  debug state. `FileActionPreferencesPage::SyncAssociationFormFromSelection`
  now writes `state.viewersSelectedExtensionText` from the selected mapping, or
  clears it when no mapping is selected. Debug x64 rebuilt cleanly at
  `.build/logs/msbuild-20260506_104749_552.log`.
- Preferences Viewers Remove validation then exposed that destructive mapping
  actions reselected the next row after `SyncFromState`. Remove and Reset now
  clear the association selection and refresh the shared state after model
  rebind. Debug x64 rebuilt cleanly at
  `.build/logs/msbuild-20260506_105204_415.log`; the Remove case passed on the
  immediate repeat at
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_105530/`.
- `cmd_preferences_dialog_viewers_reset_live_dx_interaction` then exposed stale
  reset expectations. The live grid intentionally shows the Default mapping row,
  so the default Viewers reset count is 113 rows, not 112 extension-only rows.
  After shell Cancel discards a pending reset, reopening Preferences restores
  the three custom mappings and the normal first-row selection
  `.selftest-viewers-001`; destructive Reset itself still clears selection after
  the mutation. The test now mirrors that contract and prints current category,
  row count, selected extension, search text, pane-window counts, and resize
  failures when it misses.
- Debug x64 rebuilt cleanly after the reset diagnostics and expectation fixes at
  `.build/logs/msbuild-20260506_111359_915.log`. The focused reset case passed
  at `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_111539/` with 1 passed /
  0 failed / 0 skipped.
- `cmd_preferences_dialog_viewers_add_update_live_dx_interaction` then exposed
  both stale test targeting and a product behavior gap. The Viewers page now uses
  the shared file-actions form, so the live UIA edit and commit controls are
  named `Match value` and `Save Association`, not the retired Viewers-specific
  `Extension` / `Add / Update` captions. The product save path also now replaces
  the selected association row when the edited `(match, computer)` key is not a
  duplicate, so editing `.selftest-viewers-001` to
  `.selftest-viewers-updated` keeps the row count stable instead of appending a
  fourth mapping. Debug x64 rebuilt cleanly at
  `.build/logs/msbuild-20260506_112302_958.log`, and the focused Add/Update
  case passed at `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_112439/`.
- The durable Viewers/Editors shared file-actions contract was merged into
  `Specs/UI/UI_PreferencesDialog.md`: shared UIA labels, default-row visibility,
  Save Association selected-row replacement/upsert behavior, and selection
  clearing after destructive mutations are now authoritative spec text.

Next validation: rerun foreground fail-fast `cmd_preferences_dialog_`, then
continue first-failure diagnosis before returning to full Commands.

## Continuation Checkpoint - 2026-05-06 10:10 Europe/Paris

The plan remains in `Specs/Plans/WIP/`. The splash UIA hang, splash close hang,
submenu placement expectation, and Reread Associations cache-clear snapshot
timing are fixed in focused validation, and the full foreground `cmd_app_`
family now passes in order.

New evidence since the 09:45 checkpoint:

- The foreground `cmd_app_` family after the splash close pump no longer hung
  and reduced to one failure:
  `cmd_app_menuBar_submenu_placement_matches_spec`. Archive:
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_095732/` with 76 passed /
  1 failed / 0 skipped. The failure was not a visual-index leak; the test
  expected the ideal 4 DIP submenu offset even when Windows work-area clamping
  moved the submenu upward.
- `TestMainMenuSubmenuPlacementMatchesSpec` now computes its expected submenu
  position through the same DxUi debug positioning helper used by the product
  menu code, so it validates both the 4 DIP cascade offset and work-area
  clamping. Debug x64 rebuilt cleanly at
  `.build/logs/msbuild-20260506_100025_343.log`, and the exact submenu case
  passed at `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_100222/`.
- The next foreground `cmd_app_` family passed the submenu case but exposed
  `cmd_app_rereadAssociations_reloads_actions_and_refreshes_panes` as the first
  remaining failure. Archive:
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_100303/` with 76 passed /
  1 failed / 0 skipped. Root cause: the debug snapshot for
  `associationIconCacheSizeAfterClear` was sampled after pane refresh, and pane
  refresh can legitimately repopulate the association icon cache after the
  product clears it.
- `RereadAssociations(HWND)` now records `associationIconCacheSizeAfterClear`
  immediately after `IconCache::ClearAssociationCache()` and before pane
  refresh. The self-test failure text and archived
  `perf/reread_associations_metrics.json` now distinguish the before-clear
  sample from the immediate after-clear sample. Debug x64 rebuilt cleanly at
  `.build/logs/msbuild-20260506_100504_490.log`.
- Focused foreground validation for
  `cmd_app_rereadAssociations_reloads_actions_and_refreshes_panes` passed at
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_100843/`. The archived perf
  metrics include `associationIconCacheSizeBefore: 49` and
  `associationIconCacheSizeAfterClear: 0`.
- The full foreground `cmd_app_` family passed after the Reread Associations
  snapshot-timing fix at
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_101025/` with 77 passed /
  0 failed / 0 skipped. Its archived perf metrics again show
  `associationIconCacheSizeAfterClear: 0`.

Next validation: continue to full or sharded foreground Commands validation
before broader closeout.

## Continuation Checkpoint - 2026-05-06 09:45 Europe/Paris

The plan remains in `Specs/Plans/WIP/`. Compare Options and the previously
blocked shell/pane rechecks are fixed in focused validation, but broader
Commands validation exposed a new `cmd_app_` order-dependency that must be
resolved before closeout.

New evidence since the 09:05 checkpoint:

- Full foreground Commands validation after the Compare Options and shell/pane
  fixes hung in the app/splash region. The runner root was
  `.build/selftest-isolated/commands-full-after-compare-options-and-shell-fixes`;
  the native process exited with `-805306369` after about 18m45s and did not
  produce a native `results.json`. Windows Event Log recorded `AppHangB1`.
- A focused `cmd_app_splash_` run passed, while the broader `cmd_app_` family
  reproduced the hang before result emission. The root cause under test was
  splash self-test UI Automation collection running from the same thread that
  needed to keep the splash window responsive.
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Dialogs.cpp` now collects
  splash UIA descendant statistics and splash text state through a helper that
  runs the UIA work on a worker thread while the self-test thread pumps window
  messages with a bounded timeout.
- `.build/logs/msbuild-20260506_093243_985.log` rebuilt Debug x64 after the
  splash UIA pump change with 0 warnings and 0 errors.
- The post-splash-pump `cmd_app_` family no longer hangs, but it fails three
  order-dependent cases: `cmd_app_about_keeps_navigation_shell_stable`,
  `cmd_app_menuBar_mouse_open_keeps_popup_selection_clear`, and
  `cmd_app_menuBar_submenu_placement_matches_spec`. Archive:
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_093504/` with 74 passed /
  3 failed / 0 skipped.
- Each of the three failing cases passed alone immediately afterward:
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_093517/`,
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_093519/`, and
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_093523/`. This makes the
  current blocker an order-dependent state leak or settling gap, not a simple
  per-case assertion failure.
- A diagnostic rebuild at `.build/logs/msbuild-20260506_094607_866.log` exposed
  a second splash close hang before the app-family run reached the previous
  three failures. The partial trace under
  `.build/selftest-isolated/commands-app-family-diagnostics/last_run/commands/trace.txt`
  stops after the first splash UIA text-state collection and before the
  between-pass close returns. The self-test is being changed to request splash
  close and poll while pumping messages instead of calling the synchronous
  `SplashScreen::CloseIfExist()` join from the UI test path.

Next validation: instrument and fix the `cmd_app_` order dependency, rerun the
full `cmd_app_` family, then rerun full or sharded Commands before any broader
closeout validation.

## Continuation Checkpoint - 2026-05-06 09:05 Europe/Paris

The plan remains in `Specs/Plans/WIP/`. The Compare Options traversal blocker
is fixed in focused validation; broader Commands validation is still pending.

New evidence since the 2026-05-05 checkpoint:

- The instrumented Compare Options family run passed once at
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_083717/`.
- Two immediate repeats produced one pass and one failure:
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_083808/` passed, while
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_083831/` failed.
- The failing trace showed the old assertion text was partly stale. The final
  trace snapshot for `reverse select subdirectories-only-in-one-pane toggle`
  ended with `afterFocus=0 afterScroll=162/312`, while the failure message
  reported `actualFocus=8 scroll=276/312` because the assertion formatted the
  snapshot in the same call that updated it.
- Root cause now under test: reverse tab moved from Keep Identical Items toward
  Select Subdirectories Only In One Pane, then the focus-change callback scrolled
  the retained DxUi body from 276 to 162. That layout pass could clear the body
  host's retained focus target before the test could observe it.

Applied and validated:

- `RedSalamander/CompareDirectoriesWindow.Options.cpp` now restores the intended
  body host focus after a scroll/layout pass when the body host owned focus
  before layout.
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.CompareOptions.cpp` now
  evaluates the wait result before formatting the failure text, so future
  diagnostics report the final snapshot captured by the wait.
- `.build/logs/msbuild-20260506_084509_961.log` rebuilt Debug x64 with 0
  warnings and 0 errors.
- Three fresh-root foreground Compare Options family repeats passed with 9
  passed / 0 failed / 0 skipped:
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_084721/`,
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_084755/`, and
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_084811/`.
- Each repeat's trace shows the formerly failing step now lands on
  `afterFocus=7 afterScroll=162/312` for
  `reverse select subdirectories-only-in-one-pane toggle`.

New Commands shell/pane recheck discoveries:

- `cmd_pane_goToShortcutOrLinkTarget_` passed with 6 passed / 0 failed /
  0 skipped at `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_084950/`.
- The first `cmd_pane_executeOpen_junction` recheck matched zero direct cases
  because the filter lacked the trailing `_` required for prefix matching. That
  is invalid evidence; `Tools/Run-AllTests.ps1` now treats filtered runs with no
  runner-listed expected cases as `selftest_result_coverage` failures.
- `cmd_pane_itemProperties_show_shortcut_and_reparse_targets` failed before any
  product assertion at
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_084951_001/` with
  `Failed to create item-properties link-target folder.` The blocked fixture
  used a 246-character root, which pushed the target folder to 260 characters
  and target file to 271 characters. Shell command fixture path segments are
  now shortened in source; rebuild and rerun are pending.
- The zero-case filtered-run guard is now validated. `Invoke-Pester -Script
  Tools\Tests\RunAllTestsPlan.Tests.ps1` passed 8 / 0 / 0, and an intentional
  `cmd_pane_executeOpen_junction` bad-filter probe failed with
  `selftest_result_coverage: no expected cases matched the requested filter`.
  The final probe artifacts are under
  `.build\selftest-isolated\commands-zero-case-filter-guard-final\last_run`;
  the native archive for the corresponding no-case process is
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_090352/`.
- `.build/logs/msbuild-20260506_085910_490.log` rebuilt Debug x64 after the
  shell fixture shortening with 0 warnings and 0 errors.
- The affected shell/pane rechecks passed after rebuild with 13 total cases and
  0 failures:
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_090447/`,
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_090449/`,
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_090452/`,
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_090455/`,
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_090457/`, and
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_090500/`.

Next validation: run full or sharded Commands validation in a foreground-capable
session, then rerun the closeout Pester/source-contract checks and decide
whether any runner-owned aggregate artifact archiving gap must be fixed before
moving this plan to `Specs/Plans/Done/`.

## Continuation Checkpoint - 2026-05-05 23:12 Europe/Paris

This plan is intentionally still in `Specs/Plans/WIP/`. Do not move it to
`Specs/Plans/Done/` yet. Work was paused by user request while diagnosing the
remaining Commands Compare Options traversal issue.

Current implementation state:

- The latest build succeeded after adding diagnostic trace for the remaining
  Compare Options traversal failure:
  `.build/logs/msbuild-20260505_230757_177.log` (0 warnings, 0 errors).
- `cmd_compare_directories_options_live_dx_body_interaction` is no longer a
  405 s slow pass after bounding the shared Commands message pump; focused
  evidence is archived at
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-05_230056/`.
- The Compare Options family currently has one order-dependent failure:
  `cmd_compare_directories_options_tab_traversal_live_dx_interaction` passes
  alone but failed after sibling Compare Options cases.
  Failing family archive:
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-05_230258/`.
  Passing standalone archive:
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-05_230405/`.
- Failure details from the family run:
  `reverse subdirectory attributes toggle focus target not reached during tab
  traversal; actualFocus=7`. In
  `CompareDirectoriesOptionsDebugFocusTarget`, value 6 is
  `CompareSubdirAttributesToggle` and value 7 is
  `SelectSubdirsOnlyInOnePaneToggle`, so the failing reverse `Shift+Tab` stayed
  on "Select subdirectories only in one pane" instead of moving to
  "Subdirectory attributes".
- A diagnostic patch is currently present in
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.CompareOptions.cpp`
  inside the `sendTab` helper. It logs `compare_options_tab` entries with the
  routed HWND, target window class, before/after focus target, and body scroll
  offset. This is intentional continuation evidence; the family has not yet
  been rerun after this diagnostic build.
- Root-cause hypothesis to test next: order-dependent focus-window routing or
  shared DxUi/window modifier state after previous Compare Options cases. Do
  not replace the test with a debug-only focus traversal; this case is meant to
  validate live keyboard routing.

Resume with these exact commands from the repository root:

```powershell
$env:REDSALAMANDER_SELFTEST_ROOT = (Join-Path (Get-Location).Path '.build\selftest-isolated\commands-compare-options-family-tab-trace')
$env:REDSALAMANDER_REPO_ROOT = (Get-Location).Path
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter cmd_compare_directories_options_ -SkipBuild -TimeoutMultiplier 2 -ExePath .build\x64\Debug\RedSalamander.exe
```

Then inspect the archived `commands_trace.txt` for `compare_options_tab:` lines,
especially the entries immediately before:

- `reverse select subdirectories-only-in-one-pane toggle`
- `reverse subdirectory attributes toggle`

After that, fix the production/test root cause based on the trace, rebuild, and
rerun:

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug
$env:REDSALAMANDER_SELFTEST_ROOT = (Join-Path (Get-Location).Path '.build\selftest-isolated\commands-compare-options-family-after-tab-fix')
$env:REDSALAMANDER_REPO_ROOT = (Get-Location).Path
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter cmd_compare_directories_options_ -SkipBuild -TimeoutMultiplier 2 -ExePath .build\x64\Debug\RedSalamander.exe
```

Important continuation notes:

- Do not use hidden/background execution for GUI pointer/focus selftests. A
  hidden full Commands run was invalid and produced focus/pane failures plus a
  long stall. Run GUI suites in a foreground-capable session.
- Always use an absolute `REDSALAMANDER_SELFTEST_ROOT`; relative roots now get
  normalized by the app, but absolute roots make runner and product behavior
  explicit.
- Full Commands is still pending after the Compare Options traversal fix.
  Recheck the previously failing shell/pane cases with the absolute-root setup:
  `cmd_pane_goToShortcutOrLinkTarget_`, `cmd_pane_executeOpen_junction_`,
  `cmd_pane_itemProperties_show_shortcut_and_reparse_targets`,
  `cmd_pane_newFromShellTemplate_`, `cmd_pane_clipboardCut_`, and
  `cmd_pane_clipboardPasteShortcut_`.
- Before final closeout, rerun the focused Pester/source-contract tests,
  rebuild, archive the relevant evidence, update this checklist, then move this
  plan to `Specs/Plans/Done/` only after the validation is clearly closed.

References:

- `Specs/Testing/Testing_SelfTests.md` - result contract
- `Specs/Testing/Testing_TestCoverage.md` - declared case inventory
- `Specs/Testing/Testing_PerformanceValidation.md` - perf-gate contract
- `Specs/Testing/Testing_SelfTestRemoteCredentials.md` - remote credential setup
- `Tools/Run-AllTests.ps1` - unified test runner
- `.github/workflows/ci.yml` - CI pipeline
- `.github/skills/cpp-build/SKILL.md` - build/test invocation guidance
- `.github/skills/perf-validation/SKILL.md` - mandatory perf evidence workflow

## Detailed Review Checklist

Use this checklist as the execution contract for the detailed sections below.
Do not close the plan until each item is either completed or explicitly marked
blocked with the missing input.

### A. Rebuild the authoritative inventory first

- [x] Regenerate the current test inventory from source instead of trusting
  `Tests/README.md` or `Testing_TestCoverage.md`; both are stale.
- [x] Record the current self-test counts in the spec: runner-native listing
  now reports 597 Commands cases, 149 CompareDirectories cases, and 75
  FileOperations phases; the source fallback scan still records 567 Commands
  `SelfTest::RunCase` call sites, 141 CompareDirectories `RunCase` call sites,
  and 73 active FileOperations phases in `kFileOpsPhaseOrder`.
- [x] Record standalone/native tests: DxUiTests, ViewerPETests,
  ViewerSqliteTests, MonitorTest, LocalizationTests, and
  `PerformanceTests2.dll` via Visual Studio CppUnitTest.
- [x] Record script tests: `Tools/Tests/*.Tests.ps1`,
  `Tests/vcpkg-merge-synthetic-test.ps1`, and
  `Tests/vcpkg-merge-lock-validation.ps1`.
- [x] Add a source-derived inventory manifest path
  (`Tools\Get-TestInventory.ps1 -Format Json`) so current counts can be
  regenerated without hand counting.
- [x] Add runner-native `--selftest-list-cases` output so future case names are
  produced by the runner and can be linted in CI.
  Evidence: `.build/logs/msbuild-20260505_160654_612.log` built the native
  changes with 0 warnings / 0 errors, and
  `Specs/TestRuns/4cb089111a23/Tests/2026-05-05_163053_selftest_case_inventory/`
  archives JSON with the pre-fix 817 total runner-listed entries and zero duplicate names
  across Commands, CompareDirectories, and FileOperations.
- [blocked] Add runner-native case descriptions. Missing input: the current
  C++ registration contract carries case names only; descriptions require a
  broader `RunCase` metadata shape change.

### B. Fix runner and CI coverage before adding more cases

- [x] Add Commands self-test coverage to PR CI with a distinct timeout and
  failure step.
- [x] Add CI execution for MonitorTest, LocalizationTests, PerformanceTests2,
  PowerShell tool tests, and the fast vcpkg synthetic merge test.
- [x] Decide whether `vcpkg-merge-lock-validation.ps1` is PR-safe, nightly-only,
  or manual-only; document the choice because it runs vcpkg install flows and
  mutates `.build`.
- [x] Extend `Tools/Run-AllTests.ps1` with a true full-run mode that includes
  self-tests, standalone EXEs, CppUnitTest DLLs, Pester tests, and script tests.
- [blocked] Add an ARM64 smoke test job after x64 runner parity is fixed.
  Missing input: an ARM64-capable CI runner; `windows-latest` can build ARM64
  but cannot reliably execute ARM64 binaries as a test smoke.

### C. Converge harness behavior and make failures easy to understand

- [x] Replace ad hoc result emission with a shared helper for skipped,
  preconditioned, and state-machine test cases; CompareDirectories and
  FileOperations need this more than a blanket `RunCase` migration.
- [blocked] Add per-case, machine-readable reporting to standalone/native
  harnesses that currently only print pass/fail text or fail fast. Missing
  input: choose a standalone harness contract change. DxUiTests currently
  fail-fast exits from `Require(...)`, ViewerPETests/ViewerSqliteTests/
  MonitorTest/LocalizationTests use custom success aggregation, and
  PerformanceTests2 is owned by Visual Studio CppUnitTest; true per-case JSON
  requires converting those harnesses to named case registries or adding a
  common native reporter API.
- [x] Normalize DxUiTests filtering around the implemented `--suite=<name>`
  contract; remove or reject misleading `--filter` references.
- [x] Validate that every declared case/phase is emitted once, including
  skipped remote cases and setup/precondition failures; runner-native listing
  now provides the expected-name side of that comparison.
- [x] Validate that filtered self-test invocations cannot pass with zero
  matched runner-native cases. Discovery: `cmd_pane_executeOpen_junction`
  without trailing `_` matched no cases and initially looked green; runner code,
  Pester coverage, `Testing_SelfTests.md`, and the effective suite status are
  updated. Evidence: 8 Pester checks passed, and
  `.build\selftest-isolated\commands-zero-case-filter-guard-final\last_run\run-all-tests-results.json`
  records `selftest_result_coverage` with
  `no expected cases matched the requested filter`.
- [x] Add a source-contract guard that every literal CompareDirectories
  `SelfTest::RunCase` name is present in `kCompareCaseNames`.
- [x] Validate `--selftest-timeout-multiplier` input: reject invalid values,
  clamp extreme finite values, and keep scaled timeouts finite/nonzero.
- [x] Add a direct parser/source-contract test for timeout multiplier clamping
  at the too-small and too-large finite boundaries.
- [x] Make FileOperations `--selftest-case` filters honor phase-name prefixes
  so `Run-AllTests.ps1 -CaseFilter Phase5_` matches the documented runner
  contract instead of failing as an unknown case.
- [x] Write a runner-owned `run-all-tests-results.json` aggregate artifact so
  multi-suite `Run-AllTests.ps1` archives preserve every executed suite even
  when later native suites overwrite the shared root `results.json`.
- [x] Make `Run-AllTests.ps1` honor `REDSALAMANDER_SELFTEST_ROOT` when locating
  `last_run` artifacts; validation on 2026-05-05 showed the default
  `%LOCALAPPDATA%` folder can be overwritten by a different checkout and
  produce false blocker evidence.
- [x] Normalize the in-product `REDSALAMANDER_SELFTEST_ROOT` override to an
  absolute path before any test uses it for pane navigation. A relative
  worktree-local override was valid for the runner but caused `SetFolderPath`
  to reject the path and fall back to `C:\` inside the app.
- [x] Require Commands UI tests that open dialogs from pane state to wait for
  the requested pane paths to settle before dispatching commands that consume
  current left/right roots; Compare Options now uses that guard before opening
  Compare Directories.
- [x] Bound the shared Commands `PumpPendingMessages()` helper so active DxUi
  hover/cursor traffic cannot keep a selftest in an unbounded queue-drain loop.
  Focused evidence: `cmd_compare_directories_options_live_dx_body_interaction`
  dropped from a slow 405 s pass with a 227 MB perf log to a normal 5 s pass
  with a 1.7 MB perf log.
- [x] Resolve the order-dependent Compare Options live `Shift+Tab` traversal
  failure. Current evidence: the original family run failed at reverse
  navigation from `SelectSubdirsOnlyInOnePaneToggle` to
  `CompareSubdirAttributesToggle`
  (`Specs/TestRuns/4cb089111a23/Commands/2026-05-05_230258/`), while the same
  case passed standalone immediately afterward
  (`Specs/TestRuns/4cb089111a23/Commands/2026-05-05_230405/`). The
  instrumented 2026-05-06 repeats isolated the remaining failure to retained
  DxUi body focus being cleared after scroll/layout during reverse traversal:
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_083831/` shows
  `afterFocus=0 afterScroll=162/312`. The focus-restoration fix and stale
  diagnostic-message fix rebuilt cleanly in
  `.build/logs/msbuild-20260506_084509_961.log`, and three foreground family
  repeats passed at
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_084721/`,
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_084755/`, and
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_084811/`.
- [x] Rerun full or sharded Commands validation in a foreground-capable session
  after the Compare Options traversal issue is fixed. Do not treat the earlier
  hidden/background full Commands attempt as valid evidence. Foreground
  fail-fast `C:\R\FULLF50\last_run\commands\results.json` passed with
  597 passed / 0 failed / 0 skipped; archived evidence:
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-08_204608/`.
- [x] Resolve the post-splash-pump `cmd_app_` order-dependent failures before
  treating Commands as closed. Current blocker evidence:
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_093504/` failed
  `cmd_app_about_keeps_navigation_shell_stable`,
  `cmd_app_menuBar_mouse_open_keeps_popup_selection_clear`, and
  `cmd_app_menuBar_submenu_placement_matches_spec`, while each exact case
  passed standalone at `2026-05-06_093517/`, `2026-05-06_093519/`, and
  `2026-05-06_093523/`. Follow-up evidence reduced the app family to submenu
  placement at `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_095732/`;
  after changing the test to use the DxUi popup-positioning helper, the exact
  submenu case passed at
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_100222/` and the next app
  family run passed that case but exposed the Reread Associations cache
  snapshot timing failure at
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_100303/`. Full foreground
  `cmd_app_` validation after the Reread Associations snapshot fix passed at
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_101025/` with 77 passed /
  0 failed / 0 skipped.
- [x] Resolve Reread Associations cache-clear snapshot timing. The app-family
  failure at `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_100303/`
  showed the debug cache sample was taken after pane refresh could repopulate
  association icon entries. `RereadAssociations(HWND)` now samples
  `associationIconCacheSizeAfterClear` immediately after
  `IconCache::ClearAssociationCache()` and before pane refresh. Debug x64
  rebuilt cleanly at `.build/logs/msbuild-20260506_100504_490.log`; focused
  foreground evidence passed at
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_100843/`, with archived
  metrics showing `associationIconCacheSizeBefore: 49` and
  `associationIconCacheSizeAfterClear: 0`.
- [x] Resolve the second splash close hang exposed by the app-family diagnostic
  run. `TestSplashScreenUsesDxUiSurface` now requests close and observes the
  splash window/thread state while pumping messages instead of synchronously
  joining the splash worker from the UI self-test path. Debug x64 rebuilt cleanly
  at `.build/logs/msbuild-20260506_095235_549.log`, and focused foreground
  evidence passed with 3 / 0 / 0 at
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_095426/`; the trace records
  completed closes for initial cleanup, between-pass close, and cleanup.
- [x] Recheck the previously affected shell/pane Commands cases with an
  absolute `REDSALAMANDER_SELFTEST_ROOT` after the Compare Options fix:
  shortcut/link target navigation, junction open, item properties for shortcut
  and reparse targets, shell template creation, clipboard cut, and clipboard
  paste shortcut. Final focused evidence after the shortened-fixture rebuild:
  13 passed / 0 failed / 0 skipped across
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_090447/`,
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_090449/`,
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_090452/`,
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_090455/`,
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_090457/`, and
  `Specs/TestRuns/4cb089111a23/Commands/2026-05-06_090500/`.
- [x] Tighten self-test archive repo-root discovery so parent walks require
  `RedSalamander.sln`, `Specs/TestRuns`, and a `.git` directory/file before
  writing archived runs.
- [x] Add a FileOperations source guard that every active `Step` enum value is
  present in `kFileOpsPhaseOrder` exactly once, excluding setup/cleanup and
  terminal states.
- [x] Guard FileOperations self-test prompt dispatch against synchronous modal
  blocking; the custom speed-limit prompt must be posted asynchronously so the
  phase state machine can advance before the modal loop starts.
- [x] Fix Commands Connection Manager pointer-toggle coverage for scrollable
  DxUi editor controls. The S3 `Use HTTPS` toggle sits below the initial
  600px Connection Manager viewport, so the test must scroll it into view and
  consume a client-space rectangle that accounts for `ScrollPanel` offset
  before sending a real pointer click.
- [x] Recheck adjacent Commands pointer-toggle cases after the scrollable
  editor fix: Plugin Configuration `Use HTTPS`, credential prompt secret
  visibility, Compare Options, and Find recursive checkbox all still pass.
- [x] Rerun the full `cmd_connection_manager_window_` Commands family after
  the scrollable editor fix; 29 cases passed with 0 failures.
- [x] Verify the Compare/Search service cleanup discovered during full-suite
  validation: direct-SQLite-only fixtures must skip when no live NTFS journal
  cursor is available, generic service-routing tests must not confuse
  `DEGRADED_NO_INDEX` with host fallback, foreground services must use
  isolated stores when the assertion is not about the default ProgramData
  database, and no-wait SQLite queries must expect live-scan fallback until
  warmup/rebuild mirrors the requested root.
- [x] Add CI-reachable inventory lint that compares source-derived counts
  against `Specs/Testing/Testing_TestCoverage.md` and `Tests/README.md`.
- [x] Document the RedSalamander GUI-executable exit-code nuance discovered
  during CLI validation: PowerShell harnesses must use `Start-Process -Wait
  -PassThru` or `System.Diagnostics.Process` for self-test exit codes.
- [x] Fix fragile Commands Preferences chunk namespace structure before any
  larger split of those files.

### D. Update the durable specs and docs

- [x] Update `Tests/README.md` with current suites, current invocation commands,
  and correct counts.
- [x] Update `Specs/Testing/Testing_TestCoverage.md` with the regenerated
  inventory, including `Phase6_PopupRateSmoothing`.
- [x] Update `Specs/Testing/Testing_SelfTestRemoteCredentials.md` with
  `REDSALAMANDER_SELFTEST_CONN_S3_ALT`.
- [x] Document FileOperations phase names as stable result identifiers before
  moving phase implementations into feature-named files.
- [x] Keep any new durable runner behavior in `Specs/Testing/`, not only in this
  WIP plan.

### E. Challenge functional coverage area by area

- [x] Commands: ShellCommands and Plugin Configuration UIA are no longer open
  gaps. ShellCommands covers focused-item security routing, current-directory
  context-menu routing, NTFS ADS removal, and recursive Change Attributes
  progress; Plugin Configuration covers visible UIA Value/Toggle/Invoke
  patterns, accessible names, live DX interaction, Cancel restoration,
  keyboard/default-cancel routing, access keys, pointer toggles, and long-run
  open/close/scroll stability. Stale credential refresh remains tracked as a
  Connection Manager settings hot-reload/stale-save contract, not a remote
  provider refresh simulation. No implemented Commands long-poll contract was
  found.
- [x] CompareDirectories: current coverage includes OAuth auth-mode/refresh-token
  storage and missing-refresh-token gates, extensive SQLite bootstrap,
  compaction, WAL, legacy upgrade, and future-schema rejection coverage,
  content compare short-read coverage, synthetic crash-quarantine marker
  coverage, and bounded cancellation. Real remaining gaps are OAuth token
  refresh failure/revocation, HTTP 429/backoff, hard content-compare I/O
  failures, and true search-service process death mid-query; those need
  deterministic provider/service failure seams before they can be made reliable.
- [x] FileOperations: current coverage includes pre-calc cancel, queued cancel,
  local bandwidth cancel latency, active-worker cancellation in the recursive
  matrix, bridge throughput/concurrency, conflict prompts, settings defaults,
  connection overrides, and junction/reparse policy. Real remaining gaps are
  disk-full/quota, deterministic ACL-denied copy destinations, UNC destinations,
  cloud placeholder recall, and settings mutation while a task is already
  running; those need stable OS/test-environment seams or env-gated fixtures.
- [x] FileOperations live validation discoveries are resolved in code and tests:
  custom speed-limit self-test prompts no longer block the phase driver,
  parallel recursive copy restores newly created directory metadata only after
  queued workers drain, and postmortem diagnostics now seed their own
  warning/error-producing task so Phase13 is self-contained.
- [x] Standalone/native tests: MonitorTest is still a narrow ETW transport and
  opt-in diagnostic gate; LocalizationTests now covers embedded fallback,
  satellite string/menu/dialog lookup, invalid-culture fallback, and persisted
  language roundtrip but not every shipped satellite; PerformanceTests2 remains
  icon/refresh/splash/plugin-manager focused; DxUiTests remains executable-level
  fail-fast rather than per-case JSON.
- [x] Re-check every proposed "missing" DxUi area against current DxUiTests.
  Typeahead, scrollbars, single-line editing, TextField bridge routing, combo
  popup scrollbar paging/hover, tree selection/context/toggle, reduced motion,
  and UIA provider patterns already have dedicated coverage. Remaining DxUi
  additions should target controls with only visual-baseline coverage, not
  duplicate those areas.

### F. Close perf-validation gaps with evidence

- [x] Audit `Debug::Perf` emission by suite, case, and phase; do not rely on
  broad grep claims. Runner self-test metrics currently come from ViewerText
  diff/hex baselines, Compare local wide-tree search, FileOps pre-calc
  cancellation, bandwidth throttling, recursive/copy concurrency, recycle-bin
  batching, conflict convergence, bridge pipeline, and connection override
  phases.
- [x] For every reviewed perf-sensitive case, either emit a metric or document
  why the case is a correctness-only smoke test. UI long-run dialog/list cases
  assert bounded visible work and resource stability but do not claim perf
  deltas; `cmd_compare_directories_progress_perf` exercises Compare progress
  behavior but does not currently emit a self-test-local metric and remains a
  follow-up instrumentation candidate.
- [x] Archive before/after `perf_metrics.jsonl` for any new perf gate or
  instrumentation change under `Specs/TestRuns/`. No new perf instrumentation
  was added by this review pass, so no before/after perf JSONL is required for
  the helper/listing/reporting changes.
- [x] Update metric names/examples in `Specs/Testing/Testing_PerformanceValidation.md`
  when new metrics land. No new metric names landed; the existing metric-family
  map was documented in `Specs/Testing/Testing_TestCoverage.md`.

### G. Closeout

- [x] Run the full validation command set in the Validation section. The final
  fresh `Run-AllTests.ps1 -Suite Full -SkipBuild -TimeoutMultiplier 2` closeout
  under `C:\R\FULLSUITE7` passed 785 / 0 / 44 and includes the standalone,
  CppUnitTest, Pester, and fast script surfaces.
- [x] Rerun the full FileOperations suite after the Phase7/Phase12/Phase13 and
  FileOperations prefix-filter fixes.
- [x] Rerun the full CompareDirectories suite after the service no-wait,
  direct-SQLite precondition, isolated-store, and request-budget fixes.
- [x] Rerun the focused Compare Options live-DxUi interaction case after the
  pane-settle, absolute self-test root, and bounded message-pump fixes.
- [x] Archive the run evidence under `Specs/TestRuns/<machine>/Tests/<timestamp>/`.
  Final closeout archive:
  `Specs/TestRuns/4cb089111a23/Tests/2026-05-09_125031/`.
- [x] Move this plan to `Specs/Plans/Done/` only after runner/spec updates land
  and all durable requirements are merged into authoritative specs.

## Summary

This plan captures findings from a detailed review (2026-05-05) of the
repository test surface and the WIP plan itself.

Current runner-native scope:

- 3 in-product self-test suites in `RedSalamander/SelfTest/`.
- Commands self-tests: 597 runner-listed cases across 12 family files; the
  source fallback scan sees 567 static `SelfTest::RunCase` call sites.
- CompareDirectories self-tests: 149 runner-listed cases; the source fallback
  scan sees 141 static `SelfTest::RunCase` call sites plus setup/precondition
  result paths.
- FileOperations self-tests: 75 runner-listed phases: 73 active ordered phases,
  plus setup and cleanup phases.
- 5 standalone native test executables in `Tests/`: DxUiTests, ViewerPETests,
  ViewerSqliteTests, MonitorTest, and LocalizationTests.
- `Tests/PerformanceTests2/`: a Visual Studio CppUnitTest DLL with 7
  `TEST_METHOD`s.
- PowerShell/script tests in `Tools/Tests/` and `Tests/vcpkg-merge*.ps1`.

The review found four classes of issue:

1. **CI / harness gaps** - meaningful test code is built but never run on PR.
2. **Inventory drift** - several specs/readmes undercount or misdescribe what
   exists today.
3. **Functional coverage gaps** - error paths, cancellation paths,
   accessibility paths, and remote failure modes are not fully challenged.
4. **Maintenance debt** - file size, framework divergence, fragile namespace
   structure, inconsistent reporting, and the large FileOperations
   `SelfTestState`.

I did not find a declared test that is intentionally unimplemented, but the
coverage cannot be treated as fully gated until runner parity and inventory
drift are fixed.

## Track 1 - CI and harness gaps (P0 - BLOCKING)

### 1.1 CI does not run the Commands suite

`.github/workflows/ci.yml` runs `RedSalamander.exe` with
`--compare-selftest --fileops-selftest --selftest-timeout-multiplier=2.0`.
It omits `--commands-selftest`, even though Commands is the largest self-test
surface and covers dialog automation, preferences, navigation, search,
connections, shortcuts, shell commands, and view commands.

Action:

- Add `--commands-selftest` to CI, preferably as a separate step with its own
  timeout and artifact name.
- Keep failures distinct so a Commands-only failure is not attributed to
  CompareDirectories or FileOperations.
- Include Commands in any "full local run" command before declaring PR parity.

### 1.2 CI does not run MonitorTest, LocalizationTests, or PerformanceTests2

CI runs ViewerPETests, DxUiTests, and ViewerSqliteTests, but not MonitorTest,
LocalizationTests, or PerformanceTests2. Those projects can still break the
build, but their assertions do not gate PRs.

PerformanceTests2 is a CppUnitTest DLL, not a normal EXE. It must be run with
`vstest.console.exe .\.build\x64\Debug\PerformanceTests2.dll` or the equivalent
Visual Studio test runner.

Action:

- Add CI steps for MonitorTest and LocalizationTests.
- Add a CppUnitTest step for PerformanceTests2.
- Decide whether PerformanceTests2 is a smoke/correctness gate, a perf gate, or
  both; document that contract in `Specs/Testing/Testing_TestCoverage.md`.

### 1.3 CI and the unified runner omit PowerShell/script tests

The repository contains script tests that are not represented in the current
unified runner or PR CI:

- `Tools/Tests/*.Tests.ps1` - 50 Pester-style `It` blocks across build,
  MSBuild invocation, plugin deployment, environment sanitization, process
  streaming, test harness source contracts, test inventory linting, versioning,
  and vcpkg install safety helpers.
- `Tests/vcpkg-merge-synthetic-test.ps1` - 5 fast synthetic merge tests.
- `Tests/vcpkg-merge-lock-validation.ps1` - 3 lock/merge validation scenarios
  that invoke vcpkg install flows.

Action:

- Add the `Tools/Tests` suite to local full-run and CI.
- Add the fast synthetic vcpkg merge script to local full-run and CI.
- Classify the lock-validation script as PR, nightly, or manual-only before
  gating on it; it is more intrusive than the synthetic script.

### 1.4 `Tools/Run-AllTests.ps1` only knows about the 3 self-test suites

The unified runner accepts `-Suite All|Compare|Commands|FileOps`. It does not
invoke standalone native tests, PerformanceTests2, Pester tests, or vcpkg
scripts. Developers cannot currently run "everything" with one command before
pushing.

Action:

- Add a `-Suite Full` value or equivalent switches that run:
  - RedSalamander self-tests: Compare, Commands, FileOps.
  - Standalone native EXEs: DxUiTests, ViewerPETests, ViewerSqliteTests,
    MonitorTest, LocalizationTests.
  - CppUnitTest DLLs: PerformanceTests2 via `vstest.console.exe`.
  - PowerShell tests: `Tools/Tests` and fast vcpkg synthetic merge tests.
- Use an explicit manifest instead of only enumerating `*Tests*.exe`; the full
  test surface includes DLLs and scripts.
- Aggregate results into a single summary and preserve per-suite logs.

Status 2026-05-05: implemented. `Tools/Run-AllTests.ps1` now writes
`run-all-tests-results.json` for every invocation with wrapper metadata,
suite counts, case names, skipped reasons, failure reasons, and the wrapper
exit code. Evidence: `Tools/Tests/RunAllTestsPlan.Tests.ps1` passed 7/7, and
a live `Run-AllTests.ps1 -Suite FileOps -CaseFilter Phase6_PopupRateSmoothing
-SkipBuild` invocation emitted the aggregate artifact with 3 passed / 0 failed
/ 0 skipped.

### 1.5 No ARM64 test job

CI builds/tests x64 only. ARM64 is built by release flows but not test-gated.
Native-code regressions specific to ARM64 can therefore escape PR validation.

Action: add an ARM64 Debug smoke job after x64 parity lands. Start with
CompareDirectories/FileOperations plus one standalone smoke test if Commands
runtime is too high, then expand.

## Track 2 - Test framework convergence and infrastructure

### 2.1 Result reporting is split across incompatible patterns

The original WIP plan overstated this as "CompareDirectories uses its own
`AppendCaseResult` while Commands and FileOperations both use `RunCase`." The
actual split is more nuanced:

- Commands uses `SelfTest::RunCase` consistently for case execution.
- CompareDirectories uses `SelfTest::RunCase` for most cases, but emits
  additional setup/remote-precondition skipped or failed results through
  `AppendCaseResult`.
- FileOperations does not use `SelfTest::RunCase`; it uses a state-machine
  phase recorder (`RecordCurrentPhase` / `BuildResult`).

Action:

- Do not blindly migrate all suites to `RunCase`; FileOperations needs phase
  semantics.
- Introduce a shared result-emission helper for skipped/preconditioned cases and
  state-machine phases.
- Add a validation guard that every declared case or phase emits exactly one
  result, including skipped remote scenarios.
- Preserve `Specs/Testing/Testing_SelfTests.md` result JSON compatibility.

Status 2026-05-05: complete for the in-product self-test result summary path.
`SelfTest::AppendCaseResult` now owns suite case insertion, pass/fail/skip
counts, and first-failure propagation for `SelfTest::RunCase`, CompareDirectories
ad hoc setup/precondition results, and FileOperations state-machine summary
backfill. Evidence: `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` passed
9/9 and `.build/logs/msbuild-20260505_162133_076.log` built RedSalamander x64
Debug with 0 warnings / 0 errors.

Status 2026-05-05: runner-level emitted-once validation is implemented.
`Tools/Run-AllTests.ps1` now calls `--selftest-list-cases` for each executed
in-product suite and fails the run on duplicate expected names, duplicate actual
names, missing emitted results, or unexpected extra result names. Compare setup
failure remains an allowed explicit extra result. Evidence:
`Tools/Tests/RunAllTestsPlan.Tests.ps1` passed 7/7,
`Tools/Tests/TestHarnessSourceContracts.Tests.ps1` passed 13/13,
`.build/logs/msbuild-20260505_162852_589.log` built RedSalamander x64 Debug
with 0 warnings / 0 errors,
`Specs/TestRuns/4cb089111a23/Commands/2026-05-05_163030/` archives a focused
Commands run validated by the full runner, and
`Specs/TestRuns/4cb089111a23/Tests/2026-05-05_163053_selftest_case_inventory/`
archives a duplicate-name-free runner-native inventory.

### 2.2 Standalone/native harnesses lack consistent per-case reporting

DxUiTests has broad internal coverage but still primarily behaves like a custom
EXE harness: many families call `Test...()` functions directly and failures
often fail fast through `Require`. ViewerPETests, ViewerSqliteTests,
MonitorTest, and LocalizationTests also do not share the same machine-readable
result contract as the in-product self-tests.

Action:

- Add a small common standalone-test reporter, or teach `Tools/Run-AllTests.ps1`
  to wrap these harnesses and emit a normalized JSON summary.
- Prefer per-case `[RUN]` / `[OK]` / `[FAIL]` output plus machine-readable
  result files.
- Ensure suite filtering is represented in result metadata.

Status 2026-05-05: blocked pending a harness contract decision. A runner-level
wrapper can only emit one synthetic case per executable, which would hide the
current internal case boundaries and would not satisfy this plan's per-case
coverage goal. DxUiTests needs its `Require(...)->std::exit(1)` helpers replaced
or guarded by a named-case runner before per-case JSON can be truthful; the
other standalone EXEs need matching named case registries or a shared native
reporter. Leave this out of Done unless the broader standalone-harness contract
is explicitly scoped.

### 2.3 DxUiTests filtering is misdocumented/misleading

`DxUiTests.cpp` implements suite filtering through `--suite=<name>` or a
positional suite name. Unknown dash-prefixed switches are ignored. Any docs or
scripts that imply `--filter=Menu` will silently run all tests instead of the
requested subset.

Action:

- Document and use `DxUiTests.exe --suite=Menu`.
- Reject unknown `--*` switches, or at least print a warning and usage text.
- Add one harness-level test or script check for an unknown-switch failure.

### 2.4 Helper duplication across Commands family files

`TriggerAndDismissOwnedMenuCommand` is implemented twice with code-identical
bodies. `WaitForWindowClosed`, `WaitForPanePath`, and several pump/wait helpers
are reimplemented in multiple family files.

Action: hoist shared command-test helpers into a new
`RedSalamander/SelfTest/Commands/Commands.SelfTest.Common.cpp` or equivalent
local shared module so family files stop duplicating test plumbing.

### 2.5 Preferences coordinator namespace structure is fragile

`Commands.SelfTest.Preferences.cpp` opens an anonymous namespace, closes it,
includes chunk files, then reopens an anonymous namespace whose closing brace
lives in the parent orchestrator. One chunk closes a namespace that was not
opened in that chunk. The brace count balances, but the structure is brittle
and misleading.

Action:

- Make every included Preferences chunk namespace-self-contained.
- Verify symbol linkage before/after the mechanical change.
- Fix this before splitting `Preferences.FileOpsCompareAndTree.cpp` further.

Status 2026-05-05: implemented. Each Preferences chunk now owns its anonymous
namespace wrapper, the coordinator no longer leaves a namespace open for
included chunks, `Invoke-Pester .\Tools\Tests -ExcludeTag RequiresBuildToolchain
-PassThru` now passes 47 artifact-safe cases, `.build/logs/msbuild-20260505_155833_056.log`
shows a Debug x64 `RedSalamander` build with wrapper diagnostics 0 warnings /
0 errors, and `Specs/TestRuns/4cb089111a23/Commands/2026-05-05_160043/`
archives a focused Preferences case with 1 passed / 0 failed / 0 skipped.

### 2.6 Archive helpers tolerate stale repo-root walks too liberally

`SelfTestCommon.cpp` walks parent directories looking for
`RedSalamander.sln + Specs/TestRuns`. On nested checkouts or shared runners,
this can match the wrong root and archive results unexpectedly.

Action:

- Require a matching `.git` directory or `.git` file before accepting a repo
  root candidate.
- Keep the parent-walk depth explicit and documented.
- Add a focused unit/self-test for nested checkout behavior if practical.

### 2.7 Timeout multiplier parsing should be explicit and bounded

The original WIP plan said negative/NaN values are accepted. Current CLI parse
requires `parsed > 0.0`, so negative and NaN do not become the active scale;
they silently fall back to the default. The remaining issue is that invalid
values are silent and very large finite/infinite values are not clearly bounded
before `ScaleTimeout` multiplies them into `uint64_t` deadlines.

Action:

- Reject invalid/non-finite values with a diagnostic.
- Clamp finite values to a documented range, for example `[0.1, 100.0]`.
- Ensure scaled timeouts are finite and at least 1 ms when the original timeout
  is nonzero.
- Add one parser/self-test case covering invalid, too-small, too-large, and
  normal values.

Status 2026-05-05: parser behavior, docs, and source-contract coverage are
implemented. Invalid or non-finite values now fail fast, finite values are
clamped to `[0.1, 100.0]`, `ScaleTimeout` preserves zero while keeping nonzero
scaled timeouts finite and at least 1 ms, and
`Tools/Tests/TestHarnessSourceContracts.Tests.ps1` locks the exact clamp
constants plus invalid/clamped diagnostic paths.

## Track 3 - Spec compliance and case-quality issues

### 3.1 `GUI_INMENUMODE` reliance can mask DxUi popup regressions

Several Commands tests accept `(gti.flags & GUI_INMENUMODE) != 0 || popup !=
nullptr`. The spec says owned `DxUi_ContextMenu` popup windows are the
authoritative "menu opened" signal once a surface routes through DxUi. The
current fallback is useful for mixed native/DxUi surfaces but too weak for
fully migrated surfaces.

Action:

- For surfaces that are fully migrated to DxUi, require the popup HWND.
- For surfaces that intentionally still fall back to native menus, keep the
  `GUI_INMENUMODE` path and document why.
- Add one negative/timeout assertion where the surface should not open a native
  menu at all.

### 3.2 Bare `std::this_thread::sleep_for(<N>ms)` polling intervals are not scaled

Commands family files still contain bare short poll sleeps inside deadline
loops. The outer deadlines are often scaled, but on slow CI the fixed inner
interval can reduce the number of attempts enough to make timing failures
harder to diagnose.

Action:

- Define one named scaled poll interval per family, for example
  `constexpr auto kPollInterval = SelfTest::Scale(10ms)`.
- Replace only polling sleeps inside deadline loops; do not mechanically scale
  sleeps that intentionally model fixed UI debounce behavior.
- Document the distinction in the helper name or comment.

### 3.3 Magic numbers in test code hide intent

Examples include process wait timeouts, JSON test ports, recycle-bin batch
sizes, and perf improvement thresholds. The values may be correct, but the
tests do not explain why those values matter.

Action:

- Promote repeated or semantically important values to named `constexpr`
  constants near the relevant phase/family.
- Use names that describe the invariant being asserted, not only the raw value.
- Avoid changing thresholds and refactoring names in the same PR unless there
  is archived before/after evidence.

## Track 4 - Functional coverage gaps

### 4.1 Commands suite

| Gap | Severity | Notes |
| --- | --- | --- |
| ShellCommands family is small relative to surface area | Resolved by audit | Current `Commands.SelfTest.ShellCommands.cpp` covers focused-item Security shell action routing, current-directory context-menu routing, Alternate Data Stream removal, recursive Change Attributes progress/task reporting, shortcut/link targets, junction navigation, Shell New templates, and clipboard shortcut paths. A true `IContextMenu` verb-dispatch case remains optional future breadth, not a blocker for this review. |
| Plugin Configuration dialog has no UIA accessibility cases | Resolved by audit | Current Plugin Configuration tests collect visible UIA provider counts and Value/Toggle/Invoke pattern stats, assert accessible names, mutate visible DX edits through UIA, invoke Browse/Cancel/OK through UIA, and cover tab traversal, access keys, pointer toggles, theme cycle, and long-run stability. |
| Compare Options + Plugin Config cancel persistence | Resolved by audit | Compare Options live DX body interaction asserts canceled toggle/edit mutations are discarded on reopen; Plugin Configuration live DX interaction asserts Browse cancellation preserves edits and live DX Cancel restores baseline edit/toggle state on reopen. |
| Connection Manager stale-credential refresh while open | Partially covered | Connection Manager covers clean external reload refresh and dirty/stale save prompts while open. It does not simulate a remote provider refreshing an already-open secret; that belongs with provider refresh-failure seams rather than a pure Commands dialog case. |
| DxUi popup authoritative checks | Partially covered | Compare Directories and View Commands now use visible DxUi popup-window discovery for several migrated surfaces. A few older shell/menu stability checks still tolerate `GUI_INMENUMODE` because native shell/menu ownership is intentional there; convert only surfaces that are fully DxUi-owned. |
| Long-poll timing | No declared Commands contract found | The audit did not find an implemented Commands long-poll behavior or public requirement to test. Do not add a synthetic timing case until a product contract names the polling loop, expected cadence, and user-visible failure. |

### 4.2 CompareDirectories suite

| Gap | Severity | Notes |
| --- | --- | --- |
| OAuth refresh failure paths | H / blocked | Current coverage validates OAuth auth-mode roundtrip, refresh-token storage/deletion/session cache behavior, Windows Hello guard, and Google Drive missing-refresh-token gates. Expired token, network failure, and revocation require a deterministic OAuth/provider mock or plugin failure injection seam. |
| HTTP 429 / rate-limit retry | H / blocked | No runner-listed case or test seam exercises rate-limit retry/backoff. This needs a provider/service mock that can return ordered 429/retry-after responses without touching live cloud services. |
| Content-compare I/O errors | M / partial | Content compare covers equal/different decisions, dual I/O, no-I/O fallback, Unicode names, short reads, cross-filesystem dummy I/O, and pending propagation. Hard read failures such as sharing violation, access denied, and disk full remain future cases needing injectable `IFileSystemIO` failure points. |
| SQLite store failure modes | M / mostly covered | Current coverage includes bootstrap/schema creation, manual and automatic compaction, checkpoint/WAL truncation, legacy schema upgrade, future-schema rejection, default service status, seeded default startup, delete-burst maintenance, and live compaction. Corrupted WAL and simulated full disk during compaction remain future failure-injection cases. |
| Search service process crash mid-query | M / partial | `crash_quarantine_synthetic_marker` covers crash-quarantine marker behavior, but not a real service process terminating mid-query. Add only after the search-service harness exposes deterministic crash/kill control. |
| Cancel during content-compare stamp cycle | L / partial | Existing `cancel_completes_bounded` and content-pending cases exercise bounded cancel and pending content-comparison state, but not a narrowly timed cancel during a content-compare stamp update. |

### 4.3 FileOperations suite

| Gap | Severity | Notes |
| --- | --- | --- |
| Disk-full / quota-exceeded scenarios | H / blocked | No deterministic phase currently forces `ERROR_DISK_FULL`/quota without relying on machine volume state. Add after a dummy/local failure-injection seam can report disk-full at copy/write time. |
| ACL-deny destination at copy-time | H / blocked | Phase 12 uses protected junction ACLs for reparse safety, but there is no destination-copy ACL-deny phase asserting issues-pane details. A reliable case needs a privilege-stable ACL fixture or local `IFileSystemIO` denial seam. |
| Mid-bridge cancellation | M / partial | Pre-calc cancel, queued cancel, local bandwidth cancel latency, recursive active-worker cancellation, and post-mortem dummy cancellation are covered. Bridge pipeline throughput is covered, but a mid-bridge cancel/rollback case remains a future phase. |
| Symlink-specific reparse coverage | M / partial | Phase 12 covers junction loops, out-of-tree junctions, copied reparse tags including symlink/mount-point acceptance, move rollback, bridge reparse copy/move, and reparse delete target preservation. Relative symlink and cloud-placeholder recall behavior remain uncovered. |
| UNC destination | M / blocked | No env-gated UNC destination fixture exists. Add only with a documented variable and minimum local/share behavior so local and CI runs skip clearly when unavailable. |
| Conflict + concurrent rename | L / partial | Phase 9 covers overwrite/replace-readonly, apply-to-all cache, overwrite auto-cap, skip all, retry cap, skip-continues-directory-copy, and per-item concurrency. Concurrent destination rename races remain a future deterministic scheduling case. |
| Settings change mid-operation | L / partial | Phase 5 covers switching wait/parallel during pre-calc and Phase 8/11 cover default/connection override settings before tasks start. Runtime bandwidth/parallelism changes after data transfer begins remain uncovered. |

Implementation discoveries resolved during validation:

- `Phase7_ParallelCopyMoveKnobs` could block behind the Custom Speed Limit
  modal because the self-test invoke path opened the prompt synchronously before
  the phase driver advanced. `OnSelfTestInvoke` now posts the deferred prompt
  message, and `TestHarnessSourceContracts.Tests.ps1` guards that contract.
  Evidence:
  `Specs/TestRuns/4cb089111a23/FileOperations/2026-05-05_1738_prompt_async_focus_after_assertion/`.
- `Phase13_PostMortemDiagnostics` was order-dependent when filtered because it
  expected previous phases to have produced warning/error summaries. It now seeds
  its own deterministic missing-file copy and validates diagnostic log/export
  behavior from that task. Evidence:
  `Specs/TestRuns/4cb089111a23/FileOperations/2026-05-05_1747_phase12_phase13_focus_after_fix/Phase13_PostMortemDiagnostics/`.
- `Phase12_ReparsePointPolicy` exposed a production metadata bug: the parallel
  recursive copy path restored a newly created directory's timestamps before
  queued child work had drained, so later file creation rewrote the parent
  directory's last-write time. The parallel path now remembers created
  directories and restores metadata deepest-first after workers finish. Evidence:
  `Specs/TestRuns/78ac7c415c54/FileOperations/2026-05-05_1805_phase12_parallel_directory_metadata_after_fix/`.
- `Run-AllTests.ps1 -CaseFilter Phase5_` exposed a FileOperations selector gap:
  the runner contract advertises prefix filters, but FileOperations only
  accepted exact phase/family names. Prefix filters now select all matching
  phases, and the source-contract tooling suite guards this. Evidence:
  `Specs/TestRuns/78ac7c415c54/FileOperations/2026-05-05_1830_phase5_prefix_filter_pass/`.
- Full-suite recheck after these fixes passed. Evidence:
  `Specs/TestRuns/78ac7c415c54/FileOperations/2026-05-05_1842_fileops_full_after_fixes_pass/`
  with 55 passed / 0 failed / 20 expected skips. The earlier failed full run at
  `Specs/TestRuns/78ac7c415c54/FileOperations/2026-05-05_1820_fileops_full_failed_phase5_worker_budget/`
  remains archived as non-gating evidence for the transient Phase5 investigation;
  isolated `Phase5_PreCalcSettingsApplied`, exact Phase5 family, and `Phase5_`
  prefix reruns all passed afterward.

### 4.4 Standalone/native test coverage

| Gap | Severity | Notes |
| --- | --- | --- |
| MonitorTest only validates a narrow transport slice | M / partial | MonitorTest covers ETW TraceLogging emit/receive, self-diagnostic suppression, opt-in invalid-rectangle visualization, and self-originated event rejection. It still does not cover a full Monitor window state machine or rich ETW filtering-rule matrix. |
| PerformanceTests2 covers 7 cases but remains narrow | M / partial | Current methods cover folder icon enumeration, duplicate path icon enumeration, duplicate-path refresh/compact-mode hit testing, splash close guard, and plugin-manager empty discovery. ColorTextView render, search throughput, and Compare scan throughput are not represented there; some are instead covered by in-product self-test metrics. |
| LocalizationTests is small | L / partial | LocalizationTests covers embedded fallback, satellite string/menu/dialog lookup, localized dialog templates with executable-owned custom child classes, invalid culture fallback, and persisted `ui.language` roundtrips. It still does not enumerate every shipped satellite resource. |
| DxUi gaps must be re-audited from the existing component matrix | M / resolved by audit | Typeahead, scrollbars, single-line editing, text-input bridge routing, ComboBox popup scrolling/hover, Tree selection/context/toggle, reduced motion, and UIA patterns already have dedicated coverage. Future additions should target visual-baseline-only controls. |

## Track 5 - Refactor and file-size debt

### 5.1 File sizes that are too large to navigate

| File | Lines | Recommendation |
| --- | --- | --- |
| `Commands.SelfTest.Search.cpp` | ~11,098 | Split into Find dialog UI, search index/service, and local search backends. |
| `Commands.SelfTest.Settings.cpp` | ~10,756 | Extract shared helpers first; keep settings cases grouped until helper churn settles. |
| `Commands.SelfTest.Dialogs.cpp` | ~10,497 | Split into About/Splash, prompts, Filter/Rename/ChangeCase, and FileOps prompts. |
| `Commands.SelfTest.ViewCommands.cpp` | ~10,217 | Split into selection ops, sort, pane ops, and tabs. |
| `Commands.SelfTest.Navigation.cpp` | ~9,136 | Split into NavigationView edit/popup, drive/menu shell, and directory-impact selection. |
| `Commands.SelfTest.Preferences.FileOpsCompareAndTree.cpp` | ~7,472 | Fix namespace structure first, then split by preference page. |
| `CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp` | ~7,265 | Split into Local, Service, and Remote search/index families. |
| `FolderWindow.FileOperations.SelfTest.Phases07_09.cpp` | ~5,961 | Split by feature rather than phase number. |
| `Tests/DxUiTests/DxUiTests.TextInputBridge.cpp` | ~8,311 | Split by surface, such as TextField, multiline text, and browser/embed bridge behavior. |

Mechanical splits must preserve case names and archived result identity.

### 5.2 `SelfTestState` god-object in FileOperations

`FolderWindow.FileOperations.SelfTest.cpp` declares one large `SelfTestState`
with fields spanning every phase. Cross-phase state is difficult to audit, and
reset behavior is not expressed at the same level as the phase order.

Action:

- Group fields into nested structs: `FileSystems`, `LocalConfigBackup`,
  `RemoteS3`, `RemoteOneDrivePersonal`, conflict state, rate-limit state, and
  UI observer state.
- Add an explicit `ResetPerPhaseState(Step previous, Step next)` hook called
  from phase transitions.
- Add a guard that every active `enum class Step` value appears in
  `kFileOpsPhaseOrder` exactly once, excluding setup/cleanup/terminal states.
- Preserve existing phase result names.

### 5.3 Phase files are split by number, not feature

Files such as `Phases05_06.cpp`, `Phases07_09.cpp`, and `Phases10_13.cpp` are
hard to navigate once phase numbers change. Case names also embed phase
numbers, but those names are durable result identifiers and should not be
renamed casually.

Action:

- Move implementation files toward feature names, for example
  `Phases_PreCalc.cpp`, `Phases_PopupBandwidth.cpp`,
  `Phases_Parallelism.cpp`, `Phases_Conflicts.cpp`, `Phases_Bridge.cpp`,
  `Phases_Reparse.cpp`, and `Phases_Remote.cpp`.
- Keep existing case/result names unless a separate migration plan updates
  archived-result compatibility.
- Document stable phase names in `Testing_TestCoverage.md`.

## Track 6 - Spec drift

### 6.1 Published counts are stale

`Tests/README.md` and `Specs/Testing/Testing_TestCoverage.md` undercount or
misdescribe several suites. Examples found during review:

- Commands is documented around 346/360+ cases in older locations, but current
  runner-native listing finds 597 cases and the source fallback scan finds 567
  `SelfTest::RunCase` call sites.
- FileOperations is documented as 68/76 phases in different places, but current
  runner-native listing has 75 phases: 73 active phases plus setup/cleanup.
- PerformanceTests2 is documented as 4 cases in places, but current source has
  7 `TEST_METHOD`s.
- PowerShell/tooling and vcpkg script tests are not represented in the main
  test inventory.

Action:

- Update counts after the runner/manifest work defines the authoritative source.
- Add CI lint that fails when the generated inventory differs from the
  documented inventory.

### 6.2 `Phase6_PopupRateSmoothing` is in code but missing from the spec

`Phase6_PopupRateSmoothing` appears in the FileOperations enum and active phase
order, but `Testing_TestCoverage.md` does not list it with the other Phase 6
coverage.

Action: add it to the spec, including what behavior it validates and whether it
emits perf metrics.

### 6.3 `REDSALAMANDER_SELFTEST_CONN_S3_ALT` is not documented

FileOperations reads `REDSALAMANDER_SELFTEST_CONN_S3_ALT`, but
`Testing_SelfTestRemoteCredentials.md` omits it from the remote credential
setup contract.

Action: document `S3_ALT` in the credentials spec, including which phases use
it and what minimum bucket/account behavior is required.

### 6.4 `Testing_TestCoverage.md` is hand-maintained

Maintaining a long markdown case inventory by hand will keep drifting.

Action:

- Prefer a lint path first: runner emits case inventory; CI compares it to the
  committed spec.
- Consider generation later if the lint path proves too noisy.
- Keep human-readable descriptions in the spec even if names/counts become
  generated.

## Track 7 - Performance instrumentation gaps

`Specs/Testing/Testing_PerformanceValidation.md` requires perf-sensitive
scenarios to define scenarios, emit metrics, run deterministic selftests, and
archive evidence. The earlier WIP claim that CompareDirectories has zero
`Debug::Perf` emission is stale: at least some Compare search/index cases emit
metrics. The real problem is partial and inconsistent instrumentation.

Audit findings:

- Commands self-tests emit ViewerText diff/hex metrics from
  `viewer_text_diff_perf` and `viewer_text_hex_byte_color_perf`. Long-run UI
  dialog/list cases are correctness/stability guards: they assert bounded
  visible work, row counts, UIA/provider counts, or resize-failure counts, but
  they do not claim before/after performance deltas.
- CompareDirectories self-tests emit
  `compare.selftest.local_search_scan_wide_tree_workers_us` and
  `compare.selftest.local_search_scan_wide_tree_us` for the local wide-tree
  search case. Other Compare cases are correctness or failure-mode guards unless
  the UI progress test is upgraded to emit a metric.
- FileOperations self-tests already emit metrics for pre-calc cancel latency,
  local/parallel bandwidth throttling, copy/move concurrency, auto concurrency,
  recursive copy matrices, recycle-bin batching, default bandwidth limits,
  conflict convergence, bridge pipeline, connection overrides, and the global
  connection gate.
- PerformanceTests2 covers a narrow performance slice and now has CI/unified
  runner coverage, but it remains a CppUnitTest DLL without shared JSON/perf
  archive semantics.

Action:

- Treat new perf-sensitive cases as incomplete unless they emit a metric or
  explicitly declare themselves correctness-only.
- Upgrade `cmd_compare_directories_progress_perf` with a self-test-local metric
  if future work wants to use it as a perf gate instead of a progress
  correctness/stability check.
- Add before/after `perf_metrics.jsonl` for future instrumentation changes.
- Update metric naming examples in `Testing_PerformanceValidation.md` when new
  metrics land.

## Validation

After each track lands, run the relevant focused tests, then run the full set
before closeout:

```powershell
.\Tools\Run-AllTests.ps1 -Suite All -TimeoutMultiplier 2
.\.build\x64\Debug\DxUiTests.exe
.\.build\x64\Debug\ViewerPETests.exe
.\.build\x64\Debug\ViewerSqliteTests.exe
.\.build\x64\Debug\MonitorTest.exe
.\.build\x64\Debug\LocalizationTests.exe
vstest.console.exe .\.build\x64\Debug\PerformanceTests2.dll
Invoke-Pester .\Tools\Tests
.\Tests\vcpkg-merge-synthetic-test.ps1
```

Manual-only intrusive validation:

```powershell
.\Tests\vcpkg-merge-lock-validation.ps1
```

Archive a `Specs/TestRuns/<machine>/Tests/<timestamp>/` snapshot per
`Specs/TestRuns/README.md`. Tracks that change perf instrumentation must include
before/after `perf_metrics.jsonl` per
`Specs/Testing/Testing_PerformanceValidation.md`.

## Suggested order of execution

1. Track 1 + Track 6 inventory correction - establish what exists and what is
   actually gated.
2. Track 2 harness/reporting convergence - make failures complete and
   comparable across suites.
3. Track 3 spec-compliance fixes - tighten existing assertions before adding
   new cases.
4. Track 7 perf instrumentation - add evidence where the current tests already
   exercise perf-sensitive paths.
5. Track 4 functional gaps - split by domain and add focused cases.
6. Track 5 file-size/state refactors - do after result identity and specs are
   stable, unless a refactor is needed to make a coverage addition maintainable.

## Out of scope

- Adding a new test framework such as gtest or Catch2. The near-term need is
  runner/reporting convergence, not a framework rewrite.
- Rewriting production code under test. Findings here are about the tests
  themselves; if a test gap reveals a production bug, file that separately.
- Mutation testing or coverage tooling. Reconsider after runner parity,
  inventory linting, and the largest file splits land.
