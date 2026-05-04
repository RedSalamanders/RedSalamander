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
- Checked-in archived runs may contain skipped cases; that is valid when the skip reason documents the missing precondition.
- `Tools/Run-AllTests.ps1` must tolerate both direct suite JSON (`commands_results.json`, `compare_results.json`, `fileops_results.json`) and aggregated run JSON (`selftest_run_results.json`) when summarizing archived results.
- If a self-test process exits early, crashes, or writes only partial artifacts, `Tools/Run-AllTests.ps1` must report the runner failure from the available JSON/trace data instead of failing its own parser.

## Artifact Contract

Self-test artifacts must preserve enough detail to explain why a run passed, failed, or skipped:
- `results.json` records final case status and reason,
- `trace.txt` records supporting diagnostic context,
- `perf_metrics.jsonl` must be preserved when the scenario emits performance metrics,
- archived copies under `Specs/TestRuns/` must keep those files intact.

## DxUi Popup Validation

For command selftests that validate migrated app chrome using DxUi popup menus:
- owned `DxUi_ContextMenu` popup windows are an authoritative “menu opened” signal,
- popup dismissal may be validated by observing that owned popup window close, not only by `GUI_INMENUMODE`,
- tests must not rely on `GUI_INMENUMODE` alone once the validated surface routes through the shared DxUi popup path instead of a native Win32 menu loop.

## NavigationView DxUi Text-Host Validation

For command selftests that validate NavigationView address-bar edit mode or full-path popup edit mode:
- `NavigationViewDebugSnapshot` is the authoritative contract for edit visibility, focus target, current edit text, selection range, and the active DxUi host/bridge HWNDs,
- tests must not enumerate descendant native `Edit` / `RICHEDIT50W` windows to prove edit mode, because NavigationView no longer exposes a visible native edit child and the hidden bridge is an implementation detail behind the DxUi host.
- address-bar edit clipboard coverage must verify the focused DxUi host remains active and pane-level Select All, Copy, and Paste commands mutate/copy the edit text instead of falling through to FolderView command handling.

## Preview Pane and ViewerVLC Validation

For command selftests that validate embedded preview behavior:
- `FolderWindow::PreviewPaneDebugSnapshot` is the authoritative contract for preview activity, source/host panes, selected Folder/Preview tab, hosted plugin ID, content HWND, and embedded viewer instance identity.
- Preview tests must verify that embedded viewers do not take keyboard focus from the source pane, that focus changes reuse the current embedded viewer instance when the resolved plugin is unchanged, and that preview close/replacement persists changed plugin configuration.
- When saved viewer associations are empty or resolve only the default text viewer, preview tests must cover the built-in embedded viewer defaults before `builtin/viewer-text` fallback.

For command selftests that validate `builtin/viewer-vlc`:
- `WndMsg::ViewerVlcDebugGetSnapshot` is the authoritative contract for HUD state, volume/mute state, snapshot dimensions, Fluent icon glyph usage, and filled-button HUD styling.
- Wheel-seek coverage must exercise both normal viewer surfaces and libVLC-owned child video windows so embedded preview and standalone playback keep the same wheel behavior.

## Search-Specific Coverage

Search coverage follows the same contract:
- local, fallback, indexed, and service search cases stay declared,
- ReFS validation stays declared even on machines without ReFS,
- a machine without a fixed ReFS volume records `skipped` with a reason instead of silently omitting the case.

## Source Organization

Project-owned self-test implementation must live in `.cpp` source files, with optional `.h` declarations when cross-translation-unit declarations are needed.

Project code must not introduce `.inl` self-test implementation files.

The `--commands-selftest` suite may continue to include family source files from `RedSalamander/SelfTest/Commands/Commands.SelfTest.cpp`, but those included family files must still be `.cpp` files rather than `.inl` fragments.
