---
name: red-salamander-crash-forensics
description: Preserve and investigate RedSalamander dumps, sidecar stacks, fatal dialogs, hangs, timeouts, and failed selftest runs before rerunning a focused scenario.
---

# RedSalamander crash forensics

Read [the selftest contract](../../../Specs/Testing/Testing_SelfTests.md) for run ownership, artifacts, containment, and focus. Preserve the failing verdict before retrying; a passing retry does not replace it.

## Resolve the actual run

- Record exact command/filter, source/dependency/build/profile identity, process ID plus start time/path, and runner/run ID. Match symbols and source to that build.
- Resolve test artifacts from the run receipt. Active output is beneath the existing marked X:\RedSalamander.Perf runs/<runId>/artifacts/selftest/last_run; completed runs move to evidence/completed-runs/<runId>. Follow the archive record rather than newest timestamps.
- Product crashes use %LOCALAPPDATA%\RedSalamander\Crashes. Check the exact app's paired dump/sidecar first, then targeted WER/CrashDumps locations if absent. Avoid unrelated user data and blind AppData recursion.
- Repository RedSalamander/SelfTest/last_run and .build/codex-runs are historical/manual-run fallbacks, not primary current test paths.
- New test scratch, copied logs and evidence stay in the selected marked sandbox. An environment variable does not authorize another root. A reviewed compact promotion into Specs/TestRuns is separate from test execution.

For a fatal dialog without a new dump, retain its exact text/state, process identity, runner log, trace, filter, and sidecar. Captures still require the governed harness and synthetic fixtures; triage grants no foreground access.

## Ownership and concurrency

Never terminate an independently launched application. Process names or stale PIDs do not prove ownership. For a child launched and contained by this operation, use the runner's owned cancellation/timeout/Job Object cleanup and preserve terminal artifacts first. Diagnose an externally owned hang and report the blocker.

Follow actual profile/dependency/desktop coordination: same output profiles serialize; disjoint profiles/worktrees can run independently when no narrower resource conflicts. Use the interactive lease only for marked entries after agreement to the time. Do not add a blanket parallel-runner prohibition.

## Treat signatures as hypotheses

Start at the deepest product frame owning the object or UI action; a CRT/NTDLL top frame may observe earlier corruption. Verify line numbers against the failing build.

- Perf JSONL faults near UTF conversion/scope destruction: inspect string/detail ownership in the current implementation; an old signature is not proof of the same bug.
- Modal DxUi/TSF teardown: inspect destructive cross-thread synchronous messages and UI-thread lifetime. Post destructive close commands where the host contract requires it; preserve synchronous non-destructive snapshot semantics.
- Plugin/callback faults: inspect module pins, target lifetime, cancellation, in-flight quiescence and posted-payload drains together.

Use exact current case names from source/live inventory, not log labels. Convert formatted process snapshots with Out-String before preserving them.

## Verify repair

Reproduce narrowly through the governed runner, then run the affected case/cluster and applicable qualification gates. Record command, identity, actual result, retained failure and unresolved limits in the active plan. Implemented, focused-pass, broad-pass, performance, native-platform and foreground acceptance are separate states; blocked gates stay unqualified.
