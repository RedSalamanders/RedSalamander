---
name: perf-validation
description: Performance-first development contract for RedSalamander. Use when adding features, optimizing code, instrumenting hot paths, or validating UI/startup/search/fileops/compare performance from the beginning.
metadata:
  author: RedSalamander
  version: "1.0"
---

# Performance Validation

Use this skill whenever a change can affect responsiveness, throughput, queueing, rendering, startup, memory retention, or any repeated callback path.

## Non-negotiable rule

Do not treat performance validation as a follow-up task.

For new features and optimizations:

1. Define the scenario first.
2. Add or reuse instrumentation before claiming results.
3. Add deterministic selftest coverage or a deterministic harness.
4. Archive the run under `Specs/TestRuns/`.
5. Compare baseline vs candidate and state caveats honestly.
6. If the work closes a plan, move that plan to `Specs/Plans/Done/` and merge enduring requirements into the authoritative spec or repo guidance.

See:

- `Specs/Testing/Testing_PerformanceValidation.md`
- `Specs/Testing/Testing_SelfTests.md`
- `Specs/TestRuns/README.md`

## Required workflow

### 1. Name the protected scenario

Write down the exact path being protected, for example:

- warm Find Files progress storm
- cold folder open with icon churn
- FileOps pre-calc cancel
- Compare Directories content-progress queue churn

If you cannot name the scenario, you do not yet have a valid perf task.

### 2. Choose the metric family

Prefer subsystem-prefixed metrics such as:

- `find.ui.*`
- `compare.ui.*`
- `FileOps.*`
- `render.*`
- `icons.*`

Good metrics include:

- queue coalesced/drained counts,
- input-to-visible or progress-to-visible latency,
- stage timings,
- rebuild counts,
- lock hold time,
- bytes rewritten,
- visible-work counters.

### 3. Add deterministic validation

Prefer:

- `--commands-selftest` for real UI/window paths,
- `--compare-selftest` for Compare/search engine paths,
- `--fileops-selftest` for file-operation paths.

If the existing suites do not exercise the scenario, add a focused case instead of relying on manual repro.

Commands cases run behind the same per-case isolation boundary in normal,
repeat, and shuffled order. A fixture that mutates runtime managers, deferred
plugin unload state, retained overlays, menu mode, focus, or pane state must
restore and verify that state before returning; an isolated exact-case retry is
not evidence that the broad process is clean. Retained-window metrics and
teardown assertions must obtain HWNDs from process-owned debug hooks or another
process-scoped lookup, never from class-only `FindWindowW(...)` evidence that
can resolve a different running app instance.

Broad command-dispatch smoke cases must leave whole-runtime settings/plugin
reload commands and provider/OS lifecycle commands that prepare before user
confirmation to dedicated deterministic cases. After any dispatched command can
mutate navigation or presentation state, restore the fixture's in-memory settings
snapshot, re-establish both pane providers and paths, and reapply each pane's live
display mode, sort key/direction, and visibility options. Then prove both
current-folder enumerations have settled before the next dispatch. Restoring the
settings object alone does not reapply already-mutated FolderView state, and
matching path text alone is not quiescence evidence.
Navigation-shell stability cases that first change provider, path, or focus must
pump messages and require a bounded multi-sample quiet baseline before capturing
refresh, path, focus, or selection counters.
When an owned pane modal closes, validate the real FolderView focus after the
modal loop unwinds. Product teardown should combine a synchronous best-effort
restore with the owning window's queued, foreground-safe focus restore; record
Win32 focus, active/foreground owners, and the foreground process on failure.
For asynchronous file/lease cleanup, poll to a bounded quiet point and retry
transient filesystem observation errors. Report the terminal path, existence,
error, and observation count so an observer race is not classified as retained
work.
For bidirectional tab-traversal coverage, explicitly re-establish and verify the
boundary control before testing the opposite-direction wrap, and include logical
plus native focus identities in failures. For UI redraw-churn coverage, prove
settling and render-to-invalidation attribution first; hard-bound the invalidation
source instead of using a smaller fixed render cap that rejects attributable work
under broad-process load.

For a same-process UIA close/reopen scenario, wait for the old HWND to close,
release thread-held UIA clients, pump teardown, and reacquire the current window,
page, host, and provider on every bounded mutation attempt. Treat Value, Toggle,
and Selection operations as idempotent state-setting: success requires observing
the requested state, not only a successful provider call. Preferences category
selection must remain stable for repeated samples and reassert the target when
queued initialization restores an earlier page; take render/invalidation baselines
only after any subsequent focus transition also settles.

When a failure reproduces only in broad Commands order, run the failed case with
its immediate predecessor cluster in one process and repeat that cluster after the
repair. Then run canonical Commands once. A standalone exact-case pass does not
prove process cleanliness, and increasing the timeout does not repair stale UIA
identity or lifecycle state.

### Operation Startrail planning scenarios

Changes to affected-set planning, evidence lookup, or resume must measure
`validation.plan.warm_explain_ms`, `validation.plan.cold_snapshot_ms`, and the controlled
`validation.plan.fileops_resume_avoided_ms` comparison. Emit component timings for
snapshot acquisition, hashing, plan construction, evidence verification, serialization,
and runner launch. Archive same-machine evidence under
`Specs/TestRuns/<MachineHash>/Validation/<RunId>/`; treat cross-machine comparisons as
directional and never use estimated savings as a trust decision.

Report Git discovery, content hashing, canonicalization, impact matching, schema
validation, evidence lookup, receipt-closure binding, and entry fingerprints separately.
If Commands remains a monolithic executable, archive
`validation.plan.fileops_resume_avoided_ms` as blocked with zero Full avoidance; a focused
family process is not an independently reusable payload. Use an existing same-machine
Validation archive as baseline only when the fixture/profile still match.

### 4. Archive the evidence

After the run:

- confirm a new folder exists under `Specs/TestRuns/<MachineHash>/<Area>/...`
- keep `results.json`, `trace.txt`, and `perf_metrics.jsonl` intact
- use same-machine comparisons whenever possible

### 5. Report results correctly

Always state:

- baseline run,
- candidate run,
- exact metrics that improved or regressed,
- whether the comparison is same-machine,
- any residual blocker.

If the evidence is not same-machine or not apples-to-apples, call it directional.

### 6. Close out the spec correctly

If the change completes a WIP plan:

- move the plan to `Specs/Plans/Done/`,
- update the authoritative subsystem spec under `Specs/<Domain>/` (or repo-level guidance when appropriate),
- do not leave durable behavior or validation requirements living only in the finished plan.

## Patterns that are expected in this repo

- For queue-storm work, instrument coalesced counts and drained counts, not just total time.
- For UI work, prefer visible-latency metrics over raw callback timing alone.
- For large-list/grid work, keep bounded visible-work instrumentation.
- For rendering work, separate synchronous refresh cost from downstream paint when possible.
- For storage or background services, keep correctness selftests and perf measurements together.

## Definition of done

A perf-sensitive change is not done until:

- correctness coverage exists,
- perf measurement exists,
- an archived run exists or the block is explicitly documented,
- and any finished WIP plan has been closed out into `Specs/Plans/Done/` with the lasting contract reflected in the authoritative spec.
