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
