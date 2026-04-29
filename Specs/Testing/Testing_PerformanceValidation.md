# Performance Validation Specification

## Overview

Performance work in RedSalamander is a **normative engineering requirement**, not an optional cleanup step.

Any new feature, behavior change, or optimization that can affect:

- UI responsiveness,
- startup time,
- folder enumeration,
- rendering,
- search,
- Compare Directories,
- File Operations,
- plugin I/O,
- memory retention,
- background queueing,

MUST define its validation path up front and MUST ship with measurable coverage from the beginning.

Related documents:

- `Specs/Testing/Testing_SelfTests.md`
- `Specs/TestRuns/README.md`
- `Specs/Plans/WIP/Perf_InstrumentationPlan_2026-03-26.md`

## Required Development Contract

### 1. Every perf-sensitive change MUST name its scenario

Before implementation is considered complete, the owning change MUST identify:

- the user-visible scenario being protected,
- the primary metric or metric family,
- the authoritative selftest or deterministic repro path,
- the expected direction of change.

Examples:

- “Find Files progress storm while results are visible”
- “Compare Directories content-progress queue churn”
- “Cold folder open with repeated icon indices”
- “FileOps pre-calc cancel latency”
- “Parsed ViewerText diff open, side-by-side scroll repaint, unchanged-text expansion, range-bounded referenced-file hydration, viewport rehydration, and unresolved placeholder fallback”

### 2. Every perf-sensitive change MUST have measurable evidence

A change MUST satisfy one of these:

- extend an existing metric family, or
- add new instrumentation, or
- explicitly justify why existing metrics already cover the scenario.

For optimizations, “seems faster” is not sufficient. The change MUST be supported by archived measurements or by a documented blocked reason.

### 3. New features MUST integrate tests and perf measurement from the start

When a new feature introduces a new hot path, queue, async pipeline, render surface, large-list path, or repeated callback flow:

- a deterministic selftest or deterministic repro harness MUST be added with the feature,
- at least one performance metric relevant to that path MUST be emitted with the feature,
- the expected archive path under `Specs/TestRuns/` MUST be part of the validation plan for the feature.

This requirement applies even when the first landing only establishes a baseline.

### 4. Optimizations MUST preserve correctness coverage

An optimization is incomplete if it improves a metric but lacks behavioral protection.

Any performance change MUST keep or add:

- correctness selftests,
- perf instrumentation,
- archived run evidence when the scenario is runnable on the current machine.

### 5. Claims of improvement MUST be grounded

A claimed performance improvement MUST state:

- the baseline run,
- the candidate run,
- whether the comparison is same-machine and same-suite,
- the metrics that improved,
- any material caveats.

If same-machine evidence is unavailable, the claim MUST be labeled directional rather than definitive.

## Perf Gate Template

Every perf-sensitive PR MUST include this gate in the PR description, review notes, or linked closeout spec before merge:

- Scenario: the exact user-visible path being protected.
- Metric keys: the emitted metric family or the existing metrics that cover the path.
- Deterministic validation: the selftest, focused harness, or deterministic repro that exercises the path.
- Archived evidence: the `Specs/TestRuns/<MachineHash>/<Area>/<RunId>/` folder containing `results.json`, `trace.txt`, and `perf_metrics.jsonl` when metrics are emitted.
- Before/after delta: baseline run, candidate run, same-machine status, changed metric values, and caveats.

If any field is blocked, the PR MUST state the blocker explicitly and identify the follow-up owner/task. A perf-sensitive PR is not complete with only manual timing notes or unarchived terminal output.

## Instrumentation Rules

### Metric design

Metrics SHOULD be:

- scenario-specific,
- deterministic enough for repeated selftest use,
- attributable to one stage of work,
- named consistently with the owning subsystem.

Examples:

- `find.ui.*`
- `compare.ui.*`
- `FileOps.*`
- `render.*`
- `icons.*`
- `viewer.diff.*`
Current ViewerText diff baselines include `viewer.diff.open_to_first_visible_us`, `viewer.diff.visible_rows`, `viewer.diff.semantic_row_paint_us`, `viewer.diff.visible_styled_rows`, `viewer.diff.visible_context_rows`, `viewer.diff.visible_banner_rows`, `viewer.diff.visible_split_rows`, `viewer.diff.theme_switch_repaint_us`, `viewer.diff.scroll_repaint_us`, `viewer.diff.hunk_jump_to_visible_us`, `viewer.diff.expand_context_us`, `viewer.diff.viewport_rehydrate_us`, `viewer.diff.viewport_backtrack_us`, `viewer.diff.deferred_rows`, `viewer.diff.referenced_bytes_read`, `viewer.diff.viewport_referenced_bytes_read`, `viewer.diff.viewport_referenced_bytes_delta`, `viewer.diff.viewport_backtrack_referenced_bytes_delta`, `viewer.diff.placeholder_rows`, `viewer.diff.placeholder_bands`, and `viewer.chrome.paint_us`.
The `viewer_text_diff_perf` artifact also records semantic-row paint timing, context-row visibility, pane-local side-by-side layout activation, visible split-row counts, pane column widths, rainbow-mode theme-switch repaint timing, root-shell chrome repaint timing, hunk-jump latency, and expand-time, post-viewport, and post-backtrack referenced-byte counts so clickable hidden-banner reveal, calmer base-background side-by-side structure, combo-sync behavior on scroll repaint, root-shell typography changes, horizontal-scroll presentation switches, bounded referenced-file growth, and cache reuse can be reviewed alongside the metric stream.
Preserve-viewport diff rehydrate/backtrack rebuilds should reuse the existing section-combo model and avoid full combo repopulation or header relayout unless the visible section set actually changes.

### Preferred measurements

When applicable, instrument:

- queue depth or queue drain behavior,
- message coalescing and skipped work,
- input-to-visible or progress-to-visible latency,
- rebuild/repaint counts,
- bytes processed or rewritten,
- lock hold time,
- pre-calc / sort / refresh stage timings,
- bounded visible-work counters for large lists or grids.

### Anti-patterns

Avoid relying only on:

- ad hoc logging with no archived metric,
- one-off manual stopwatch measurements,
- broad “total duration” metrics when the decision requires stage attribution.

## Test and Archive Requirements

### Selftests

Perf-sensitive features SHOULD prefer:

- `--commands-selftest` for window/UI behavior,
- `--compare-selftest` for Compare/search engine behavior,
- `--fileops-selftest` for File Operations.

If none of the existing suites can exercise the scenario, a new deterministic selftest case or equivalent harness MUST be added.

### Archival

Meaningful validation runs MUST be archived under `Specs/TestRuns/` per `Specs/TestRuns/README.md`.

If the repo archive path is unavailable, the blocked reason MUST be documented explicitly.

### Baselines

Once a scenario becomes important for ongoing work, a same-machine archived baseline SHOULD be recorded and reused for future comparisons.

## Spec Ownership Requirement

The owning normative spec for a subsystem MUST describe:

- the important user-visible performance scenarios,
- the expected test entrypoints,
- the authoritative metric families when they exist.

Performance requirements MUST NOT live only in a WIP plan when the subsystem has already adopted them as standard practice.

When a WIP plan is finished:

- the plan MUST move to `Specs/Plans/Done/`,
- any durable performance contract, verification entrypoint, or workflow requirement discovered in that work MUST be merged into the authoritative subsystem spec or repo-level guidance,
- the finished plan remains historical evidence, not the authoritative contract.

## Completion Criteria

A feature or optimization is performance-complete when:

1. the scenario is named,
2. the metrics exist or were shown to already exist,
3. deterministic coverage exists,
4. archived evidence exists or the block is documented,
5. the owning authoritative spec notes the lasting result when the change establishes or updates a baseline, and any finished WIP plan has been moved to `Specs/Plans/Done/`.
