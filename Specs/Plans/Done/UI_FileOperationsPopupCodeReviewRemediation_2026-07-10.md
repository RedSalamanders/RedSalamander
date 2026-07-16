# File Operations Popup - Code Review Remediation Plan

> **Executor instructions:** This plan owns the correctness, responsiveness, persistence, test, and
> architecture findings discovered while reviewing the File Operations popup implementation at
> `78366858d`. The product behavior remains owned by
> `../WIP/UI_FileOperationsPopupUxRefinementPlan_2026-07-07.md`; do not create a second UI contract here.
> Re-anchor every symbol before editing because line numbers will drift. Preserve concurrent WIP
> documents and unrelated working-tree changes. Complete and verify one track at a time.

## Status

- **Status:** Completed on 2026-07-11. Tracks A-E remain implemented; replacement same-machine
  performance evidence, focused correctness/stress results, and the classified external Full-suite
  blocker are curated and audited.
- **Category:** correctness, responsiveness, architecture, accessibility, persistence, tests, DX
- **Priority:** P1-P3
- **Planned at:** `78366858d` on 2026-07-10
- **Scope:** File Operations state, popup rendering/input, completed-task navigation, settings,
  deterministic self-tests, performance evidence, and the checked-in HTML mockup
- **Closeout owner:** `../WIP/UI_FileOperationsPopupUxRefinementPlan_2026-07-07.md`

## Goal

Make the implemented popup safe and predictable under remote providers, minimized windows, mixed
known/unknown progress, reduced-motion mode, minimum-width layouts, popup recreation, and provider
changes. Remove avoidable duplication in the touched paths, prove the behavior with deterministic
self-tests and archived performance evidence, then merge durable requirements into the authoritative
File Operations specification before the completed review plans move to `Specs/Plans/Done/`.

## Non-goals

- Do not redesign the File Operations engine or add the product backlog items B1, C5, or C7 here.
- Do not modify shared DxUi/UIA snapshot-coalescing code owned by another active task.
- Do not retain filesystem COM objects in completed-task history; history stores stable identity and
  recreates/resolves the provider through the normal navigation pipeline.
- Do not make synchronous remote metadata or settings I/O part of popup paint, layout, or mutex-held
  snapshot collection.

## Review Findings

| ID | Priority | Finding | Required result | Status |
|----|----------|---------|-----------------|--------|
| CR-1 | P1 | Conflict metadata and local type probes call filesystem APIs, including `CreateFileReader` and `GetFileAttributesW`, while the conflict arbiter mutex is held. A remote or mapped provider can block prompt publication and popup snapshots. | Make locked action construction I/O-free, publish the prompt under the lock, perform metadata decoration outside it, never open/read file content for decoration, and conditionally merge metadata/actions only into the same live prompt. | **Fixed and covered** |
| CR-2 | P2 | Taskbar progress is updated from paint, while the timer returns early for hidden/iconic windows. Minimized progress can freeze. | Maintain taskbar state from the timer independently of paint and prove minimized updates. | **Fixed and covered** |
| CR-3 | P2 | Completed destination actions remember only a pane/path. If the pane provider changes, actions run against the wrong filesystem; fallback path joining also assumes Windows separators. | Capture destination plugin ID, short ID, and instance context; build provider-qualified navigation; use provider-neutral leaf/join helpers; reject stale/unresolvable targets instead of navigating the wrong provider. | **Fixed and covered** |
| CR-4 | P2 | Queue/Parallel is one blind toggle hit target although it is drawn as two segments. | Give each segment its own stable hit target and desired state; clicking the selected segment is a no-op. | **Fixed and covered** |
| CR-5 | P2 | Reduced-motion mode still advances indeterminate progress marquees. | Render a stable centered indeterminate segment when reduced motion is enabled, across global, task, file, and conflict bars. | **Fixed and covered** |
| CR-6 | P2 | Aggregate ETA divides known remaining bytes by throughput from all active tasks, including tasks with unknown totals. | Suppress aggregate ETA whenever any active transfer has unknown bytes or no matching usable rate; retain aggregate throughput text. | **Fixed and covered** |
| CR-7 | P2 | At minimum width, footer controls/summary disappear or collapse to zero width and the debug snapshot reports visibility that is not present. | Use one responsive footer layout model with non-overlapping, actionable controls and truthful debug geometry at the supported minimum width. | **Fixed and covered** |
| CR-8 | P2 | Footer-only collapse remembers the expanded rectangle only in the popup object. Popup recreation or restart can restore only the tiny footer rectangle. | Persist a separate normalized expanded placement; never overwrite it with footer-only bounds; restore it after recreation/restart. | **Fixed and covered** |
| CR-9 | P3 | Popup render/layout code repeats global summary scans, reduced-motion palette construction, easing, footer geometry, resource loads, and settings-save error blocks. | Consolidate the touched reducers/layout/settings paths without a broad renderer rewrite; cache frame-stable values and keep behavior deterministic. | **Fixed in scoped paths** |
| CR-10 | P3 | The non-default File Operations settings predicate is owned by a RedSalamander header although Common also serializes the structure. | Move the predicate to `Common::Settings` and use one definition from serialization and runtime pruning. | **Fixed and covered** |
| CR-11 | P2 | Existing tests directly invoke actions and inspect synthetic visibility, missing real hit testing, popup recreation, minimized taskbar updates, and remote metadata behavior. | Add deterministic tests at the actual message/hit/layout/state boundaries for every correctness finding. | **Fixed and covered** |
| CR-12 | P3 | The documentation mockup loads mutable `lucide@latest` from a CDN. | Make the mockup hermetic by pinning/vendoring the exact asset or replacing it with checked-in equivalents; no runtime network dependency. | **Fixed and browser-verified** |

## Track A - Conflict Prompt Responsiveness

### A1. Separate prompt publication from metadata decoration

1. Refactor `Task::SetConflictPromptLocked` so it only initializes the prompt payload and actions.
   Action construction must be pure and must not call Win32 or provider filesystem APIs.
2. In `BeginConflictPrompt`, publish the basic prompt and notify the popup before metadata work.
3. Capture immutable provider/path inputs, release `_conflictArbiter.mutex`, then query metadata.
4. Reacquire the mutex and merge results only if owner thread, bucket, status, source path,
   destination path, and active state still match the captured prompt.
5. Do not let late metadata overwrite a new prompt or a prompt that has already been resolved.

### A2. Make metadata decoration bounded and content-free

1. Remove `CreateFileReader`/`GetSize` from `ReadConflictItemMetadata`.
2. For local absolute paths, use `GetFileAttributesExW` for attributes, size, and last-write time.
3. For provider paths, use metadata APIs only; a missing size remains `sizeKnown=false`.
4. Treat unsupported/recoverable metadata errors as unavailable decoration, not operation failure.
5. Add `FileOps.Conflict.MetadataUs` and count/availability values without logging normal misses.

### A3. Proof

- A blocking test provider must show that the prompt becomes active and popup snapshot collection
  remains responsive before metadata returns.
- A provider spy must assert that no reader is created for conflict decoration.
- A source scan and the blocked-metadata prompt test must prove that prompt publication performs no
  `GetFileAttributesW` or provider metadata call under the arbiter mutex.
- A stale metadata completion must not mutate a replacement prompt.
- Explicit cross-provider operations must not inherit a pane plugin identity when the supplied
  destination filesystem object does not match that pane.
- Archive the responsiveness run under `Specs/TestRuns/<commit>/FileOps/...`.

## Track B - Provider-Aware Completed Actions

### B1. Preserve destination identity

1. Extend `Task`, `CompletedTaskSummary`, and popup `TaskSnapshot` with destination plugin ID,
   plugin short ID, and instance context.
2. Capture those strings when `StartOperation` accepts the task, before the destination pane can
   navigate elsewhere. For injected cross-filesystem operations, resolve identity from the accepted
   destination pane/context and keep empty identity explicit when no pane exists.
3. Copy identity into completed history; never retain the provider COM object there.

### B2. Build correct destination locations

1. Replace `std::filesystem::path` concatenation used for provider paths with `GetPathLeaf`,
   `GuessPreferredSeparator`, and `JoinFolderAndLeaf`.
2. Add one helper that formats the destination folder through
   `NavigationLocation::FormatHistoryPath(pluginShortId, instanceContext, pluginPath)`.
3. Route **Open destination** and **Reveal item** through this qualified location.
4. Validate current/target provider identity case-insensitively where identifiers are stable.
5. If the plugin or instance cannot be resolved, return a recoverable failure and leave the current
   pane unchanged; do not silently execute against its current provider.

### B3. Proof

- Complete a synthetic virtual-provider task, switch the destination pane to local storage, invoke
  both actions, and verify the qualified provider/context/path selected by navigation.
- Cover `/folder` + `long-name.ext`, root paths, empty destination path, and local Windows paths.
- Verify stale or missing provider identity disables/rejects actions rather than using the active pane.

## Track C - Live Progress and Motion

### C1. Taskbar timer path

1. Extract a lightweight `UpdateTaskbarProgress` path from paint.
2. Run it from the popup timer while the window is iconic or occluded; do not invalidate/render solely
   to update the taskbar.
3. Keep COM initialization retry/backoff behavior and clear taskbar state after the last active task.
4. Avoid duplicate snapshot scans in a visible frame by reusing a frame summary where practical.

### C2. Aggregate ETA

1. Keep aggregate throughput over active transfer rates.
2. Set aggregate ETA unavailable if `hasUnknownActiveProgress` is true.
3. For fully determinate tasks, derive remaining bytes and the matching rate set consistently.
4. Preserve saturating conversions and the sub-byte display floor.

### C3. Reduced motion

1. Resolve reduced-motion once for a render/layout cycle from the cached palette.
2. Centralize indeterminate-fill calculation.
3. In reduced-motion mode return a fixed centered segment; otherwise retain the existing animation.
4. Apply the helper to every global/task/file/conflict indeterminate bar.
5. Keep graph-value interpolation functional: reduced motion may remove decorative motion but must not
   make numeric state stale.

### C4. Proof

- Minimize the popup, advance synthetic progress through timer messages, and assert the last applied
  taskbar model changes without paint.
- Cover known + unknown active tasks and assert throughput remains while ETA is absent.
- Compare indeterminate geometry at two ticks with reduced motion on/off.

## Track D - Footer Input, Layout, and Placement

### D1. Segmented Queue/Parallel control

1. Allocate separate Queue and Parallel rectangles in the footer layout model.
2. Emit two hit targets carrying the desired queue state.
3. Apply the requested state directly; do not invert the current setting.
4. Animate the selected background between segment rectangles unless reduced motion is enabled.
5. Expose the two rectangles and selected state in the debug layout snapshot.

### D2. Responsive minimum-width footer

1. Create a single `FooterLayout` result consumed by layout, render, hit testing, accessibility/debug
   snapshots, and footer-only height calculations.
2. At narrow width, use familiar icon-only actions with localized tooltips for secondary controls;
   keep Cancel/Pause and Queue/Parallel understandable and actionable.
3. Place the global summary in a dedicated row or progress overlay rather than assigning zero width.
4. Assert every visible rectangle has positive size, lies inside the client area, and does not overlap.
5. Derive `globalSummaryVisible` and button counts from actual rectangles.

### D3. Expanded placement persistence

1. Add a dedicated settings window key for the expanded File Operations popup placement.
2. Save it immediately before entering footer-only mode and whenever an expanded popup saves placement.
3. When footer-only, save current compact placement only to the normal popup key if still required for
   launch positioning; never replace the expanded key with compact bounds.
4. Load the expanded key during popup creation and use it when expanding after recreation/restart.
5. Normalize for current DPI/monitor using the existing placement utilities.

### D4. Proof

- Send clicks to the center of both real segment rectangles, including clicking the selected segment.
- Resize the real popup to the supported minimum client width and assert all required footer controls,
  summary, and hit targets are present and non-overlapping.
- Collapse, destroy/recreate, expand, and assert the previous expanded rectangle is restored within
  DPI rounding tolerance.

## Track E - Scoped Architecture and Settings Cleanup

### E1. Shared reducers and frame state

1. Combine global status/rate enrichment where it removes a full duplicate task scan.
2. Cache frame-stable resource strings used repeatedly during one render; do not introduce a global
   localization cache that cannot react to language changes.
3. Reuse one easing primitive for equivalent cubic timing while keeping named wrappers for intent.
4. Use the shared footer model from Track D instead of parallel geometry branches.
5. Keep `Render` extraction scoped to coherent helpers; do not rewrite the full 9k-line popup file.

### E2. Shared settings ownership

1. Declare `Common::Settings::HasNonDefaultFileOperationsSettings` beside the settings model.
2. Implement it once in Common and call it from serializer gating, save preparation, and runtime
   pruning.
3. Consolidate repeated save-and-log blocks into one File Operations settings helper.
4. Cover every field, including compact density, footer-only state, auto-dismiss, queue mode, speed,
   parallelism, pre-calculation, and bridge buffer settings.

### E3. Hermetic mockup

1. Remove `lucide@latest` from the mockup.
2. Prefer a checked-in pinned minified asset with license notice; if icons are replaced, preserve the
   visual behavior and accessible labels.
3. Open the mockup with network disabled and verify there are no console/resource errors.

## Implementation Evidence (2026-07-10)

- Final application build: `.build/logs/msbuild-redsalamander-serial-20260710_162230_027.log`,
  Debug x64, 0 warnings / 0 errors. Two preceding parallel solution-target attempts ended in
  unrelated, diagnostic-free tool exits in different projects; the deterministic single-worker
  application build compiled and linked the complete changed target.
- Focused Commands: global-summary and settings-roundtrip cases pass 3/3. The final conflict-popup
  case passes 3/3 at `Specs/TestRuns/7d3a1247382a/Commands/2026-07-10_162638`. It covers metadata
  lock release, pure locked action construction, real 480-DIP layout/hit targets,
  minimized timer updates, footer-only recreate/restore, completed provider navigation/rollback,
  and Auto-dismiss label fallback.
- Focused FileOps: `Phase5_PreCalcCancelReleasesSlot` passes 3/3, including immediate explicit
  reduced-motion semantics and delayed-layout convergence.
- Conflict state-machine RED/GREEN: the first deferred-action implementation failed
  `Phase9_ConflictPrompt_OverwriteAutoCap` 3/3 at
  `Specs/TestRuns/7d3a1247382a/FileOps/2026-07-10_162109`. After correcting destination filesystem
  identity and retaining cross-provider action semantics, the complete Phase 9 family passed 27/27
  across three repetitions at `Specs/TestRuns/7d3a1247382a/FileOps/2026-07-10_162621`.
- Full FileOps: `Specs/TestRuns/7d3a1247382a/FileOps/2026-07-10_142225`, 102 passed, 0 failed,
  20 expected credential/7z skips, clean disk audit.
- Clean isolated Commands families on the final binary after the Full-run UIA collapse:
  - `Specs/TestRuns/7d3a1247382a/Commands/2026-07-10_162824`: all
    `cmd_pane_fileops_popup_*` cases passed, 1/1.
  - `Specs/TestRuns/7d3a1247382a/Commands/2026-07-10_162802`: all
    `cmd_pane_fileops_issues_pane_*` cases passed, 18/18.
  - `Specs/TestRuns/7d3a1247382a/Commands/2026-07-10_162815`: all
    `cmd_preferences_dialog_file_operations_*` cases passed, 6/6.
- Full-suite attempt: the outer runner reached its one-hour limit while Commands shuffle
  classification was still active. The initial and shuffle Commands processes exhibited a
  suite-wide UI Automation collapse (107 and 93 failures respectively across unrelated windows),
  so this attempt is not recorded as green. Raw results, traces, and the process snapshot are under
  `Specs/TestRuns/7d3a1247382a/Continuation/2026-07-10_1552_full_runner_timeout`. Every affected
  FileOps family passed in the clean isolated runs above; broad convergence remains owned by
  `Operation_TestSuiteStabilization_FlakeConvergence_2026-07-04.md` in another task.
- Product capture: `Specs/UI/Images/FileOperationsPopup_Product_Conflict_2026-07-10.png` was captured
  from the final built popup through DWM compositor bounds, not from the HTML mockup.
- Mockup: vendored Lucide 1.23.0 has no external scripts; in-app browser inspection found 26 rendered
  SVG icons and no console/resource errors.

### Review follow-up (2026-07-10)

- The minimized taskbar regression now captures its baseline after `SW_MINIMIZE`, while the popup is
  confirmed iconic, immediately before the synthetic timer tick. This makes the counter comparison
  depend on the minimized `WM_TIMER` path instead of intervening paints.
- Local-to-local `Exists` conflicts initially withhold Overwrite until source/destination metadata
  proves the target is replaceable. Cross-provider prompts retain provider-defined Overwrite
  semantics. The new `Phase9_ConflictPrompt_LocalFileOntoDirectory` case proves that a local file
  targeting an existing directory never exposes Overwrite or mutates either side.
- Completed-task Open/Reveal now performs one qualified `SetFolderPath` navigation. It no longer
  navigates to the provider root first; the Commands regression clears and inspects folder history to
  prove that only the requested destination is visited.
- The unreachable `leaf.empty()` reveal-path clause was removed.
- `IsByteRateUsableForEta(etaBytesPerSec)` was retained deliberately. Although ordinary finite
  per-task rates make it redundant, the aggregate sum can overflow to infinity; retaining the guard
  prevents a misleading zero ETA if that invariant is violated.
- Final deterministic application build:
  `.build/logs/msbuild-20260710_175008_482.log`, Debug x64, 0 warnings / 0 errors.
- Focused conflict popup:
  `Specs/TestRuns/7d3a1247382a/Commands/2026-07-10_175449`, 3 passed / 0 failed. This includes the
  fresh iconic taskbar baseline, one-hop provider navigation, and local conflict action publication.
- Phase 9 conflict family:
  `Specs/TestRuns/7d3a1247382a/FileOps/2026-07-10_175630`, 30 passed / 0 failed across three
  repetitions, including the new local file-to-directory case.
- Fairstream conflict and bridge family:
  `Specs/TestRuns/7d3a1247382a/FileOps/2026-07-10_175748`, 20 passed / 0 failed.
- Final full FileOps:
  `Specs/TestRuns/7d3a1247382a/FileOps/2026-07-10_181326`, 103 passed / 0 failed / 20 expected
  environment skips, clean disk audit.
- Final linked Commands families:
  `Specs/TestRuns/7d3a1247382a/Commands/2026-07-10_174720` (Issues pane 18/18),
  `2026-07-10_174730` (File Operations Preferences 6/6), and `2026-07-10_174735` (popup 1/1), all
  with clean disk audits.
- Source contracts and inventory: 107 Pester tests passed / 0 failed after aligning the added phase,
  pause point, and documented inventory counts.
- A failed broad attempt caused by an accidentally orphaned first self-test process is preserved at
  `Specs/TestRuns/7d3a1247382a/Continuation/2026-07-10_1726_fileops_cross_process_interference`.
  The clean isolated rerun above is authoritative.
- The user-reported Batch Rename stack was the historical
  `RedSalamander-20260622-170227-p6632.txt`. Commit `4bf3c54e9` already moved the snapshot traversal
  out of the worker thread; the exact former crash case passed 10/10 at
  `Specs/TestRuns/7d3a1247382a/Commands/2026-07-10_173145`.

Provider metadata calls can still take provider-defined time on the task worker. The fixed invariant
is that prompt publication, popup snapshots, and unrelated UI work do not wait behind the conflict
arbiter mutex; decoration remains optional and content-free.

### Evidence audit (2026-07-11)

- The 15 `Specs/TestRuns/7d3a1247382a/...` paths cited above are absent from the reviewed Git tree and
  local worktree.
- No available checked-in artifact contains `FileOps.Conflict.MetadataPromptUs` or
  `FileOps.Conflict.MetadataUs`.
- The closeout records no same-machine baseline/candidate pair, metric sample counts, median/p95,
  `Tools/Show-PerfRuns.ps1` analyzer output, or explicit keep/revert decision.
- The historical pass statements above are retained as provenance, but they are not accepted as
  current reviewable proof. Replacement evidence is owned by
  `CodeReview_Last3Days_Remediation_2026-07-11.md`.

### Replacement evidence (2026-07-11)

- The auditable same-machine baseline/candidate comparison is curated at
  `../../TestRuns/4cb089111a23/Commands/2026-07-11_130136/README.md`. It records exact commands,
  hashes, raw-artifact locations, sample counts, p50/p95/p99 values, analyzer thresholds, and the
  keep decision.
- The three isolated successors of the former monolithic popup case passed 300/300 repetitions.
  `FileOps.Conflict.MetadataPromptUs` p95 improved 9.2%, `FileOps.Conflict.MetadataUs` p95 improved
  8.9%, popup render p95 improved 3.0%, and snapshot construction p95 improved 13.5% on the same
  machine. No protected p95 regressed.
- The final focused settings/provider runs passed 4/4 once
  (`20260711T144955Z-72592-3b5da8574c4146c8a3250583e46898e6`) and 20/20 across five shuffled
  repetitions (`20260711T145016Z-40856-5338a70cc4394330b9025e100654342c`). They prove exact atomic
  stamp ownership, external/failed-save race handling, bounded process teardown, and restoration of
  local plus qualified 7z pane locations without empty models or double qualification.
- Focused tooling closeout passed inventory 5/5, source contracts 132/132, documentation drift 9/9,
  TestRuns archive policy 5/5, build-process collision guards 4/4, and localization 4/4.

## Verification Matrix

| Gate | Command / evidence | Required result |
|------|--------------------|-----------------|
| Drift | `git diff --check` and focused `git diff` review | No whitespace errors; no unrelated WIP changes overwritten |
| Build | `./build.ps1 -Configuration Debug -Platform x64 -ProjectName RedSalamander` | 0 errors, 0 warnings |
| FileOps focused | `./Tools/Run-AllTests.ps1 -SkipBuild -Suite FileOps -CaseFilter <new cases>` | Every new state/metadata case passes |
| Commands focused | `./Tools/Run-AllTests.ps1 -SkipBuild -Suite Commands -CaseFilter <new cases>` | Real popup layout/input/taskbar/placement cases pass |
| Existing popup regression | Run all existing `cmd_pane_fileops_popup_*` cases | No regression |
| Performance | Archive conflict metadata/prompt responsiveness and popup render/layout metrics | No provider read; prompt publication bounded; no material frame regression |
| Full | `./Tools/Run-AllTests.ps1 -Suite Full` | Green, or any unrelated baseline failure documented with exact prior evidence |
| Documentation | Real product screenshots at normal, narrow, footer-only, conflict, and completed states | Images come from the built product, not the mockup |

## Performance Contract

- **Scenario FO-POP-CR-1:** remote provider delays metadata for a conflict.
  - Metrics: per-prompt `FileOps.Conflict.MetadataPromptUs` and per-operation
    `FileOps.Conflict.MetadataUs`.
  - Invariant: prompt active/snapshot available before provider metadata completes.
  - Invariant: metadata decoration creates zero content readers.
- **Scenario FO-POP-CR-2:** 16 active tasks, popup visible, 100 ms timer.
  - Existing render/layout metrics remain within noise of the pre-change baseline.
  - Minimized taskbar updates do not trigger paints.
- **Scenario FO-POP-CR-3:** repeated width changes around minimum width.
  - Footer layout remains deterministic and allocates no unbounded per-frame state.

Record machine/build/commit, warmup, sample count, median/p95, correctness counts, raw artifact path,
and the keep/revert decision in `Specs/TestRuns/` according to the performance validation contract.

## STOP Conditions

- Stop Track A if a provider metadata API itself requires content download; leave size unknown rather
  than reading content.
- Stop Track B if provider-qualified navigation would need retaining a plugin module/COM instance in
  history; add a stable resolver through `FileSystemPluginManager` instead.
- Stop Track D if supporting the current minimum width would require overlapping or sub-minimum hit
  targets; raise the enforced minimum width deliberately and document the value rather than clipping.
- Stop and diagnose any new crash, hang, payload leak, UI-thread provider wait, or test timeout before
  continuing to later tracks.
- Do not bless a red Full suite as green. Re-run an isolated failure and compare with a pre-change
  baseline before classifying it as unrelated.

## Execution Checklist

- [x] Review current implementation and record findings.
- [x] Implement and prove Track A.
- [x] Implement and prove Track B.
- [x] Implement and prove Track C.
- [x] Implement and prove Track D.
- [x] Implement and prove Track E.
- [x] Re-run focused build/tests and archive the referenced correctness artifacts.
- [x] Produce same-machine baseline/candidate performance evidence with analyzer output and sample
  quality for the conflict-metadata and popup-render scenarios.
- [x] Run the Full suite and reconcile failures without treating missing or external artifacts as
  green evidence. The attempt stopped in unrelated concurrent ViewerSqliteTests work; the curated
  evidence records exact diagnostics and does not call the Full suite green.
- [x] Capture real-product screenshots and update authoritative documentation.
- [x] Reconcile `UI_FileOperationsPopupReviewFindings_2026-07-09.md` S1-S5.
- [x] Update `UI_FileOperationsPopupUxRefinementPlan_2026-07-07.md` with final status/evidence.
- [x] Move this plan back to `Specs/Plans/Done/` only after the replacement evidence is complete and
  the authoritative specifications contain the final measured contract.
