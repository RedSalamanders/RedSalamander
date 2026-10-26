# Waiting File Operations cards and fixed destinations

## Checklist

- [x] Trace the independent counting, amount, source mini-progress and destination-menu paths.
- [x] Add an owned, no-activation production-popup regression fixture with synthetic paths.
- [x] Attempt baseline qualification: blocked by build-receipt rejection and a source change during build. No baseline scenario or performance improvement is claimed.
- [x] Suppress transfer/discovery facts and animation while queued, overlap-waiting or queue-paused; measure exactly the rows that paint.
- [x] Remove queued destination changes from the card, scheduler, worker, prompts, breadcrumbs, resources and obsolete tests.
- [x] Preserve queue ordering, Start now, speed settings, cancellation and ordinary collision/identity protections; retain their regressions and run private admission contracts in the no-activation review case.
- [x] Build and run affected, noninteractive regression coverage; inspect all eight actual harness screenshots.
- [x] Archive candidate frame metrics, results and failed-attempt evidence; record the unavailable baseline; update the authoritative UI and File Operations specs.
- [x] Record qualification limits and move this completed implementation plan to Done.

## Accepted behavior

A Waiting Copy/Move card shows its operation/selection, Waiting chip, source and fixed destination,
and applicable actions. It shows no Counting text, Written: 0 B, percentages, ETA, concurrency,
discovery activity, whole-task progress or source-row mini-progress. This also applies to stale
stream/discovery facts retained during a queue pause. Manual pause of an already running task
retains its existing truthful progress. Starting work restores the ordinary progress presentation.

The user explicitly requested removal of destination changes from both UI and code. Once accepted,
an operation keeps its destination. A different destination requires canceling and submitting a new
operation. Source/destination inspection remains an inspection action.

## Validation

Extend `FileOps_ReviewRemediationContracts` with production-popup witnesses for Copy/Move,
Queue/interlock/queue-paused state, English/French, narrow/light and wide/dark presentation. Assert
painted text, hosted accessibility regions, mini-bar count, status, paths, actions and unchanged
foreground. Save synthetic waiting-card captures and scenario metadata through the canonical WGC
harness. Emit `FileOps.Popup.WaitingCard.FrameUs` for 12 redraws per fixture and compare baseline and
candidate on the same machine. Existing card-hierarchy coverage protects running, discovery,
verification, terminal and compact states. Do not represent filtered results as Full qualification.

## Results and qualification limits

The [curated archive](../../TestRuns/4cb089111a23/FileOps/2026-10-20_waiting-cards/README.md)
retains passing review and decision receipts, the failed documentation expectation, failed compile
log, workspace identities, screenshot provenance/digests and all 288 waiting redraw metric rows.

- `FileOps_ReviewRemediationContracts`: Setup, phase and Cleanup passed; 24 waiting scenarios and
  124 broader gallery scenarios passed. The private admission/interlock probes now use their own
  engine, allowing the real stale clipboard-token witness to retain and pass its strict idle-owner
  assertion. English/French and narrow/light plus wide/dark captures were visually inspected.
- `FileOps_RepeatDecisionContracts`: Setup, phase and Cleanup passed on its own recorded artifact.
- Resource/localization contracts: 10 passed, none failed or skipped.
- Final x64 Debug build: zero warnings/errors; receipt
  `ffb48f219067b67f43a5a79e7e2a379de9c6b7c6dd0f525da7d946d3d46b0fae`.
  Its final test-only change re-fetches a queued neighbor after potentially reentrant UIA invocation.
  Production source bytes match the recorded review snapshot; the interactive Phase5 journey itself
  was not run. Earlier focused receipts are not promoted into a Full/current-binary gate.
- Candidate redraw timings: 288 samples, median 8,072 us, nearest-rank p95 8,645 us, max 26,172 us.
  There is no usable baseline and no speedup, throughput or performance-budget pass claim.
- Archive explicit pre-stage validation passed for 28 files (about 1.5 MB); whole-inventory
  validation passed for 15,901 files. Full repository, Release/ASan/ARM64 and focus-taking
  qualification remain pending.

The implementation is complete within this scoped request. This plan does not close the broader
File Operations work or establish repository-wide production qualification.
