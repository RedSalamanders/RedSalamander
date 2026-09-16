# File Operations: direct provider API discovery contract

## Status and ownership

- Status: COMPLETE; implemented and verified on 2026-09-14. See Implementation record.
- Index owner: I23.
- Priority: P3. No current host route reaches the defective paths, so this is an API-contract repair, not a user-visible popup repair.
- Effort: M.
- Change risk: low for the bulk entry points, medium for Delete because its reporting shape is shared with the recycle route.
- Planned at: `63d78f55`, 2026-09-14. Opened by I21 (`FileOperations_DiscoveryScopeAndLeafProgress_2026-09-13.md`) as the routed owner of its D5 residual.
- Dependencies: I21 owns the host bridge and the host-side discovery contract. I22 owns non-Local provider parity. This plan owns only the built-in Local plugin's direct API surface.
- Authoritative contracts: `Specs/Plugins/Plugins_VirtualFileSystem.md`, `Specs/FileSystem/FileSystem_FileOperations.md`.

This plan is self-contained and non-normative.

Run this drift check before implementation:

```powershell
git diff --stat 63d78f55..HEAD -- Plugins/FileSystem RedSalamander/FolderWindow.FileOperations.State.cpp Specs/Plugins/Plugins_VirtualFileSystem.md
git status --short
```

## Problem and required outcome

The built-in Local plugin is the only provider that reports discovery in production, and its per-item entry points are correct. Its bulk and Delete entry points are not. They are reachable by any consumer of the public `IFileSystem` surface, and the host itself rejects bulk transfer before I/O today (`RedSalamander/FolderWindow.FileOperations.State.cpp:17343`), so the gap is latent rather than live.

Required outcome: every Local mutation entry point either reports cumulative discovery for its governed call and closes it exactly once, or documents in the authoritative spec why that call has no discovery scope of its own. A consumer must not have to know which entry point it picked to get a truthful denominator.

## Prioritized findings

Anchors verified at `63d78f55` in `Plugins/FileSystem/FileSystem.FileOps.cpp` unless stated otherwise.

| ID | Finding and evidence | Impact | Effort / risk / confidence | Disposition |
|---|---|---|---|---|
| A1 | `CopyItems` never calls `ReportTopLevelDiscovery`. The sequential branch at `:9741` and the parallel worker at `:9864` go straight from `SetItemPaths` into `CopyPathInternalWithDirectoryParallelism`. The only four call sites of the helper are `:9243` (`CopyItem`), `:9394` (`MoveItem`) and `:10010`/`:10134` (`MoveItems`). | A bulk Copy consumer gets no discovery at all: no totals and no closure. | S / low / high | Report per root through the shared call scope, then close once at call level. |
| A2 | `MoveItems` reports each root closed through one shared `OperationContext` (`:10010` sequential, `:10097-10099` parallel workers sharing `sharedOptionsState`). Each root emits `traversalClosed=TRUE` for the whole call. | The first root closes the call's scope while later roots are still being moved. This is the same shape as the defect I21 repaired in the host. | M / medium / high | One cumulative accumulator per call; exactly one closure after the last root. |
| A3 | Direct Delete reports open discovery and never closes it. `ReportDeleteDiscoveryObject:4499` always passes `traversalClosed=false` at `:4529`; its callers at `:7532` (recycle) and `:7571` (permanent) are the only producers, and `DeleteItem:9511` has no closing report anywhere in its body. | A direct Delete consumer sees a scope that never closes. The host survives only because its own per-item terminal guard closes the task. | M / medium / high | Close at the end of the governed call, including the error and partial paths. |
| A4 | `CopyItem` suppresses the root record when `context.recursive` and the root is a real directory (`:4486`), deliberately leaving the recursive walker as the owner of root and descendant totals. | Correct by design, and the same rule I21 wrote into the host. | n/a | Keep. Record it in the spec so it is not mistaken for A1. |
| A5 | Bound leaf `DeleteIfUnchanged` reports one closed record (`Plugins/FileSystem/FileSystem.cpp:3195`), gated on `_kind != FILESYSTEM_BOUND_DIRECTORY`, so bound directory cleanup reports nothing. | Correct for a standalone bound delete. It was only harmful when a host traversal passed its own root cookie into it, which I21 fixed on the host side. | n/a | Keep. Document the asymmetry so the next reader does not "fix" it. |

## Scope

- `Plugins/FileSystem/FileSystem.FileOps.cpp` and `Plugins/FileSystem/FileSystem.cpp` discovery reporting only.
- The Local plugin's existing contract tests, plus a new direct-API discovery case registered in `kFileOpsFamilyDefinitions`.
- `Specs/Plugins/Plugins_VirtualFileSystem.md` and this plan.

Do not enable bulk Copy or Move in the host. The rejection at `State.cpp:17343` is deliberate and stays. Do not change delete semantics, recycle routing, identity pinning, or the parallel scheduling shapes; this plan changes what is reported, not what is mutated.

## Execution checklist

### B0 — Characterize

- [x] Add a direct-API test that calls `CopyItems`, `MoveItems` and `DeleteItem` with a recording operation control, and asserts the current event streams. Expect: no events for `CopyItems`, an early closure for `MoveItems`, and no closure for `DeleteItem`.
- [x] Register the new case in a family; an unregistered step is skipped silently by unfiltered runs.

### B1 — Bulk Copy and Move

- [x] Give `CopyItems` per-root reporting through the shared call scope, matching `CopyItem`'s recursive-directory carve-out exactly.
- [x] Replace `MoveItems`' per-root closure with one cumulative accumulator and a single call-level closure after the last root, in both the sequential and parallel branches. The accumulator must be safe under the parallel workers that already share the call's options.
- [x] Keep counters monotonic within the call's cookie so a retry cannot be read as new discovery.

### B2 — Delete

- [x] Close the governed Delete call's discovery scope exactly once at the end, on success, failure, partial and cancellation.
- [x] Keep direct, recycle and bound permanent Delete distinct. Shell-owned recursive recycling may legitimately expose only the selected roots; say so rather than inventing descendant counts.

### B3 — Contract and closeout

- [x] Extend the provider-contract test so a Local entry point that mutates without a closing discovery report fails.
- [x] Merge the resulting obligation, including the two by-design exceptions A4 and A5, into `Specs/Plugins/Plugins_VirtualFileSystem.md`.

## Commands and verification gates

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter FileOpsFamily_DiscoveryScope -TestRoot X:\RedSalamander.Perf -AllowAlternateVolumeTestRoot
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter FileOpsFamily_Phase05_Discovery -TestRoot X:\RedSalamander.Perf -AllowAlternateVolumeTestRoot
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter FileOpsFamily_Phase10_DeleteValidation -TestRoot X:\RedSalamander.Perf -AllowAlternateVolumeTestRoot
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -Configuration Debug -TestRoot X:\RedSalamander.Perf -AllowAlternateVolumeTestRoot
git diff --check
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
```

## Completion checklist

- [x] A1, A2 and A3 are repaired with failing-before and passing-after witnesses.
- [x] A4 and A5 are recorded in the authoritative spec as deliberate, with their reasons.
- [x] Host behavior is unchanged: bulk transfer stays rejected and `FileOps.Discovery.GrowthAfterClose` stays 0 across the FileOps suite.
- [x] Fresh Full passes; `git diff --check` and the spec inventory pass; the completed plan moves to `Specs/Plans/Done/` and I23 leaves the active index.

## Implementation record

Implemented and verified on 2026-09-14 from `63d78f55`. Every anchor in this plan was re-verified against the source before the work began; A1, A2, A3, A4 and A5 all reproduced exactly as written.

### What changed

All in `Plugins/FileSystem/FileSystem.FileOps.cpp`.

- Two new provider-local types. `CallDiscoveryAggregator` holds one governed call's cumulative totals behind a mutex. `RootDiscoveryTranslation` holds the last cumulative values one selected root reported, so a root that restarts its own stream at zero contributes a delta rather than a regression. `OperationContext` carries a `discoveryTranslation` pointer that only a bulk entry point sets.
- The raw emitter is now `ReportRawDiscoveryProgress`. The name `ReportDiscoveryProgress` belongs to a translating wrapper that every traversal already called, so no traversal needed changing: a single-root call passes its own totals straight through and owns its closure, while a root inside a bulk call is folded into the call's aggregate and has its closure withheld. Only the bulk entry point knows when the last root is done.
- **A1.** `CopyItems` now calls `ReportTopLevelDiscovery(context, source.extended, context.recursive)` per root in both the sequential and the parallel branch, with the same recursive-directory carve-out `CopyItem` uses, and calls `rootDiscovery.BeginRoot()` before each root.
- **A2.** `MoveItems` no longer lets each root close the call. Both branches share one aggregator, each root begins its own translation, and the single closure runs from a `wil::scope_exit`.
- **A3.** `DeleteItem` and `DeleteItems` close their governed scope from a `wil::scope_exit` calling `ReportDeleteDiscoveryClosed`, so error, partial and cancellation paths cannot leave a caller waiting on a scope forever. Delete already accumulated into one call-level `DeleteDiscoveryState`; only the closure was missing.
- The aggregate is advanced **and emitted** under the same `std::scoped_lock`. This is load-bearing, not stylistic — see the witness table.
- **A4 and A5** are unchanged in code and are now written into `Specs/Plugins/Plugins_VirtualFileSystem.md` as deliberate silences, so neither is later mistaken for the defect above.

`AddSaturatingDiscoveryTotal` keeps the aggregate from wrapping. `queuedItems` takes the maximum rather than a sum, because it is a queue depth, not a cumulative count.

### Regressions added

| Case | What it proves |
|---|---|
| `DiscoveryScope_LocalDirectApi` | Drives `CopyItems`, `MoveItems` and `DeleteItem` against the provider's own API with no host in the loop, through a recording `IFileSystemOperationControl`. For each call it asserts the stream never decreases, that there is exactly one closure, that the closure is the last event, and that it closed on the exact expected totals. It reports every defect it finds in one run rather than stopping at the first. |

Registered in `kFileOpsFamilyDiscoveryScope`, so unfiltered FileOps and Full runs execute it.

A source contract, `closes the discovery scope of every Local mutation entry point exactly once`, fails any of the six Local mutation entry points whose body lacks its closing construct, asserts both aggregate types exist, and asserts that the aggregate's emission sits inside the `scoped_lock` that advances it.

### Failing-before and passing-after witnesses

Against the original provider the new case named all three defects in a single run:

| Finding | Recorded failure |
|---|---|
| A1 | `CopyItems closed on wrong totals (bytes=32780/131115 files=4/7).` |
| A2 | `MoveItems cumulative totals went backwards at report 1 (49165/1/0 after 0/0/1).` |
| A3 | `DeleteItem closed its governed scope 0 times.` |

With the repair in place the case passes, as do the other eight discovery cases.

One further defect was found by this case during implementation and is worth recording because it was mine, not the original provider's: a first cut computed the aggregate under the lock and emitted it afterwards. Concurrent roots then delivered their snapshots out of order and the case reported `MoveItems cumulative totals went backwards at report 2 (49170/2/0 after 98335/3/0)`. Moving the emission inside the lock fixed it. The host would never have surfaced this: it clamps a regressing cumulative value to a zero delta instead of failing.

### Notes and deviations

- B0 predicted `CopyItems` would report **no events at all**. That was wrong, and the correction matters for anyone re-reading A1: the recursive walker inside the per-path copy does report its children, so a bulk Copy did emit a stream and did close it. What it never emitted was a top-level root record, so it closed on totals short by the selected roots themselves — `files=4/7` above. The defect is a wrong denominator, not silence.
- Move discovery describes what the call had to discover, not what it moved. A same-endpoint directory Move is one rename that relocates the subtree without enumerating it, so that root contributes one directory and no descendant files or bytes. The case's first expectations wrongly reused the Copy totals and had to be split into separate Copy, Move and Delete expectations. This is now stated in the provider spec so the next reader does not "fix" it back.
- Delete semantics, recycle routing, identity pinning and the parallel scheduling shapes are untouched; this change alters what is reported, not what is mutated. The host's bulk-transfer rejection at `State.cpp:17343` stays as it was.
- No ABI change, no capability advertisement change, no new setting, no hardcoded UI string.

### Validation gates

Fresh Full on 2026-09-14 from `caad89d8` with a clean working tree: **status passed**, all 20 plan
entries promoted, **2,136 passed / 0 failed / 52 declared skips** in 66 minutes 27.7 seconds, with
flaky, regression, isolation-suspect and unclassified-failure classifications all zero. FileOperations
rose from I21's 186 to 187 passes for the one case this plan adds. `DiscoveryScope_LocalDirectApi`
passed at 125 ms and every I21 discovery case, the alternate-volume control, the wide-tree overlap
witness, the Beeline rename cases and the Curl native safety corpus stayed green. The suite-wide
invariant holds: `FileOps.Discovery.GrowthAfterClose` is 0 across all 376 recorded task completions.
Evidence is archived and validated at
`../../TestRuns/4cb089111a23/FileOps/2026-09-14_181429_i23_direct_api_discovery_fresh_full/README.md`.

Gotchas for the next owner, all paid for during this closeout:

- This repository's working tree is shared with other sessions, and a Fresh Full validates the
  **working tree**, not your commits. One attempt failed `ToolsPesterTests` on two
  `Specs/Plans/Done/` files another session had edited: `Test-RSProtectedDoneHistory` hashes the
  working file with `git hash-object`, Done history has no modification allowlist, and the Pester
  summary names only the test, never the file. Triage a gate failure with `git status` before your
  own diff.
- A commit landing on master mid-run is fatal, not cosmetic. The runner records
  `SOURCE MUTATION: workspace changed during <entry>` and forces the run's exit code to 1
  (`Tools/Run-AllTests.ps1:1770`) whatever the suites did, so the run cannot produce a passing verdict.
- Interrupting a run leaves `.build\artifact-operations\*.contaminated.json` and the next run is
  refused with `BLOCKED ATTESTATION` until `build.ps1 -Rebuild` clears it.
- Run the Full gate with `-AllowAlternateVolumeTestRoot`, as this plan's Commands block specifies.
  Without it `R4A19_DiscoveryIndependentVolumes` fails outright rather than skipping, because the
  authorization is an environment variable the runner sets only for that switch.
- Do not run anything else between the gate and archiving its evidence. A later run's stale-run
  cleanup reclaims `RedSalamander.Perf\runs\<id>`; only the `evidence\runs\<id>` copy survives, and it
  does not carry the aggregate `results.json`, `trace.txt` or perf rows.

### Post-gate closeout step

Admitting this plan's Done path to the protected-history allowlist in
`Tools/Modules/Tooling/SpecInformationArchitecture.psm1` is the only `Tools/` change outside the
repair scope, and it landed ahead of the move in `bdb777f2` together with its paired test, matching
the pattern I21 established.

## Stop and reconciliation conditions

Pause and report if closing the Delete scope changes an existing host-visible Delete total, if the shared `MoveItems` accumulator cannot be made safe under the existing parallel workers without changing their scheduling, or if a consumer outside this repository depends on the current per-root closure shape.
