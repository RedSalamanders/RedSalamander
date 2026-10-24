# Plan 001: Qualify private libghostty and record a real winner

> **Status — Completed 2026-08-03.** Ghostty is qualified and selected as the private Terminal runtime. The decision
> and upgrade boundary are authoritative in `../../Terminal/TerminalEngineDecision.md` and
> `../../Terminal/Terminal_EmbeddedPlugin.md`. Exact x64/ARM64 identities, offline reproduction, notices, and rollback
> inputs are verified; later WIP/checkpoint language below is historical.

## Implemented outcome (2026-08-01)

**COMPLETE for the first-implementation engine gate.** Ghostty commit
`b2d44625908eb56aee299d8511d803bc11ab79fc` with Zig 0.16.0 is the selected
winner under the bounded private-runtime policy. Reproducible native engine
outputs were produced for both architectures:

- x64: `a080beaa964f7a10022ebc0132515bd7399454ed13ce8fbd152da69004638bfa`;
- ARM64: `b32b4854d6d57494576c2579c13582a84ab21ef1efd82c7ffe7c287fca6e2e75`;
- tracked ARM64 shared-only overlay:
  `7cdc35d18bee7aead13ce8e68845ddab5c0a5cf0de873b60b8a23b18b5bf1ff7`.

`Tools/Build-GhosttyTerminalRuntime.ps1` now owns source/tool verification,
the closed export/import checks, identity JSON, license output, and the
repeatable upgrade entry point. Do not rerun this selection plan unless the
source, Zig pin, build overlay, feature set, or mandatory engine requirements
change. Product-release hardening remains in Plan 005.

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. Do not start product implementation. If anything in the STOP
> conditions occurs, stop and report; do not weaken a gate or edit historical
> evidence. When done, update this plan's row in `Specs/Plans/WIP/Terminal_FirstImplementationExecutionIndex_2026-08-01.md`.
>
> **Drift check (run first)**:
> `git diff --stat aac3ed260..HEAD -- Specs/Terminal Specs/Plans/WIP Specs/Plans/Done Tools/TerminalEngine Tools/Tests Tools/Evaluate-TerminalEngines.ps1 Tools/Build-TerminalEngine.ps1 Tools/Restore-TerminalEngineInputs.ps1 Tools/Verify-TerminalEngine.ps1 External LICENSE.txt`
> If any in-scope path changed, compare it with Current state. STOP if a new
> Round 5 universe/evidence collection already exists or if the protected
> Round 4 block no longer has its recorded digest.

## Status

- **Execution state**: COMPLETE for first implementation
- **Priority**: P1
- **Effort**: L (15-20 owner-days, hard cap)
- **Risk**: HIGH
- **Depends on**: none
- **Category**: direction
- **Planned at**: commit `aac3ed260`, 2026-08-01

## Why this matters

The active Round 5 plan preserves a zero-private-runtime-leaf requirement that
contradicts the product architecture already approved in the original Gate 0
plan. It also biases the next proof toward WezTerm even though Ghostty now
offers the closest engine-shaped C surface. This plan corrects the candidate-
neutral policy, runs a bounded native proof, and records either a real winner
or a precise no-winner result before product code is permitted.

## Current state

- `Specs/Terminal/TerminalEngineDecision.md:91` starts the immutable Round 4
  historical block. `:130` records the old `New non-system runtime leaves | 0`
  result. Never edit bytes in that protected block.
- `Specs/Plans/Done/Terminal_EngineSelectionAndDependencyProof_2026-07-22.md:25`
  already permits an engine to be linked into or privately loaded only by
  `Terminal.dll`; `:97` and `:220` explicitly define the locked private
  `Plugins\TerminalRuntime\` package topology.
- `Specs/Plans/WIP/Terminal_EngineSelectionAndDependencyProof_Round5_WezTermSynchronousCallerOwned_2026-07-31.md`
  is unexecuted as of this plan. It carries the obsolete zero-leaf cap and must
  be superseded, not silently repurposed after evidence exists.
- `Specs/Terminal/TerminalEngineCandidateUniverse.v4.json` is the latest
  universe. No `v5` universe, transition, collection, or decision artifact
  exists at planned-at commit.
- `TerminalEngineCandidateAssessment.v1.json`,
  `TerminalEngineCandidateReview.v1.json`, `TerminalEngineCandidateSet.v1.json`,
  `TerminalEngineRequirements.v1.json`, `TerminalEngineMeasurementProtocol.v1.json`,
  and `TerminalEngineScoreAnchors.v1.json` are Round 4-bound frozen inputs or
  records. Create new versioned files; never edit them for the new round.
- `Tools/Build-TerminalEngine.ps1`,
  `Tools/Restore-TerminalEngineInputs.ps1`, and
  `Tools/Verify-TerminalEngine.ps1` already implement lock-bound restore,
  clean/offline build, output identity, and smoke hooks.
- `Tools/TerminalEngine/TerminalEngineGate0Harness.cpp` and
  `Tools/TerminalEngine/TerminalEngineGate0Probe.cpp` already provide the
  neutral native harness/probe. Extend contracts and adapters; do not replace
  them with a Ghostty-only test program.
- The current official Ghostty checkout inspected for this plan was commit
  `b2d44625908eb56aee299d8511d803bc11ab79fc` on 2026-08-01. Its
  `build.zig.zon` identifies `1.3.2-dev` and minimum Zig `0.16.0`. This is a
  candidate pin, not yet an accepted lock.
- Current Ghostty headers establish these load-bearing facts:
  - VT effects execute synchronously during `ghostty_terminal_vt_write`; they
    must not block or re-enter the terminal.
  - render update is caller-owned/two-phase; only begin requires the terminal
    lock, and dirty rows/global state must be cleared deliberately;
  - terminal/Kitty handles and pixel pointers are borrowed and invalidated by a
    mutating call;
  - PNG and log callbacks are process-global;
  - the custom allocator covers Ghostty allocations and supplies allocation
    sizes/alignment, enabling an aggregate hard cap;
  - the C API is usable on Windows but upstream still describes signatures as
    evolving and does not publish a versioned libghostty release.
- Ghostling is not Windows proof. Its sample application currently describes
  Windows support as unimplemented/untested and pins a different Ghostty/Zig
  combination.

## Policy to freeze before measuring candidates

The replacement Gate 0 round must use the following candidate-neutral rules.
Do not change the caps after seeing a candidate result.

| Policy | Required value |
|---|---|
| Engine product boundary | linked into or dynamically loaded only by `Plugins\Terminal.dll` |
| Dynamic runtime location | exact package-relative `Plugins\TerminalRuntime\<leaf>` only |
| Private non-system runtime leaves | maximum 4 per architecture/configuration |
| Private runtime payload | maximum 67,108,864 uncompressed bytes per architecture/configuration |
| Global/system installation | forbidden: no PATH, registry, app-root engine DLL, service, COM registration, or installer mutation |
| Toolchain families added | maximum 1, fully pinned and restored offline |
| Runtime identity | SHA-256, size, PE machine, imports, exports, source pin/tree, toolchain, flags, notices |
| Loading | exact absolute path; no basename/PATH fallback |
| Runtime replacement | atomic lock/adapter/runtime/notices update; restart or complete quiet point required |
| Memory proof | hard aggregate Ghostty allocator cap plus separate hard plugin queue/cache/image/GPU budgets |
| Kitty media v1 | direct transmission only; file, temp-file, and shared-memory media disabled |
| Owner effort | maximum 20 owner-days for this proof; record daily by category |
| Native lanes | x64 Debug, x64 Release, x64 ASan Debug, ARM64 Debug, ARM64 Release |

Private leaf count and bytes below their caps remain scored maintenance costs;
they are not automatic disqualifiers. System Windows DLLs do not count as
private leaves, but every non-system transitive dependency does.

## Commands you will need

Run from repository root in PowerShell 7.

| Purpose | Command | Expected on success |
|---|---|---|
| Confirm branch/clean baseline | `git status --short --branch; git rev-parse HEAD` | expected branch; no unreviewed changes; SHA shown |
| List v5 artifacts | `rg --files Specs/Terminal Specs/TestRuns | rg 'TerminalEngine.*v5|TerminalEngine.*Round5'` | no evidence/universe result before Step 1 |
| Focused Pester | `Invoke-Pester -Path Tools/Tests/TerminalEngineCandidateUniverseTransition.Tests.ps1,Tools/Tests/TerminalEngineCandidateAssessment.Tests.ps1,Tools/Tests/TerminalEngineEvaluator.Tests.ps1,Tools/Tests/TerminalEngineGate0Harness.Tests.ps1,Tools/Tests/TerminalEngineGate0Probe.Tests.ps1,Tools/Tests/TerminalDependencyLifecycle.Tests.ps1 -Output Detailed` | all tests pass |
| Restore exact inputs | `.\Tools\Restore-TerminalEngineInputs.ps1 -LockFile External\terminal-engine.lock.json` | every locked input reports restored/verified |
| Offline input check | `.\Tools\Restore-TerminalEngineInputs.ps1 -LockFile External\terminal-engine.lock.json -Offline` | every input reports offline verification |
| Build one lane | `.\Tools\Build-TerminalEngine.ps1 -Platform x64 -Configuration Release -LockFile External\terminal-engine.lock.json -Clean` | exit 0; build manifest emitted |
| Verify one lane | `.\Tools\Verify-TerminalEngine.ps1 -Platform x64 -Configuration Release -LockFile External\terminal-engine.lock.json -RunSmoke` | exit 0; identity and smoke manifest emitted |
| Full repository gate | `.\Tools\Run-AllTests.ps1 -Suite Full` | exit 0; no failed cases |

Offline build/verification requires the repository's exclusive disposable
runner. Use `-Offline -ExclusiveDisposableRunner`; never disable isolation just
to obtain a pass.

## Suggested executor toolkit

- Read `.github/skills/cpp-build/SKILL.md`,
  `.github/skills/perf-validation/SKILL.md`,
  `.github/skills/plugin-callbacks/SKILL.md`, and
  `.github/skills/compiler-warnings/SKILL.md` before changing their domains.
- Use the official Ghostty repository and documentation only for upstream API,
  build, and release claims. Archive exact pages/commits used by the evidence.
- Start from the official [Ghostty repository](https://github.com/ghostty-org/ghostty),
  whose current roadmap says `libghostty-vt` supports Windows but API signatures
  remain in flux and no library version is tagged. Cross-check the official
  [Ghostty 1.3 libghostty release note](https://ghostty.org/docs/install/release-notes/1-3-0)
  rather than equating a desktop-app release with a library release.
- Treat [Ghostling](https://github.com/ghostty-org/ghostling) only as a C API
  sample: its own README says Windows is not tested/implemented in Ghostling and
  that consumers still own windowing, rendering, tabs, sessions, and settings.
- Use `dumpbin /headers /imports /exports` or the existing PE identity helper;
  do not infer runtime closure from build files alone.

## Scope

**In scope**:

- `Specs/Terminal/TerminalEngineDecision.md`
- `Specs/Terminal/TerminalEngineCandidateUniverse.v5.json` (new)
- `Specs/Terminal/TerminalEngineCandidateUniverseTransition.v5.json` (new)
- `Specs/Terminal/TerminalEngineCandidateUniverseTransition.v5.schema.json` (new)
- `Specs/Terminal/TerminalEngineCandidateAssessment.v2.json` (new)
- `Specs/Terminal/TerminalEngineCandidateAssessment.v2.schema.json` (new if the
  existing schema is frozen to v1)
- `Specs/Terminal/TerminalEngineCandidateReview.v2.json` (new)
- `Specs/Terminal/TerminalEngineCandidateReview.v2.schema.json` (new if needed)
- `Specs/Terminal/TerminalEngineCandidateSet.v2.json` (new)
- `Specs/Terminal/TerminalEngineCandidateSet.v2.schema.json` (new if needed)
- `Specs/Terminal/TerminalEngineRequirements.v2.json` (new)
- `Specs/Terminal/TerminalEngineMeasurementProtocol.v2.json` (new)
- `Specs/Terminal/TerminalEngineScoreAnchors.v2.json` (new)
- `Specs/Terminal/TerminalEngineLock.schema.json` only through an additive
  versioned branch that continues to validate the frozen v1 corpus
- `Specs/Plans/WIP/README.md`
- `Specs/Plans/WIP/Terminal_EngineSelectionAndDependencyProof_Round5_WezTermSynchronousCallerOwned_2026-07-31.md`
- a new replacement Round 5 WIP plan under `Specs/Plans/WIP/`
- the superseded plan moved under `Specs/Plans/Done/`
- `Tools/TerminalEngine/TerminalEngineCandidateAssessment.psm1`
- `Tools/TerminalEngine/TerminalEngineLifecycle.psm1`
- `Tools/TerminalEngine/TerminalEngineGate0Adapter.h`
- `Tools/TerminalEngine/TerminalEngineGate0AdapterContract.md`
- `Tools/TerminalEngine/TerminalEngineGate0Harness.cpp`
- `Tools/TerminalEngine/TerminalEngineGate0Probe.cpp`
- a new Ghostty Gate 0 adapter project under `Tools/TerminalEngine/`
- `Tools/Build-TerminalEngine.ps1`
- `Tools/Restore-TerminalEngineInputs.ps1`
- `Tools/Verify-TerminalEngine.ps1`
- `Tools/Evaluate-TerminalEngines.ps1`
- the matching `Tools/Tests/TerminalEngine*.Tests.ps1` tests
- `External/terminal-engine.lock.json` only after a winner exists
- `LICENSE.txt` only after a winner exists
- new immutable evidence under `Specs/TestRuns/<commit>/Terminal/`

**Out of scope**:

- any product source under `RedSalamander/`, `Common/PlugInterfaces/`, or
  `Plugins/Terminal/`;
- editing the protected Round 4 decision bytes or rewriting old evidence;
- editing any existing v1 candidate assessment/review/set/requirements/
  measurement/anchor record or a v1-v4 universe/transition record;
- declaring Ghostty the winner based on README claims or Ghostling;
- creating a private fork, adding a runtime path setting, or installing a
  global Zig/runtime dependency;
- weakening a required native lane, fixture, resource +1 test, or unload test.

## Git workflow

- Continue on `codex/terminal-engine-gate0` unless the operator directs a new
  branch. Do not push or open a PR without instruction.
- Use one commit for the candidate-neutral policy/tool changes and a later
  commit for immutable evidence/decision. Do not mix product implementation.
- Preserve the repository's imperative commit style, for example
  `Harden Terminal Gate 0 Round 5 implementation plan`.

## Steps

### Step 1: Prove the Round 5 replacement is safe

1. Confirm there is no `v5` candidate universe, transition, evidence-set,
   collection, or finalization record and no `External/terminal-engine.lock.json`.
2. Re-run the Round 4 closeout verifier and record its protected digest.
3. If Round 5 evidence exists, STOP. A versioned Round 6 policy transition is
   required; do not relabel existing evidence.
4. Move the unexecuted WezTerm-only WIP plan to `Specs/Plans/Done/` with a
   prominent `Superseded before execution` status and reason. Do not alter its
   historical content beyond status/front-matter and links.
5. Create a replacement Round 5 WIP plan whose policy is exactly the table
   above and update `Specs/Plans/WIP/README.md` routing.

**Verify**:

```powershell
.\Tools\Test-TerminalEngineRound4Closeout.ps1
rg -n "Superseded before execution|privateRuntimeLeafCap|privateRuntimeBytesCap" Specs\Plans\WIP Specs\Plans\Done
```

Expected: Round 4 verifier exits 0; exactly one active replacement Round 5 WIP
contains the new caps; the old plan is under Done and marked superseded.

### Step 2: Encode a versioned, candidate-neutral policy transition

1. Add `TerminalEngineCandidateUniverseTransition.v5.schema.json` by versioning
   the v4 transition contract. Do not edit v3/v4 schemas or records.
2. Add fields for `privateRuntimeLeafCap`, `privateRuntimeBytesCap`,
   `privateRuntimeRoot`, `globalInstallForbidden`, and
   `runtimeCostScoredBelowCap` with the exact frozen values above.
3. Add `TerminalEngineCandidateUniverseTransition.v5.json` explaining that the
   old zero-leaf constraint contradicted the already-approved private-runtime
   architecture. This is a requirements correction, not a retrospective
   reinterpretation of Round 4.
4. Add `TerminalEngineCandidateUniverse.v5.json`. Include Ghostty at exact
   commit `b2d44625908eb56aee299d8511d803bc11ab79fc`; retain all candidates that
   satisfy the new neutral discovery rules. Do not create a Ghostty-only
   universe.
5. Create v2 requirements, measurement, anchors, assessment, review, and
   candidate-set files/schemas for this round. Keep every prior version
   byte-for-byte unchanged and teach the evaluator to select the schema by the
   record's explicit format version.
6. Make runtime leaf count and bytes measured/scored fields with hard rejection
   only when the frozen caps are exceeded.
7. Extend transition/evaluator/assessment Pester cases for 0, cap, cap+1,
   case-insensitive path collision, rooted/escaping private path, and runtime
   outside `Plugins/TerminalRuntime`.

**Verify**:

```powershell
Invoke-Pester -Path Tools\Tests\TerminalEngineCandidateUniverseTransition.Tests.ps1,Tools\Tests\TerminalEngineCandidateAssessment.Tests.ps1,Tools\Tests\TerminalEngineEvaluator.Tests.ps1 -Output Detailed
```

Expected: all tests pass; negative cap+1 and escaping-path fixtures are rejected
with stable categories; old v3/v4 fixtures still pass unchanged.

### Step 3: Freeze exact Ghostty source, toolchain, and build inputs

1. Review the pinned Ghostty commit's `LICENSE`, `build.zig.zon`, `build.zig`,
   `src/build/`, `include/ghostty/`, and CMake wrapper. Record source repository,
   commit, source-tree/archive SHA-256, version text, and every lazy dependency
   actually fetched by the lib-vt build.
2. Pin the official Zig 0.16.0 Windows x64 archive by URL, byte length, and
   SHA-256. If native ARM64 Zig is required, pin it separately. Never use an
   unversioned PATH `zig`.
3. Express the build entirely through `External/terminal-engine.lock.json` and
   existing restore/build tooling. The lock must bind target, optimization,
   CRT, SIMD choice, output names, build commands, environment allowlist, and
   expected outputs for all five required lanes.
4. Test both upstream shared-library variants needed to answer closure risk:
   SIMD enabled and `-Dsimd=false`. Select one before final evidence based on
   native fixture/performance results and import closure, not DLL count alone.
5. Do not commit the lock as a production lock until Step 8 declares a winner.
   Use a clearly named candidate lock under the Gate 0 evidence work root while
   iterating; finalization atomically promotes it.

**Verify**:

```powershell
.\Tools\Restore-TerminalEngineInputs.ps1 -LockFile <candidate-lock>
.\Tools\Restore-TerminalEngineInputs.ps1 -LockFile <candidate-lock> -Offline
```

Expected: online restore and offline verification both exit 0; every source,
toolchain, and dependency record matches exact hash and length; no PATH tool is
used.

### Step 4: Build and inspect the private Windows closure first

1. Build x64 Debug and Release shared `ghostty-vt.dll` from the candidate lock.
2. Enumerate exact PE machine, imports, exports, file size, SHA-256, PDB policy,
   and all non-system transitive leaves. Resolve forwarded imports as the
   existing PE helper requires.
3. Repeat natively for ARM64 Debug and Release. Do not accept cross-compilation
   alone as native runtime proof.
4. Build/consume the x64 ASan Debug boundary according to the locked lane. It
   may reuse a Debug engine only if the lock explicitly records and the probe
   proves that mapping; never label an ordinary Debug consumer as ASan.
5. Fail if any artifact loads from app root, PATH, a build-tool directory, or
   an undeclared adjacent file.

**Verify**:

```powershell
$lanes = @(
  @{ Platform='x64'; Configuration='Debug' },
  @{ Platform='x64'; Configuration='Release' },
  @{ Platform='x64'; Configuration='ASan Debug' },
  @{ Platform='ARM64'; Configuration='Debug' },
  @{ Platform='ARM64'; Configuration='Release' }
)
foreach ($lane in $lanes) {
  .\Tools\Build-TerminalEngine.ps1 -Platform $lane.Platform -Configuration $lane.Configuration -LockFile <candidate-lock> -Clean
  if ($LASTEXITCODE -ne 0) { throw "build failed: $($lane.Platform) $($lane.Configuration)" }
  .\Tools\Verify-TerminalEngine.ps1 -Platform $lane.Platform -Configuration $lane.Configuration -LockFile <candidate-lock> -RunSmoke
  if ($LASTEXITCODE -ne 0) { throw "verify failed: $($lane.Platform) $($lane.Configuration)" }
}
```

Expected: five lane manifests pass; primary machine matches lane; private
closure is at most four leaves and 64 MiB; no undeclared non-system import.

### Step 5: Implement the neutral Ghostty Gate 0 adapter

Create a Gate 0 adapter under `Tools/TerminalEngine/` that satisfies the
existing adapter contract while using Ghostty only behind its `.cpp` file.
It must prove the exact integration strategy planned for `Terminal.dll`:

1. Resolve all required symbols through a closed function table. Do not link
   the harness/probe or any RedSalamander host target against the import library.
2. Register one process-global log callback and one WIC PNG decoder callback.
   The decoder initializes COM MTA on worker threads, reads dimensions before
   decode, checks `width * height * 4` with overflow-safe arithmetic, enforces
   hard width/height/pixel/decoded-byte caps, and allocates the returned bytes
   with Ghostty's supplied allocator.
3. Supply a custom Ghostty allocator that atomically reserves before every
   allocation/remap, rejects cap+1 without transient overshoot, and records
   current/peak bytes and allocation failures. `std::bad_alloc` is fatal.
4. Serialize terminal mutation. VT effect callbacks may only copy bounded data
   into a plugin-owned queue. They must not block, write the PTY, call a host
   callback, or re-enter Ghostty while `vt_write` is active.
5. During snapshot, hold the session lock for render begin and Kitty metadata/
   changed-pixel copying only. Copy all borrowed data needed after unlock. End
   the render update outside the lock. Never store borrowed terminal, image,
   placement, row, or pixel addresses.
6. Clear global callbacks/userdata only after all terminals, render states,
   images, placement iterators, and worker callbacks are destroyed. Prove a
   second clean load/use/unload cycle in the same harness process.
7. Disable Kitty file, temp-file, and shared-memory transmission. Exercise
   direct PNG and raw formats only in v1.

**Verify**:

```powershell
Invoke-Pester -Path Tools\Tests\TerminalEngineGate0Harness.Tests.ps1,Tools\Tests\TerminalEngineGate0Probe.Tests.ps1,Tools\Tests\TerminalEnginePeIdentity.Tests.ps1,Tools\Tests\TerminalEngineNetworkIsolation.Tests.ps1 -Output Detailed
```

Expected: all structural and negative tests pass; adapter contains no static
Ghostty import in the harness/probe; cap+1, malformed PNG, stale pointer,
callback-after-unload, reentrant effect, and second-load tests have deterministic
passing results.

### Step 6: Run the complete functional and adversarial corpus

For every required lane, collect the versioned fixture corpus through the
existing neutral driver. Required coverage includes:

- UTF-8 split sequences, invalid bytes, combining marks, wide/ambiguous cells,
  emoji/color glyph metadata, OSC 8, alternate screen, scroll regions, tabs,
  SGR, cursor shapes, resize/reflow, and large scrollback;
- Kitty raw RGB/RGBA/gray/gray-alpha, PNG, zlib, cropping, scaling, placement,
  delete, animation if required by the existing contract, negative/nonnegative
  z-order, malformed/truncated/oversized payloads, and every disabled medium;
- effect queue ordering, bounded overflow behavior, terminal response delivery
  after unlock, and no callback reentrancy;
- allocator exactly-at-cap and cap+1, image storage cap, snapshot/cache/GPU
  budget boundaries, checked arithmetic, and recovery after rejected input;
- repeated resize/feed/render/close races, 1,000 create/destroy cycles, process
  exit, and load/use/unload/load/use/unload;
- native ARM64, x64 ASan, WARP/no-hardware rendering seams where applicable.

Record performance counters even though this is a candidate gate: feed
throughput, snapshot p50/p95/max, lock hold p50/p95/max, peak allocator bytes,
image decode/upload, teardown time, DLL/package bytes, and owner-days.

**Verify**:

```powershell
.\Tools\Evaluate-TerminalEngines.ps1 -Mode Collect -Candidate ghostty -Platform x64 -Configuration Debug -EvidenceRoot <evidence-root>
.\Tools\Evaluate-TerminalEngines.ps1 -Mode Collect -Candidate ghostty -Platform x64 -Configuration Release -EvidenceRoot <evidence-root>
.\Tools\Evaluate-TerminalEngines.ps1 -Mode Collect -Candidate ghostty -Platform x64 -Configuration 'ASan Debug' -EvidenceRoot <evidence-root>
.\Tools\Evaluate-TerminalEngines.ps1 -Mode Collect -Candidate ghostty -Platform ARM64 -Configuration Debug -EvidenceRoot <evidence-root>
.\Tools\Evaluate-TerminalEngines.ps1 -Mode Collect -Candidate ghostty -Platform ARM64 -Configuration Release -EvidenceRoot <evidence-root>
```

Expected: each collection exits 0 and emits a schema-valid manifest bound to
the same candidate pin, lock, fixture corpus, adapter, harness, and source SHA.

### Step 7: Complete maintenance, legal, and upgrade evidence

1. Generate notices from the actual lib-vt build closure, not the whole Ghostty
   repository and not a marketing statement. Account for every shipped source,
   generated file, toolchain redistribution, and runtime leaf.
2. Define production discovery as the official Ghostty repository plus an exact
   approved commit/ref selection rule. Because libghostty is not independently
   tagged, never interpret `latest` as a library version.
3. Run `Tools/Prepare-TerminalEngineUpgrade.ps1` against the current candidate
   and one later/alternate commit without changing the tracked candidate lock.
   Evidence must diff headers, exports, ABI layouts, build info, imports,
   fixtures, performance, licenses, toolchain, and output hashes.
4. Record monthly observation, pre-release, and advisory checks. Missing owner,
   redirect, ambiguous ref, prerelease-only head, or advisory uncertainty is a
   blocked observation, not an automatic upgrade.
5. Demonstrate rollback by switching the isolated work root between two exact
   lock/adapter/runtime/notices sets. Mixed versions must fail closed.

**Verify**:

```powershell
Invoke-Pester -Path Tools\Tests\TerminalEngineNotices.Tests.ps1,Tools\Tests\TerminalDependencyLifecycle.Tests.ps1,Tools\Tests\TerminalEngineArchive.Tests.ps1,Tools\Tests\TerminalEngineCandidatePatch.Tests.ps1,Tools\Tests\TerminalEngineSourceReview.Tests.ps1 -Output Detailed
```

Expected: all tests pass; actual closure has complete provenance/notices; the
upgrade candidate remains isolated; mixed-version and stale-lock fixtures fail.

### Step 8: Finalize the decision without advocacy

1. Finalize only from the five immutable collection manifests using
   `-RequireAllEvidence`. Exit 0 means a winner. Exit 20 is a complete no-winner
   result and must remain a valid outcome.
2. A Ghostty winner requires every hard gate, not merely the highest score.
   Record the selected SIMD choice, dynamic private closure, aggregate allocator
   requirement adjustment, owner-days, known limitations, and upstream status.
3. If Ghostty wins, atomically add `External/terminal-engine.lock.json`, the
   exact notices update, selected adapter identity, and current summary above
   the protected historical block in `TerminalEngineDecision.md`. Do not edit
   the Round 4 block.
4. If no candidate wins, do not create a production lock. Add the measured
   decision matrix showing which single requirement changes would make each
   candidate pass and their security/product costs.
5. Archive all evidence under the exact tested repository commit and update the
   replacement Round 5 plan status. A completed WIP plan moves to Done.

**Verify**:

```powershell
.\Tools\Evaluate-TerminalEngines.ps1 -Mode Finalize -EvidenceManifest <x64-debug>,<x64-release>,<x64-asan>,<arm64-debug>,<arm64-release> -EvidenceRoot <evidence-root> -RequireAllEvidence
$decisionExit = $LASTEXITCODE
if ($decisionExit -notin 0,20) { throw "unexpected finalization exit $decisionExit" }
.\Tools\Test-TerminalEngineRound4Closeout.ps1
.\Tools\Run-AllTests.ps1 -Suite Full
```

Expected: finalization exits 0 (winner) or 20 (complete no-winner); Round 4
digest remains unchanged; full suite exits 0. Plans 002-005 may start only for
exit 0 with `ghostty` recorded as winner.

## Test plan

Add or extend tests for:

- v5 transition schema/policy, old-history immutability, and no post-measurement
  cap change;
- private leaf/path/byte rules at 0, cap, and cap+1;
- exact source/toolchain/offline restoration and command/environment binding;
- x64/ARM64 PE identity, imports, exports, and undeclared closure;
- dynamic function-table completeness, build-info mismatch, wrong machine,
  corrupt/missing runtime, and no search-path fallback;
- process-global callback registration/clear and same-process reload;
- aggregate allocator cap/current/peak/failure, image caps, decoder preflight,
  and malformed input recovery;
- effects callback non-reentrancy/queue bounds and post-unlock delivery;
- borrowed render/Kitty data lifetime, generation changes, dirty clearing, and
  mutation races;
- exact required fixture corpus across all five native lanes;
- license/notice completeness, isolated upgrade diff, atomic rollback, and
  mixed-set rejection.

Use existing `TerminalEngine*.Tests.ps1` files as structural patterns. New
native C++ cases belong in the neutral harness/probe projects, not in a
Ghostty-branded product target.

## Done criteria

- [ ] Protected Round 4 verifier still exits 0 and its bytes are unchanged.
- [ ] The unexecuted WezTerm-only Round 5 plan is marked superseded under Done.
- [ ] A candidate-neutral v5 transition/universe freezes the exact private
      runtime policy before collection.
- [ ] Exact Ghostty source, Zig, build inputs, commands, outputs, imports,
      exports, hashes, sizes, and notices are reproducible offline.
- [ ] All five required lanes have native, immutable, schema-valid evidence.
- [ ] Aggregate allocation, PNG, Kitty, effects, snapshot, race, and unload
      adversarial tests pass.
- [ ] Gate 0 effort is at most 20 owner-days with category-level records.
- [ ] Finalization exits 0 with a recorded Ghostty winner, or exits 20 with an
      explicit no-winner matrix and no production lock.
- [ ] `External/terminal-engine.lock.json` exists only for a real winner.
- [ ] `.\Tools\Run-AllTests.ps1 -Suite Full` exits 0.
- [ ] No product terminal code was added.
- [ ] WIP/Done indexes and `Specs/Plans/WIP/Terminal_FirstImplementationExecutionIndex_2026-08-01.md` are updated.

## STOP conditions

Stop and report; do not improvise if:

- any Round 5 collection/finalization evidence already exists at Step 1;
- the protected Round 4 closeout digest changes;
- Ghostty cannot produce a reproducible Windows DLL for one required native
  lane, including ARM64;
- the closure exceeds four leaves/64 MiB or requires global installation;
- the proof needs more than one new toolchain family or 20 owner-days;
- an essential Ghostty API cannot be represented by a closed runtime function
  table or changes during the locked proof;
- callbacks cannot be unregistered before DLL unload;
- borrowed data must remain alive after unlock/mutation;
- aggregate allocator cap cannot reserve-before-allocate and reject cap+1;
- WIC/decoder limits cannot be checked before unbounded decoded allocation;
- passing requires enabling Kitty file/shared-memory media in v1;
- any required native lane, ASan boundary, fixture, race, or unload test is
  unavailable or fails twice after a reasonable correction;
- actual closure provenance/license cannot be established;
- selecting Ghostty requires a long-lived private fork. A single small,
  upstreamable patch is reported with cost and maintenance delta; it is not
  silently adopted.

## Maintenance notes

- The raw SHA-256 in the lock binds the final shipped DLL bytes. The current
  release workflow signs the MSIX package, not individual DLLs. If individual
  PE signing is later added, sign `ghostty-vt.dll` before generating the
  runtime manifest/`Terminal.dll`, or move to an authenticated detached
  manifest; otherwise signing invalidates the digest.
- Treat Ghostty `ghostty_build_info` as diagnostics, not the sole ABI identity.
  Keep exact file hash, export set, header/layout fingerprint, and exercise
  probe.
- Monthly observation is informational. An upgrade is a reviewed atomic change
  only after all five lanes and product regression/performance evidence pass.
- Do not erase no-winner evidence. It is the baseline for any later requirement
  change and prevents repeating failed experiments.
