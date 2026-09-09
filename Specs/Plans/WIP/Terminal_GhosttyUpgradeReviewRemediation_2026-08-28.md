# Terminal Ghostty upgrade review remediation

> **NON-NORMATIVE WORKING PLAN.** The admitted Ghostty pin and the completed
> upgrade plan remain frozen in
> `Specs/Plans/Done/Terminal_LibGhosttyVt_UpgradeAndContinuousReadiness_2026-08-25.md`.
> Durable runtime, testing, performance, and tooling behavior is owned by the
> authoritative specifications named below.

## Status

| Field | Value |
|---|---|
| State | **ACTIVE** |
| Index owner | **I15** |
| Priority | **P1** validation correctness and lifecycle safety; **P2** observability, tooling, and closeout durability |
| Planned at | `388f31695bf1bdf8317c4e55da88972c664b6359` (`388f31695`) |
| Updated | 2026-08-28 |
| Scope | Accepted findings from the post-admission Ghostty VT review; no pin change and no generic lifecycle retirement |
| Drift check | `git diff 388f31695bf1bdf8317c4e55da88972c664b6359 -- Plugins/Terminal Tests/PluginContractTests Common/StringConversion.h Tools/Test-TestRunArchive.ps1 Tools/Modules/Testing Tools/Tests .github/workflows Specs/Terminal Specs/Testing Specs/Core/Core_SharedHelpers.md` |

## Progress checklist

`[x]` means implemented, covered, and verified. `[+]` marks the active package.

- [x] **R0. Establish the accepted remediation boundary**
  - [x] Recheck the review against the admitted headers, runtime implementation,
        tests, tooling, workflows, and specifications.
  - [x] Reject false positives: production render locking, dormant generic-tool
        retention, decoded-PNG rejection, `TakeReady` semantics, and PNG decoder
        type placement.
  - [x] Record focused baseline results and exact affected files.
- [x] **R1. Make Terminal initialization transactional and diagnosable**
  - [x] Roll back every Ghostty resource in reverse order after any post-load
        initialization failure while preserving the documented Diagnostic `Open`
        contract.
  - [x] Refuse a second engine initialization on a live instance.
  - [x] Preserve a stage-specific localized diagnostic after the loader succeeds.
  - [x] Add deterministic failure-boundary tests and prove unload/refcount cleanup.
- [x] **R2. Harden callback and worker/UI boundaries**
  - [x] Replace allocating UTF-16 clipboard validation with canonical in-place
        strict UTF-8 validation; keep conversion on the UI thread.
  - [x] Route output-ready and Kitty-ready wakes through the posted-payload
        registry with generation checks and teardown drain; remove worker-thread
        `InvalidateRect` fallbacks.
  - [x] Exercise the production clipboard callback against the admitted DLL and
        prove one reply/callback per write plus pending-message coalescing.
  - [x] Surface WIC/COM PNG decoder initialization failure once without making
        the text terminal unavailable.
- [+] **R3. Repair VT snapshot evidence correctness**
  - [x] Complete render-state update before reading rows/cells/styles in the VT
        upgrade snapshot helper.
  - [x] Add a styled-cell mutation regression that fails under the old ordering.
  - [x] Validate the admitted historical product/candidate pair with the new
        schema-aware comparator.
  - [ ] Re-run the paired 200-sample corpus and validate corrected archives.
- [x] **R4. Enforce paired evidence and archive schemas in tooling**
  - [x] Extend the thin public archive command with baseline/candidate comparison
        backed by its reusable Testing module; do not add a redundant entrypoint.
  - [x] Recompute nearest-rank p50/p95 from raw samples, bind controls and runtime
        identities, and enforce candidate p95 at or below baseline x1.10.
  - [x] Make archive validation dispatch to the Terminal VT evidence schema and
        parse every nonblank performance JSONL row.
  - [x] Update tool inventory/help and add positive/negative Pester coverage.
- [x] **R5. Preserve failure reports in CI/release**
  - [x] Upload Ghostty status/preflight reports with `always()` before the final
        gate step fails the job.
  - [x] Add workflow-policy tests for ordering and failure preservation.
- [ ] **R6. Merge durable contracts and close out**
  - [x] Update Terminal, performance-validation, validation-evidence, selftest,
        tooling-governance, shared-helper, and TestRuns documentation.
  - [x] Run focused Debug checks, Pester, inventories, archive validation, and
        `git diff --check`.
  - [ ] Run test-enabled Release checks after the independently launched
        `.build/x64/Release/RedSalamander.exe` releases exact output ownership.
  - [ ] Obtain a green final Fresh Full; the 2026-08-28 run failed only its
        repository-wide Commands lane, with broad non-Terminal UI/state failures.
  - [ ] Move this plan to `Specs/Plans/Done/` and remove I15 from the WIP index.

## Ownership and coordination

- I15 owns only the accepted post-admission Ghostty remediation listed above.
- I10 retains ownership of the broader Tools inventory and dormant generic
  TerminalEngine lifecycle. This plan may add/update directly required active
  Ghostty evidence commands but will not delete compatibility entrypoints.
- I4 retains repository-wide test-contract migration. I15 adds focused runtime
  tests and only narrow source-policy guards needed for workflow/tool contracts.
- I5 owns repository-wide performance infrastructure. I15 reuses the existing
  Terminal VT evidence format and adds only its missing paired admission gate.
- The admitted Ghostty commit, tree, Zig version, ABI inventory, and overlay are
  out of scope. Any required pin change is a STOP and requires a new upgrade plan.

## Accepted findings

| ID | Priority | Remediation |
|---|---:|---|
| GVR-01 | P1 | Move VT snapshot traversal after `renderStateEndUpdate`; regenerate evidence. |
| GVR-02 | P1 | Add reverse-order initialization rollback, re-entry refusal, and stage diagnostics. |
| GVR-03 | P1 | Put output/Kitty wakes on the posted-payload registry and remove worker UI calls. |
| GVR-04 | P1 | Add a machine-enforced paired 1.10/200-sample/control-digest comparator. |
| GVR-05 | P2 | Validate clipboard UTF-8 without allocating UTF-16 under `_terminalMutex`. |
| GVR-06 | P2 | Make archive validation honor area schemas and validate every JSONL row. |
| GVR-07 | P2 | Upload status/release reports before failing the workflow gate. |
| GVR-08 | P2 | Add production-callback, teardown/posting, and PNG decoder observability coverage. |
| GVR-09 | P2 | Merge the VT performance/evidence contract into authoritative testing specs. |

## Rejected findings

- Production correctly releases `_terminalMutex` before
  `ghostty_render_state_end_update`; only the test snapshot helper violates the
  admitted render-state ordering contract.
- The generic TerminalEngine commands are classified compatibility/history and
  have named consumers and retention policy. Removing them is owned by I10, not
  this runtime remediation.
- The admitted public API returns decoded RGBA image data; rejecting a public
  `FORMAT_PNG` result remains fail-closed behavior, not a current-pin bug.
- `TakeReady` consumes the requested generation intentionally; retaining a
  mismatched result would blur request ownership.
- `TerminalPngDecodeThreadContext` placement is cosmetic and does not justify
  churn during a safety remediation.

## Verification and done criteria

The plan is complete only when all checklist rows are checked, authoritative
specs describe the lasting contracts, corrected paired evidence passes the new
comparator, tool/spec inventories have zero findings, and a Fresh Full run is
green for the final worktree. If the runtime cannot expose the production
clipboard callback safely in a deterministic selftest, leave that row open and
report the exact missing seam rather than substituting a source-shape assertion.

## Validation evidence (2026-08-28)

| Surface | Result |
|---|---|
| Terminal Debug build | PASS; zero warnings/errors in `.build/logs/msbuild-20260828_100937_644-pid82520-913cf82d.log`. |
| PluginContractTests Debug build | PASS; zero warnings/errors in `.build/logs/msbuild-20260828_103513_600-pid90236-ddd98936.log`. |
| Terminal runtime selftests | PASS, 162/162, including all 24 initialization rollback stages, exact admitted-DLL clipboard callback behavior, typed wake/drain behavior, and PNG decoder initialization. |
| VT upgrade selftests | PASS, 14/14, including the corrected render-state snapshot ordering and styled mutation fixture. |
| DxUiTests | PASS, including strict allocation-free UTF-8 validation coverage. |
| Focused Pester | PASS: Terminal source contracts 34/34; archive contracts 14/14; release workflow policy 22/22; tool inventory 17/17; resource localization 6/6; Ghostty runtime status 15/15. |
| Tool/spec/TestRuns inventories | PASS with zero blocking findings; complete TestRuns inventory validated 1,216 files. |
| Historical admitted VT pair | PASS through `Tools/Test-TestRunArchive.ps1 -BaselineRunPath Specs/TestRuns/4cb089111a23/Terminal/20260826_071812_pair2-product -CandidateRunPath Specs/TestRuns/4cb089111a23/Terminal/20260826_071820_pair2-candidate`; all recomputed p95 values remain below their baseline x1.10 ceilings. |
| Fresh Full Debug x64 | FAILED only Commands: overall 1,476 passed, 475 failed, 53 skipped; every other lane passed, including Tools Pester 707/707, Compare 229/229, File Operations 123/123, DxUiTests, PluginContractTests, viewer suites, monitor latency, and PerformanceTests2. Evidence: `.build/TestSandbox/runs/20260828T083546Z-84164-6c31c47be52d4ef99058e5a343695f4e/artifacts/selftest/last_run`. Commands failures are broad non-Terminal UI/state failures (FolderView, Connection Manager, dialogs, and action launching), so they are not accepted as Terminal remediation regressions. |
| Test-enabled Release x64 | BLOCKED: PID 79948 is an independently launched `.build/x64/Release/RedSalamander.exe`; build preflight correctly refuses to replace owned output and kills nothing. |
| Corrected current paired corpus | BLOCKED: the historical isolated product/candidate fixtures were cleaned after the admitted run, and Release output is independently owned. Do not substitute the historical passing pair for newly generated corrected snapshot evidence. |

## STOP conditions

Stop and request direction if remediation requires changing the admitted pin or
private ABI, weakens clipboard/Kitty security policy, changes the public
Diagnostic `Open` behavior, deletes a compatibility tool owned by I10, or cannot
preserve complete teardown before unloading `ghostty-vt.dll`.
