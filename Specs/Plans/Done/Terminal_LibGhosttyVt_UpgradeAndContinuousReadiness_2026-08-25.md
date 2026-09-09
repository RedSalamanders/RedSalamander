# Ghostty lib-vt Upgrade and Continuous Readiness

> **NON-NORMATIVE EXECUTION PLAN.** Durable product, build, security,
> performance, packaging, and maintenance contracts must be merged into the
> authoritative specifications named below before this plan is closed.
>
> **Executor instructions:** Read this plan completely before changing code.
> Execute the phases in order. Run every verification and retain its result.
> Do not publish a candidate into the product runtime path until every admission
> gate in Phase 10 passes. If a STOP condition occurs, stop and report the exact
> command, output, and affected files; do not improvise.
>
> **Drift check — run first:**
>
> ~~~powershell
> git diff --stat e7475d3aaacc6dda0eeab58cd486c65d4074be4c..HEAD -- External/TerminalEngine Tools/Build-GhosttyTerminalRuntime.ps1 Tools/Get-GhosttyTerminalRuntimeStatus.ps1 Tools/TerminalEngine Tools/Modules/Build/BuildEvidence.psm1 Tools/Modules/Packaging/PortablePackageSmoke.psm1 Tools/Tests Plugins/Terminal Tests/PluginContractTests Specs/Terminal Specs/Testing Specs/Installer .github/workflows/build-reusable.yml .github/workflows/release.yml .github/workflows/ghostty-runtime-status.yml .github/workflows/ghostty-runtime-native-qualification.yml .github/workflows/ghostty-runtime-product-qualification.yml RuntimeDependencies.props
> ~~~
>
> If an in-scope path changed, compare the live implementation with the
> Current-state snapshot. Reconcile harmless drift explicitly in this file.
> Treat changes to the pin, patch, loader-required exports, clipboard policy,
> Kitty image lifecycle, runtime identity, or package-smoke path as a STOP.

## Live execution checklist — authoritative WIP-to-Done gate

This is the single progress checklist for the upgrade. The executor must update
it as development proceeds; do not wait until the end of the upgrade and do not
maintain a competing checklist in a chat, pull request, issue, or scratch file.

### Checklist update protocol

1. At the start of every implementation session, read this checklist and resume
   at the first unchecked item whose prerequisites are complete.
2. Change State below from READY FOR EXECUTION to IN PROGRESS when the first
   implementation item starts. Update Current progress in the same change.
3. Check an item only after its named verification has run and produced the
   expected result. Code presence, compilation alone, or an executor statement
   is not evidence.
4. Every checked leaf item must have a matching Evidence log row containing the
   date, implementation commit or working-tree identifier, exact command/result,
   and durable run/receipt/report path when one exists.
5. Check a bold phase item only after every nested item is checked. Conditional
   work is never silently skipped: check it with an Evidence log disposition
   such as not applicable — premise disproved, followed by the proving command.
6. If a later change or drift invalidates earlier evidence, immediately uncheck
   the affected item and its phase item, record the invalidation, and rerun it.
7. On failure, leave the item unchecked and add
   BLOCKED <date>: <short reason> beneath it. Do not weaken or delete the item.
8. Criteria may change only through a reviewed amendment to this plan that
   explains why. Never edit the checklist merely to match an implementation.
9. At the end of every implementation session, update Current progress, this
   checklist, the Evidence log, and the next unchecked item before handing off.
10. The plan remains in Specs/Plans/WIP while any box is unchecked, any BLOCKED
    marker remains, any required evidence is missing, or Fresh Full/rollback/
    normative closeout is incomplete.
11. P0 through P12 each require their own coherent implementation commit. Do not
    combine two phases in one commit and do not begin the next phase with
    uncommitted work from the previous phase.
12. After a phase commit succeeds, capture its full SHA with
    git rev-parse HEAD. Make a plan-only checkpoint commit that records that SHA
    in the Evidence log, checks the phase-commit leaf and parent phase box, and
    updates Current progress before starting the next phase.
13. Additional focused commits inside a phase are allowed when needed for safe
    review or a verified repair, but the required phase commit must remain a
    coherent boundary containing that phase's final code, tests, specs, and
    bounded evidence.
14. Never amend or rewrite a completed phase commit after the next phase starts.
    If later work invalidates it, uncheck the affected items, record the
    invalidation, and add a new corrective commit in the active phase.

### Current progress

| Field | Value |
|---|---|
| Updated | 2026-08-27 |
| State | DONE |
| Current phase | DONE — P12 implementation committed and plan archived |
| Last verified item | P12.8 — P12 SHA recorded, plan moved to Done, I13 removed from the WIP index, and final spec inventory passed |
| Next unchecked item | None. |
| Blockers | None. |
| Latest evidence | Required P12 phase commit `8f919e8fbba904e7502c22f3e830ff117abb5c1a` has exact subject `Close Ghostty runtime upgrade`. It retains the green isolated rollback, admitted/native qualification, fresh-process Locked repair, zero-finding inventories/archive/diff gates, final receipt validation, and all 18 immutable Fresh promotions. This separate checkpoint moved the fully checked plan to Done, removed I13 from the active index, registered this exact new Done path in the protected-history allowlist, and passed final spec inventory. |
| Last phase commit | 8f919e8fbba904e7502c22f3e830ff117abb5c1a — Close Ghostty runtime upgrade |

### Required phase commits

The Subject column is mandatory. The SHA is filled in after each commit by the
plan-only checkpoint described above.

| Phase | Required commit subject | Full commit SHA |
|---|---|---|
| Initial plan / C0 preparation | Plan Ghostty lib-vt upgrade | e0cc093b27a9ef4d63f7f03a64ec24707731c610 |
| P0 | Establish Ghostty upgrade baseline | 95f3d5a02a6d519751166e2e7c6239deaede2dba |
| P1 | Add Ghostty runtime product lock | f830294e4ab6fd997937e03ab036d20a75920376 |
| P2 | Refactor Ghostty runtime lifecycle | 1e135cbbe2323722cbaa1e50c72fdf123e455838 |
| P3 | Add Ghostty update observation | 43a6bf6214126141b0b2f86f4968b74771d64d07 |
| P4 | Freeze Ghostty runtime candidate | 2347ddbb5835bec235e9de5ce0f37d262a5e0934 |
| P5 | Rederive Ghostty Windows runtime overlay | 551a0a6ef54e1a9e41430f33fced6427f146cfa3 |
| P6 | Qualify isolated Ghostty candidate | 90e277a6e72da1ddab4726b41b3f1488afbde2e6 |
| P7 | Adapt Terminal to Ghostty mode ABI | 39f3ed0888912531a6ad7d621896c0654c5fe620 |
| P8 | Adapt Ghostty clipboard security contract | 3e91f6e8bb106cba20cbc8e99e3d2bf660a7f16d |
| P9 | Handle pending Ghostty Kitty images | 655e45ac2a0f6fc01f51233a1b72fb0ef3c3f615 |
| P10 | Admit reviewed Ghostty runtime candidate | 0ff71930e1edcbd548677ceacea97f959820593f |
| P11 | Qualify Ghostty shipping boundaries | 05cc5ee506cee2c17bfca4e432a8dba3d44adff2 |
| P12 | Close Ghostty runtime upgrade | 8f919e8fbba904e7502c22f3e830ff117abb5c1a |

### C0 — Control, drift, and ownership

- [x] **C0 — Control, drift, and ownership complete**
  - [x] C0.1 Run the planned-at drift command and reconcile every in-scope
        change against Current-state snapshot.
  - [x] C0.2 Record HEAD, worktree status, branch, Windows build, CPU,
        architecture, compiler, SDK, Zig, and vcpkg identity without secrets or
        absolute user-profile paths.
  - [x] C0.3 Confirm I10 has no conflicting Tools-governance change and I12 has
        no concurrent overlapping Plugins/Terminal edit.
  - [x] C0.4 Confirm the allowed Scope and STOP conditions remain sufficient;
        amend and review this plan before touching any additional production
        path.

### P0 — Current-pin baseline and package premise

- [x] **P0 — Current-pin baseline and package premise complete**
  - [x] P0.1 Force-build and validate the current x64 runtime.
  - [x] P0.2 Force-build and validate the current ARM64 runtime.
  - [x] P0.3 Complete x64 Offline Force and Offline reuse without network.
  - [x] P0.4 Build Debug x64 and test-enabled Release x64; run existing Terminal
        and embedded-terminal Commands tests.
  - [x] P0.5 Add the deterministic measurement-only VT upgrade corpus while the
        canonical lock still names the current pin.
  - [x] P0.6 Archive a current-pin baseline with exact corpus/output/snapshot
        digests and at least 200 samples for every p95 claim.
  - [x] P0.7 Run a clean tests-disabled Release x64 portable-package build and
        prove whether package smoke uses a harness from the same receipt.
  - [x] P0.8 If the package-smoke mismatch is confirmed, land and test the
        production-safe repair; otherwise record a checked not-applicable
        disposition with proof.
  - [x] P0.9 Run every Phase 0 verification command and validate the resulting
        Terminal/package evidence archives.
  - [x] P0.10 Create the required Establish Ghostty upgrade baseline commit,
        then record its full SHA in the phase table and Evidence log through a
        plan-only checkpoint commit before P1 starts.

### P1 — Active product lock no-op migration

- [x] **P1 — Active product lock no-op migration complete**
  - [x] P1.1 Create and validate GhosttyRuntimeLock.schema.json with strict
        required properties, bounded values, sorted unique sets, and no
        additional properties.
  - [x] P1.2 Create the canonical current-pin lock from every active builder
        constant without changing source, toolchain, dependency, overlay,
        import/export, license, build, or pairing identity.
  - [x] P1.3 Add all positive and negative lifecycle/schema fixtures named in
        Phase 1.
  - [x] P1.4 Replace the nonexistent TerminalEngine.lock.json BuildEvidence
        input with the real active lock and complete direct identity closure.
  - [x] P1.5 Add the active lock/schema/module to the existing reusable-build
        cache closure; do not create another cache.
  - [x] P1.6 Update tool inventory, help, and owning specs for the lock surface.
  - [x] P1.7 Prove current x64 and ARM64 semantic/runtime identities have no
        unexplained delta after lock migration.
  - [x] P1.8 Create the required Add Ghostty runtime product lock commit, then
        record its full SHA in the phase table and Evidence log through a
        plan-only checkpoint commit before P2 starts.

### P2 — Lock-driven builder and candidate isolation

- [x] **P2 — Lock-driven builder and candidate isolation complete**
  - [x] P2.1 Implement GhosttyRuntimeLifecycle.psm1 by reusing the active
        archive, identity, import/export, and contained-process helpers while
        retaining the reviewed Ghostty repository/canonical-drive leases.
  - [x] P2.2 Keep Build-GhosttyTerminalRuntime.ps1 as the stable public Product
        entrypoint with the canonical lock as its default.
  - [x] P2.3 Implement explicit candidate discovery/build output only beneath
        .build/TerminalEngineUpgrade.
  - [x] P2.4 Prove candidate mode cannot write Product, overwrite pairing
        headers, fall back to an omitted descriptor, or escape its output root.
  - [x] P2.5 Prove fail-closed behavior for lock, schema, overlay, license,
        source, import/export, machine, size, process-ownership, and path errors.
  - [x] P2.6 Repeat current-pin x64/ARM64 Force, x64 Offline Force, and Offline
        reuse through the refactored builder with exact equivalence.
  - [x] P2.7 Run every Phase 2 lifecycle/dependency/import/export test.
  - [x] P2.8 Create the required Refactor Ghostty runtime lifecycle commit, then
        record its full SHA in the phase table and Evidence log through a
        plan-only checkpoint commit before P3 starts.

### P3 — Read-only observation and advisory process

- [x] **P3 — Read-only observation and advisory process complete**
  - [x] P3.1 Create the strict Ghostty runtime status schema and canonical
        serialization tests.
  - [x] P3.2 Implement Locked mode with zero network and exact canonical-lock
        validation.
  - [x] P3.3 Implement Online mode with factual current, update-observed,
        security-review, required-manual, and unreachable results.
  - [x] P3.4 Implement Candidate mode requiring an explicit full commit and
        refusing substitution with default-branch head.
  - [x] P3.5 Prove all output stays under .build and no status mode mutates the
        lock or any tracked file.
  - [x] P3.6 Add deterministic mocked-network tests for every success/failure
        case listed in Phase 3.
  - [x] P3.7 Add the pinned, read-only daily/monthly/manual workflow with no
        issue, branch, pull-request, or lock-write permissions.
  - [x] P3.8 Add the lock/source-bound fail-closed release preflight.
  - [x] P3.9 Run status, workflow-policy, tool-governance, and inventory
        verification with zero failures/findings.
  - [x] P3.10 Create the required Add Ghostty update observation commit, then
        record its full SHA in the phase table and Evidence log through a
        plan-only checkpoint commit before P4 starts.

### P4 — Exact candidate freeze and review

- [x] **P4 — Exact candidate freeze and review complete**
  - [x] P4.1 Authenticate candidate commit, tree, archive SHA-256, version, and
        official-source provenance.
  - [x] P4.2 Prove the candidate is exactly 717 commits ahead and zero behind
        the current pin.
  - [x] P4.3 Regenerate and classify the complete path-filtered change inventory
        under the Phase 4 categories.
  - [x] P4.4 Diff every public VT header, exported C declaration, enum/value,
        and all 72 loader-required exports.
  - [x] P4.5 Complete the bounded allocation/pointer/OSC/DCS/APC/clipboard/
        snapshot security review without unsupported vulnerability claims.
  - [x] P4.6 Create the exact untracked candidate descriptor beneath the
        candidate .build root and bind all observation/input digests.
  - [x] P4.7 Create the required Freeze Ghostty runtime candidate commit, then
        record its full SHA in the phase table and Evidence log through a
        plan-only checkpoint commit before P5 starts.

### P5 — Candidate overlay rederivation

- [x] **P5 — Candidate overlay rederivation complete**
  - [x] P5.1 Map every existing overlay hunk to a retained, replaced, or removed
        concern with source evidence and a focused test.
  - [x] P5.2 Reimplement reviewed Windows shared-runtime linker/strip behavior
        against pristine candidate source and prove the unused static companion
        remains outside staging and packages.
  - [x] P5.3 Reallocate private enum values after the complete candidate public
        ranges and prove C/Zig uniqueness and parity.
  - [x] P5.4 Preserve 4,096 placement count, 1 MiB placement-map allocation,
        atomic all-screen validation, unchanged rejection, replacement, reset,
        screen, pin-release, and query contracts.
  - [x] P5.5 Generate one normalized patch and update its ordered concern/digest
        identity.
  - [x] P5.6 Prove the patch applies cleanly to pristine candidate source and
        diff/overlay tests pass.
  - [x] P5.7 Confirm the overlay remains within 16 files, 2,500 changed lines,
        and five independent product concerns.
  - [x] P5.8 Create the required Rederive Ghostty Windows runtime overlay
        commit, then record its full SHA in the phase table and Evidence log
        through a plan-only checkpoint commit before P6 starts.

### P6 — Isolated x64 and native ARM64 candidate qualification

- [x] **P6 — Isolated candidate qualification complete**
  - [x] P6.1 Hash Product and pairing inputs before the candidate builds.
  - [x] P6.2 Build the candidate x64 and ARM64 only under their isolated roots.
  - [x] P6.3 Prove Product DLL/import library/manifests/pairing hashes are
        unchanged after candidate builds.
  - [x] P6.4 Pass upstream test-lib-vt, test-lib-vt-build, and
        test-lib-vt-schema on x64.
        REVALIDATED 2026-08-26: after the corrective Kitty clipboard-write gate
        invalidated the earlier result, all three exact targets passed from a
        fresh isolated cache and the schema target validated 153 ABI types.
  - [x] P6.5 Pass the same upstream targets on native ARM64 Windows.
        VERIFIED 2026-08-26: run 32930638145 passed the exact build and all
        three upstream targets on native ARM64. Canonical evidence validated
        the 153-type schema, frozen runtime identity, diagnostic module clone,
        and `codeFiveRetryUsed=false`.
  - [x] P6.6 Record and explain exact candidate exports, loader subset, imports,
        machine, size, build info, dependencies, and licenses.
  - [x] P6.7 Create a complete proposed candidate lock under .build and pass a
        second clean isolated build using its exact expectations.
  - [x] P6.8 Create the required Qualify isolated Ghostty candidate commit, then
        record its full SHA in the phase table and Evidence log through a
        plan-only checkpoint commit before P7 starts.

### P7 — Loader, mode, and parser ABI adaptation

- [x] **P7 — Loader, mode, and parser ABI adaptation complete**
  - [x] P7.1 Add the candidate path that replaces
        ghostty_terminal_mode_get with the frozen two-field
        ghostty_terminal_get/GHOSTTY_TERMINAL_DATA_MODE contract while the
        canonical Product compatibility path remains active through P9.
  - [x] P7.2 Update the candidate loader-required exports and preserve atomic partial-load
        rollback and old-header/new-DLL rejection.
  - [x] P7.3 Add structure-size guards for every versioned candidate structure
        read by the adapter.
  - [x] P7.4 Pass exact key, mouse, paste, focus, alternate-screen, bracketed
        paste, mouse-mode, and terminal-mode tests.
  - [x] P7.5 Add and pass the DCS byte-above-0x7F regression.
  - [x] P7.6 Prove the candidate build/runtime path no longer requires
        ghostty_terminal_mode_get. The canonical Product path deliberately
        retains its locked symbol until the candidate is admitted in P10; P10
        removes that compatibility branch and updates the canonical lock.
  - [x] P7.7 Create the required Adapt Terminal to Ghostty mode ABI commit, then
        record its full SHA in the phase table and Evidence log through a
        plan-only checkpoint commit before P8 starts.

### P8 — Clipboard security and reply adaptation

- [x] **P8 — Clipboard security and reply adaptation complete**
  - [x] P8.1 Convert production/test callbacks to the candidate void,
        size-versioned request/reply ABI.
  - [x] P8.2 Set and enforce both the 256 KiB public runtime clipboard-write
        ceiling and private option 41=false during terminal initialization;
        fail initialization if either set fails.
  - [x] P8.3 Preserve bounded allow/deny/ask behavior without retaining upstream
        pointers, blocking, or calling UI/clipboard APIs under the terminal
        lock.
  - [x] P8.4 Prove the engine denies Kitty OSC 5522 with `ENOSYS` before
        buffering/callback while OSC 52 and OSC 1337 still reach the bounded
        host callback; invoke any callback reply at most once.
  - [x] P8.5 Keep clipboard-read callback null and fail closed if it becomes
        required.
  - [x] P8.6 Pass every boundary, encoding, consent, reply, teardown, Kitty, and
        no-read case listed in Phase 8.
  - [x] P8.7 Create the required Adapt Ghostty clipboard security contract
        commit, then record its full SHA in the phase table and Evidence log
        through a plan-only checkpoint commit before P9 starts.

### P9 — Pending and animated Kitty image lifecycle

- [x] **P9 — Pending and animated Kitty image lifecycle complete**
  - [x] P9.1 Separate engine-pending, adapter-queued, ready, rejected, replaced,
        deleted, evicted, and animated state.
  - [x] P9.2 Retry engine-pending content on storage-generation change and
        invalidate cached frames on image-generation change.
  - [x] P9.3 Reject stale worker results by request/storage/image generation
        without resurrecting removed content.
  - [x] P9.4 Preserve queue coalescing, cancellation, one-worker, memory
        accounting, Close quiet point, and resize reuse.
  - [x] P9.5 Pass every pending/complete/replace/delete/evict/retransmit/
        animation/reset/screen/stale-result deterministic regression.
  - [x] P9.6 Pass placement-limit, replacement, pin-release, zero-byte Close,
        cache/pipeline memory, and no-reupload invariants.
  - [x] P9.7 Create the required Handle pending Ghostty Kitty images commit,
        then record its full SHA in the phase table and Evidence log through a
        plan-only checkpoint commit before P10 starts.

### P10 — Performance review and manual Product admission

- [x] **P10 — Performance review and manual Product admission complete**
  - [x] P10.1 Run the frozen corpus against the isolated candidate on the same
        machine/configuration as the current-pin baseline.
  - [x] P10.2 Prove exact non-intentional output/snapshot digests and record the
        named DCS correction.
  - [x] P10.3 Pass hosted-render, VT-write, and snapshot p95 ratios with at least
        200 samples plus every fixed memory/resource/Close/resize gate.
  - [x] P10.4 Complete human review of provenance, source diff, overlay, ABI,
        security, clipboard, Kitty, upstream/native ARM64, licenses, packages,
        and performance evidence.
  - [x] P10.5 Replace the canonical lock only after P10.1 through P10.4 are
        checked; bind overlay/adapter/spec changes in the same coherent review.
  - [x] P10.6 Pass canonical x64/ARM64 Force, x64 Offline Force, and Offline
        reuse with manifests bound to the admitted lock.
  - [x] P10.7 Validate both baseline and candidate archives and record exact run
        IDs in the Evidence log.
  - [x] P10.8 Create the required Admit reviewed Ghostty runtime candidate
        commit, then record its full SHA in the phase table and Evidence log
        through a plan-only checkpoint commit before P11 starts.

### P11 — Shipping-boundary qualification

- [x] **P11 — Shipping-boundary qualification complete**
  - [x] P11.1 Build Debug x64, test-enabled Release x64, normal Release x64 and
        ARM64, ASan Debug x64, and Debug ARM64.
  - [x] P11.2 Pass all x64 Debug and test-enabled Release Terminal/Commands
        cases.
  - [x] P11.3 Pass ASan Terminal cases with no sanitizer finding.
  - [x] P11.4 Pass the admitted-product native ARM64 workflow, including runtime
        schema/load, Terminal, representative Commands cases, production build,
        and native portable smoke; retain its required artifact evidence.
  - [x] P11.5 Build and validate x64/ARM64 ZIP, x64 MSI, and x64/ARM64 MSIX
        serially with the approved build identity.
  - [x] P11.6 Prove every package has the exact runtime architecture, identity,
        required exports/imports, licenses, native DLL load/free when executable,
        plugin load, and receipt-bound smoke. Prove release builds depend on the
        admitted-product native qualification workflow or an explicit reviewed
        waiver; the default release path has no waiver.
  - [x] P11.7 Pass the complete focused Pester set with zero failures.
        Amendment 2026-08-26: the byte-sealed historical
        `TerminalDependencyLifecycle.Tests.ps1` is not an active-runtime gate
        and must not be edited or executed against the admitted six-notice
        specification. Active dependency/notice coverage remains in
        `GhosttyRuntimeLifecycle.Tests.ps1` and
        `TerminalPluginSourceContracts.Tests.ps1`; this P11 repair also adds
        `VcpkgInstallSafety.Tests.ps1` and
        `TestHarnessSourceContracts.Tests.ps1` to the focused set.
  - [x] P11.8 Pass Fresh Full x64 Debug with terminal status passed and
        repository status passed.
  - [x] P11.9 Create the required Qualify Ghostty shipping boundaries commit,
        then record its full SHA in the phase table and Evidence log through a
        plan-only checkpoint commit before P12 starts.

### P12 — Rollback, durable contracts, and archive transition

- [x] **P12 — Rollback, durable contracts, and archive transition complete**
  - [x] P12.1 Rebuild the previous x64 and ARM64 runtime from the pre-admission
        parent in isolation; do not copy Product cache files.
  - [x] P12.2 Pass previous-lock Offline, imports/exports/licenses/pairing,
        Terminal, and package-smoke rollback checks.
  - [x] P12.3 Update every authoritative Terminal/Build/Testing/Installer/Tools
        contract named in Phase 12 and NormativeConsistency when applicable.
  - [x] P12.4 Run Locked status, tool inventory, spec inventory, diff check, and
        final receipt/archive validation with zero findings.
  - [x] P12.5 Confirm the reviewed diff contains only in-scope implementation,
        tests, bounded evidence, normative specs, this plan, and WIP index
        changes.
  - [x] P12.6 Confirm every C0/P0–P11 leaf and phase box is checked, every
        Evidence log row is complete, no BLOCKED marker remains, and Current
        progress says READY TO ARCHIVE.
  - [x] P12.7 Create the required Close Ghostty runtime upgrade commit while the
        plan remains in WIP, then capture its full SHA.
  - [x] P12.8 In a separate archive/checkpoint commit, record the P12 SHA in the
        phase table and Evidence log, move this plan to Specs/Plans/Done, set its
        State to DONE, check P12/P12.7/P12.8 in the moved file, remove I13 from
        the WIP index, and rerun spec inventory.

### Evidence log

Append one row whenever a leaf item is checked. Use repository-relative paths
for durable evidence and never include secrets or absolute user-profile paths.

| Checklist item | Date | Change / commit | Verification result | Durable evidence / receipt |
|---|---|---|---|---|
| Initial plan / C0 preparation | 2026-08-25 | e0cc093b27a9ef4d63f7f03a64ec24707731c610 | Get-SpecInventory.ps1 -FailOnFindings: 0 blocking findings; git diff --check: exit 0 | .build/SpecInventory/spec-inventory.json |
| C0.1 | 2026-08-25 | 3a2716b447295c7223f979850585140e6d7e7d9d | Planned-at path-scoped git diff: no source/build/spec drift; worktree clean | Git history and this plan |
| C0.2 | 2026-08-25 | 3a2716b447295c7223f979850585140e6d7e7d9d | Windows 11 build 26200; AMD Ryzen 9 9950X3D; x64; VS 2026 18.9.1; MSVC 19.51.36252; SDK 10.0.26100.0; Zig/vcpkg executable caches absent before baseline build; locked vcpkg checkout a51bb4d1434e6d0927ff79db8033bed8522b85df | Console evidence summarized here; exact non-user-path hashes retained in task log |
| C0.3 | 2026-08-25 | 3a2716b447295c7223f979850585140e6d7e7d9d | I10/I12 ownership reviewed; no concurrent executor or overlapping dirty edit in this worktree | Specs/Plans/WIP/README.md; clean git status |
| C0.4 | 2026-08-25 | 3a2716b447295c7223f979850585140e6d7e7d9d | Scope and STOP conditions reviewed; no amendment required before P0 | This plan |
| C0.4 P6 scope amendment | 2026-08-26 | working tree after 08329cbe7a153edcb97ba92f7fd8b2471b2c6fa7 | Added the already-reviewed candidate freeze/overlay/native-qualification schemas, thin commands, internal modules, focused tests, and read-only native workflow to the explicit surface. This records the P4–P6 implementation paths and adds the native workflow to the recurring drift command; no product-admission scope was broadened. | Scope and drift-command sections in this plan |
| P0.1 | 2026-08-25 | 19cd99f8e7c297c9ef9b2429c37e95c1a1cbe73a | Current pin x64 Force: exit 0; AMD64; 180 exports; export hash 596894726504f7d398ee2a87b8a53116c8d727e5390b1e066293e1770e40fced; imports kernel32.dll/ntdll.dll | .build/TerminalEngine/Product/x64/terminal-runtime-identity.json |
| P0.2 | 2026-08-25 | 19cd99f8e7c297c9ef9b2429c37e95c1a1cbe73a | Current pin ARM64 Force: exit 0; ARM64; 180 exports; same export hash; imports kernel32.dll/ntdll.dll | .build/TerminalEngine/Product/ARM64/terminal-runtime-identity.json |
| P0.3 | 2026-08-25 | 19cd99f8e7c297c9ef9b2429c37e95c1a1cbe73a | x64 Offline Force and Offline reuse: exit 0 without restore/download; reuse retained raw SHA-256 c33e337d7a581b1eb135559a53b4dc8061c3a3b5fd2bd57a402f7cfe88a94e8d | .build/TerminalEngine/Product/x64/terminal-runtime-identity.json |
| P0.4 blocked attempt 1 | 2026-08-25 | 19cd99f8e7c297c9ef9b2429c37e95c1a1cbe73a | Debug x64: 0 warnings/0 errors, receipt 0a1f3bb9f7020d98633cb38675dc6bec29a3e4bc5f6582e445a3cdd5f0acf23d; Terminal 99/99; Commands filter exit 0. Test-enabled Release failed with five C1090 errors. | .build/logs/msbuild-20260825_201850_329-pid64200-d9822c19.log; .build/logs/msbuild-20260825_202320_952-pid14904-6d9bcf00.log |
| P0.4 blocked attempt 2 | 2026-08-25 | 19cd99f8e7c297c9ef9b2429c37e95c1a1cbe73a | Test-enabled Release Rebuild with MaxCpuCount 4 failed with six C1090 PDB API/RPC errors across FileSystem, Curl, MicrosoftDrive, MTP, S3, and DxUi. Read-only process evidence found an independent MSBuild rebuild rooted at Z:\src\RedSalamander plus active mspdbsrv; possible cross-worktree PDB interference, not proved and not terminated. | .build/logs/msbuild-20260825_202505_149-pid21224-bd695926.log |
| P0.4 | 2026-08-25 | 6be623578015366816c5d7643fd427657a1e3628 | After the independently owned builds exited, test-enabled Release x64 Rebuild with MaxCpuCount 1 passed with 0 warnings/0 errors and receipt df7143c2ba1ff656600cff47442ef54430768bb8cf0436d95868e8aa3b977a3a; Release Terminal selftests passed 99/99 and the embedded-terminal Commands filter exited 0. | .build/logs/msbuild-20260825_203224_952-pid92252-a81f34eb.log |
| P0.5 | 2026-08-25 | 1517b44ede1f3ae06f72d291255804e4db330d6b | Added the declarative corpus, strict evidence schema, test-only Terminal export, 5 warmups plus 200 measured iterations, digest-only semantic snapshots, Release measurements, Kitty pending/completion/accounting probes, and source-contract coverage. Focused Pester passed 31/31; focused Terminal and PluginContractTests Release builds passed with 0 warnings/0 errors. | Specs/Terminal/TerminalVtUpgradeCorpus.v1.json; Specs/Terminal/TerminalVtUpgradeEvidence.schema.json; .build/logs/msbuild-20260825_211410_677-pid33940-5b9ec506.log; .build/logs/msbuild-20260825_211709_936-pid93244-2c2e0c9d.log |
| P0.6 | 2026-08-25 | 1517b44ede1f3ae06f72d291255804e4db330d6b | Current-pin Release corpus passed 11 assertions; 200/200 samples per p95 metric; corpus SHA-256 e457df11809c0f039b1e09fcf94fa940d389d51972e54503b25ddcadbed1b7ba; PTY SHA-256 4b2c3c1d5bc7784cac1d05215e046d912a43658e8d816cc065ecb3d4b4bcc3e4; snapshot SHA-256 dc3a1ad8367ae51a75ffff481d6831c5bfabc14b8c162489d147ed75c96a0de1; p95 values 12175/4077/6867 us for VT write/snapshot/hosted render; archive contract passed for all 3 files. | Specs/TestRuns/4cb089111a23/Terminal/20260825_211846_current-pin/results.json; Specs/TestRuns/4cb089111a23/Terminal/20260825_211846_current-pin/trace.txt; Specs/TestRuns/4cb089111a23/Terminal/20260825_211846_current-pin/perf/perf_metrics.jsonl |
| P0.7 confirmed defect | 2026-08-25 | 07068abcb83e65623bbc164ce67da6e92b480d6f | A clean tests-disabled Release x64 Rebuild compiled with 0 warnings/0 errors but failed receipt publication because disabled test-project CppClean replay removed shared blake3.dll/yyjson.dll; the same profile did not build a current package-smoke harness. | .build/logs/msbuild-20260825_211959_034-pid29140-df543f18.log; Tests/Directory.Build.targets; .build/Intermediate/x64/Release/Tests/PluginContractTests/PluginContractTests.Build.CppClean.log |
| P0.7 | 2026-08-25 | e5edd4b02d48d6cd8fb95b29c9b7c78b7b00a58b | Clean tests-disabled Release x64 Rebuild passed with 0 warnings/0 errors and receipt abe2af24268d9c5b2fb3280f054dd289b49c543eba2cdbdb2d1db7090fe178f4. The receipt has tests_enabled=false, attests one PluginContractTests.exe with SHA-256 b124c699942d7f3faaae3ed49261d0bf032614410781c5cebc8e732cfaa57096, and that digest matched the invoked file. Fresh-extraction x64 ZIP smoke passed and produced SHA-256 45e8bfc400ee193584bb7a2181f9bc6b10f1f8938175f9a1fca2a76a744c72ea. | .build/logs/msbuild-20260825_213715_960-pid29324-1b61c742.log; .build/x64/Release/build-receipt.json; .build/AppPackages/RedSalamander-7.0.1-x64-Portable.zip |
| P0.8 | 2026-08-25 | e5edd4b02d48d6cd8fb95b29c9b7c78b7b00a58b | Confirmed premise repaired: PluginContractTests is the sole tests-disabled non-shipping package host; package mode skips test-only exports; other disabled test nodes clean only project-named outputs and cannot remove shared app-local dependencies. Focused BuildReproducibility passed 20/20; tests-disabled focused build passed with 0 warnings/0 errors; direct and fresh-extraction package smoke passed. | Tests/Directory.Build.targets; Tests/PluginContractTests/PluginContractTests.cpp; Specs/Build/Build_Toolchain.md; Specs/Installer/Installer_PortableZip.md; .build/logs/msbuild-20260825_213630_679-pid39280-5af0b0ce.log |
| P0.9 | 2026-08-25 | working tree at 1537015ede42c4e7b4f4b60b2f8e17171ec8a250 | All Phase 0 gates passed. Runtime commands exited 0 for x64/ARM64 Force and x64 Offline Force/reuse; the first online x64 Force exposed the already-scoped non-deterministic rebuild delta (raw 5b18a0b2f3775baac7ec6a1beab98a331855d2bfa00ed60c0a2dd3450fd16e8d, normalized 9443d9b916e6bfa947df0d6527dc450008f72e2c1f75f4f8af12f2b09134e82d), while Offline Force/reuse restored and retained the canonical x64 identity (raw c33e337d7a581b1eb135559a53b4dc8061c3a3b5fd2bd57a402f7cfe88a94e8d, normalized 420c19583a5f64eaa93008644e33c7c701a5642ccac4050ac34acb13157e1f10); ARM64 remained raw 2893ba637dc7d2ee6b29cc3bb77a8025cb3f9cf0207cfcec4289c3ab3e314b3f, normalized 9e96a3318dc430e8868d3f14938ccb0f43e9cdcbe18c7d69aac8b78e411c1525, and both platforms retained 180 exports with export-set SHA-256 596894726504f7d398ee2a87b8a53116c8d727e5390b1e066293e1770e40fced. The literal Debug x64 build passed with 0 warnings/0 errors and receipt 6e79f7c916f98e74a0a276a99b43dfa8de3f0105b6770f3805d8b0a0b93ef7fe; Terminal passed 99/99 and Commands exited 0. The literal Pester invocation correctly failed closed without a selected artifact profile; after adding the required x64/Debug variables to this plan, BuildReproducibility plus targeted deployment passed 21/21 and Terminal source contracts passed 31/31. Evidence schema and three-file archive validation passed. Fresh-extraction package smoke passed all 19 required-file checks, app startup, and package contracts; ZIP SHA-256 remained 45e8bfc400ee193584bb7a2181f9bc6b10f1f8938175f9a1fca2a76a744c72ea, and the tests-disabled Release receipt's sole harness SHA-256 b124c699942d7f3faaae3ed49261d0bf032614410781c5cebc8e732cfaa57096 matched the invoked file. | .build/logs/msbuild-20260825_221546_478-pid31080-6f7c430d.log; .build/x64/Debug/build-receipt.json; .build/TerminalEngine/Product/x64/terminal-runtime-identity.json; .build/TerminalEngine/Product/ARM64/terminal-runtime-identity.json; Specs/TestRuns/4cb089111a23/Terminal/20260825_211846_current-pin/results.json; .build/x64/Release/build-receipt.json; .build/AppPackages/RedSalamander-7.0.1-x64-Portable.zip |
| P0.10 | 2026-08-25 | 95f3d5a02a6d519751166e2e7c6239deaede2dba | Required phase boundary commit created with exact subject `Establish Ghostty upgrade baseline`; `git rev-parse HEAD` and `git show -s --format=%s` matched the recorded SHA and subject. | Git history and Required phase commits table |
| P1.1 | 2026-08-25 | working tree at 1cbc8336b37d2ae65dba354a8c59682fedfac3c7 | Added strict draft-2020-12 schema version 1 with closed objects, bounded fields, lowercase digest/pin/path rules, and runtime/toolchain/dependency/overlay/build/ABI/notice/pairing sections; canonical validation passed. | Specs/Terminal/GhosttyRuntimeLock.schema.json; External/TerminalEngine/GhosttyRuntimeLock.v1.json |
| P1.2 | 2026-08-25 | working tree at 1cbc8336b37d2ae65dba354a8c59682fedfac3c7 | Migrated all active builder product identities into the canonical current-pin lock; source remains b2d44625908eb56aee299d8511d803bc11ab79fc/tree 9b5e7ca2fb3dfaf7c8329645dfb51a1d3baa0325, lock digest is 7063130976d50ef5c48aadd3ea1b896bf0aab3ba5b63adbcfadbde20d4cf0f5d, and the builder parser reported zero errors. | External/TerminalEngine/GhosttyRuntimeLock.v1.json; Tools/Build-GhosttyTerminalRuntime.ps1 |
| P1.3 | 2026-08-25 | working tree at 1cbc8336b37d2ae65dba354a8c59682fedfac3c7 | Canonical positive validation plus every named missing/extra/hash/order/duplicate/derived/subset/machine/path/schema/overlay/license negative fixture passed with field-specific diagnostics; Ghostty lifecycle suite passed 15/15. | Tools/Tests/GhosttyRuntimeLifecycle.Tests.ps1 |
| P1.4 | 2026-08-25 | working tree at 1cbc8336b37d2ae65dba354a8c59682fedfac3c7 | BuildEvidence now binds the 15-file direct Ghostty product identity closure, excludes the nonexistent generic lock, and its mutation proof changed the identity for every direct input. | Tools/Modules/Build/BuildEvidence.psm1; Tools/Tests/BuildEvidence.Tests.ps1 |
| P1.5 | 2026-08-25 | working tree at 1cbc8336b37d2ae65dba354a8c59682fedfac3c7 | The existing reusable-build Terminal cache hashes the canonical lock, schema, seven direct modules, overlay, notices, and staging manifest; Release workflow source contracts passed. | .github/workflows/build-reusable.yml; Tools/Tests/ReleaseWorkflowPolicy.Tests.ps1 |
| P1.6 | 2026-08-25 | working tree at 1cbc8336b37d2ae65dba354a8c59682fedfac3c7 | Public help, tool inventory, README maps, and authoritative Build/Terminal specs were updated; focused lifecycle/evidence/release/governance Pester passed 61/61; tool inventory and spec inventory each reported 0 blocking findings. | Tools/tool-inventory.json; Tools/README.md; Tools/TerminalEngine/README.md; Specs/Build/Build_Toolchain.md; Specs/Terminal/Terminal_EmbeddedPlugin.md; .build/SpecInventory/spec-inventory.json |
| P1.7 | 2026-08-25 | working tree at 1cbc8336b37d2ae65dba354a8c59682fedfac3c7 | Lock-driven Offline reuse validated x64 AMD64/1560576 bytes and ARM64/1337856 bytes with exact kernel32.dll/ntdll.dll imports, 180 exports, export-set SHA-256 596894726504f7d398ee2a87b8a53116c8d727e5390b1e066293e1770e40fced, and unchanged notice identities. ARM64 retained raw/normalized 2893ba637dc7d2ee6b29cc3bb77a8025cb3f9cf0207cfcec4289c3ab3e314b3f/9e96a3318dc430e8868d3f14938ccb0f43e9cdcbe18c7d69aac8b78e411c1525. x64 raw/normalized 6c6af73525019129917beca9959c5502faf3ece65eb3db9bafa68ee061dd0780/c4b15003eb85d6f29c2e25c48a30160e21f989b7e1108bd70c6fc6d71667dce4 is the explained Zig/LLD layout variance already demonstrated at P0; source, toolchain, build arguments, overlay, ABI, imports, notices, and pairing contract are unchanged. | .build/TerminalEngine/Product/x64/terminal-runtime-identity.json; .build/TerminalEngine/Product/ARM64/terminal-runtime-identity.json |
| P1.8 | 2026-08-25 | f830294e4ab6fd997937e03ab036d20a75920376 | Required phase boundary commit created with exact subject `Add Ghostty runtime product lock`; `git rev-parse HEAD` and `git show -s --format=%s` matched the recorded SHA and subject. | Git history and Required phase commits table |
| P2.1 | 2026-08-25 | working tree at b9b4706f7f68a033e991998b584445f97b95063b | Moved lock parsing, restore, overlay, build, inspection, notice, identity, and publication behavior into GhosttyRuntimeLifecycle.psm1 while reusing archive, PE/text/import/export, canonical-JSON, and contained-process helpers. The reviewed coordination amendment retains the Ghostty repository mutex and rebuild-only R: lease because the nested dependency builder is outside app artifact-profile roots. | Tools/TerminalEngine/GhosttyRuntimeLifecycle.psm1; reviewed P2 coordination amendment in this plan |
| P2.2 | 2026-08-25 | working tree at b9b4706f7f68a033e991998b584445f97b95063b | Build-GhosttyTerminalRuntime.ps1 is a 72-line documented wrapper that delegates to Invoke-RSGhosttyRuntimeBuild; no candidate arguments select the canonical lock and Product output. Product x64/ARM64 Offline reuse exited 0. | Tools/Build-GhosttyTerminalRuntime.ps1; Tools/Tests/GhosttyRuntimeLifecycle.Tests.ps1 |
| P2.3 | 2026-08-25 | working tree at b9b4706f7f68a033e991998b584445f97b95063b | Resolve-RSGhosttyRuntimeBuildPlan requires an explicit candidate lock/output pair and confines both the state and output to children of .build/TerminalEngineUpgrade with pairing publication disabled. | Tools/TerminalEngine/GhosttyRuntimeLifecycle.psm1; Tools/Tests/GhosttyRuntimeLifecycle.Tests.ps1 |
| P2.4 | 2026-08-25 | working tree at b9b4706f7f68a033e991998b584445f97b95063b | Candidate tests passed for contained paths, Product exclusion, upgrade-root exclusion, missing/incomplete descriptor rejection, no canonical fallback, pairing disabled, and unchanged Product DLL/header/identity bytes after rejected candidate invocation. | Tools/Tests/GhosttyRuntimeLifecycle.Tests.ps1 |
| P2.5 | 2026-08-25 | working tree at b9b4706f7f68a033e991998b584445f97b95063b | Lifecycle/schema negative fixtures, generic dependency/source/archive/path tests, exact import/export tests, machine/size/notice checks, and atomic publication source contracts passed. A held exact output DLL failed before Product deletion with an independently-owned diagnostic, retained identical bytes, and confirmed no termination path. | Tools/Tests/GhosttyRuntimeLifecycle.Tests.ps1; Tools/Tests/TerminalDependencyLifecycle.Tests.ps1; Tools/Tests/TerminalRuntimeExportIdentity.Tests.ps1; Tools/Tests/TerminalEngineRuntimeImportPolicy.Tests.ps1 |
| P2.6 | 2026-08-25 | working tree at b9b4706f7f68a033e991998b584445f97b95063b | Final-code x64 Force passed at AMD64/1560576 bytes, raw c33e337d7a581b1eb135559a53b4dc8061c3a3b5fd2bd57a402f7cfe88a94e8d, normalized 420c19583a5f64eaa93008644e33c7c701a5642ccac4050ac34acb13157e1f10. ARM64 Force retained raw 2893ba637dc7d2ee6b29cc3bb77a8025cb3f9cf0207cfcec4289c3ab3e314b3f, normalized 9e96a3318dc430e8868d3f14938ccb0f43e9cdcbe18c7d69aac8b78e411c1525, 1337856 bytes. x64 Offline Force produced the other already-explained valid Zig/LLD layout 5b18a0b2f3775baac7ec6a1beab98a331855d2bfa00ed60c0a2dd3450fd16e8d/9443d9b916e6bfa947df0d6527dc450008f72e2c1f75f4f8af12f2b09134e82d and Offline reuse matched it byte-for-byte. Both platforms retained the exact imports, 180 exports, export hash, notices, build tuple, and pairing contract. | .build/TerminalEngine/Product/x64/terminal-runtime-identity.json; .build/TerminalEngine/Product/ARM64/terminal-runtime-identity.json |
| P2.7 | 2026-08-25 | working tree at b9b4706f7f68a033e991998b584445f97b95063b | Required lifecycle/dependency/export/import Pester passed 34/34; the expanded lifecycle/build-evidence/release/governance cohort passed 80/80; tool inventory and spec inventory each reported 0 blocking findings; git diff --check exited 0. | Tools/Tests; .build/SpecInventory/spec-inventory.json |
| P2.8 | 2026-08-25 | 1e135cbbe2323722cbaa1e50c72fdf123e455838 | Required phase boundary commit created with exact subject `Refactor Ghostty runtime lifecycle`; `git rev-parse HEAD` and `git show -s --format=%s` matched the recorded SHA and subject. | Git history and Required phase commits table |
| P3.1 | 2026-08-25 | working tree at 004b9231c8b0e3fb8a3af613e0b4a856ce470171 | Added closed draft-2020-12 status schema version 1, canonical atomic publication, repeated fixed-time byte equality, and additional-property rejection; schema/canonical cases passed. | Specs/Terminal/GhosttyRuntimeStatus.schema.json; Tools/Tests/GhosttyRuntimeStatus.Tests.ps1 |
| P3.2 | 2026-08-25 | working tree at 004b9231c8b0e3fb8a3af613e0b4a856ce470171 | Final Locked run returned current for b2d44625908eb56aee299d8511d803bc11ab79fc with zero network attempts and canonical lock SHA-256 e507e5aed353551768f4451553a06a20836d80ba7925e87108370faa194a2a70. | .build/TerminalEngineStatus/p3-final/locked.json |
| P3.3 | 2026-08-25 | working tree at 004b9231c8b0e3fb8a3af613e0b4a856ce470171 | Final real Online run authenticated official head 683d8db643b95cf229bfb5fe9fab9ae677920343/tree 1274f47578cc98a8285aeb3732ade0b5dd1872f2 as update-observed, 719 ahead/0 behind, with four GitHub plus one OSV HTTP 200 attempts and zero advisories; deterministic cases cover every factual result. | .build/TerminalEngineStatus/p3-final/online.json; Tools/Tests/GhosttyRuntimeStatus.Tests.ps1 |
| P3.4 | 2026-08-25 | working tree at 004b9231c8b0e3fb8a3af613e0b4a856ce470171 | Final real Candidate run resolved exactly requested commit 9c3ec931d64561a8407dde7ac984ce156ae91539/tree 359872ab74fe1ecc96428133c0546be3cb969fdd at 717 ahead/0 behind; substitution and deleted/refused candidate cases passed. | .build/TerminalEngineStatus/p3-final/candidate.json; Tools/Tests/GhosttyRuntimeStatus.Tests.ps1 |
| P3.5 | 2026-08-25 | working tree at 004b9231c8b0e3fb8a3af613e0b4a856ce470171 | Status publication rejected an output escape, wrote only ignored evidence below .build, and left tracked porcelain unchanged across Online fixture execution. | Tools/TerminalEngine/GhosttyRuntimeStatus.psm1; Tools/Tests/GhosttyRuntimeStatus.Tests.ps1 |
| P3.6 | 2026-08-25 | working tree at 004b9231c8b0e3fb8a3af613e0b4a856ce470171 | Status suite passed 14/14, covering default advance, unrelated history, exact/substituted/refused candidate, affected/unknown/not-affected advisory, advisory-only binding, malformed/rate-limit/timeout, Locked offline, path escape, release source binding, and canonical serialization. | Tools/Tests/GhosttyRuntimeStatus.Tests.ps1 |
| P3.7 | 2026-08-25 | working tree at 004b9231c8b0e3fb8a3af613e0b4a856ce470171 | Added full-SHA-pinned read-only workflow with daily advisory-only, monthly full Online, and manual Locked/Online/Candidate dispatch; source contracts prove contents/actions read permissions and no issue/PR/branch/lock-write path. | .github/workflows/ghostty-runtime-status.yml; Tools/Tests/ReleaseWorkflowPolicy.Tests.ps1 |
| P3.8 | 2026-08-25 | working tree at 004b9231c8b0e3fb8a3af613e0b4a856ce470171 | Release preflight now runs Online before build, requires checked-out HEAD, recomputes and compares canonical lock SHA-256, warns on update-observed, and fails closed on security-review, required-manual, or unreachable; source contract passed. | .github/workflows/release.yml; Tools/Tests/ReleaseWorkflowPolicy.Tests.ps1 |
| P3.9 | 2026-08-25 | working tree at 004b9231c8b0e3fb8a3af613e0b4a856ce470171 | Final status/release/governance/tool-inventory Pester cohort passed 63/63; tool inventory classified 177/177 with 0 findings; spec inventory reported 0 blocking findings; git diff --check exited 0. | Tools/Tests; Tools/tool-inventory.json; .build/SpecInventory/spec-inventory.json |
| P3.10 | 2026-08-25 | 43a6bf6214126141b0b2f86f4968b74771d64d07 | Required phase boundary commit created with exact subject `Add Ghostty update observation`; `git rev-parse HEAD` and `git show -s --format=%s` matched the recorded SHA and subject. | Git history and Required phase commits table |
| P4.1 | 2026-08-25 | working tree at 2d478ce57cac7921cb662eba5dd3be1fe13f4558 | Authenticated official Ghostty commit 9c3ec931d64561a8407dde7ac984ce156ae91539, tree 359872ab74fe1ecc96428133c0546be3cb969fdd, version 1.3.2-dev, pristine official origin, and GitHub verification `valid`; two independent official archives matched SHA-256 ece97aa2b3c6afe638849a67bc56073207d5599af3b1dedb53bf7990f1930fe8. | .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/candidate-review.json |
| P4.2 | 2026-08-25 | working tree at 2d478ce57cac7921cb662eba5dd3be1fe13f4558 | `merge-base --is-ancestor` exited 0 and exact `rev-list --left-right --count` was 0 behind / 717 ahead from the current product pin. | .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/candidate-review.json |
| P4.3 | 2026-08-25 | working tree at 2d478ce57cac7921cb662eba5dd3be1fe13f4558 | Regenerated the complete six-root path filter: 122 changed files, 40,419 insertions, and 4,984 deletions. Exact category counts are allocation/memory 5, build/toolchain 4, clipboard/security 12, Kitty graphics 18, parser/state 42, public ABI 16, and snapshot/serialization 25. | .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/candidate-review.json |
| P4.4 | 2026-08-25 | working tree at 2d478ce57cac7921cb662eba5dd3be1fe13f4558 | Reviewed every public header and declaration: 185 to 195 declarations (23 added, 13 removed, 0 signature changes) and 465 to 514 explicit enum values (49 added, 0 removed/changed). All 72 loader-required names were checked; only `ghostty_terminal_mode_get` is absent, with reviewed replacement `ghostty_terminal_get`. Candidate-review plus lifecycle/status/governance/inventory Pester passed 47/47. | Tools/Tests/GhosttyCandidateReview.Tests.ps1; .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/candidate-review.json |
| P4.5 | 2026-08-25 | working tree at 2d478ce57cac7921cb662eba5dd3be1fe13f4558 | Bounded primary-source review covered 220 path-relevant commits and queries for allocation/pointer bounds (153), clipboard/paste (36), Kitty images (96), OSC/DCS/APC parsing (45), and snapshot serialization (75). Six curated findings record mode, clipboard bounds, Kitty lifecycle/resource hardening, parser bounds, and snapshot ownership effects without unsupported vulnerability labels. | .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/candidate-review.json |
| P4.6 | 2026-08-25 | working tree at 2d478ce57cac7921cb662eba5dd3be1fe13f4558 | Published schema-valid canonical untracked evidence under the exact candidate root. Review SHA-256 is da748ac962a42e69f394fe47e09a348e783a07cf777b50ccbb1cdbd5cda5e429; descriptor SHA-256 is 8f8f1faaf55de2b00f5a0f7e03c06829d3469d3cb9cdddc7d4970381a22cd1d2; observation SHA-256 is 9bda5c7c7b39f806f11c009d7aa52ab8eaceb397a6b6f364a52176cf03afb180; derived input digest is 8b5a2c87ff4d726a9d4954cbefdf014412b1343644565cd089ef3ce2b71faff5. Tool inventory classified 180/180 with 0 findings, spec inventory had 0 blocking findings, and `git diff --check` exited 0. | .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/candidate-review.json; .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/candidate-descriptor.json; .build/SpecInventory/spec-inventory.json |
| P4.7 | 2026-08-25 | 2347ddbb5835bec235e9de5ce0f37d262a5e0934 | Required phase boundary commit created with exact subject `Freeze Ghostty runtime candidate`; `git rev-parse HEAD` and `git show -s --format=%s` matched the recorded SHA and subject. | Git history and Required phase commits table |
| P5.1 | 2026-08-26 | working tree at 585b04f69eb12bbd2b71fdadc2e4574441bde74e | Reviewed all seven current concerns. Shared-only and reproducible-link concerns are retained; placement accounting/capacity are replaced by candidate-aware admission; retained-image accounting, virtual-placeholder snapshots, and compression activity are removed from the private overlay because the candidate now owns those behaviors. Every disposition names source evidence and focused tests. | .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/candidate-overlay-concerns.json; .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/candidate-overlay-review.json |
| P5.2 | 2026-08-26 | working tree at 585b04f69eb12bbd2b71fdadc2e4574441bde74e | Reimplemented Windows-only shared lib-vt emission while retaining upstream static behavior on other targets; the private Windows build explicitly uses LLVM/LLD and stripping. | .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/Ghostty-WindowsPrivateRuntime.patch |
| P5.3 | 2026-08-26 | working tree at 585b04f69eb12bbd2b71fdadc2e4574441bde74e | Candidate public terminal values end at 39 and private placement limits use 40; public Kitty data values end at 2 and private values use 3 through 6. Compile-time Zig/C parity guards and collision-negative Pester fixtures passed. | .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/candidate-overlay-review.json; Tools/Tests/GhosttyCandidateOverlay.Tests.ps1 |
| P5.4 | 2026-08-26 | working tree at 585b04f69eb12bbd2b71fdadc2e4574441bde74e | Candidate storage validates all requested screens and placement count/allocation before reap, map, generation, or pin mutation; preserves the 4,096/1 MiB limits across reset; releases replaced pins; and retains exact screen/query behavior. Upstream storage tests cover count, bytes, rejection-state equality, and reset preservation. | .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/Ghostty-WindowsPrivateRuntime.patch |
| P5.5 | 2026-08-26 | working tree at 585b04f69eb12bbd2b71fdadc2e4574441bde74e | Generated one full-index normalized 23,705-byte patch with SHA-256 3f60d9436cec65521f2d29f1810d3308e49eb06abe95f48af6e0e0ea6eea8674. The ordered five-concern review has SHA-256 a96ca7a212841d9f2ea36b2bca2d45633ab3808f334fb42380bee9dc507664f7; the overlay-bound descriptor has SHA-256 e1a5bb3b18ca8da6f943d849c06eac35dd9c0f0514a51ef6556f359ed746ad3d and input digest 061a17c43e24401a7c9af7db5cab3206a69d58d9c42a8b6bcb4a53e464bc0651. A same-time rerun reproduced every hash and preserved the P4 descriptor at SHA-256 8f8f1faaf55de2b00f5a0f7e03c06829d3469d3cb9cdddc7d4970381a22cd1d2. | .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/candidate-overlay-review.json; .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/candidate-descriptor.json; .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/candidate-descriptor.p4.json |
| P5.6 | 2026-08-26 | working tree at 585b04f69eb12bbd2b71fdadc2e4574441bde74e | `git apply --check`, independent apply, `git diff --check`, and exact regenerated-patch comparison passed. Candidate `zig build test-lib-vt -Dtarget=x86_64-windows-msvc -Doptimize=Debug -Dsimd=false` passed. Overlay/lifecycle/source-contract/governance/inventory Pester passed 84/84; tool and spec inventories reported zero blocking findings; repository `git diff --check` exited 0. | .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/Ghostty-WindowsPrivateRuntime.repeat.patch; Tools/Tests/GhosttyCandidateOverlay.Tests.ps1; .build/SpecInventory/spec-inventory.json |
| P5.7 | 2026-08-26 | working tree at 585b04f69eb12bbd2b71fdadc2e4574441bde74e | Measured overlay is 7 files, 343 insertions plus 48 deletions (391 changed lines), and five independent concerns, within all mandatory review limits. | .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/candidate-overlay-review.json |
| P5.2–P5.7 corrective revalidation | 2026-08-26 | working tree after 08329cbe7a153edcb97ba92f7fd8b2471b2c6fa7 | P6 security review proved callback metadata cannot distinguish unnamed Kitty OSC 5522 from OSC 52. The candidate-only static-skip optimization was removed and replaced, within the frozen five-concern ceiling, by private option 41 that returns `ENOSYS` before Kitty buffering/callback while leaving OSC 52 unaffected. The regenerated patch is 7 files, 387 insertions, 27 deletions, SHA-256 `06847f6594427cc584cef0f0ec66a54080f236b2414fe40e61b381accaf38264`; schema-v2 overlay review enumerates both private terminal options 40/41 and all four private Kitty data values. Clean apply, exact regenerated diff, focused Zig test, schema validation, and focused Pester 5/5 passed. | .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/candidate-overlay-review.json; .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/Ghostty-WindowsPrivateRuntime.patch; Tools/Tests/GhosttyCandidateOverlay.Tests.ps1 |
| P5.8 | 2026-08-26 | 551a0a6ef54e1a9e41430f33fced6427f146cfa3 | Required phase boundary commit created with exact subject `Rederive Ghostty Windows runtime overlay`; `git rev-parse HEAD` and `git show -s --format=%s` matched the recorded SHA and subject. | Git history and Required phase commits table |
| P6.1 | 2026-08-26 | bc402bd8535f72adf7cc82fe47430198d9f31866 | Captured SHA-256 for the x64/ARM64 Product DLLs, import libraries, runtime identities, and generated pairing headers plus the canonical Product lock and overlay before candidate execution. The snapshot contains eight paired files and has SHA-256 2f042a139d5dbc4eea8d3adf9e8f6917b8344941bc3f47c71554e38cde72e20b. | .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/product-pair-before.json |
| P6.2 | 2026-08-26 | bc402bd8535f72adf7cc82fe47430198d9f31866 | Candidate-only Force builds exited 0 under isolated x64 and ARM64 roots. x64 is AMD64/2,141,184 bytes/raw 1359467b6eb4610d7e42801aa62c6997c60c58e76c1336e1c8354d134a109aec; ARM64 is ARM64/1,881,600 bytes/raw 06a1d5116ef19725b69a8c64b477b35c58d860dcc37e0e135733a68b557e5d79. Both have 191 exports and export-set SHA-256 85d134a5fa8c29d83f9c22559632600d474eb433c12a8bd3d55da9863ddeef15. | .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/candidate-qualified/x64/terminal-runtime-identity.json; .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/candidate-qualified/ARM64/terminal-runtime-identity.json |
| P6.3 | 2026-08-26 | bc402bd8535f72adf7cc82fe47430198d9f31866 | The post-build comparison found all eight Product/pairing files byte-identical to the pre-build snapshot; the Product lock and tracked overlay hashes were also unchanged. Post snapshot SHA-256 is aaf19e295d3f18e40b06926e93c4d6aab41cce31e04f6f61f0ef2fa361736f36. | .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/product-pair-before.json; .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/product-pair-after.json |
| P6.4 | 2026-08-26 | bc402bd8535f72adf7cc82fe47430198d9f31866 | Candidate `zig build test-lib-vt test-lib-vt-build -Dtarget=x86_64-windows-msvc -Doptimize=Debug -Dsimd=false` passed with 6,225 executed and 100 skipped tests. After supplying isolated bundled Python/jsonschema through the candidate host-tools/cache roots, `test-lib-vt-schema` passed and validated the x86_64-windows ABI manifest with 153 types. | .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/host-tools; .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/python-packages; .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/candidate-qualification.json |
| P6.4 invalidated | 2026-08-26 | working tree after 08329cbe7a153edcb97ba92f7fd8b2471b2c6fa7 | The corrective option-41 candidate patch changed upstream C/Zig test source, so the earlier full x64 result no longer closes P6.4. The item was unchecked pending a complete exact-target rerun. | Live checklist P6.4; corrective patch SHA-256 `06847f6594427cc584cef0f0ec66a54080f236b2414fe40e61b381accaf38264` |
| P6.4 corrective rerun | 2026-08-26 | working tree after 08329cbe7a153edcb97ba92f7fd8b2471b2c6fa7 | `zig build --seed 0x97d35bc7 -j1 --cache-dir ../x64-upstream-cache-corrective test-lib-vt test-lib-vt-build test-lib-vt-schema -Dtarget=x86_64-windows-msvc -Doptimize=Debug -Dsimd=false` exited 0 from the exact corrective overlay; schema output validated 153 x86_64-windows ABI types. A first attempt using the stale prior cache failed before tests because its cached `uucode_generate.exe` was absent; the fresh contained cache removed that local artifact and the complete unchanged command graph passed. | .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/x64-upstream-cache-corrective; .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/Ghostty-WindowsPrivateRuntime.patch |
| P6.5 pre-test correction | 2026-08-26 | bf473320fed3cf8f4496e32f3cb7bae7103bb49d | Native workflow run 32920173950 passed the exact corrective ARM64 candidate build and stopped before launching any upstream target. The pre-test cleanliness check rejected the same two Ghostty fuzz-corpus paths that the builder deliberately removes because Windows cannot materialize them. The workflow now compares pre-test and post-cleanup status to that exact two-path allowlist and still fails every added, modified, untracked, or other deleted path. | https://github.com/DualTail/RedSalamander/actions/runs/32920173950; .github/workflows/ghostty-runtime-native-qualification.yml |
| P6.5 bounded first-launch retry | 2026-08-26 | c2b8c9e20c394680142fd68e2f1635710c06e186 | Native workflow run 32920756183 passed the exact ARM64 candidate build, validated the 153-type ARM64 ABI schema, and completed one suite with 2,904 passed, 50 skipped, and 0 failed. The sibling newly linked Debug test executable returned process code 5 before emitting test output; no matching Application Error event existed. Because serialized execution and the simple runner did not remove this first-launch condition, the workflow now authenticates the complete one-code-5/one-zero-failure-suite/schema-pass/no-FAIL shape, resolves and hashes the exact cache-contained executable, waits two seconds, and permits exactly one same-seed retry. Any other shape or retry failure remains fatal. | https://github.com/DualTail/RedSalamander/actions/runs/32920756183; .github/workflows/ghostty-runtime-native-qualification.yml; Specs/Terminal/GhosttyRuntimeNativeQualification.schema.json |
| P6.5 retry crash | 2026-08-26 | 3f1bab6b514ba3b1b68e267fbfcb82783930a771 | Native workflow run 32923375705 passed the exact corrective ARM64 candidate build and 153-type ABI schema. The Zig suite reported 2,904 passed, 50 skipped, and 0 failed. The sibling C-ABI test executable first returned code 5 before output; its authenticated exact-image retry exited `-1073741819` (`0xC0000005`) with zero summaries and zero reported test failures. This is crash evidence, not a passing retry, so P6.5 remains blocked while a WER dump and exact executable are captured. | https://github.com/DualTail/RedSalamander/actions/runs/32923375705; Specs/TestRuns/GitHubARM64/Continuation/20260826_030920_ghostty_native_test_access_violation/README.md |
| P6.5 retained crash and repair | 2026-08-26 | working tree after 53c1efe080c2228eda149ffec7b8fd9129ca7e35 | Native workflow run 32925672178 reproduced `0xC0000005` and retained the exact 55,152,640-byte stripped C-ABI `test.exe`, SHA-256 `58885dc091476573ddf69b1843f14c955857436f946121291ddc5f8fc26bbba9`; WER produced no dump or event. Source-graph inspection proved `GhosttyLibVt.zig` mutates the reused `mod.vt_c` to `strip=true` for the shared runtime before `build.zig` links the Debug C-ABI test. A test-only module clone with strip disabled, frame pointers retained, and synchronous unwind tables produced a 66,208,768-byte ARM64 test plus PDB while leaving the reviewed overlay and production identities unchanged. The native workflow and evidence schema now encode that diagnostic-only clone; final native execution remains required. | https://github.com/DualTail/RedSalamander/actions/runs/32925672178; Specs/TestRuns/GitHubARM64/Continuation/20260826_030920_ghostty_native_test_access_violation/README.md; .github/workflows/ghostty-runtime-native-qualification.yml; Specs/Terminal/GhosttyRuntimeNativeQualification.schema.json |
| P6.5 native pass | 2026-08-26 | 3d36d9f3f9d7c1bb61c31b8eb29067fd313ce6c7 | Native ARM64 workflow run 32930638145 passed in 47m41s. The exact candidate build and `test-lib-vt`, `test-lib-vt-build`, and `test-lib-vt-schema` all passed; the canonical report validated source commit `3d36d9f3...`, the exact candidate/tree/lock/overlay, native Arm64 OS/process/Python, the test-only unstripped module clone, ARM64 runtime identity SHA-256 `2bd9dfa4a3408d60a1279d69fe5c1276687061a4e515c4705332db33a9b02da6`, 191 exports, and `codeFiveRetryUsed=false`. Report SHA-256 is `2176df6545c35c1560284f82af8aa08e5a610cf12253d92522b6a6d917bef72d`. All five temporary GitHub candidate variables/secrets were deleted and verified absent after download. | https://github.com/DualTail/RedSalamander/actions/runs/32930638145; .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/native-run-32930638145/native-arm64-qualification.json |
| P6 closeout validation | 2026-08-26 | working tree after 3d36d9f3f9d7c1bb61c31b8eb29067fd313ce6c7 | Fresh-cache x64 `test-lib-vt`, `test-lib-vt-build`, and `test-lib-vt-schema` passed with the fixed seed and 153-type schema validation. Lock-driven x64 closeout build passed with 191 exports and the frozen import/dependency/notice closure. Deterministic overlay review reproduced review SHA-256 `0f829fed06b6e866328158ded318b9aaaa9e567e30cdd942884615c8505973b5`, descriptor SHA-256 `bfd336bd7e7acf9968f7597458668b98e19ccc6ca453f02d46c6d6f80bed90e1`, and input digest `7890aede47c536ab2232b6fa71b8820dc1772d77cf7b81c8f18ba40a8972d93c` twice. All eight Product/pairing files remained byte-identical. Final candidate qualification report SHA-256 is `35dd3a5c37e5ec705f9392049614ea79f03c49171d52c0ede07ddaf866533594` with `releaseReady=true`; focused Pester passed 178/178, tool inventory classified 183/183 with zero findings, spec inventory reported zero blocking findings, and `git diff --check` exited 0. | .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/candidate-qualification.json; .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/x64-upstream-cache-final-p6; .build/SpecInventory/spec-inventory.json |
| P6.8 | 2026-08-26 | 90e277a6e72da1ddab4726b41b3f1488afbde2e6 | Required phase boundary commit created with exact subject `Qualify isolated Ghostty candidate`; `git rev-parse HEAD` and `git show -s --format=%s` matched the recorded SHA and subject. | Git history and Required phase commits table |
| P6.5 blocked | 2026-08-26 | bc402bd8535f72adf7cc82fe47430198d9f31866 | Host probes report OS and process architecture X64/AMD64. `zig build test-lib-vt-build -Dtarget=aarch64-windows-msvc -Doptimize=Debug -Dsimd=false` passed as a cross-compile, but this host cannot execute the resulting ARM64 DLL/tests or satisfy the native schema/load gate. Native ARM64 Windows evidence is required; cross-compilation is intentionally not accepted as a substitute. | .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/candidate-qualification.json |
| P6.6 | 2026-08-26 | bc402bd8535f72adf7cc82fe47430198d9f31866 | Recorded the exact 191-name export set/hash, 20 additions, 9 removals, 71/72 current loader names present, and the Phase-7 replacement of missing `ghostty_terminal_mode_get` with `ghostty_terminal_get` plus the frozen two-field `GhosttyTerminalModeConfig`. Both architectures import the same six modules; candidate Wuffs compilation with Windows `link_libc` explains the added UCRT/VCRuntime closure and most size growth. Build info remains SIMD false, Kitty true, tmux false, ReleaseFast, lib-vt 0.1.0-dev. Exact translate-c/uucode/Wuffs dependency and six-notice identities are locked. Qualification report SHA-256 is 1cf03063f9cdb5f65e48159df74aa62fcd21b9eb59af72302b6d760ab3f69244. | .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/candidate-runtime-lock.v1.json; .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/candidate-qualification.json |
| P6.7 | 2026-08-26 | bc402bd8535f72adf7cc82fe47430198d9f31866 | Proposed lock file SHA-256 175cfed668d3f3ad748c732352e9825dd63554707d47fb00f0c312d898a93fb9 and canonical digest ae209cc272dd46c219334e8671bfcdde71b34ec6d13886659d8afb28b5433105 passed schema and lifecycle validation. Clean lock-driven x64 and ARM64 builds passed all exact machine/size/import/export/dependency/notice checks. A separate x64 `-Offline -Force` build then non-Force `-Offline` reuse passed without network and retained raw SHA-256 6f1ba61e45e4fa854d733cff118fa717962b999d5acb2f765693747c25f1dbad byte-for-byte; its alternate valid Zig/LLD code layout is covered by the generated pairing boundary and the documented no-source-level-DLL-hash rule. Final focused Pester passed 127/127; tool inventory classified 183/183 with zero findings; spec inventory had zero blocking findings; `git diff --check` exited 0. | .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/candidate-runtime-lock.v1.json; .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/candidate-qualified-offline/x64/terminal-runtime-identity.json; .build/SpecInventory/spec-inventory.json |
| P6.2/P6.7 corrective exact-lock revalidation | 2026-08-26 | working tree after 08329cbe7a153edcb97ba92f7fd8b2471b2c6fa7 | Corrective lock SHA-256 `d1bad7bd545164b94f4cf6d10c3143fa5e25db6d0bd62ffa6aa36fecfbcc9078`, canonical digest `6409f257f4efaf40ea8e33e92a57e75edfb670430731576f9658b419a44ce331`, and patch SHA-256 `06847f6594427cc584cef0f0ec66a54080f236b2414fe40e61b381accaf38264` passed schema/lifecycle validation. Exact-lock clean x64 and ARM64 builds passed at unchanged sizes 2,141,184/1,881,600 with 191 exports and export-set SHA-256 `85d134a5fa8c29d83f9c22559632600d474eb433c12a8bd3d55da9863ddeef15`; x64 raw/normalized SHA-256 are `cb84ae8fc116e2721d0e1b1d82396c8b898c5364ddb98faf47434ddc22048088`/`a65a45989fb75954489538ff37a98e0722022c08798c5c0a0c2ab0d775dc2599`, ARM64 raw/normalized SHA-256 are `425c23fb45e8685e497b2c8f1d99bc5773bef86aa2d4c3fb84a6e5babc3bc906`/`fe61836dfaa32388a297971d26f7452a2177944cc27f97c2e44e7e6eae2f887f`. | .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/candidate-runtime-lock.v1.json; .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/candidate-qualified-final/x64/terminal-runtime-identity.json; .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/candidate-qualified-final/ARM64/terminal-runtime-identity.json |
| P7.1/P7.2 | 2026-08-26 | working tree after 4b8c4d4f6bc1727f102f8e1bd4f7cf1dd9f32b56 | Added explicit Product/Candidate header lanes. The candidate loader translation unit compiled against the qualified candidate headers and resolves mode data only through `ghostty_terminal_get`; the default Product build retains its canonical lock and legacy required export until P10. Conditional resolution/reset remains inside the existing all-or-nothing load/unload path. | Plugins/Terminal/Terminal.vcxproj; Plugins/Terminal/TerminalRuntimeLoader.h; Plugins/Terminal/TerminalRuntimeLoader.cpp |
| P7.3 | 2026-08-26 | working tree after 4b8c4d4f6bc1727f102f8e1bd4f7cf1dd9f32b56 | Candidate compile-time guards prove `GhosttyTerminalModeConfig` is the frozen non-sized standard-layout record and that `value` immediately follows `mode`. Source contracts enumerate every versioned Ghostty adapter record and require its `size` initialization; focused source-contract Pester passed 34/34. | Plugins/Terminal/TerminalRuntimeLoader.cpp; Tools/Tests/TerminalPluginSourceContracts.Tests.ps1 |
| P7.4 | 2026-08-26 | working tree after 4b8c4d4f6bc1727f102f8e1bd4f7cf1dd9f32b56 | Debug x64 Product build completed with 0 warnings and 0 errors. Focused Terminal tests passed 102/102, including key encoding, paste encoding, mouse encoding, selection, and new exact false/true/false assertions for cursor-key, normal/SGR mouse, focus, alternate-screen-save, and bracketed-paste modes. A separately compiled candidate ABI probe repeated all six mode transitions against the qualified candidate DLL and exited 0. | .build/x64/Debug/PluginContractTests.exe; .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/candidate-qualified-p6-closeout/bin/candidate-mode-probe.exe |
| P7.5 | 2026-08-26 | working tree after 4b8c4d4f6bc1727f102f8e1bd4f7cf1dd9f32b56 | Added a candidate-only host regression whose DCS payload contains `0xC3 0x9B` followed by bytes that would become a live CSI under the old parser. The direct candidate probe formatted only the post-ST `Y` and no payload `X`, exiting 0. The frozen corpus also retains explicit `0x80`, `0xFE`, and `0xFF` coverage, and P6 passed the candidate's exact upstream DCS table regression on x64 and native ARM64. | Plugins/Terminal/Terminal.cpp; Plugins/Terminal/TerminalVtUpgradeTests.cpp; .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/p7-candidate-mode-probe.cpp |
| P7.6 | 2026-08-26 | working tree after 4b8c4d4f6bc1727f102f8e1bd4f7cf1dd9f32b56 | Candidate loader code, security call sites, candidate compile, and runtime probe no longer require `ghostty_terminal_mode_get`. The reviewed phase amendment preserves that name only inside the default Product compatibility branch and canonical Product lock until P10 admission. Source/export Pester passed 35/35 and `git diff --check` exited 0. | Plugins/Terminal/TerminalSecurity.cpp; Tools/Tests/TerminalPluginSourceContracts.Tests.ps1; Tools/Tests/TerminalRuntimeExportIdentity.Tests.ps1; reviewed P7 staged-admission amendment in this plan |
| P7.7 | 2026-08-26 | 39f3ed0888912531a6ad7d621896c0654c5fe620 | Required phase boundary commit created with exact subject `Adapt Terminal to Ghostty mode ABI`; `git rev-parse HEAD` and `git show -s --format=%s` matched the recorded SHA and subject. | Git history and Required phase commits table |
| P8.1 | 2026-08-26 | working tree after 186eb1ffa18951d6e0e94162222667ab0e8f06bb | Added Product/Candidate callback lanes around one bounded handler. The candidate lane validates that the sized request contains the final `reply` field before any processing, emits one zero-initialized sized reply with `remember=false`, and returns without processing when the field or function is absent. The complete Terminal project compiled cleanly against the qualified candidate headers. | Plugins/Terminal/Terminal.h; Plugins/Terminal/TerminalSecurity.cpp; Plugins/Terminal/TerminalInternal.h; .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/p8-candidate-identity |
| P8.2 | 2026-08-26 | working tree after 186eb1ffa18951d6e0e94162222667ab0e8f06bb | Candidate initialization now explicitly sets clipboard-read null, the public decoded-write ceiling to 256 KiB, and private Kitty clipboard option 41 to false before input. Any failed set returns initialization failure. Candidate compile and the exact-runtime probe both accepted all three settings. | Plugins/Terminal/Terminal.cpp; .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/p8-candidate-clipboard-probe.cpp |
| P8.3 | 2026-08-26 | working tree after 186eb1ffa18951d6e0e94162222667ab0e8f06bb | Extracted sized request validation into the ABI-specific Terminal helper, retained only a bounded owned UTF-8 copy, and preserved deny/allow/ask plus helper-owned UI-thread posting and teardown drain. Source contracts prove no clipboard/UI API occurs in the synchronous handler. | Plugins/Terminal/TerminalInternal.h; Plugins/Terminal/TerminalSecurity.cpp; Tools/Tests/TerminalPluginSourceContracts.Tests.ps1; Specs/Terminal/Terminal_EmbeddedPlugin.md |
| P8.4/P8.5 | 2026-08-26 | working tree after 186eb1ffa18951d6e0e94162222667ab0e8f06bb | The direct exact-candidate probe exited 0 after proving OSC 52 and OSC 1337 each invoke and receive exactly one successful reply, disabled Kitty writes return one `ENOSYS` without callback or buffering, trailing write packets remain ignored, OSC 52 reads do not invoke the write callback, and a null Kitty read returns `EPERM`. | .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/p8-candidate-clipboard-probe.cpp; .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/p8-candidate-clipboard-probe.exe |
| P8.6 | 2026-08-26 | working tree after 186eb1ffa18951d6e0e94162222667ab0e8f06bb | Product Debug x64 completed with 0 warnings/0 errors and Terminal selftests passed 108/108. Focused source-contract Pester passed 34/34. Tracked deterministic cases cover undersized and missing-reply records, null data, 256 KiB minus/exact/plus boundaries, invalid UTF-8, consent structure, reply-once, teardown drain, Kitty denial, and null read. A complete candidate adapter build succeeded; its ahead-of-P9 full run passed 114/115, while the isolated P8 protocol probe passed every P8 candidate behavior, leaving the separate Kitty image-lifecycle adaptation assigned to P9 unchecked. | .build/logs/msbuild-20260826_080115_841-pid62664-83a6060a.log; .build/x64/Debug/PluginContractTests.exe; Tools/Tests/TerminalPluginSourceContracts.Tests.ps1; .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/p8-candidate-fixture |
| P8.7 | 2026-08-26 | 3e91f6e8bb106cba20cbc8e99e3d2bf660a7f16d | Required phase boundary commit created with exact subject `Adapt Ghostty clipboard security contract`; `git rev-parse HEAD` and `git show -s --format=%s` matched the recorded SHA and subject. | Git history and Required phase commits table |
| P9.1 | 2026-08-26 | working tree after b950db2f3ef243516f3524bc374f3d184d7767f6 | Split engine-owned `NO_VALUE` pending from adapter conversion-pending state, retained distinct ready/deferred/rejected outcomes, and added cumulative superseded/removed plus per-capture engine-pending metrics. The shared query helper maps only `GHOSTTY_NO_VALUE` to engine-pending; other failures reject without retaining an upstream pointer. | Plugins/Terminal/Terminal.h; Plugins/Terminal/TerminalVt.cpp; Tools/Tests/TerminalPluginSourceContracts.Tests.ps1 |
| P9.2/P9.3 | 2026-08-26 | working tree after b950db2f3ef243516f3524bc374f3d184d7767f6 | Storage-generation change invalidates the latest pipeline request and resets both engine-pending and conversion-pending entries for a fresh query. Per-image generation remains the exact cache identity; a visible new generation erases every older generation for that image ID before work admission. Result merge requires the current storage generation, exact key, conversion-pending state, and request ID, so deleted or superseded content cannot be resurrected. | Plugins/Terminal/TerminalVt.cpp; Specs/Terminal/Terminal_EmbeddedPlugin.md |
| P9.4 | 2026-08-26 | working tree after b950db2f3ef243516f3524bc374f3d184d7767f6 | Existing latest-only queue invalidation, one `std::jthread` worker, cancellation, bounded 64 MiB source/converted/cache accounting, nonvisible LRU, Close stop/join ordering, and resize/no-reupload behavior remain on the same production path. Focused source contracts passed 34/34 and require the distinct pending helpers, exact-key supersession, stale-result guard, metrics, one-worker queue, and quiet point. | Plugins/Terminal/TerminalKittyImagePipeline.cpp; Tools/Tests/TerminalPluginSourceContracts.Tests.ps1 |
| P9.5 | 2026-08-26 | working tree after b950db2f3ef243516f3524bc374f3d184d7767f6 | Added production-path state-machine checks for `NO_VALUE`, storage retry, error rejection, and request mismatch plus a visible same-ID/same-size retransmission that replaces the explicit placement, waits through any engine-pending interval, proves a new generation, removes the old cache key, retains unrelated ready content, and completes conversion. Final Product Debug passed 114/114; the full candidate adapter/runtime fixture passed 121/121. P6 exact upstream x64/native-ARM64 qualification already covers candidate pending completion, replacement, delete, eviction, retransmission, animation frame generation, reset, and screen behavior. | Plugins/Terminal/Terminal.cpp; .build/x64/Debug/PluginContractTests.exe; .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/p8-candidate-fixture; P6.4/P6.5 evidence |
| P9.6 | 2026-08-26 | working tree after b950db2f3ef243516f3524bc374f3d184d7767f6 | Final Product Debug x64 build completed with 0 warnings/0 errors and attestation `4cbcd296d3f7c9cc6569761828079aaf606a6c7434e9dc2cb517b771374277c3`; Product Release x64 completed with 0 warnings/0 errors and attestation `25d07352427308069162783c2cc73f62d2be7afd71343f4e59a7a1bbcd5c2ebe`. Candidate Debug selftests passed 121/121 and candidate Release x64 compiled with Terminal SHA-256 `267302e7405011fa2d710e2b6da78f130fc99c947ac5f320ffd8f6af940b00c6`. Placement 4,096/1 MiB limits, replacement pin release, cache/pipeline ceilings, stale-result rejection, zero-byte Close, resize reuse, and zero reupload remain covered. Canonical Product Debug/Release DLLs were restored with SHA-256 `a829cb757c7af478846017cd2b57997de288b3d3cd5276836621a4cb981bd25d` and `ba65c99736feff58108d555abad38cc1c73622217acb392f2bab861d013421bd`. | .build/logs/msbuild-20260826_085120_590-pid106772-d3cb068b.log; .build/logs/msbuild-20260826_085348_825-pid43448-929328f6.log; .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/p9-product-restore; .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/p9-product-release-restore |
| P9.7 | 2026-08-26 | 655e45ac2a0f6fc01f51233a1b72fb0ef3c3f615 | Required phase boundary commit created with exact subject `Handle pending Ghostty Kitty images`; `git rev-parse HEAD` and `git show -s --format=%s` matched the recorded SHA and subject. | Git history and Required phase commits table |
| P10.1 | 2026-08-26 | working tree after d6a30f0ae3aea62d39c19401dbaa11334d558b2f | Built isolated test-enabled Release x64 Product and candidate fixtures from the same checkout, configuration, machine, compiler, SDK, Zig, vcpkg, and frozen 1,048,965-byte corpus. Because standalone hosted-render timings exposed environmental noise, ran three immediately interleaved Product/candidate pairs instead of selecting one favorable standalone result. | .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/p10-product-release; .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539/p10-candidate-release; Specs/TestRuns/4cb089111a23/Terminal/20260826_071800_pair1-product; Specs/TestRuns/4cb089111a23/Terminal/20260826_071807_pair1-candidate; Specs/TestRuns/4cb089111a23/Terminal/20260826_071812_pair2-product; Specs/TestRuns/4cb089111a23/Terminal/20260826_071820_pair2-candidate; Specs/TestRuns/4cb089111a23/Terminal/20260826_071825_pair3-product; Specs/TestRuns/4cb089111a23/Terminal/20260826_071832_pair3-candidate |
| P10.2 | 2026-08-26 | working tree after d6a30f0ae3aea62d39c19401dbaa11334d558b2f | Corrected the semantic snapshot contract to exclude Ghostty-private placement allocation bytes/capacity, which are ABI storage rather than terminal behavior. Added an untimed same-corpus control that omits only the named DCS step. Product and candidate controls then matched exactly: PTY SHA-256 `e8f76ba921507c4df938eccff8d84d067f74c648ec2840b4efcff35c6ed15271` and snapshot SHA-256 `d737da86401710c72989649fec770ac53bcd54b6f6be373f94dcb1a96980b9ac`; the primary digests differ only when `dcs-high-byte-preservation` is present. The evidence schema keeps the additive control optional so the frozen P0 archive remains valid. | Plugins/Terminal/TerminalVtUpgradeTests.cpp; Plugins/Terminal/TerminalVtUpgradeTestContract.h; Tests/PluginContractTests/PluginContractTests.cpp; Specs/Terminal/TerminalVtUpgradeEvidence.schema.json; Specs/Terminal/TerminalVtUpgradeCorpus.v1.json; Specs/Terminal/Terminal_EmbeddedPlugin.md |
| P10.3 | 2026-08-26 | working tree after d6a30f0ae3aea62d39c19401dbaa11334d558b2f | Each lane contains 200 nearest-rank samples. Pair 1 Product/candidate VT-write, snapshot, and hosted-render p95 values were `11434/996`, `4015/3651`, and `7252/7356` microseconds (ratios 0.0871, 0.9093, 1.0143); pair 2 was `11470/1033`, `4111/3870`, and `7384/7301` (0.0901, 0.9414, 0.9888); pair 3 was `12031/1127`, `4126/3810`, and `8773/7093` (0.0937, 0.9234, 0.8084). All are at or below 1.10. Every archive also passed Kitty pending/completion, bounded memory, zero queued/active/ready/resident bytes at close, and the fixed placement/cache/resource assertions. Pair 2 is the designated admission comparison; all three pairs remain archived to show the observed variance. | Six P10 run directories under Specs/TestRuns/4cb089111a23/Terminal; terminal evidence results and perf_metrics.jsonl |
| P10.4 | 2026-08-26 | working tree after d6a30f0ae3aea62d39c19401dbaa11334d558b2f | Completed manual admission review across authenticated upstream commit/tree/archive provenance, the bounded source inventory, all five overlay concerns and patch SHA-256 `06847f6594427cc584cef0f0ec66a54080f236b2414fe40e61b381accaf38264`, the 191-export/six-import ABI and mode replacement, 256 KiB clipboard bound/null-read/Kitty-denial/reply-once behavior, pending Kitty lifecycle and quiet point, exact upstream x64 plus native ARM64 qualification, translate-c/uucode/Wuffs dependency and notice closure, and the three-pair performance result. No unreviewed delta or unsupported security claim remains. | P4-P9 Evidence log; candidate qualification report; canonical lock and overlay; six P10 run directories |
| P10.5 | 2026-08-26 | working tree after d6a30f0ae3aea62d39c19401dbaa11334d558b2f | Admitted commit `9c3ec931d64561a8407dde7ac984ce156ae91539` and tree `359872ab74fe1ecc96428133c0546be3cb969fdd` in the canonical lock with the reviewed overlay, 191-name export set/hash, six-module import set, exact x64/ARM64 sizes, dependencies, and notices. Removed the temporary Product/candidate ABI branches and obsolete `ghostty_terminal_mode_get` dependency so the candidate mode/clipboard/Kitty contracts are now the only Product path. Added Wuffs notice staging and updated authoritative specs/source guards coherently. The complete focused admission Pester set passed 86/86 after replacing stale old-pin expectations with canonical-lock-derived status fixtures and a synthetic future candidate, and making the multi-dependency fixture additive to the now three-dependency Product lock. | External/TerminalEngine/GhosttyRuntimeLock.v1.json; Tools/TerminalEngine/Patches/Ghostty-WindowsPrivateRuntime.patch; Plugins/Terminal; RuntimeDependencies.props; Specs/Terminal/Terminal_EmbeddedPlugin.md; Tools/Tests/GhosttyRuntimeLifecycle.Tests.ps1; Tools/Tests/GhosttyRuntimeStatus.Tests.ps1; Tools/Tests/TerminalPluginSourceContracts.Tests.ps1 |
| P10.6 | 2026-08-26 | working tree after d6a30f0ae3aea62d39c19401dbaa11334d558b2f | Canonical x64 Force passed at 2,141,184 bytes with raw/normalized SHA-256 `78c92163f920c9e8cefe7b3acc35a7de3762da56131957648790b6982d23ae00`/`3711a6edd1b3d291e86c83804bd6ebba30ca026f3aca8c53f86b720ba1baf316`; native ARM64 Force passed at 1,881,600 bytes with `8fee8e63741c4d3410f0af31baaa9fb67f7f00a02c20483e1e7cff1c27caf09b`/`9847f3637bc7d21ebb7e3bd16b1a25efb3f075f5f889e315488bae47e581ce06`. x64 Offline Force reproduced the raw DLL exactly and Offline reuse passed. Locked status reported `current` with zero network attempts. Canonical Debug x64 then built with 0 warnings/errors, attestation `a6978d0d40cf6a360df044f4925c2b04b818e79daaa4a8cc29a152d3270e9694`, and passed 121/121 Terminal selftests. | .build/x64 and .build/ARM64 runtime identities; .build/TerminalEngineStatus/20260826T073011Z/status.json; .build/logs/msbuild-20260826_093033_610-pid113844-bb2b6c73.log; .build/x64/Debug/PluginContractTests.exe |
| P10.7 | 2026-08-26 | working tree after d6a30f0ae3aea62d39c19401dbaa11334d558b2f | The frozen current-pin archive `20260825_211846_current-pin` and all six exact P10 run IDs remained schema-valid. Each P10 Product/candidate archive independently passed Test-TestRunArchive.ps1 with all 18 required files accepted. The designated admission pair is `20260826_071812_pair2-product` versus `20260826_071820_pair2-candidate`; pairs 1 and 3 are retained as independent variance checks. | Specs/TestRuns/4cb089111a23/Terminal/20260825_211846_current-pin; six P10 run directories; Specs/Terminal/TerminalVtUpgradeEvidence.schema.json; Tools/Test-TestRunArchive.ps1 |
| P10.8 | 2026-08-26 | 0ff71930e1edcbd548677ceacea97f959820593f | Required phase boundary commit created with exact subject `Admit reviewed Ghostty runtime candidate`; `git rev-parse HEAD` and `git show -s --format=%s` matched the recorded SHA and subject. | Git history and Required phase commits table |
| P11 review remediation | 2026-08-26 | working tree after e52a0521e395ad92d881e6e9878612a7898f664f | Focused post-admission review confirmed premature clipboard `SUCCESS`, one-shot engine-pending Kitty queries, permanent same-key rejection, fail-open scheduled status outcomes, a wrong native artifact path, a release/native-gate gap, package-runtime authentication weakness, and the 1 MiB configuration/256 KiB engine-cap mismatch. The stale-result `TakeReady` claim was rejected because dropping superseded request IDs is intentional. The corrective Debug x64 Terminal build passed with 0 warnings/errors and attestation `ac54f2446b297b3f5daf40c864e7c27d74cb2cc7049b1b64b77774a5b1b4323a`; Terminal selftests passed 129/129. The production Release x64 package harness compiled with 0 warnings/errors and attestation `30bb937a583970a167ea85f685433ea7be41985ce0ea79bd2d2a4e746d2a9b40`; its fresh extraction proved native Ghostty load/free plus `ghostty_build_info` under the packaged application/system-only search policy. The complete focused set passed 145/145. A stale mixed Release fixture still failed a pre-existing read-only-destination package assertion after app startup, so no package or P11 leaf is checked until the full corrected Release profile and packages rerun. | Plugins/Terminal; Tests/PluginContractTests/PluginContractTests.cpp; Tools/Modules/Packaging/PortablePackageSmoke.psm1; nine focused Pester files; `.build/logs/msbuild-20260826_110457_899-pid25328-73185e1f.log`; `.build/logs/msbuild-20260826_112000_840-pid81192-da04cf56.log` |
| P11 qualification repair | 2026-08-26 | working tree after 36494a80566567c966b22f1134c48f3ae7c156de | Clean x64 Debug, test-enabled Release, ASan Debug, and production Release builds passed with attestations `c77edc1abb5b2d30a4a3a6fd28c2d689bfbf2ac06fc7f7fd054b2d200e40ee00`, `dc6947b895e1265126a0c49f40ba29179a5dcf5a3b673d620b6cb824e3b1e6c2`, `c18c53184b60361606bb8d26eb917706da899e3d3e863b7fe4e3f1f268ad23be`, and `ad8ea64feac261dc8faf61174c4c1c1587adc432130d5dd95283c89e3fe89a25`; Terminal passed 129/129 in both test-enabled profiles and the seeded ASan detector plus ordinary contracts passed. Production rebuild exposed a retained package-host `CppClean` log deleting shared shipping DLLs; the bounded disabled-test clean now covers that host and focused build/workflow contracts pass 43/43. Native chain run `32952824878` passed exact candidate qualification but exposed the admitted-product lane's fresh-runner dependency gap through missing `wil/com.h`; the lane now bootstraps the exact `vcpkg-tool.json` commit and invokes governed `vcpkg-install.ps1 -Platform ARM64` before compilation. No P11 leaf is checked until the corrected native rerun and remaining build/package matrix pass. | `.build/logs/msbuild-20260826_112515_861-pid62648-43315342.log`; `.build/logs/msbuild-20260826_113356_158-pid30208-db0cf7da.log`; `.build/logs/msbuild-20260826_115818_642-pid111576-f817914c.log`; `.build/logs/msbuild-20260826_122326_081-pid111032-6c86ee53.log`; GitHub Actions run 32952824878; Tests/Directory.Build.targets; .github/workflows/ghostty-runtime-product-qualification.yml; Tools/Tests/BuildReproducibility.Tests.ps1; Tools/Tests/ReleaseWorkflowPolicy.Tests.ps1 |
| P11.1/P11.2 | 2026-08-26 | 9edb5e1893878e5d4350aeed5129c0d672b9925c | The complete clean matrix passed: x64 Debug, test-enabled Release, production Release, and ASan Debug attestations are `c77edc1abb5b2d30a4a3a6fd28c2d689bfbf2ac06fc7f7fd054b2d200e40ee00`, `dc6947b895e1265126a0c49f40ba29179a5dcf5a3b673d620b6cb824e3b1e6c2`, `ad8ea64feac261dc8faf61174c4c1c1587adc432130d5dd95283c89e3fe89a25`, and `c18c53184b60361606bb8d26eb917706da899e3d3e863b7fe4e3f1f268ad23be`. ARM64 Debug passed with 0 errors, 43 existing compiler-policy warnings, and attestation `a129cd0fb98ab74c3a3bed662903cbcdca160a147b3a3ae533a3f81a3d51ecf5`; production ARM64 Release passed with 0 warnings/errors and attestation `7187615117a4630fec3fe676b1e7dae0b680d71309fb8acdbd9787712136e7cd`. Both x64 test-enabled profiles passed Terminal 129/129 and the embedded-terminal Commands filters. | `.build/logs/msbuild-20260826_112515_861-pid62648-43315342.log`; `.build/logs/msbuild-20260826_113356_158-pid30208-db0cf7da.log`; `.build/logs/msbuild-20260826_115818_642-pid111576-f817914c.log`; `.build/logs/msbuild-20260826_122326_081-pid111032-6c86ee53.log`; `.build/logs/msbuild-20260826_123615_919-pid37480-6fd5d8f9.log`; `.build/logs/msbuild-20260826_124920_187-pid67636-f92fa386.log` |
| P11.3 | 2026-08-26 | 9edb5e1893878e5d4350aeed5129c0d672b9925c | The x64 ASan Debug profile first proved the seeded heap overflow is detected, then passed the ordinary PluginContract and Terminal 129/129 contracts without an unexpected sanitizer report. | `.build/logs/msbuild-20260826_115818_642-pid111576-f817914c.log`; `.build/x64/ASan Debug/PluginContractTests.exe` |
| P11.5 | 2026-08-26 | 9edb5e1893878e5d4350aeed5129c0d672b9925c | Approved build identity 1 produced the complete serial package matrix. SHA-256 values are x64 ZIP `c0541ee2529511d55b357a78a749df3e3e7771c1f30cceb8e0090aa9bea7a012`, ARM64 ZIP `d51a03128d2d144a52c60e59a3df0f4c784b8c9014a7bd4bd672062b29f9a4e5`, x64 MSI `2f3fdff31bd1d156be7851ca3e65d22a18d414e023bd1ffe468399e0558f3dd0`, x64 MSIX `df0b3a51104a94e5fa2cd221c15cb6c8ce96c1a131bf76116a18701a272459c0`, and ARM64 MSIX `e1456d4b09f7da0be2ff5947a6e32de57ebf7acb34d1487dd04f522a537abbd0`. MSI administrative extraction and both MSIX extractions matched the receipt-bound runtime identities/notices; MSIX manifests report `7.0.1.0` and exact x64/arm64 architectures. | `.build/AppPackages`; `.build/P11MsiValidation-a16fe9fa6a434d90b5ed0d85dd01dc3f`; `.build/P11PackageValidation-d415ea52db2c4c18be6779663945e004` |
| P11.7 | 2026-08-26 | 9edb5e1893878e5d4350aeed5129c0d672b9925c | The original nine-file focused Pester gate passed 145/145. The later sealed-profile repair correctly restored `TerminalDependencyLifecycle.Tests.ps1` to its historical five-notice bytes, so that historical suite is superseded as an active-runtime gate by the amended current-runtime cohort recorded below. | Tools/TerminalEngine/HistoricalTerminalProfiles.json; Tools/Tests/TerminalPluginSourceContracts.Tests.ps1 |
| P11 native dependency-pin repair | 2026-08-26 | working tree after 9edb5e1893878e5d4350aeed5129c0d672b9925c | Native chain `32959056475` passed the exact candidate job, native ARM64 host guard, pinned vcpkg tool bootstrap, and tool-identity check. Dependency installation then failed before compilation because the tool checkout at `a51bb4d1434e6d0927ff79db8033bed8522b85df` did not contain version-database entries required by the independently pinned ports baseline `9e593bb18ea69cc5095e012465dcd675a822ed0d`. The corrected workflow fetches both commits, keeps the executable checkout on the tool pin, materializes the ports baseline in a separate worktree, assigns that worktree to `VCPKG_ROOT`, and still invokes governed `vcpkg-install.ps1 -Platform ARM64`. Focused release-workflow policy passed 22/22. | GitHub Actions run 32959056475; .github/workflows/ghostty-runtime-product-qualification.yml; Tools/Tests/ReleaseWorkflowPolicy.Tests.ps1; Specs/Testing/Testing_SelfTests.md; Specs/Terminal/Terminal_EmbeddedPlugin.md |
| P11 sealed-profile repair | 2026-08-26 | working tree after ee219ec0a1185cad6ecdc0f06a059eb10e962dc0 | The first Fresh Full diagnostic passed targeted deployment and automatic re-attestation, then broad Pester found one historical-profile hash mismatch. Restored the exact sealed Gate0/Round4 test bytes and moved the active six-notice assertion into `HistoricalTerminalProfiles.json`'s current-runtime cohort. The sealed file again matches SHA-256 `11154a9e66a8bb26a94795e00cf716669d00e5690db9a232f1f8b71fbff97d1e`; historical/current focused contracts pass 38/38 and the full non-build-toolchain Pester profile passes 693/693. The later Commands diagnostic's 13 path-related failures included `ERROR_FILENAME_EXCED_RANGE` under the deep default worktree sandbox, so the final Fresh run will use the documented `C:\RSPerf` root while its Pester entry remains repository-local by contract. | Tools/TerminalEngine/HistoricalTerminalProfiles.json; Tools/Tests/HistoricalTerminalProfile.Tests.ps1; Tools/Tests/TerminalPluginSourceContracts.Tests.ps1; Specs/Testing/Testing_SelfTests.md; `.build/ValidationEvidence/runs/20260826T114923Z-44984-2e2daca306824620a1d6f8e29c66fbf1` |
| P11 native explicit-root repair | 2026-08-26 | working tree after f806671b482cdeb840964ef46d740610a9f71ab8 | Native chain `32965426965` passed candidate input authentication, exact native candidate build, all upstream native ARM64 targets, canonical evidence publication, and artifact upload. The admitted-product continuation then failed dependency installation before compilation because `VCPKG_ROOT` is ignored when `vcpkg.exe` already resides in the independently pinned tool checkout; the old checkout therefore supplied obsolete versions. Added a validated `vcpkg-install.ps1 -VcpkgRoot` surface that requires `ports` plus `versions` and passes vcpkg's explicit `--vcpkg-root` option while retaining `vcpkg-tool.json` executable attestation. Focused workflow, install-safety, and reproducibility contracts pass 49/49; tool inventory and spec inventory report zero blocking findings. | GitHub Actions run 32965426965; vcpkg-install.ps1; .github/workflows/ghostty-runtime-product-qualification.yml; Tools/Tests/VcpkgInstallSafety.Tests.ps1; Tools/Tests/ReleaseWorkflowPolicy.Tests.ps1; Tools/Tests/BuildReproducibility.Tests.ps1; Specs/Build/Build_Toolchain.md; Specs/Testing/Testing_SelfTests.md |
| P11 Fresh Full diagnostic | 2026-08-26 | working tree at f806671b482cdeb840964ef46d740610a9f71ab8 | Fresh Full x64 Debug run `20260826T124723Z-73692-418eeba216d148febc2a3ae7f9f6d48b` completed 1,978 cases: 1,913 passed, 13 failed, 52 skipped. Both Pester profiles had zero assertion failures (2/2 and 693/693) but correctly rejected seven stale repository-local scratch artifacts; the final disk audit correctly rejected 71 older `C:\RSPerf` diagnostics. Commands failed ten unrelated UI interaction cases and DxUi failed one menu light-dismiss case; every Terminal contract passed 129/129 and all other product/performance suites passed. Preserved both dirty roots by moving them to `.build/TestSandbox-quarantine-p11-fresh1` and `C:\RSPerf-quarantine-p11-fresh1`, recreated clean roots, and verified both disk audits at zero issues before isolated UI replay. | `.build/ValidationEvidence/runs/20260826T124723Z-73692-418eeba216d148febc2a3ae7f9f6d48b`; `.build/logs/msbuild-20260826_145821_885-pid73692-b6faf1ed.log`; quarantined TestSandbox roots |
| P11 Fresh Full UI isolation | 2026-08-26 | working tree after e409a85f43ecf4da2323125104ff5ddea6d0ec7a | Strict isolated replay passed the complete DxUi suite. Nine of the ten failed Commands cases passed serially without product changes. The remaining `cmd_pane_fileops_completed_group_and_navigation` failed repeatably because its fixture obtained distinct left/right pane filesystem interfaces but passed `nullptr` for the destination filesystem in three admissions while expecting right-pane qualified navigation. The durable UI contract requires a supplied destination filesystem to match the pane before those actions are exposed, so production remained fail-closed and the fixture now passes its existing `rightFileSystem`. Attested Debug x64 build `4b685e1f8bca073217c476289449347b7ba8519e697ce48c959841e3c1e70d55` completed with zero diagnostics; the exact formerly failing case then passed 1/1 in 7.160 seconds. | `RedSalamander/SelfTest/Commands/Commands.SelfTest.FileOps.cpp`; `.build/logs/msbuild-20260826_161342_573-pid123492-623e3cb4.log`; preserved P11 isolation scratch root; isolated DxUi log `rs-p11-dxui-isolation.log` |
| P11 native tool-schema repair | 2026-08-26 | working tree after e409a85f43ecf4da2323125104ff5ddea6d0ec7a | Native chain `32977391011` passed exact candidate authentication/build, all upstream native ARM64 targets, canonical evidence publication, and artifact upload. The admitted-product continuation reached the explicit registry root but failed before compilation: tool commit `a51bb4d1434e6d0927ff79db8033bed8522b85df` is the 2026-07-13 vcpkg tool release and its `scripts/vcpkg-tools.json` uses schema 1; ports baseline `9e593bb18ea69cc5095e012465dcd675a822ed0d` is the 2026-07-27 tool release and uses schema 2. Updated `vcpkg-tool.json` to the compatible baseline release. Both reusable bootstrap workflows now read the two pinned metadata files from exact Git objects, require positive schema versions and executable schema greater than or equal to registry schema, and fail before checkout/bootstrap with an update-specific diagnostic. Durable Build/Testing/Terminal/Installer contracts now own that process; focused build/release policy tests pass 43/43. | GitHub Actions run 32977391011; `vcpkg-tool.json`; `.github/workflows/build-reusable.yml`; `.github/workflows/ghostty-runtime-product-qualification.yml`; `Tools/Tests/BuildReproducibility.Tests.ps1`; `Tools/Tests/ReleaseWorkflowPolicy.Tests.ps1`; official microsoft/vcpkg commit metadata for both pins |
| P11 Fresh Full diagnostic 2 | 2026-08-26 | working tree at 9887a7ec8f0089127c592f3c0bbfe2f7bb4df91e | Fresh Full x64 Debug run `20260826T151217Z-111912-d70bd17930fe4c61a66600ab9877eee4` completed 1,976 cases: 1,915 passed, 9 failed, and 52 skipped. Every Terminal case passed 129/129; both Pester profiles passed 2/2 and 693/693; Compare passed 228 with 30 intentional skips; File Operations passed 122 with 20 intentional skips; DxUi and the native/performance suites passed. Eight Commands failures shifted with same-process predecessor order and passed isolated replay. The only non-Commands failure, `TestViewerShellComboHostsLongRunOpenCloseStayStable`, passed a direct six-cycle long-run rerun including forced-parent teardown with no threshold change. The run is diagnostic, not P11.8 evidence. | `.build/ValidationEvidence/runs/20260826T151217Z-111912-d70bd17930fe4c61a66600ab9877eee4`; external runner artifact run `20260826T151217Z-111912-d70bd17930fe4c61a66600ab9877eee4`; exact ViewerPE long-run replay |
| P11 native manifest-registry repair | 2026-08-26 | working tree after 9887a7ec8f0089127c592f3c0bbfe2f7bb4df91e | Native run `32983706970` again passed exact candidate authentication/build and every upstream native ARM64 target. The admitted-product continuation selected the compatible tool and registry roots, then vcpkg failed before product compilation because `vcpkg.json` requested `aws-sdk-cpp` 1.11.860 while the pinned registry contains 1.11.840. `Specs/FileSystem/FileSystem_S3.md` binds generated serializer ownership to 1.11.840 until a dedicated SDK re-audit, so the manifest floor is restored rather than silently advancing the baseline. The installer now proves every object-form `version>=` floor has an exact selected-registry entry before dependency locks or staging; focused install-safety/build-reproducibility/tool-inventory tests passed 46/46. A clean dependency install replaced cached 1.11.860 with 1.11.840, and the complete Debug x64 solution built with zero warnings/errors and attestation `65ed04d503e19b37fa15b0ad9a6be6eb704200860e043534a926ce2ca3f75b35`. | GitHub Actions run 32983706970; `vcpkg.json`; `vcpkg-install.ps1`; `Tools/Modules/Build/VcpkgInstallSafety.psm1`; `Tools/Tests/VcpkgInstallSafety.Tests.ps1`; `.build/logs/msbuild-20260826_182314_117-pid122604-43e246d1.log` |
| P11 native manifest-baseline repair | 2026-08-26 | working tree after 4beb3962d3aa333a78d754fc0f45ed6bf68dc9c6 | Native run `32993024766` passed the exact candidate build and every upstream native ARM64 target. The admitted-product continuation then exercised the new pre-mutation constraint audit and correctly rejected `ftxui >= 7.0.3`: the pinned ports baseline `9e593bb18ea69cc5095e012465dcd675a822ed0d` contains only 7.0.0. Upstream vcpkg commit `f941c7e94d289cf53903405bed7f41c5b197ca6c` is the reviewed 7.0.3 version-database addition, and an exact upstream version-file audit proves that it also contains every other declared manifest floor. Advanced only `builtin-baseline`; the independently pinned executable tool already supports the registry schema. | GitHub Actions run 32993024766; `vcpkg.json`; official `microsoft/vcpkg` version database at `f941c7e94d289cf53903405bed7f41c5b197ca6c` |
| P11 exact AWS dependency-pin repair | 2026-08-26 | working tree after 1e3b595efcce0115345aa8ccbcbb7cc583b2dea6 | Final-tree Fresh Full dependency resolution exposed that `version>= 1.11.840` is a floor, not an exact pin: the corrected `f941c7e9` baseline selected its newer default `aws-sdk-cpp` 1.11.860, contrary to the S3 serializer audit. Stopped the local run before product compilation and canceled native run `32998094279` before admitted-product dependency installation. Added a manifest override for exactly 1.11.840, expanded the pre-mutation registry audit to exact overrides, and added a repository contract that requires the floor and override together until the serializer boundary is re-audited. Focused install/build/release contracts passed 54/54; tool and spec inventories reported zero blockers; vcpkg dry-run resolved exact version 1.11.840 from tree `3058f6cf9c4b0cec8044e644f59680739a7d47e3`. | `vcpkg.json`; `Tools/Modules/Build/VcpkgInstallSafety.psm1`; `Tools/Tests/VcpkgInstallSafety.Tests.ps1`; `Specs/FileSystem/FileSystem_S3.md`; GitHub Actions run 32998094279; vcpkg dry-run |
| P11 broad UI isolation repair | 2026-08-26 | working tree after 9887a7ec8f0089127c592f3c0bbfe2f7bb4df91e | Cluster replay proved the remaining Commands failures were accumulated directed-input/focus state rather than eight deterministic product defects. The shared per-case boundary now proves left FolderView focus, pumps a bounded input quiet point, clears any navigation/search state delivered during settling, and re-proves focus before dispatch. The File Operations completion-layout case now polls to the existing bounded asynchronous completion contract instead of sampling once. Source contracts passed 176/176; the final Debug x64 solution built with zero warnings/errors and attestation `bde52ef29dd1d9573904b08bd80c34e65092674e42bd3366ce9238d97a1ea421`. The authoritative same-process Commands suite then passed all 866 executed cases with 0 failures and only the two intentional opt-in skips. | `RedSalamander/SelfTest/Commands/Commands.SelfTest.Settings.cpp`; `RedSalamander/SelfTest/Commands/Commands.SelfTest.FileOps.cpp`; `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`; `.build/logs/msbuild-20260826_183431_762-pid81372-139f9614.log`; external TestSandbox run `20260826T163842Z-120432-f869919432cb455ab25296a276e3da58` |
| P11.7 revalidation | 2026-08-26 | working tree after 9887a7ec8f0089127c592f3c0bbfe2f7bb4df91e | The focused gate was amended to preserve byte-sealed historical tests and cover the two active P11 repair surfaces. The sealed `TerminalDependencyLifecycle.Tests.ps1` SHA-256 still matches `11154a9e66a8bb26a94795e00cf716669d00e5690db9a232f1f8b71fbff97d1e`; the historical-profile integrity suite passed 4/4. The ten active focused Pester files passed 317/317 with zero failed, skipped, pending, or inconclusive tests; tool inventory classified 183/183 with zero findings and spec inventory reported zero blocking findings. | `Tools/TerminalEngine/HistoricalTerminalProfiles.json`; `Tools/Tests/HistoricalTerminalProfile.Tests.ps1`; ten active focused Pester files; `Tools/tool-inventory.json`; `.build/SpecInventory/spec-inventory.json` |
| P11 Fresh Full diagnostic 3 | 2026-08-26 | working tree at e2858d0e6f8c0e7596c9fbfa59f34e8c4e41d861 | Fresh Full x64 Debug run `20260826T184510Z-115104-3722188961da4695b056fc727af14c60` reached 1,893 cases: 1,860 passed, 1 failed suite, and 32 skipped. Commands passed all 866 executed cases with two intentional skips, source contracts passed 698/698, Compare passed all 228 executed cases with 30 intentional skips, and every Terminal gate passed. File Operations exited during `Phase7_CopyItemsSingleFolderRecursiveParallelism`; WER report `e6baff53-1191-4353-ab9f-c9e685978573` and dump SHA-256 `7CDDE52F7677208C1D8E2F460814FDF5A437F3B364D8FDDE86E92C9E6BE34D57` prove an access-violation write. LLDB symbolication located the corrupted path in `MakeCopyTempSiblingPath` under the recursive parallel-copy producer. | `.build/ValidationEvidence/runs/20260826T184510Z-115104-3722188961da4695b056fc727af14c60`; `Specs/TestRuns/4cb089111a23/Continuation/20260826_213953_fileops_recursive_copy_crash/README.md`; local WER report and dump |
| P11 recursive-copy crash repair | 2026-08-26 | working tree after e2858d0e6f8c0e7596c9fbfa59f34e8c4e41d861 | The parallel recursive producer retained path strings, find data, directory authority, and retry selection in each recursion frame; the 96-level deterministic case exhausted the effective Debug worker stack before the explicit traversal ceiling. Those traversal-owned objects now live in RAII-owned heap frames while depth, metadata, conflict, cancellation, queue, and cleanup semantics remain unchanged. The complete Debug x64 solution built with zero warnings/errors and attestation `3f21b9b502f980b18be58434135cd28b89ad5f42be80134601b641a42a06a7f8`. Exact replay `20260826T200200Z-107868-35596a52f33a45b397598d0ac2b99545` passed 3/3; broad File Operations run `20260826T200236Z-90056-04e05f6cca8c4a59b215954726cf90cf` passed 122/122 executed with only 20 intentional remote/opt-in skips and zero disk-audit findings. | `Plugins/FileSystem/FileSystem.FileOps.cpp`; `.build/logs/msbuild-20260826_215842_662-pid72832-61b6e06b.log`; both `C:\RSPerf-Fix` validation runs; compact continuation record |
| P11 native workspace-isolation repair | 2026-08-26 | working tree after 66e6a3aa29373ba07a4ad590902305bb31006dcf | Native run `33001401332` passed exact candidate authentication/build, all upstream native ARM64 targets, canonical evidence publication, the pinned admitted-product registry/tool bootstrap, and the complete 47-minute exact ARM64 dependency installation. The first admitted Debug product build then failed before MSBuild because repository identity validation correctly rejected untracked root path `vcpkg-registry/`; the required evidence upload also failed because no product artifact had been created. The workflow now materializes and consumes that auxiliary Git worktree under ignored `.build/vcpkg-registry`, preserving fail-closed BuildEvidence semantics and a clean repository identity. Focused release-workflow policy passes 22/22. | https://github.com/DualTail/RedSalamander/actions/runs/33001401332; `.github/workflows/ghostty-runtime-product-qualification.yml`; `Tools/Tests/ReleaseWorkflowPolicy.Tests.ps1` |
| P11 Fresh Full diagnostic 4 | 2026-08-26 | 66e6a3aa29373ba07a4ad590902305bb31006dcf | Fresh Full x64 Debug run `20260826T201917Z-113072-9bb2cd87d72147439d381b2162082c00` completed 1,981 cases: 1,927 passed, 2 failed executable entries, and 52 intentional skips. Build-toolchain Pester passed 2/2, source contracts passed 698/698, Compare passed 228/228 executed, Commands passed 866/866, File Operations passed 122/122 without reproducing the repaired recursive-copy crash, and all Terminal and performance gates passed. DxUiTests failed one menu shadow-margin timing assertion and ViewerPETests failed one localized ViewerText clipboard-warning observation in the broad ordering; complete direct isolated replays of both executables then exited zero and passed the exact assertions without source changes. The run is diagnostic and does not satisfy P11.8; final Fresh Full remains required on the final committed tree. | `.build/ValidationEvidence/runs/20260826T201917Z-113072-9bb2cd87d72147439d381b2162082c00`; `C:\RSPerf-Final\runs\20260826T201917Z-113072-9bb2cd87d72147439d381b2162082c00`; complete isolated DxUiTests and ViewerPETests replays |
| P11 Fresh Full diagnostic 5 | 2026-08-27 | e52d059846f070e24e8b7bdbc6c4faeb53db9ebc | Fresh Full x64 Debug run `20260826T212935Z-84192-806e1dd17c704fa5905e9f439a9f776e` completed 1,981 cases: 1,904 passed, 4 failed, and 73 skipped. Build-toolchain Pester passed 2/2, source contracts passed 698/698, Compare passed all 228 executed cases, all Terminal and performance gates passed, and the repaired recursive-copy path completed. Commands failed only `cmd_pane_copy_text`; File Operations failed `Causeway_BridgeFailureStatusAndPausedReader` and `Phase5_DiscoveryCancelReleasesSlot`; ViewerVLC failed its first dependency-load/focus harness but passed the same complete harness in later churn cycles. On the unchanged binary the copy-text case passed 20/20, and each File Operations case passed 10 repeats with 30/30 recorded steps. The copy-text failure nevertheless exposed a durable contract violation: `UI_CommandMenuKeyboard.md` required bounded clipboard-open retries while `Common::Clipboard::TrySetUnicodeText` attempted once. This diagnostic run is not P11.8 evidence. | `.build/ValidationEvidence/runs/20260826T212935Z-84192-806e1dd17c704fa5905e9f439a9f776e`; `C:\RSPerf-P11Final\runs\20260826T212935Z-84192-806e1dd17c704fa5905e9f439a9f776e`; focused runs `20260826T223802Z-15792-c66f80116aad415db2a3148a4b42f655`, `20260826T223858Z-118276-96d7d603655146d99cb0cb16f1ac6e3a`, and `20260826T223944Z-113100-af7500ddbcd74d70972d71eb3382c96b` |
| P11 clipboard-contention repair | 2026-08-27 | working tree after e52d059846f070e24e8b7bdbc6c4faeb53db9ebc | The canonical shared Unicode writer now prepares the complete WIL-owned payload, then makes at most 20 `OpenClipboard(...)` attempts separated by 10 ms sleeps without translating, dispatching, or removing UI messages. The shared-helper catalog now owns that retry contract, and the existing no-reentrant-pump source guard covers both the FolderView file-transfer helper and `Common::Clipboard::TrySetUnicodeText`. A serialized Debug x64 build passed with zero warnings/errors and attestation `eeebf07575d41027946cf2e0f2bae6767897b39ff069319b03c8934120fb0113`; the exact source guard exited zero and `cmd_pane_copy_text` passed 50/50 against the repaired binary. | `Common/UnicodeClipboard.h`; `RedSalamander/SelfTest/Commands/Commands.SelfTest.PluginConfig.cpp`; `Specs/Core/Core_SharedHelpers.md`; `.build/logs/msbuild-20260827_004250_395-pid113604-d64739b0.log`; `C:\RSPerf-P11ClipboardGuard`; `C:\RSPerf-P11ClipboardRepeat` |
| P11 native tool-workspace repair | 2026-08-27 | working tree after 7f33e3af9547228e1b38bdc64c525bfade04d374 | Native run `33015576024` passed exact candidate authentication/build, every upstream native ARM64 target, canonical evidence publication, and the exact pinned dependency installation. The admitted-product Debug build then failed before MSBuild because BuildEvidence correctly rejected the remaining untracked root `vcpkg/` executable checkout; no product artifact existed for the required upload. The workflow now creates `.build`, clones the pinned executable tool to `.build/vcpkg-tool`, materializes the ports baseline as sibling `.build/vcpkg-registry`, and passes those explicit ignored paths to bootstrap and governed installation. Focused release-workflow policy passes 22/22. | GitHub Actions run `33015576024`; `.github/workflows/ghostty-runtime-product-qualification.yml`; `Tools/Tests/ReleaseWorkflowPolicy.Tests.ps1` |
| P11 Fresh Full qualification before native integration repair | 2026-08-27 | 5e38a79b6a34ed088d7e61dea0b913b2f7b50393 | Fresh Full x64 Debug run `20260826T232559Z-90732-47589b7655304602aac9f34eb3c17abe` passed every executable and policy entry: 1,981 total, 1,929 passed, zero failed, 52 intentional skips, zero flaky/regression/isolation classifications, zero quarantine entries, and zero disk-audit issues. Commands passed 866/866 executed, File Operations 122/122, Compare 228/228, source contracts 698/698, and all Terminal, DxUi, Viewer, native helper, and performance gates passed. The final attested build receipt is `ff3d84560f8cc9f2b7641d71792593e7e0f881dac4096c49200a73dae1b24329`. A later workflow/spec-only native integration repair invalidates this as the final source-contract snapshot, so P11.8 remains unchecked pending one final Fresh rerun. | `.build/ValidationEvidence/runs/20260826T232559Z-90732-47589b7655304602aac9f34eb3c17abe`; `C:\RSPerf-P11Final-5e38a79b\runs\20260826T232559Z-90732-47589b7655304602aac9f34eb3c17abe` |
| P11 native MSBuild-integration repair | 2026-08-27 | working tree after 5e38a79b6a34ed088d7e61dea0b913b2f7b50393 | Native run `33022375261` passed exact candidate qualification, canonical evidence publication, pinned tool/registry bootstrap, and the 44-minute exact ARM64 dependency install. The installer proved `.build/vcpkg_installed/arm64/arm64-windows/include/wil/com.h` existed, but Debug compilation failed with missing WIL/Blake3/7-Zip headers because the hosted native Visual Studio image lacks optional vcpkg MSBuild integration. The product lane now validates and process-locally imports the pinned checkout's `vcpkg.props`/`vcpkg.targets`, disables manifest installation, binds `VcpkgInstalledDir` to the repository-owned ARM64 root, and supplies the same include root through `CL`; it never mutates user-wide MSBuild. Release-workflow policy passes 22/22 and spec inventory reports zero blockers. | GitHub Actions run `33022375261`; artifact `ghostty-product-native-arm64-33022375261`; `.github/workflows/ghostty-runtime-product-qualification.yml`; `Tools/Tests/ReleaseWorkflowPolicy.Tests.ps1`; `Specs/Build/Build_Toolchain.md`; `Specs/Testing/Testing_SelfTests.md`; `Specs/Terminal/Terminal_EmbeddedPlugin.md` |
| P11 exact-tree Fresh Full before app-local repair | 2026-08-27 | 0f5a90aa14fa04bee04098f91a1d6c04fa6a3a1c | Fresh Full x64 Debug run `20260827T011211Z-18848-fe034c5667c04208bdc8d8b5eae307f6` completed with every one of its 18 planned entries promoted and final status `passed`. The mutation state is stable, every Commands/Compare/File Operations/DxUi/standalone/performance/Pester entry passed, and the build receipt is `752a90d1f2f53a6f45cebd37ca267f9b02338b48768e1d4d0a6fe4967e71641d`. This proves commit `0f5a90aa`, but the later app-local workflow/spec/test repair requires one final exact-tree Fresh rerun for P11.8. | `.build/ValidationEvidence/runs/20260827T011211Z-18848-fe034c5667c04208bdc8d8b5eae307f6`; `C:\RSPerf-P11Final-0f5a90aa\runs\20260827T011211Z-18848-fe034c5667c04208bdc8d8b5eae307f6` |
| P11 native app-local evidence repair | 2026-08-27 | working tree after 0f5a90aa14fa04bee04098f91a1d6c04fa6a3a1c | Native run `33029337780` passed candidate qualification and the exact pinned dependency install. Process-local vcpkg integration fixed the prior header failure: the entire 53-minute Debug ARM64 solution compile completed with 43 warnings and zero errors, including native `RedSalamander.exe`. Receipt publication then rejected `.build/vcpkg_installed/arm64/arm64-windows/debug/bin/blake3.dll` from `Common.write.1u.tlog` because the pinned vcpkg defaults `VcpkgXUseBuiltInApplocalDeps=true`, whose built-in copier logs package source paths even after copying to `.build/ARM64/Debug`. Both official reusable build paths now set the property false before MSBuild, selecting the same pinned checkout's PowerShell copier, which records copied destination paths compatible with the fail-closed receipt boundary. Durable Build/Installer/Testing/Terminal contracts and focused source-policy tests own the requirement. | GitHub Actions run `33029337780`; artifact `ghostty-product-native-arm64-33029337780` ID `9632387117`, SHA-256 `244f7f2182a218dc40f8fffcac857862cc971939a13990ace6b26d83e7ee7554`; `.github/workflows/build-reusable.yml`; `.github/workflows/ghostty-runtime-product-qualification.yml`; `Tools/Tests/BuildReproducibility.Tests.ps1`; `Tools/Tests/ReleaseWorkflowPolicy.Tests.ps1`; `Specs/Build/Build_Toolchain.md`; `Specs/Installer/Installer_PortableZip.md`; `Specs/Testing/Testing_SelfTests.md`; `Specs/Terminal/Terminal_EmbeddedPlugin.md` |
| P11 Fresh Full diagnostic 6 | 2026-08-27 | 0da1668838a6201a86be22383e9cf9634b88b29b | Fresh Full run `20260827T065028Z-120016-a1c78d8f699d49499be352c5e752c83d` completed 1,981 cases: 1,905 passed, 3 failed, 73 skipped, and zero disk-audit issues. Build-toolchain Pester passed 2/2, source/tooling contracts 698/698, Compare 228/228 executed, all Terminal/plugin, DxUi, Viewer, and performance entries passed. Commands failed only `cmd_pane_find_dialog_restored_combined_view_state_copy_follows_visible_columns`; File Operations failed the same `Causeway_BridgeFailureStatusAndPausedReader` and `Phase5_DiscoveryCancelReleasesSlot` broad-order cases that had already passed the unchanged binary's isolated 122/122 replay. A simultaneously active independent Full run held successive interactive leases and makes this diagnostic unsuitable for P11.8. | `.build/ValidationEvidence/runs/20260827T065028Z-120016-a1c78d8f699d49499be352c5e752c83d`; `C:\RSPerf-P11Final-0da16688-r3`; isolated File Operations run `20260827T063408Z-95780-071436d1c42a4ee79dc79a5358705894`; direct ViewerPE replay |
| P11 native hosted-desktop/TestSandbox repair | 2026-08-27 | working tree after 0da1668838a6201a86be22383e9cf9634b88b29b | Native run `33041113316` passed the exact candidate job and reached the admitted product's representative Commands smoke after successful native dependency bootstrap, Debug ARM64 build, runtime contracts, and package preparation. Three of four embedded-terminal cases passed; `cmd_pane_embedded_terminal_plugin_lifecycle` failed because the hosted desktop allowed child focus without granting app-root foreground, where production correctly rejects pointer-follow. The artifact's only disk-audit issue was the preceding `terminal_` filter's valid sibling run under the shared root. The lifecycle fixture now binds its expectation to the production foreground guard and proves fail-closed source-focus preservation when foreground is unavailable. Every native filter receives a distinct `REDSALAMANDER_TEST_ROOT`. Focused policy suites pass 199/199, spec inventory reports zero blockers, the full Debug x64 solution builds with zero diagnostics and receipt `afb8d099ab2dd0f84440cdfc123b5c745de591b862e69408eb8834c3a3833b29`, and exact lifecycle runner case `20260827T083205Z-116220-debe4b5e90774f78b72444d1401b7456` passes 1/1 with zero disk issues. | GitHub Actions run `33041113316`; artifact `ghostty-product-native-arm64-33041113316` ID `9637209386`, SHA-256 `b235930e19d583776560059d5f9927c5596f72e4a8bed66b91eb866fd382833c`; `.github/workflows/ghostty-runtime-product-qualification.yml`; `RedSalamander/SelfTest/Commands/Commands.SelfTest.Navigation.cpp`; focused Pester; `.build/logs/msbuild-20260827_102802_946-pid115460-d59b7512.log`; `C:\RSPerf-P11NativeHarnessFix-x64-r2` |
| P11.8 | 2026-08-27 | 05cc5ee506cee2c17bfca4e432a8dba3d44adff2 | Final Fresh Full x64 Debug run `20260827T083522Z-29788-a876a1efe2d54da993787f9f5ab524a4` promoted every one of its 18 planned entries and completed with canonical status `passed`: 1,982 total cases, 1,930 passed, zero failed, 52 intentional capability skips, zero flaky/regression/isolation classifications, and zero active/invalid quarantine entries. Commands passed 866/866 executed, File Operations 122/122, Compare 228/228, both Pester profiles 2/2 and 699/699, and every Terminal, DxUi, Viewer, standalone, and performance gate passed. The final attested build receipt is `f073ed777cf84e60ff2175489861419ab1d0f8009612bf29370d3e05ceb810d2`; validation plan digest is `b996b39bea73fde9e70405fa819091a3288b53bbd323c49602eb673261094e7c`. | `.build/ValidationEvidence/runs/20260827T083522Z-29788-a876a1efe2d54da993787f9f5ab524a4`; external TestSandbox run `20260827T083522Z-29788-a876a1efe2d54da993787f9f5ab524a4` |
| P11.4 | 2026-08-27 | 05cc5ee506cee2c17bfca4e432a8dba3d44adff2 | Exact-tree native run `33054673670` passed both jobs. Candidate qualification authenticated commit/tree/lock/overlay, built on an ARM64 host, and passed `test-lib-vt`, `test-lib-vt-build`, and `test-lib-vt-schema`. The admitted-product continuation passed pinned dependency installation, Debug ARM64 build, runtime and Terminal contracts, 15/15 filtered Commands cases, production Release ARM64 build, native portable smoke, and required evidence upload. All three Commands roots were isolated, exited zero, had zero disk-audit issues, zero flaky/regression/isolation classifications, and no quarantine. The canonical Product identity binds ARM64, commit `9c3ec931d64561a8407dde7ac984ce156ae91539`, 191 exports, export hash `85d134a5fa8c29d83f9c22559632600d474eb433c12a8bd3d55da9863ddeef15`, six imports, six notices, and overlay hash `06847f6594427cc584cef0f0ec66a54080f236b2414fe40e61b381accaf38264`. | GitHub Actions run `33054673670`; candidate artifact `ghostty-native-arm64-9c3ec931d64561a8407dde7ac984ce156ae91539` ID `9640574448`, SHA-256 `c16356115ab49015860656e93a4a3447f5798e8d93f04b8932980f154fa4d217`; product artifact `ghostty-product-native-arm64-33054673670` ID `9645240373`, SHA-256 `bbf4c7e8f64f398b4cc36363fe0f292a09fff2d5acc339dfc9c360ebabfae061` |
| P11.6 | 2026-08-27 | 05cc5ee506cee2c17bfca4e432a8dba3d44adff2 | The serial x64/ARM64 ZIP, x64 MSI, and x64/ARM64 MSIX matrix from P11.5 already proved receipt-bound architecture, identities, notices, extraction, plugin loading, and executable x64 smoke. Final native run `33054673670` completed the remaining ARM64 execution boundary: its production ZIP contains the admitted `ghostty-vt.dll`, runtime identity, Ghostty license, and aggregate notices, then passed clean extraction plus receipt-bound identity/license/native DLL load/free smoke. The uploaded ZIP is 23,553,882 bytes with SHA-256 `72e26be127f0e64d6e3552fbe1d351ba05bc80384267a875267fcaf115a4467b`. Release policy on this exact tree requires the admitted-product workflow and provides no default waiver. | P11.5 package hashes/extractions; GitHub Actions run `33054673670`; product artifact `9645240373`; `AppPackages/RedSalamander-7.0.1-ARM64-Portable.zip`; `.github/workflows/release.yml`; `.github/workflows/ghostty-runtime-product-qualification.yml` |
| P11.9 | 2026-08-27 | 05cc5ee506cee2c17bfca4e432a8dba3d44adff2 | Required phase commit has exact subject `Qualify Ghostty shipping boundaries`; this plan-only checkpoint records its full SHA only after final Fresh Full and native ARM64 product/package qualification passed. Temporary `RS_GHOSTTY_CANDIDATE_LOCK_B64` and `RS_GHOSTTY_CANDIDATE_PATCH_B64` repository secrets were deleted after artifact inspection and verified absent. | `git show -s --format=%H%n%s 05cc5ee506cee2c17bfca4e432a8dba3d44adff2`; GitHub Actions run `33054673670`; repository secret-name inventory |
| P12.1 | 2026-08-27 | P12 working tree after 9dc15fb42364ca724e41c4f9561d5510ef11a87a | Detached isolated checkout of pre-admission parent `d6a30f0ae3aea62d39c19401dbaa11334d558b2f` received no current Product cache. Previous-lock x64 Force, ARM64 Force, x64 Offline Force, ARM64 Offline Force, and x64 Offline reuse all exited zero. | `Specs/Terminal/Terminal_EmbeddedPlugin.md`; isolated checkout runtime identities summarized in this plan |
| P12.2 | 2026-08-27 | P12 working tree after 9dc15fb42364ca724e41c4f9561d5510ef11a87a | Both final rollback identities authenticated prior commit `b2d44625908eb56aee299d8511d803bc11ab79fc`, 180 exports, export hash `596894726504f7d398ee2a87b8a53116c8d727e5390b1e066293e1770e40fced`, imports `kernel32.dll`/`ntdll.dll`, five notices, overlay, DLL, and pairing. Focused policy passed 143/143; Debug x64 built with zero diagnostics and receipt `04d80cb3820e4e2a49443a38cf7ce2a70211bf4e31d345864e3cc33691a78b0b`; Terminal passed 114/114. Release x64 plus ZIP built with zero diagnostics and receipt `d10cfbd7ea47ae91c40b474d3d6f69454fbe8c185fde38a3b0588f1dfedef66d`; fresh package smoke passed and ZIP SHA-256 is `9cb68260e6dfa715f4a669bc3111e74ce64f969c51fe7db08638a66f0ee82876`. | `Specs/Terminal/Terminal_EmbeddedPlugin.md`; rollback build logs/receipts and package retained in the isolated checkout |
| P12.3 | 2026-08-27 | P12 working tree after 9dc15fb42364ca724e41c4f9561d5510ef11a87a | Re-reviewed Phase 12's authoritative Terminal, Build, Testing, Installer, Tools, and NormativeConsistency targets. Earlier phases already own the admitted pin/ABI/security/Kitty/build/evidence/package/process contracts; final Terminal authority now records exact admitted qualification, rollback commands/results, the five-notice prior unit, and fresh-process Locked-status behavior. Spec inventory reports zero blockers and NormativeConsistency matches the live lock. | `Specs/Terminal/Terminal_EmbeddedPlugin.md`; `Specs/Build/Build_Toolchain.md`; `Specs/Testing/Testing_PerformanceValidation.md`; `Specs/Testing/Testing_SelfTests.md`; `Specs/Testing/Testing_ValidationEvidence.md`; `Specs/Installer/Installer_PortableZip.md`; `Tools/README.md`; `Tools/tool-inventory.json`; `Specs/NormativeConsistency.json` |
| P12.4 | 2026-08-27 | P12 working tree after 9dc15fb42364ca724e41c4f9561d5510ef11a87a | The first genuinely fresh `pwsh -File ... -Mode Locked` exposed an uninitialized `$LASTEXITCODE` dependency; the status helper now captures Git output/exit before pipeline selection and a fresh-process regression passes. Final Locked status returns `current`; Ghostty status Pester passes 15/15; tool inventory classifies 183/183 with zero findings; spec inventory has zero blockers; TestRuns whole inventory passes for 1,140 files; diff check exits zero. Receipt `f073ed777cf84e60ff2175489861419ab1d0f8009612bf29370d3e05ceb810d2` rereads byte/semantic-equivalent to its immutable record and verifies all 177 outputs; every one of final Fresh run `20260827T083522Z-29788-a876a1efe2d54da993787f9f5ab524a4`'s 18 promotions rereads schema/digest/artifact/receipt-valid and terminal green. | `.build/TerminalEngineStatus/P12/locked-status.json`; `.build/SpecInventory/spec-inventory.json`; `.build/ValidationEvidence/runs/20260827T083522Z-29788-a876a1efe2d54da993787f9f5ab524a4`; `Specs/TestRuns` inventory |
| P12.5 | 2026-08-27 | P12 working tree after 9dc15fb42364ca724e41c4f9561d5510ef11a87a | Reviewed `git diff --name-status e7475d3aaacc6dda0eeab58cd486c65d4074be4c` across 99 changed paths plus the P12 working tree. Every path is the planned Ghostty implementation/test/spec/evidence surface or a narrowly recorded P11 validation-remediation exception below; no candidate cache, runtime binary, secret, unrelated feature, or sealed historical file is present. | Git diff/name-status review; Scope and P11 validation-remediation exception in this plan |
| P12.6 | 2026-08-27 | P12 working tree after 9dc15fb42364ca724e41c4f9561d5510ef11a87a | Mechanical checklist audit found 101 checked C0/P0–P11 leaf IDs, 140 pre-P12 Evidence rows, zero missing leaf evidence mappings, and no actual BLOCKED marker. Every C0/P0–P11 phase box is checked; after this update only P12, P12.7, and P12.8 remain unchecked, Current progress is READY TO ARCHIVE, and the next action is the required P12 implementation commit. | Live execution checklist and Evidence log in this plan |
| P12.7 | 2026-08-27 | 8f919e8fbba904e7502c22f3e830ff117abb5c1a | Required P12 phase commit was created while this plan remained in WIP and has exact subject `Close Ghostty runtime upgrade`; its full SHA was captured without amend or history rewrite. | `git show -s --format=%H%n%s 8f919e8fbba904e7502c22f3e830ff117abb5c1a` |
| P12.8 | 2026-08-27 | archive/checkpoint working tree after 8f919e8fbba904e7502c22f3e830ff117abb5c1a | Recorded the P12 SHA, checked P12/P12.7/P12.8, set State to DONE, moved the plan from WIP to Done, and removed I13 plus its ownership rule from the active index. The first archive inventory correctly rejected the previously unknown Done path; the checkpoint added only this exact path to the protected-history allowlist, then final spec inventory passed with zero blockers. | `Specs/Plans/Done/Terminal_LibGhosttyVt_UpgradeAndContinuousReadiness_2026-08-25.md`; `Tools/Modules/Tooling/SpecInformationArchitecture.psm1`; `.build/SpecInventory/spec-inventory.json` |

### WIP-to-Done transition rule

The executor may move this plan from WIP to Done only after C0, P0 through P11,
and P12.1 through P12.7 are checked and evidenced. The separate
archive/checkpoint commit must then perform P12.8 atomically: record the P12
implementation SHA, move the file, set State to DONE, check P12.8 and the P12
phase box in the moved file, remove I13 from Specs/Plans/WIP/README.md, and rerun
Get-SpecInventory.ps1 -FailOnFindings. If that final inventory fails, restore
the plan to WIP with P12/P12.8 unchecked and record the blocker.

## Status

- **State:** DONE
- **Archived from active-work rank:** I13
- **Priority:** P0 dependency/security correctness, P1 performance and process
- **Effort:** L, expected as reviewable phase commits rather than one batch
- **Risk:** HIGH
- **Category:** migration, security, performance, tests, build, developer experience
- **Planned at:** commit e7475d3aaacc6dda0eeab58cd486c65d4074be4c, 2026-08-25
- **Frozen candidate for this plan:** 9c3ec931d64561a8407dde7ac984ce156ae91539
- **Current product pin:** 9c3ec931d64561a8407dde7ac984ce156ae91539
- **Depends on:** I10 owns repository-wide Tools governance; follow it and do not create a parallel inventory system
- **Coordinates with:** I12 owns its existing Terminal focus/input/quiet-point follow-up; serialize overlapping edits under Plugins/Terminal

## Why this matters

The selected Ghostty lib-vt runtime is 717 upstream commits behind the frozen
candidate. The interval changes the public C surface, terminal mode access,
clipboard callback contract, Kitty graphics lifecycle, parser behavior, build
inputs, and multiple allocation and pointer-safety paths. This is therefore a
controlled runtime migration, not a pin-only dependency bump.

The current product build is reproducible and strict, but much of the selected
runtime definition is embedded directly in one build script. There is no active,
Ghostty-specific read-only observation path that can say an update exists,
freeze an explicit candidate, build it away from Product, and carry the same
review into admission. This plan upgrades the runtime and installs that process
without reactivating the sealed multi-engine selection machinery.

## Required outcome

When this plan is complete:

1. x64 and ARM64 product runtimes are reproducibly built from the reviewed
   candidate, exact toolchain, dependencies, overlay, import allowlist, export
   set, licenses, and identity contract recorded in one active product lock.
2. The C++ adapter uses the candidate ABI correctly, including the new terminal
   mode query, sized clipboard request/reply contract, and pending/animated Kitty
   image semantics.
3. Existing RedSalamander security and resource limits remain enforced:
   clipboard read remains disabled, remote clipboard writes remain consented and
   bounded, retained Kitty placements remain capped at 4,096 and 1 MiB of map
   allocation, decoded BGRA cache remains capped at 64 MiB, and applicable
   pipeline high-water remains within 128 MiB.
4. Upstream lib-vt tests, deterministic RedSalamander terminal tests, package
   smoke, x64 execution, native ARM64 execution, import/export checks,
   performance gates, and Fresh Full are green with archived evidence.
5. A read-only maintenance workflow can report the locked state, observe the
   upstream default branch, freeze an explicit candidate, and flag security
   review without editing the lock, opening issues, creating branches, or
   silently repinning.
6. The prior lock remains a complete rollback unit and a rollback rehearsal
   proves that the previous runtime can be rebuilt rather than copied from an
   unverified cache.

## Non-goals

- Do not reconsider the historical Ghostty-versus-Contour-versus-WezTerm
  selection. TerminalEngineDecision and the Gate-0/Round-4 artifacts are sealed
  history.
- Do not reactivate Get-TerminalDependencyStatus.ps1,
  Prepare-TerminalEngineUpgrade.ps1, Build-TerminalEngine.ps1,
  Verify-TerminalEngine.ps1, or TerminalEngineLifecycle.psm1 as the current
  product path. They remain compatibility or historical surfaces.
- be sure to remove useless scripts and workflows from the repository, but do not revive them as a
  current product path.
- Do not automatically choose the newest upstream commit, edit the product lock,
  push a branch, open a pull request, or create an issue.
- Do not add clipboard-read support or a new clipboard-consent product flow.
- Do not ship test hooks in normal Release binaries merely to make package smoke
  pass.
- Do not broaden this migration into unrelated Terminal UI, command-surface,
  renderer, or File Operations refactoring.
- Do not weaken exact runtime identity, import, export, license, package, or
  performance gates to admit the candidate.

## Current-state snapshot

These facts were measured at the planned-at commit. They are inputs to the plan,
not claims that remain true after drift. **This entire snapshot was superseded
by the checked P0 through P10 evidence and the admitted canonical lock. Historical
values in this section cannot trigger a STOP after P10; the live lock, checklist,
and Evidence log are authoritative for continuation.**

### Provenance and change size

| Fact | Current | Frozen candidate |
|---|---|---|
| Commit | b2d44625908eb56aee299d8511d803bc11ab79fc | 9c3ec931d64561a8407dde7ac984ce156ae91539 |
| Git tree | 9b5e7ca2fb3dfaf7c8329645dfb51a1d3baa0325 | 359872ab74fe1ecc96428133c0546be3cb969fdd |
| Version reported by source | 1.3.2-dev | 1.3.2-dev |
| Minimum Zig | 0.16.0 | 0.16.0 |
| Ancestry | base | official descendant: 717 ahead, 0 behind |

The path-filtered interval contains 122 VT/build-relevant changed files, about
40,425 insertions and 4,982 deletions. It includes 39 commits touching public
headers, 20 touching lib-vt, 202 touching terminal implementation, and 11
touching the build boundary. Those counts establish review scale; regenerate
them from the frozen commits and do not copy them into normative specs.

### Current build and identity

- Tools/Build-GhosttyTerminalRuntime.ps1 is the active product builder. It embeds
  the source commit and archive, Zig 0.16.0 archive and SHA-256, uucode inputs,
  build arguments, import allowlist, exact export count/hash, license inputs,
  size envelope, and runtime identity values.
- The Zig x86_64 Windows archive SHA-256 is
  68659eb5f1e4eb1437a722f1dd889c5a322c9954607f5edcf337bc3684a75a7e.
- The current runtime exports 180 exact names with export-set SHA-256
  596894726504f7d398ee2a87b8a53116c8d727e5390b1e066293e1770e40fced.
  Header declaration counts are not DLL export counts.
- Plugins/Terminal/TerminalRuntimeLoader.cpp and
  Plugins/Terminal/TerminalRuntimeLoader.h resolve a 72-name loader-required
  subset. Only one required name is removed in the candidate:
  ghostty_terminal_mode_get.
- The supported replacement is ghostty_terminal_get with the frozen two-field
  GhosttyTerminalModeConfig request and GHOSTTY_TERMINAL_DATA_MODE; this
  structure has no size field.
- Tools/Modules/Build/BuildEvidence.psm1 currently hashes a nonexistent root
  TerminalEngine.lock.json together with the patch and RuntimeDependencies.props.
  The new active lock must replace that dead input.
- The build-reusable workflow cache hashes the builder, its six direct modules,
  the patch, and notices. The active lock, schema, and any new direct module must
  join this same cache dependency closure. Do not create a second cache.

### Public ABI and overlay

- A declaration-oriented header comparison estimates 185 current versus 195
  candidate public declarations: 23 added and 13 removed. This is a review aid
  only; build both DLLs and compute their actual sorted exports.
- Tools/TerminalEngine/Patches/Ghostty-WindowsPrivateRuntime.patch changes 10
  upstream files with roughly 357 additions and 75 deletions.
- A clean apply check against the frozen candidate fails in build.zig,
  include/ghostty/vt/terminal.h, src/terminal/c/terminal.zig,
  src/terminal/kitty/graphics_exec.zig,
  src/terminal/kitty/graphics_storage.zig,
  src/terminal/kitty/graphics_unicode.zig, and
  src/terminal/snapshot/snapshot.zig.
- The current private terminal option assigns
  GHOSTTY_TERMINAL_OPT_KITTY_PLACEMENT_LIMITS the numeric value 31. Candidate
  upstream now owns public option values 31 through 39. This is a hard ABI
  collision. The private C and Zig enum values must be reallocated after the
  frozen candidate public range, with parity and uniqueness tests.
- The overlay also makes the Windows build shared-library-only, strips production
  debug information, and owns exact Kitty retained-placement count/allocation
  limits plus query data. Preserve those concerns unless an explicit reviewed
  replacement is recorded.
- Reviewed P6 correction: the candidate-only static-companion skip is removed;
  the extra static library remains unshipped. Its concern slot is replaced by
  the security-required Kitty clipboard-write disable, preserving the frozen
  five-concern limit.

### Clipboard contract

- The callback changes from returning GhosttyClipboardWriteResult to returning
  void.
- The candidate request is sized and carries name, granted, can_remember,
  context, and a reply function. A callback must validate the received size
  before reading each optional field and must invoke the supplied reply at most
  once when the field is present.
- The callback now also receives Kitty OSC 5522 transactions. Candidate upstream
  defaults the Kitty clipboard transaction buffer to 64 MiB unless
  GHOSTTY_TERMINAL_OPT_CLIPBOARD_WRITE_MAX_BYTES is set.
- RedSalamander currently caps remote OSC 52 clipboard data at 256 KiB and can
  ask asynchronously. A callback cannot truthfully acknowledge an asynchronous
  ask before the user decides.
- Request metadata cannot distinguish an unnamed/no-password Kitty write from
  OSC 52. The scoped upgrade policy is therefore: preserve current consented
  OSC 52 and supported image-write behavior, keep the clipboard-read callback
  unset, set the public runtime write-byte option to the existing 256 KiB
  product ceiling, and set private candidate option 41=false so the engine
  rejects Kitty OSC 5522 with `ENOSYS` before buffering/callback. A full Kitty
  grant, remember, and asynchronous reply UX requires a separately approved
  product plan; it is not silently introduced here.
- No Windows clipboard call, consent UI, or message wait may occur while the
  terminal state lock is held.

### Kitty image lifecycle

- In the candidate, GHOSTTY_KITTY_GRAPHICS_DATA_PTR can return
  GHOSTTY_NO_VALUE while an image payload is still loading.
- Current adapter behavior treats any non-success or null data as permanently
  Rejected and resets only its own Pending work when storage generation changes.
  That can permanently discard an engine-pending image.
- Candidate payload completion changes storage generation while preserving image
  generation. Animations can change image generation. Cache and worker state
  must distinguish engine-pending, adapter-queued, ready, rejected, replaced,
  deleted, and animated content.
- Required regressions cover pending-to-complete, replacement, delete, eviction,
  retransmission, screen/reset, stale worker completion, and animation frame
  changes.

### Upstream correctness and security-relevant changes

- The candidate includes bounded OSC, grapheme, PNG, and Kitty allocation work,
  plus pointer and metadata corrections. Do not claim a published vulnerability
  or CVE unless the status process records a primary advisory source.
- Candidate head fixes DCS high bytes being misclassified as payload. Add a
  deterministic regression with byte values above 0x7F.
- Upstream exposes test-lib-vt, test-lib-vt-build, and test-lib-vt-schema. The
  schema target loads the produced DLL, so ARM64 must execute on a native ARM64
  Windows host rather than being declared green from cross-compilation alone.

### Existing product budgets and evidence

- Hosted render p95 after the upgrade may be at most 10 percent slower than a
  same-machine, same-configuration current-pin baseline.
- Normal percentile claims require at least 200 samples for p95 and 1,000 for
  p99. Existing five-process Kitty ranges are not p95 claims.
- Retained Kitty placements: 4,096 maximum; admitting one additional unique
  placement must fail without mutating retained state; replacement remains
  allowed.
- Placement map allocation: 1 MiB maximum. Decoded BGRA cache: 64 MiB maximum.
  Applicable end-to-end pipeline high-water: 128 MiB maximum.
- Terminal close: at most 5 seconds, with queued, active-source,
  active-converted, and ready bytes all returning to zero.
- Resize must not recreate the terminal or reupload unchanged images.
- Evidence belongs under
  Specs/TestRuns/<MachineHash>/Terminal/<RunId> and
  Specs/TestRuns/<MachineHash>/Commands/<RunId>. Each file must be at most
  2 MiB and each run at most 5 MiB.

### Package-smoke risk to resolve before migration

The source currently suggests, but has not yet execution-proved, this mismatch:

1. Release defaults RSBuildEnableTests to false.
2. Test projects become no-output when tests are disabled.
3. x64 portable-package validation invokes PluginContractTests.exe
   with --package-smoke.
4. The release workflow does not visibly opt the packaged Release profile into
   test-enabled output.
5. Package-smoke validates plugin load and contracts but deliberately excludes
   Terminal debug/runtime selftests.

Phase 0 must reproduce or disprove this from a clean Release profile. If
confirmed, repair the production-safe smoke route. Do not rely on a stale
harness and do not enable ENABLE_TESTS in shipped product binaries.

## Fixed architecture decisions

1. **One active product lock.** Create
   External/TerminalEngine/GhosttyRuntimeLock.v1.json and validate it with
   Specs/Terminal/GhosttyRuntimeLock.schema.json. It is the sole current source
   for source/tool/dependency/build/overlay/license/import/export identity.
2. **No-op migration first.** Move the current pin into the lock and prove
   byte/semantic equivalence before changing to the candidate.
3. **Stable product entrypoint.** Tools/Build-GhosttyTerminalRuntime.ps1 and
   .build/TerminalEngine/Product/<Platform> remain the product surface.
4. **Isolated observation.** Candidate materialization and discovery go only
   beneath .build/TerminalEngineUpgrade/<candidate>/<platform>. They never
   publish Product or overwrite the product pairing header.
5. **Manual admission.** Only a reviewed commit may replace the canonical lock,
   overlay, adapter, and normative specs. Observation never edits tracked files.
6. **Exact ABI evidence.** The lock stores the full sorted export list, derived
   count and SHA-256, plus the sorted loader-required subset. Count/hash alone
   are insufficient diagnostics.
7. **Keep current runtime pairing unless deliberately revised.** Retain runtime
   identity schema/version 5 and existing pairing semantics if they can
   represent the candidate. Do not edit sealed historical tests merely to make
   a new schema number pass.
8. **Factual status language.** Status values are current, update-observed,
   security-review, required-manual, and unreachable. Never report latest when
   upstream could not be reached or when only an explicit candidate was checked.
9. **Read-only automation.** The status workflow has contents-read and
   actions-read permissions only, uses pinned actions, writes reports only to
   .build and workflow artifacts, and cannot mutate issues, branches, pull
   requests, or the lock.
10. **Cadence.** Run an advisory-only observation daily, a full upstream
    observation monthly, allow manual explicit-candidate dispatch, and run a
    fail-closed online preflight before release. An observed ordinary update is
    a warning; an applicable undisposed advisory fails scheduled/release
    checks; unreachable fails release preflight and never clears a prior
    security-review result.
11. **Scoped clipboard support.** Deny Kitty OSC 5522 transactions in the engine
    before buffering/callback through private option 41=false while preserving
    current bounded consent behavior for OSC 52/OSC 1337. Request metadata is
    not a protocol discriminator. Do not fake a success acknowledgement for
    asynchronous consent.
12. **No second generic lifecycle.** The Ghostty-specific module may reuse
    active archive, PE identity, import, export, text identity, sanitized
    process, and artifact-lock helpers. It must not copy or revive the dormant
    five-lane generic lock contract.

## Authoritative references

Read these before implementation and update the applicable normative files at
closeout:

- AGENTS.md
- .github/skills/cpp-build/SKILL.md
- .github/skills/cpp-modern-style/SKILL.md
- .github/skills/wil-raii/SKILL.md
- .github/skills/async-threading/SKILL.md
- .github/skills/error-handling/SKILL.md
- .github/skills/perf-validation/SKILL.md
- .github/skills/tooling-governance/SKILL.md
- Tools/README.md
- Tools/TerminalEngine/README.md
- Specs/Testing/Testing_ToolingGovernance.md
- Specs/Testing/Testing_PerformanceValidation.md
- Specs/Testing/Testing_SelfTests.md
- Specs/Testing/Testing_ValidationEvidence.md
- Specs/Build/Build_Toolchain.md
- Specs/Terminal/Terminal_EmbeddedPlugin.md
- Specs/Terminal/TerminalEngineDecision.md
- Specs/Installer/Installer_PortableZip.md
- Specs/Installer/Installer_Msi.md
- Specs/Installer/Installer_Msix.md

## Commands and expected results

| Purpose | Command | Expected success |
|---|---|---|
| Drift | git diff --stat e7475d3aaacc6dda0eeab58cd486c65d4074be4c..HEAD -- <paths from preamble> | no unexplained in-scope drift |
| Current runtime x64 | pwsh -File Tools/Build-GhosttyTerminalRuntime.ps1 -Platform x64 -Force -PassThru | exit 0; validated x64 Product identity |
| Current runtime ARM64 | pwsh -File Tools/Build-GhosttyTerminalRuntime.ps1 -Platform ARM64 -Force -PassThru | exit 0; validated ARM64 Product identity |
| Offline x64 | pwsh -File Tools/Build-GhosttyTerminalRuntime.ps1 -Platform x64 -Offline -Force -PassThru | exit 0 without network |
| Offline reuse | pwsh -File Tools/Build-GhosttyTerminalRuntime.ps1 -Platform x64 -Offline -PassThru | exit 0 and verified reuse |
| Debug x64 | .\build.ps1 -Configuration Debug -Platform x64 | exit 0, zero errors |
| Test-enabled Release x64 | set RSBuildEnableTests=true for the process, then .\build.ps1 -Configuration Release -Platform x64 | exit 0, zero errors |
| ASan x64 | .\build.ps1 -Configuration "ASan Debug" -Platform x64 | exit 0, zero errors |
| Debug ARM64 | .\build.ps1 -Configuration Debug -Platform ARM64 | exit 0, zero errors |
| Terminal contracts | .\.build\x64\Debug\PluginContractTests.exe --terminal-selftests | all Terminal cases pass |
| Focused commands | .\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_embedded_terminal_ | all matching cases pass |
| Pester | Invoke-Pester Tools\Tests\GhosttyRuntimeLifecycle.Tests.ps1,Tools\Tests\GhosttyRuntimeStatus.Tests.ps1,Tools\Tests\TerminalRuntimeExportIdentity.Tests.ps1,Tools\Tests\TerminalEngineRuntimeImportPolicy.Tests.ps1,Tools\Tests\TerminalPluginSourceContracts.Tests.ps1,Tools\Tests\BuildEvidence.Tests.ps1,Tools\Tests\BuildReproducibility.Tests.ps1,Tools\Tests\ReleaseWorkflowPolicy.Tests.ps1,Tools\Tests\VcpkgInstallSafety.Tests.ps1,Tools\Tests\TestHarnessSourceContracts.Tests.ps1 | zero failed |
| Tool inventory | .\Tools\Get-ToolInventory.ps1 -FailOnFindings | zero findings |
| Spec inventory | .\Tools\Get-SpecInventory.ps1 -FailOnFindings | zero findings |
| Full closeout | .\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -Configuration Debug -Platform x64 | terminal status passed and repository status passed |
| Diff hygiene | git diff --check | exit 0 |

For test-enabled Release, use process-scoped environment assignment and remove
it in a finally block. Do not make it a machine/user environment setting.

## Suggested executor toolkit

- Use the cpp-build skill for build and test orchestration.
- Use the perf-validation skill before adding the baseline scenario and before
  accepting performance evidence.
- Use the tooling-governance skill before any Tools or workflow change.
- Use cpp-modern-style, wil-raii, async-threading, and error-handling for C++
  adapter work.
- If a crash, Fatal Error, minidump, selftest hang, or unexplained timeout occurs,
  switch to red-salamander-crash-forensics before changing code.

## Scope

The following is the maximum authorized implementation surface. A phase should
touch only its listed subset. If another production path is required, STOP,
record why, and amend this plan before editing it.

### Create

- External/TerminalEngine/GhosttyRuntimeLock.v1.json
- Specs/Terminal/GhosttyRuntimeLock.schema.json
- Specs/Terminal/GhosttyRuntimeStatus.schema.json
- Specs/Terminal/GhosttyRuntimeCandidateReview.schema.json
- Specs/Terminal/GhosttyRuntimeCandidateDescriptor.schema.json
- Specs/Terminal/GhosttyRuntimeCandidateOverlayReview.schema.json
- Specs/Terminal/GhosttyRuntimeNativeQualification.schema.json
- Tools/TerminalEngine/GhosttyRuntimeLifecycle.psm1
- Tools/TerminalEngine/GhosttyCandidateReview.psm1
- Tools/TerminalEngine/GhosttyCandidateOverlay.psm1
- Tools/Get-GhosttyTerminalRuntimeStatus.ps1
- Tools/New-GhosttyTerminalRuntimeCandidateReview.ps1
- Tools/New-GhosttyTerminalRuntimeCandidateOverlayReview.ps1
- Tools/Tests/GhosttyRuntimeLifecycle.Tests.ps1
- Tools/Tests/GhosttyRuntimeStatus.Tests.ps1
- Tools/Tests/GhosttyCandidateReview.Tests.ps1
- Tools/Tests/GhosttyCandidateOverlay.Tests.ps1
- .github/workflows/ghostty-runtime-status.yml
- .github/workflows/ghostty-runtime-native-qualification.yml
- .github/workflows/ghostty-runtime-product-qualification.yml

### Update: product lock/build/status

- Tools/Build-GhosttyTerminalRuntime.ps1
- Tools/TerminalEngine/Patches/Ghostty-WindowsPrivateRuntime.patch
- Tools/TerminalEngine/README.md
- Tools/TerminalEngine/TerminalEngineArchive.psm1 only if a missing reusable
  archive primitive is proved
- Tools/TerminalEngine/TerminalEnginePeIdentity.psm1 only if candidate evidence
  requires a missing PE primitive
- Tools/TerminalEngine/TerminalRuntimeExportIdentity.psm1 only if full sorted
  export diagnostics are not already exposed
- Tools/TerminalEngine/TerminalRuntimeImportPolicy.psm1 only if an exact
  candidate import diagnostic is missing
- Tools/Modules/Build/BuildEvidence.psm1
- Tools/tool-inventory.json
- Tools/README.md
- .github/workflows/build-reusable.yml
- .github/workflows/release.yml

### Update: adapter and tests

- Plugins/Terminal/Terminal.cpp
- Plugins/Terminal/Terminal.h
- Plugins/Terminal/TerminalInternal.cpp
- Plugins/Terminal/TerminalInternal.h
- Plugins/Terminal/TerminalRuntimeLoader.cpp
- Plugins/Terminal/TerminalRuntimeLoader.h
- Plugins/Terminal/TerminalSecurity.cpp
- Plugins/Terminal/TerminalSecurity.h
- Plugins/Terminal/TerminalVt.cpp
- Plugins/Terminal/TerminalVt.h
- Plugins/Terminal/TerminalKittyImagePipeline.cpp
- Plugins/Terminal/TerminalKittyImagePipeline.h
- Tests/PluginContractTests/PluginContractTests.cpp
- Tools/Tests/TerminalDependencyLifecycle.Tests.ps1
- Tools/Tests/TerminalRuntimeExportIdentity.Tests.ps1
- Tools/Tests/TerminalEngineRuntimeImportPolicy.Tests.ps1
- Tools/Tests/TerminalPluginSourceContracts.Tests.ps1
- Tools/Tests/BuildEvidence.Tests.ps1
- Tools/Tests/BuildReproducibility.Tests.ps1
- Tools/Tests/ReleaseWorkflowPolicy.Tests.ps1
- Tools/Tests/HistoricalTerminalProfile.Tests.ps1 only for a separation
  assertion; never change sealed historical identities or expected hashes

### Conditional package repair

Touch these only if Phase 0 reproduces the clean Release/package-smoke mismatch:

- Tests/PluginContractTests/PluginContractTests.vcxproj
- Tools/Modules/Packaging/PortablePackageSmoke.psm1
- Installer/zip/build-zip.ps1
- Tools/Tests/RedSalamanderPluginDeployment.Tests.ps1

The repair must provide a production-safe, package-local loader/contract smoke
artifact or equivalent host that is built intentionally for packaging. It must
not turn on product test code or consume a stale executable from another
profile.

### Update: durable contracts and evidence

- Specs/Terminal/Terminal_EmbeddedPlugin.md
- Specs/Build/Build_Toolchain.md
- Specs/Testing/Testing_PerformanceValidation.md
- Specs/Testing/Testing_SelfTests.md
- Specs/Testing/Testing_ValidationEvidence.md
- Specs/Installer/Installer_PortableZip.md only if package behavior changes
- Specs/NormativeConsistency.json
- Specs/TestRuns/<MachineHash>/Terminal/<RunId> for new bounded evidence
- Specs/TestRuns/<MachineHash>/Commands/<RunId> only when a Commands scenario is
  measured
- Specs/Plans/WIP/README.md while active
- This plan while executing; move it to Specs/Plans/Done only at closeout

### P0/P11 validation-remediation exception reconciled at P12 review

The complete P12 diff review found that earlier gate-driven repairs were fully
evidenced in this plan but not copied into the maximum-surface list when their
blocking runs occurred. They are retained as narrowly authorized upgrade
qualification work, not as a general expansion into these domains:

- baseline VT corpus/test integration:
  `Plugins/Terminal/Terminal.vcxproj`,
  `Plugins/Terminal/TerminalVtUpgradeTestContract.h`,
  `Plugins/Terminal/TerminalVtUpgradeTests.cpp`,
  `Specs/Terminal/TerminalVtUpgradeCorpus.v1.json`, and
  `Specs/Terminal/TerminalVtUpgradeEvidence.schema.json`;
- receipt-bound shipping/runtime staging and clean build profiles:
  `Directory.Build.targets`, `RuntimeDependencies.props`,
  `Tests/Directory.Build.targets`,
  `Tools/Modules/Packaging/RuntimeDependencies.psm1`, `README.md`, and
  `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`;
- native-runner dependency pin/bootstrap qualification:
  `vcpkg-install.ps1`, `vcpkg-tool.json`, `vcpkg.json`,
  `Tools/Modules/Build/VcpkgInstallSafety.psm1`,
  `Tools/Tests/VcpkgInstallSafety.Tests.ps1`,
  `Tools/Tests/ToolingGovernance.Tests.ps1`, and
  `Specs/FileSystem/FileSystem_S3.md`;
- failures first exposed only by the mandatory P11 Fresh Full gate:
  `Common/UnicodeClipboard.h`, `Specs/Core/Core_SharedHelpers.md`,
  `Plugins/FileSystem/FileSystem.FileOps.cpp`, the four touched
  `RedSalamander/SelfTest/Commands/Commands.SelfTest.*.cpp` fixtures, and the
  bounded continuation record under
  `Specs/TestRuns/4cb089111a23/Continuation/20260826_213953_fileops_recursive_copy_crash/`;
- native qualification crash diagnosis: the bounded continuation record under
  `Specs/TestRuns/GitHubARM64/Continuation/20260826_030920_ghostty_native_test_access_violation/`.

Each exception has a matching P11 Evidence row, focused verification, and final
Fresh Full coverage. No other File Operations, clipboard, dependency, or test-
harness work is admitted by this reconciliation.

### P12 archive-transition mechanics

- `Tools/Modules/Tooling/SpecInformationArchitecture.psm1` may add only this
  exact completed-plan path to the protected Done-history additional-path list.
  The checkpoint must not weaken baseline comparison or admit another path.

### Explicitly out of scope

- Specs/Plans/Done and all sealed Terminal engine selection evidence
- External/TerminalEngine candidate-universe, evaluation, and decision JSON
- Tools/Evaluate-TerminalEngines.ps1 and Gate-0/Round-4 scripts/workflows
- Product clipboard-read support
- New terminal emulator selection
- Unrelated Terminal UI or File Operations code
- Automated lock updates, issue creation, branch creation, or pull requests

## Ownership and sequencing

- I10 owns global Tools classification, help, inventory semantics, and placement.
  This plan adds only Ghostty runtime entries and tests under that contract.
- I12 owns its current Terminal focus/input/reaper fixes. Before Phase 6, confirm
  I12 has no concurrent edits in Plugins/Terminal. Rebase or serialize; do not
  resolve semantic conflicts by choosing whichever side compiles.
- The performance measurement contract I5 owns repository-wide evidence
  semantics. This plan adds a Terminal scenario using that contract, not a
  competing evidence format.
- Packaging scripts take a repository-wide packaging lock. Run ZIP, MSI, and
  MSIX serially. x64 and ARM64 runtime/build work may use their existing
  profile-scoped coordination.

## Git workflow

- Create branch codex/ghostty-lib-vt-upgrade-20260825 from the reviewed base.
- Do not push or open a pull request unless the operator requests it.
- P0 through P12 each require the separate implementation commit and exact
  subject listed in Required phase commits. Never combine phase commits.
- After each phase implementation commit, make the required plan-only checkpoint
  commit that records its full SHA and checks the verified phase before starting
  the next phase.
- Start each phase from a clean worktree. Focused repair commits inside a phase
  are allowed, but record all of them under that phase and finish with its named
  phase boundary commit.
- Preserve imperative commit messages and do not amend a completed phase after
  its checkpoint is recorded.
- Never commit .build candidate material, downloaded archives, tokens, workflow
  responses containing credentials, or unbounded logs.

## Execution phases

### Phase 0 — Establish a clean baseline and resolve the package-smoke premise

**Goal:** prove the current pin, current packages, and pre-upgrade performance
from clean, receipt-bound inputs before changing any runtime identity.

1. Run the drift command and record HEAD, worktree status, Windows build, CPU,
   architecture, compiler, SDK, Zig, and vcpkg identities in the baseline run
   metadata. Do not include usernames or absolute home paths.
2. Force-build the current x64 and ARM64 runtimes, then perform x64 offline
   force and offline reuse. Save the returned runtime identities.
3. Build Debug x64 and a test-enabled Release x64. Run the existing Terminal
   selftests and embedded-terminal Commands cluster.
4. Add the deterministic measurement-only VT upgrade corpus before changing the
   pin. It must include:
   - printable UTF-8 and wide-cell text;
   - SGR/color and hyperlink sequences;
   - cursor, erase, scroll, alternate-screen, resize, and snapshot operations;
   - OSC 52 payloads below, at, and above 256 KiB;
   - split Kitty image transfers and pending completion;
   - DCS bytes above 0x7F;
   - repeated render-state/snapshot reads.
5. The corpus must record input byte count, output/PTY byte digest, terminal
   snapshot digest, VT write duration, snapshot duration, hosted render
   duration, Kitty queue/conversion/upload/accounting, and peak retained memory.
   Behavior that is intentionally corrected must have a named expected delta;
   all other output/snapshot digests must match.
6. Collect at least 200 samples for each p95 claim. Freeze the current-pin p50,
   p95, sample count, exact corpus digest, and thresholds in the baseline
   Terminal run before candidate measurement.
7. From a clean Release x64 build with RSBuildEnableTests absent/false, run:

   ~~~powershell
   .\Installer\zip\build-zip.ps1 -Configuration Release -Platform x64 -BuildNumber <approved-build-number>
   ~~~

   Inspect the fresh extraction and prove whether the invoked
   PluginContractTests.exe came from that exact Release receipt. A missing or
   stale harness is a confirmed defect, not a reason to skip smoke.
8. If confirmed, implement the conditional package repair now and add a test
   proving the smoke executable/host is produced by and bound to the packaged
   Release profile while product binaries remain tests-disabled. If disproved,
   add a focused test documenting the actual valid production path.

**Verify**

~~~powershell
pwsh -File Tools/Build-GhosttyTerminalRuntime.ps1 -Platform x64 -Force -PassThru
pwsh -File Tools/Build-GhosttyTerminalRuntime.ps1 -Platform ARM64 -Force -PassThru
pwsh -File Tools/Build-GhosttyTerminalRuntime.ps1 -Platform x64 -Offline -Force -PassThru
pwsh -File Tools/Build-GhosttyTerminalRuntime.ps1 -Platform x64 -Offline -PassThru
.\build.ps1 -Configuration Debug -Platform x64
.\.build\x64\Debug\PluginContractTests.exe --terminal-selftests
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_embedded_terminal_
$env:RS_VALIDATION_PLATFORM = 'x64'
$env:RS_VALIDATION_CONFIGURATION = 'Debug'
Invoke-Pester Tools\Tests\BuildReproducibility.Tests.ps1,Tools\Tests\RedSalamanderPluginDeployment.Tests.ps1
~~~

The explicit validation profile is part of the command contract: the targeted
deployment test mutates only that selected build profile and fails closed when
either variable is omitted.

Expected: every command exits 0; current product identities validate; offline
uses no network; baseline archive validates; package smoke is either proved
receipt-correct or repaired with a focused green regression. Do not advance with
an unresolved package premise.

### Phase 1 — Introduce the active lock as a no-op migration

**Goal:** remove duplicated product-definition constants from the builder
without changing the pin or accepted runtime.

1. Define GhosttyRuntimeLock.schema.json with additionalProperties false,
   bounded strings/arrays, lowercase full SHA-1/SHA-256 patterns, unique sorted
   lists, explicit schemaVersion 1, and these required sections:
   - upstream repository, commit, tree, archive URL/name/SHA-256, version;
   - Zig version, platform archives, URLs, SHA-256 values, executable identity;
   - dependency archives/content identities and license extraction paths;
   - overlay path, normalized-text SHA-256, ordered concern IDs;
   - exact build target, arguments, environment allowlist, supported platforms;
   - exact import modules and full sorted exports;
   - derived export count/hash and full sorted loader-required exports;
   - platform machine values and reviewed size envelopes;
   - license/notice inputs and output names;
   - runtime identity schema/pairing values.
2. Create the lock from the current embedded constants. Do not update the
   Ghostty commit yet.
3. Add positive validation and one negative fixture per invariant in
   GhosttyRuntimeLifecycle.Tests.ps1: missing property, extra property, uppercase
   hash, unsorted/duplicate export, derived hash mismatch, loader subset not in
   exports, bad machine, path escape, unsupported schema, overlay digest
   mismatch, and missing license input.
4. Make BuildEvidence hash the real lock, overlay, RuntimeDependencies.props,
   and any other direct identity input. Remove the nonexistent
   TerminalEngine.lock.json reference.
5. Add the lock/schema/module to the existing build-reusable cache closure.
6. Update tool inventory/help/spec entries in the same commit.

**Verify**

~~~powershell
Invoke-Pester Tools\Tests\GhosttyRuntimeLifecycle.Tests.ps1,Tools\Tests\BuildEvidence.Tests.ps1,Tools\Tests\ReleaseWorkflowPolicy.Tests.ps1
.\Tools\Get-ToolInventory.ps1 -FailOnFindings
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
~~~

Expected: zero failed and zero findings; a normalized digest of the lock is
stable across repeated reads; all negative fixtures fail with a field-specific
message.

### Phase 2 — Refactor the product builder and prove no-op equivalence

**Goal:** have the stable product entrypoint consume the active lock through
shared Ghostty-specific lifecycle functions.

1. Put lock parsing, schema validation, path containment, normalized identities,
   source/tool/dependency restore, overlay validation, candidate isolation,
   build invocation, artifact inspection, license staging, and manifest creation
   in GhosttyRuntimeLifecycle.psm1.
2. Reuse active TerminalEngineArchive, TerminalEnginePeIdentity,
   TerminalRuntimeImportPolicy, TerminalRuntimeExportIdentity,
   TerminalEngineTextIdentity, and SanitizedEnvironment helpers. Retain the
   reviewed Ghostty repository mutex plus session canonical-drive lease rather
   than acquiring an artifact-profile lock. Do not duplicate them locally.
3. Keep Build-GhosttyTerminalRuntime.ps1 as the documented public wrapper.
   Its default lock is the canonical tracked lock. Product mode alone may write
   .build/TerminalEngine/Product and the pairing header.
4. Add an internal candidate-discovery operation that requires:
   - an explicit full commit and tree;
   - an output root under .build/TerminalEngineUpgrade;
   - a non-Product output;
   - no product pairing-header write;
   - no fallback to the canonical lock when a candidate descriptor is missing.
5. Fail closed on lock/schema/overlay/license/import/export/machine/size mismatch,
   source archive drift, uncontained output, or an independently running
   executable that owns the exact product DLL path. Never kill an independent
   process.
6. Force-build current x64 and ARM64 from the lock twice. Compare the old and
   new semantic identities, import/export sets, license identities, pairing
   values, and deterministic manifests. Explain any raw PE byte delta; no
   unexplained delta is allowed.

**Reviewed P2 coordination amendment (2026-08-25):** the initial generic
artifact-lock wording was incorrect for this nested dependency builder.
`Directory.Build.targets` launches the Ghostty command beneath an MSBuild
operation that already owns the app artifact profile; acquiring that same
profile from the nested process would conflict with the immediate-child
delegation boundary, while the Ghostty outputs are outside the profile roots.
The repository tooling contract specifically assigns `.build/TerminalEngine`
to the repository-derived Ghostty mutex and reserves the session `R:` lease for
the rebuild interval. P2 therefore retains those two reviewed locks and reuses
the archive, identity, import/export, and contained-process helpers.

**Verify**

~~~powershell
pwsh -File Tools/Build-GhosttyTerminalRuntime.ps1 -Platform x64 -Force -PassThru
pwsh -File Tools/Build-GhosttyTerminalRuntime.ps1 -Platform ARM64 -Force -PassThru
pwsh -File Tools/Build-GhosttyTerminalRuntime.ps1 -Platform x64 -Offline -Force -PassThru
pwsh -File Tools/Build-GhosttyTerminalRuntime.ps1 -Platform x64 -Offline -PassThru
Invoke-Pester Tools\Tests\GhosttyRuntimeLifecycle.Tests.ps1,Tools\Tests\TerminalDependencyLifecycle.Tests.ps1,Tools\Tests\TerminalRuntimeExportIdentity.Tests.ps1,Tools\Tests\TerminalEngineRuntimeImportPolicy.Tests.ps1
~~~

Expected: all commands exit 0; the canonical lock still names the old commit;
current semantic identities are unchanged; candidate tests prove Product and
pairing-header paths remain untouched.

### Phase 3 — Add the read-only update and advisory process

**Goal:** make the next update observable and reproducible without turning
observation into mutation.

1. Define GhosttyRuntimeStatus.schema.json. Required fields:
   - schemaVersion, generatedUtc, mode, result, and ordered reason codes;
   - canonical lock path/SHA-256 and locked upstream/toolchain summary;
   - observation source, commit, tree, version, ancestry counts, and timestamps;
   - explicit candidate commit/tree when mode is Candidate;
   - advisory source records, affected/not-affected/unknown disposition,
     identifiers, URLs, and retrieval timestamps;
   - network attempt outcome without credentials or response headers;
   - nextAction chosen from none, review-update, review-security,
     freeze-candidate, retry-observation, and manual-admission.
2. Implement Get-GhosttyTerminalRuntimeStatus.ps1 with modes Locked, Online, and
   Candidate. Candidate requires a full explicit commit. All output paths must
   resolve beneath .build; default to
   .build/TerminalEngineStatus/<timestamp>/status.json.
3. Locked performs no network and validates only canonical inputs. Online
   observes the upstream default branch and official advisory sources.
   Candidate resolves exactly the supplied commit and never substitutes default
   branch head.
4. Use factual result semantics:
   - current: observation equals the lock and no applicable advisory exists;
   - update-observed: a descendant/different commit is observed, with no
     automatic recommendation or lock edit;
   - security-review: a primary advisory may affect the lock or disposition is
     unresolved;
   - required-manual: provenance, ancestry, patch, ABI, or policy needs a human;
   - unreachable: a required online source could not be authenticated/reached.
5. Add deterministic mocked-network tests for default-branch advance,
   unrelated history, deleted/refused candidate, advisory applies, advisory
   does not apply, malformed response, rate limit, timeout, offline mode,
   output escape, and repeated canonical serialization.
6. Add ghostty-runtime-status.yml:
   - daily advisory-only scheduled job;
   - monthly full observation;
   - manual Locked/Online/Candidate dispatch with explicit candidate input;
   - pinned action SHAs and permissions limited to contents read and actions read;
   - artifact upload and step summary, no issue/PR/branch permissions;
   - warning for update-observed;
   - failure for security-review on scheduled and release preflight;
   - failure for unreachable on release preflight.
7. Add a release preflight call before artifact build. It must bind the emitted
   report to the current lock SHA and current source commit so a report from
   another run cannot be replayed.

**Verify**

~~~powershell
pwsh -File Tools/Get-GhosttyTerminalRuntimeStatus.ps1 -Mode Locked
pwsh -File Tools/Get-GhosttyTerminalRuntimeStatus.ps1 -Mode Candidate -CandidateCommit 9c3ec931d64561a8407dde7ac984ce156ae91539
Invoke-Pester Tools\Tests\GhosttyRuntimeStatus.Tests.ps1,Tools\Tests\ReleaseWorkflowPolicy.Tests.ps1,Tools\Tests\ToolingGovernance.Tests.ps1
.\Tools\Get-ToolInventory.ps1 -FailOnFindings
~~~

Expected: Locked is network-free and current; Candidate records exactly the
frozen commit/tree and update-observed or required-manual as appropriate; tests
prove zero tracked-file mutation and zero write outside .build.

### Phase 4 — Freeze and classify the exact candidate

**Goal:** turn the observed commit into a reviewable, immutable candidate
without touching Product.

1. Re-resolve the candidate from the official Ghostty repository. Verify the
   full commit, tree, archive SHA-256, version, signature/provenance evidence
   available from the official source, and ancestry from the current lock.
2. Regenerate the path-filtered diff inventory and classify every changed file
   under public ABI, parser/state, clipboard/security, Kitty graphics,
   allocation/memory, snapshot/serialization, build/toolchain, licenses, tests,
   or irrelevant-to-lib-vt. Store the bounded report under .build, not Specs.
3. Diff all public include/ghostty/vt headers and all exported C declarations.
   Generate added/removed/changed symbol and enum/value reports. Explicitly
   compare every one of the 72 loader-required names.
4. Review commits that touch bounds, pointer arithmetic, ownership, integer
   conversion, compressed image data, OSC/DCS/APC parsing, paste/clipboard, and
   snapshot encoding. Record primary commit links and affected product paths;
   do not label them vulnerabilities without an advisory.
5. Create an untracked candidate descriptor under
   .build/TerminalEngineUpgrade/9c3ec931d64561a8407dde7ac984ce156ae91539.
   It must bind commit, tree, archive, toolchain, dependency inputs, current
   proposed overlay digest, and observation report digest.

**Verify**

~~~powershell
git -C <isolated-ghostty-source> merge-base --is-ancestor b2d44625908eb56aee299d8511d803bc11ab79fc 9c3ec931d64561a8407dde7ac984ce156ae91539
git -C <isolated-ghostty-source> rev-list --left-right --count b2d44625908eb56aee299d8511d803bc11ab79fc...9c3ec931d64561a8407dde7ac984ce156ae91539
git -C <isolated-ghostty-source> rev-parse 9c3ec931d64561a8407dde7ac984ce156ae91539^{tree}
~~~

Expected: ancestry exits 0; counts are 0 and 717; tree is
359872ab74fe1ecc96428133c0546be3cb969fdd; all reports identify the exact
candidate and remain under .build.

### Phase 5 — Rederive the Windows private-runtime overlay

**Goal:** preserve each product-owned concern against candidate source rather
than mechanically forcing the old patch.

1. Start from a pristine candidate tree. For each current overlay hunk, record
   concern, upstream change, retained need, new implementation, and focused
   test. The concern list must include:
   - Windows shared lib-vt LLVM/LLD/strip production identity behavior;
   - an unshipped static companion is acceptable and is not an independent
     product concern;
   - fail-closed Kitty OSC 5522 clipboard-write disable before buffering;
   - 4,096 placement-count bound;
   - 1 MiB placement-map allocation bound;
   - unchanged-state rejection at limit plus replacement allowance;
   - screen/reset preservation and pin release;
   - placement count/allocation/limit queries;
   - C header and Zig enum parity.
2. Reimplement against candidate APIs. Do not preserve dead code shapes merely
   to minimize diff size.
3. Allocate private option/data values only after scanning the complete frozen
   candidate public ranges. Private terminal options must not use 31 through
   39; placement limits use 40 and the Kitty clipboard-write gate uses 41. Add
   compile-time Zig checks and C++/Pester checks that enumerate every new private
   option and prove uniqueness plus exact C/Zig parity.
4. Use candidate storage APIs for capacity/accounting. Preserve atomic
   apply-to-all-screens behavior: validation of all screens precedes mutation of
   any screen.
5. Preserve replacement semantics and ensure failed new placement admission
   does not increment IDs, alter generation, leak pins, or mutate maps.
6. Generate one normalized patch from the pristine candidate and update the
   ordered concern list/digest. No second patch, fork, or prebuilt DLL is
   allowed.
7. Keep the overlay reviewable: if it exceeds 16 upstream files, 2,500 changed
   lines, or five independent product concerns, STOP for architecture review.

**Verify**

~~~powershell
git -C <pristine-candidate-source> apply --check <absolute-path-to-Ghostty-WindowsPrivateRuntime.patch>
git -C <pristine-candidate-source> apply <absolute-path-to-Ghostty-WindowsPrivateRuntime.patch>
git -C <pristine-candidate-source> diff --check
Invoke-Pester Tools\Tests\GhosttyRuntimeLifecycle.Tests.ps1,Tools\Tests\TerminalPluginSourceContracts.Tests.ps1
~~~

Expected: clean apply, clean diff, no public/private enum collision, exact C/Zig
parity, and all retained-limit tests green.

### Phase 6 — Build and test the candidate in isolation

**Goal:** obtain trustworthy candidate artifacts without changing the current
product pair.

1. Build candidate discovery x64 and ARM64 only beneath the candidate root.
   Assert before and after that Product manifests, DLLs, import libraries, and
   pairing headers have unchanged hashes.
2. Run upstream test-lib-vt, test-lib-vt-build, and test-lib-vt-schema for x64.
3. Run the same targets on native ARM64 Windows. Cross-compiling the ARM64 DLL
   is useful but cannot satisfy schema/load execution.
4. Inspect actual DLL exports. Save full sorted names, count, derived hash,
   removed/added names, loader-required subset, PE machine, imports, binary size,
   build info, dependency and license identities.
5. Explain every import/export/size/license change relative to the current pin.
   Never update an expectation to the observed result without a source-level
   explanation and reviewer acknowledgement.
6. Produce a proposed complete candidate lock under the candidate .build root.
   Rebuild from that proposed lock in a clean isolated directory and require all
   exact checks to pass.

**Verify**

~~~powershell
pwsh -File Tools/Build-GhosttyTerminalRuntime.ps1 -Platform x64 -CandidateLockPath <candidate-lock-under-.build> -CandidateOutputRoot <x64-candidate-root> -Force -PassThru
pwsh -File Tools/Build-GhosttyTerminalRuntime.ps1 -Platform ARM64 -CandidateLockPath <candidate-lock-under-.build> -CandidateOutputRoot <arm64-candidate-root> -Force -PassThru
~~~

Expected: both isolated builds exit 0; candidate lock checks are exact; Product
and pairing hashes are unchanged; upstream x64 and native ARM64 tests pass.
Parameter names shown here are the required public semantics; if implementation
uses one equivalent Candidate switch/object, update this plan and command table
before execution.

### Phase 7 — Adapt loader, mode access, and ABI guards

**Goal:** compile against and correctly call the candidate C API.

1. Stage a candidate ABI build property and replace terminalModeGet with the
   candidate terminalGet entrypoint in TerminalRuntimeLoader.h/.cpp. Keep the
   locked Product compatibility branch and exact loader failure diagnostics
   through P9 so P7 does not admit an unreviewed runtime ahead of P10.
2. Implement mode reads using the frozen two-field
   `GhosttyTerminalModeConfig { mode, value }`: zero-initialize it, set `mode`,
   call `ghostty_terminal_get(..., GHOSTTY_TERMINAL_DATA_MODE, &config)`, and
   validate success before reading `value`. Do not invent a size field for this
   non-versioned structure.
3. Update the candidate required-export subset and all source contracts. The
   canonical Product subset continues to require its old symbol until P10.
   Do not remove an export merely because a test says it is absent; map every
   removed loader symbol to an intentional replacement.
4. Add compile-time or selftest checks for all size-versioned structures used by
   the adapter. Read optional candidate fields only when the incoming size
   covers them.
5. Retain runtime loader atomicity: a partially resolved DLL never becomes
   active, all function pointers reset on failure, and the default Product
   build still rejects an old-header/new-DLL mismatch. Compile the isolated
   loader translation unit against the candidate headers, then execute a
   direct candidate ABI probe; the full candidate Terminal build follows only
   after the clipboard and Kitty adaptations in P8/P9.
6. Run exact key/mouse/paste/mode selftests, including alternate screen,
   bracketed paste, focus reporting, mouse modes, and DCS high-byte regression.

**Verify**

~~~powershell
.\build.ps1 -Configuration Debug -Platform x64
.\.build\x64\Debug\PluginContractTests.exe --terminal-selftests
Invoke-Pester Tools\Tests\TerminalPluginSourceContracts.Tests.ps1,Tools\Tests\TerminalRuntimeExportIdentity.Tests.ps1
# Also compile TerminalRuntimeLoader.cpp with RSTerminalGhosttyAbi=Candidate and
# RSTerminalGhosttyIncludeRoot=<qualified-candidate-include>, then execute the
# isolated candidate mode/DCS probe recorded in the Evidence log.
~~~

Expected: zero errors/failures; the candidate loader/runtime path has no
dependency on ghostty_terminal_mode_get; the locked Product compatibility path
is unchanged; DCS high-byte case passes. P10 removes the Product-only legacy
reference when it atomically admits the candidate lock.

### Phase 8 — Adapt clipboard writes without weakening consent

**Goal:** implement the candidate callback contract with truthful,
nonblocking, bounded behavior.

1. Change Terminal::GhosttyClipboardWrite and CaptureRuntimeClipboard to the
   candidate void callback type.
2. Validate write pointer and size before each field. Copy only bounded data
   needed after callback return; never retain upstream pointers.
3. Before accepting terminal input, set
   `GHOSTTY_TERMINAL_OPT_CLIPBOARD_WRITE_MAX_BYTES` to 256 KiB and private
   `GHOSTTY_TERMINAL_OPT_KITTY_CLIPBOARD_WRITE_ENABLED` (41) to false. Treat
   inability to set either option as terminal initialization failure.
4. Preserve current OSC 52/approved image behavior:
   - deny means no clipboard mutation and a negative reply when supported;
   - allow writes exactly the bounded copied bytes;
   - ask posts a helper-owned payload to the UI and returns without blocking;
   - no clipboard API or UI wait occurs under the terminal lock.
5. Request metadata is not a protocol discriminator: an unnamed/no-password
   Kitty OSC 5522 write has the same callback metadata as OSC 52. Therefore the
   private engine option must reject Kitty writes with `ENOSYS` before buffering
   or callback invocation. Prove OSC 52/OSC 1337 still reach the callback. The
   callback may use bounded asynchronous consent for those legacy no-ack paths
   and must invoke a present reply at most once; it must never enqueue a denied
   Kitty operation or later report success for it.
6. Keep clipboard read callback null. A runtime attempt to require it is
   unsupported and must fail closed.
7. Add cases for undersized request, missing reply field, null data, limit-1,
   limit, limit+1, invalid encoding, allow, deny, ask, teardown before UI
   response, reply-at-most-once, Kitty denial, and no-read callback.

**Verify**

~~~powershell
.\build.ps1 -Configuration Debug -Platform x64
.\.build\x64\Debug\PluginContractTests.exe --terminal-selftests
Invoke-Pester Tools\Tests\TerminalPluginSourceContracts.Tests.ps1
~~~

Expected: all cases pass; above-limit input never reaches a UI payload or system
clipboard; unsupported Kitty transactions receive engine-level `ENOSYS` before
the host callback; no lock/UI deadlock or teardown leak occurs.

### Phase 9 — Correct pending and animated Kitty image state

**Goal:** ensure candidate asynchronous image semantics never become permanent
false rejection or stale rendering.

1. Model engine-pending separately from adapter conversion pending and permanent
   rejection. GHOSTTY_NO_VALUE with a valid image is retryable, not Rejected.
2. Track both storage generation and image generation. On storage-generation
   change, retry engine-pending images. On image-generation change, invalidate
   converted/cache content and queue the new frame.
3. Keep stale worker results rejectable by request ID plus storage/image
   generation. A stale completion must not resurrect a replaced/deleted/evicted
   image.
4. Preserve queue coalescing, stop-token cancellation, maximum one active worker,
   memory accounting, and the Close quiet point.
5. Add deterministic tests for:
   - split transfer pending then complete;
   - pending replaced with same image ID;
   - pending deleted and evicted;
   - retransmit after rejection;
   - animation image-generation advance;
   - reset and alternate-screen changes;
   - stale conversion completion after replace/delete;
   - placement limit and exact allocated-byte queries;
   - 4,096 accepted, 4,097th unique rejected with unchanged state;
   - replacement at the limit;
   - pin release and zero queued/active/ready bytes on Close;
   - resize without recreate or unchanged reupload.
6. Do not query or copy large image payloads while holding a UI/render lock
   longer than the existing bounded capture path permits.

**Verify**

~~~powershell
.\build.ps1 -Configuration Debug -Platform x64
.\.build\x64\Debug\PluginContractTests.exe --terminal-selftests
~~~

Expected: all lifecycle cases pass deterministically; no pending image becomes
permanently rejected merely because payload data was not ready; all memory and
quiet-point invariants hold.

### Phase 10 — Measure, review, and admit the candidate

**Goal:** replace the canonical lock only after behavior, security, ABI, and
performance evidence are complete.

1. Run the frozen VT corpus against the isolated candidate on the same machine
   and configuration as the baseline.
2. Require:
   - exact non-intentional PTY output and snapshot digests;
   - named expected delta for corrected DCS high-byte behavior;
   - hosted render p95 no more than 1.10 times baseline;
   - VT write and snapshot p95 no more than 1.10 times baseline unless a
     separately reviewed budget is added before admission;
   - at least 200 samples for every p95;
   - placement/cache/pipeline/close/resize invariants from Current state.
3. Review the source diff, overlay concerns, exact exports/imports, clipboard
   denial policy, pending-image behavior, upstream test results, native ARM64
   result, license closure, package-smoke resolution, and performance archive.
4. Only after review, replace the canonical lock with the complete candidate
   lock and update the overlay/adapter/specs in the same coherent commit.
5. Force-build canonical x64 and ARM64 Product. Then offline-force x64 and
   offline-reuse x64. Require identity manifests to bind the new lock digest.
6. Confirm the loader-required export list is a subset of actual exports, exact
   import modules are allowed, machine values are correct, sizes are inside
   explained reviewed envelopes, and every license file is present.
7. The old lock and overlay remain recoverable from the parent commit. Do not
   delete shared caches as part of admission.

**Verify**

~~~powershell
pwsh -File Tools/Build-GhosttyTerminalRuntime.ps1 -Platform x64 -Force -PassThru
pwsh -File Tools/Build-GhosttyTerminalRuntime.ps1 -Platform ARM64 -Force -PassThru
pwsh -File Tools/Build-GhosttyTerminalRuntime.ps1 -Platform x64 -Offline -Force -PassThru
pwsh -File Tools/Build-GhosttyTerminalRuntime.ps1 -Platform x64 -Offline -PassThru
.\Tools\Test-TestRunArchive.ps1 -RunPath <terminal-baseline-run>
.\Tools\Test-TestRunArchive.ps1 -RunPath <terminal-candidate-run>
~~~

Expected: all commands exit 0; canonical lock names the candidate commit/tree;
both archives validate; every threshold passes without waiver.

### Phase 11 — Product, sanitizer, package, and native ARM64 qualification

**Goal:** prove the admitted runtime through every shipping boundary.

1. Build test-enabled Debug x64, test-enabled Release x64, ASan Debug x64, and
   Debug ARM64. Also build normal tests-disabled Release x64 and ARM64 for
   packaging.
2. Run all Terminal deterministic selftests and embedded-terminal Commands cases
   on x64 Debug and test-enabled Release.
3. Run ASan Terminal cases and require no sanitizer report.
4. Run `.github/workflows/ghostty-runtime-product-qualification.yml` on native
   ARM64 and retain its required runtime identity, receipts, test evidence, and
   portable artifact. It must run runtime schema/load tests,
   `PluginContractTests --terminal-selftests`, representative embedded-terminal
   Commands cases, the production Release build, and native package smoke. The
   release workflow must require this gate before artifact builds.
5. With an approved build number, package serially:

   ~~~powershell
   .\Installer\zip\build-zip.ps1 -Configuration Release -Platform x64 -BuildNumber <approved-build-number>
   .\Installer\zip\build-zip.ps1 -Configuration Release -Platform ARM64 -BuildNumber <approved-build-number>
   .\Installer\msi\build-msi.ps1 -Configuration Release -Platform x64 -BuildNumber <approved-build-number>
   .\Installer\msix\build-msix.ps1 -Version <major.minor.approved-build-number> -Configuration Release -Platform x64
   .\Installer\msix\build-msix.ps1 -Version <major.minor.approved-build-number> -Configuration Release -Platform ARM64
   ~~~

6. For each clean extraction/package, verify the packaged `ghostty-vt.dll`
   architecture, hash/identity pairing against the receipt-bound shipping
   profile, canonical lock commit/export/import identity, complete notice hashes,
   native load/free where executable, plugin load, and package-smoke result.
   Execute ARM64 package smoke natively.
7. Run the focused Pester set, then Fresh Full. Fresh Full is the final
   repository gate and cannot be replaced by Affected, Commands, Resume, or
   package smoke.

**Verify**

~~~powershell
Invoke-Pester Tools\Tests\GhosttyRuntimeLifecycle.Tests.ps1,Tools\Tests\GhosttyRuntimeStatus.Tests.ps1,Tools\Tests\TerminalRuntimeExportIdentity.Tests.ps1,Tools\Tests\TerminalEngineRuntimeImportPolicy.Tests.ps1,Tools\Tests\TerminalPluginSourceContracts.Tests.ps1,Tools\Tests\BuildEvidence.Tests.ps1,Tools\Tests\BuildReproducibility.Tests.ps1,Tools\Tests\ReleaseWorkflowPolicy.Tests.ps1,Tools\Tests\VcpkgInstallSafety.Tests.ps1,Tools\Tests\TestHarnessSourceContracts.Tests.ps1
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -Configuration Debug -Platform x64
~~~

Expected: zero focused failures; package checks pass for every requested format
and architecture; Fresh Full reports terminal passed and repository passed.

### Phase 12 — Rehearse rollback and close the plan

**Goal:** prove reversibility and move lasting knowledge out of WIP.

1. In an isolated checkout of the pre-admission parent, rebuild the previous x64
   and ARM64 runtimes from its lock in Force and Offline modes. Do not copy DLLs
   out of the current Product cache.
2. Validate their exact previous imports, exports, pairing, licenses, and package
   smoke. Record the rollback command and result in the final Terminal evidence.
3. Return to the candidate branch and rerun Locked status, tool/spec inventories,
   git diff --check, and Fresh Full receipt validation. No tracked file may be
   changed by status or rollback rehearsal.
4. Merge durable behavior into:
   - Terminal_EmbeddedPlugin for the candidate pin, ABI, clipboard, Kitty,
     limits, evidence, observation, admission, and rollback contracts;
   - Build_Toolchain for lock/cache/receipt binding;
   - Testing_PerformanceValidation, Testing_SelfTests, and
     Testing_ValidationEvidence for the new scenario and gates;
   - Installer_PortableZip only if the package-smoke contract changed;
   - Tools README/inventory and TerminalEngine README for command ownership.
5. Update NormativeConsistency implementation anchors if the new lock/status
   surfaces are normative anchors.
6. Mark every checklist item complete, replace placeholders with run IDs and
   reviewed results, move this file to Specs/Plans/Done, remove I13 from the WIP
   index, and rerun spec inventory. The Done file is immutable history after
   merge.

**Verify**

~~~powershell
pwsh -File Tools/Get-GhosttyTerminalRuntimeStatus.ps1 -Mode Locked
.\Tools\Get-ToolInventory.ps1 -FailOnFindings
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
git diff --check
git status --short
~~~

Expected: Locked returns current; inventories and diff check pass; git status
shows only the intended reviewed implementation/evidence/spec changes before
commit and is clean after commit.

## Test plan

### New lifecycle/status tests

- Schema accepts the canonical current and admitted candidate locks.
- Canonical serialization and digests repeat byte-for-byte.
- All required hashes are lowercase/full-width and all sets are sorted/unique.
- Full exports, derived count/hash, and loader-required subset agree.
- Path containment rejects absolute/external/traversal outputs.
- Candidate mode cannot publish Product or pairing headers.
- Locked mode makes no network call.
- Online/Candidate never mutate tracked files or the lock.
- Update, unrelated history, advisory, malformed/unreachable, and rate-limit
  outcomes use exact factual vocabulary and exit policy.
- Workflow permissions and pinned actions are enforced.

### Overlay/upstream tests

- Clean candidate patch apply.
- Upstream test-lib-vt, test-lib-vt-build, test-lib-vt-schema on x64 and native
  ARM64.
- Every new private/public enum collision and C/Zig name/value parity.
- Placement count/allocation limits, atomic all-screen set, unchanged rejection,
  replacement, reset, screen, pin release, and exact queries.
- Windows shared runtime and reviewed strip/linker behavior; the unused static
  companion remains unshipped.

### Adapter/security tests

- Exact loader resolution and partial-load rollback.
- Frozen two-field mode get for every mode used by the product.
- Clipboard sized fields, at-most-once reply, 256 KiB bound, allow/deny/ask,
  private option 41 fail-closed Kitty denial before callback, unaffected OSC 52/
  OSC 1337, null read callback, teardown, and no lock-held UI.
- DCS bytes above 0x7F.
- Pending-to-complete image, replace/delete/evict/retransmit/animation, stale
  worker result, reset/screen, limits, Close zeroing, and resize reuse.

### Build/package tests

- x64/ARM64 machine, exact exports/imports, size envelope, source/tool/dependency
  identity, licenses, pairing header, manifest and receipt binding.
- Force, Offline Force, and Offline reuse.
- Debug, test-enabled Release, normal Release, ASan, native ARM64.
- Clean x64/ARM64 ZIP, x64 MSI, x64/ARM64 MSIX.
- Package smoke is bound to the same Release receipt and does not require
  ENABLE_TESTS in shipped product binaries. It authenticates the staged Ghostty
  runtime and notices against the canonical lock and performs native load/free
  on a matching host.

### Performance tests

- Same corpus/config/machine before and after.
- At least 200 samples for p95; no p99 claim without at least 1,000.
- Hosted render, VT write, and snapshot p95 ratios.
- Exact behavior digests except named intentional corrections.
- Placement, cache, pipeline, close, resize, queue, and resident-memory
  invariants.
- Archive size/shape/sensitive-data validation.

## Final completion audit

The Live execution checklist at the beginning of this plan is authoritative.
Before declaring completion, reread every checked item against its Evidence log
row and the live repository. Do not create a shorter substitute checklist here.
The plan is eligible for the P12.8 archive transition only when P12.7 is checked.

## Done criteria

All Live execution checklist items are checked with exact run IDs/receipts. The
canonical active lock names candidate commit
9c3ec931d64561a8407dde7ac984ce156ae91539 and tree
359872ab74fe1ecc96428133c0546be3cb969fdd. Product x64 and ARM64 manifests bind
that lock, exact reviewed ABI/import/license identities, and the compatible
adapter. All security/resource/performance/package gates pass without waiver.
The previous lock rebuild is proven. Normative specs own the resulting contract,
Fresh Full is green, and this plan is archived under Specs/Plans/Done.

## STOP conditions

Stop and report; do not improvise if any of these occurs:

- Candidate commit/tree/archive provenance or old-to-new ancestry cannot be
  authenticated exactly.
- The observed candidate differs from the frozen commit/tree; do not silently
  move to a newer head.
- An applicable advisory cannot be dispositioned, an online release preflight
  is unreachable, or evidence is based only on secondary claims.
- The overlay requires a second patch, fork, prebuilt binary, more than 16
  upstream files, more than 2,500 changed lines, or more than five independent
  product concerns.
- Any private enum collides with candidate public values or C/Zig values differ.
- Existing 4,096-count, 1 MiB placement allocation, 64 MiB BGRA, 128 MiB
  applicable pipeline, unchanged rejection, pin release, screen/reset, Close,
  or resize invariants cannot be preserved.
- Clipboard support would require a false synchronous success, UI blocking,
  clipboard access under the terminal lock, clipboard read, an unbounded
  transaction, or retained upstream pointers.
- An engine-pending image can remain permanently rejected, or a stale
  conversion can resurrect replaced/deleted/evicted content.
- x64 or ARM64 machine/import/export/size/license identity changes lack a
  source-level explanation.
- Native ARM64 test-lib-vt-schema and product load tests cannot be run. The work
  may remain WIP but is not release-ready.
- Package smoke depends on a stale/missing harness, enables test hooks in shipped
  product binaries, or cannot load/free the exact packaged runtime without
  shell, PTY, user-profile, or clipboard mutation.
- Baseline was not frozen before candidate measurement; sample counts are too
  small; evidence exceeds size limits or includes secrets/absolute user paths;
  any fixed performance/resource gate fails.
- Candidate mode writes Product, changes the pairing header, or edits tracked
  files.
- A focused verification fails twice after one reasonable correction, Fresh
  Full fails, rollback rebuild fails, or rollback/product cache identities mix.
- An implementation step requires a production file outside Scope. Amend and
  review the plan before continuing.

## Rollback procedure

If a post-admission defect is found before merge:

1. Revert the coherent admission commits that changed the canonical lock,
   overlay, adapter, tests, and specs. Do not revert unrelated work.
2. Force-build the previous lock for x64 and ARM64; do not copy a cached DLL.
3. Run offline x64 reuse, exact import/export/license/pairing validation,
   Terminal selftests, package smoke, and focused Pester.
4. If any rollback check fails, keep release blocked and report both candidate
   and previous-lock identities.

If a defect is found after merge, create a dedicated rollback change containing
the previous complete lock/overlay/adapter contract from version control, repeat
the same builds/tests, and run Fresh Full before release. The observation
workflow may continue reporting update-observed; it must never re-admit the
candidate automatically.

## Maintenance notes

- The normal next-update flow is:
  **observe → freeze explicit commit/tree/archive → classify diff/advisories →
  rederive the one overlay → build/test in candidate isolation → review exact
  ABI/security/performance → manually admit the canonical lock → qualify
  Product/packages → rehearse rollback**.
- A monthly update observation is not approval. A candidate is not a product
  runtime until the canonical lock changes in a reviewed commit.
- Whenever upstream adds terminal options or Kitty data values, rerun the full
  enum collision/parity tests before applying the private overlay.
- Whenever Ghostty changes clipboard or image callbacks, review structure-size
  gates, pointer lifetime, bounded copies, UI-thread boundary, reply semantics,
  cancellation, and teardown as one unit.
- Whenever the builder gains a direct input, add it to the active lock when
  appropriate, BuildEvidence identity closure, existing workflow cache key,
  tool inventory, focused tests, and owning normative spec in the same change.
- Do not make the status tool a hidden package manager. It observes and records;
  humans select and admit.
- Reviewers should focus first on provenance, enum allocation, reply truthfulness,
  pending-image state, full export diagnostics, Product isolation, native ARM64,
  package receipt binding, and rollback evidence.
