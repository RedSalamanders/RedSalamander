# File Operations Global Behavior and Implementation Specification

| Field | Value |
|---|---|
| Status | **DONE — PHASES 0–6 COMPLETE (2026-08-27).** The approved behavior is implemented in the durable owners, all work packages and stop conditions are closed, same-machine performance evidence is archived, and Fresh Full plus repository inventories are green. This file is frozen implementation history; current behavior is owned by the normative specifications named below. |
| Scope | Copy, Move, Rename, Delete, Recycle, Create Directory, conflicts, links, identity, metadata, recovery, results, clipboard, provider capabilities, cancellation, performance, and tests |
| Priority | **P0/P1** — destructive safety first; bounded performance and product completeness in the same program |
| Historical program owner | This file records the completed program; current behavior is owned by the durable normative owners below |
| Audit baselines | Product ledger dated 2026-08-16; bridge/security/performance audit dated 2026-08-15; code evidence originally reviewed at `d54217a26`, then reconciled with the current tree before each phase |
| Supersedes | `FileOperations_FileBridgeAndCopySecurityPerformanceAudit_2026-08-15.md` and the former question/decision-ledger form of this file |
| Durable owners | `Specs/FileSystem/FileSystem_FileOperations.md`, `Specs/Core/Core_FileSystemBridge.md`, `Specs/Plugins/Plugins_VirtualFileSystem.md`, provider specs, UI specs, settings schema/spec, and testing specs |

## 0. Implementation progress checklist

This checklist is the durable source of implementation progress. Internal agent/task plans are
temporary coordination only and MUST NOT replace this section. Update the relevant checkbox,
evidence line, authoritative domain spec, tests, and commit ID in the same work-package commit.

Status convention:

- `[ ]` not complete. Add `**ACTIVE**` only to the single work package currently being implemented.
- `[x]` implemented, tested, documented, and committed. A code-only or spec-only landing is not done.
- `[blocked]` is written in the item text with the exact blocker, file/contract location, and required
  decision or external condition; the checkbox remains unchecked.
- Parent phase checkboxes become complete only after every child package and the phase stop-condition
  audit are complete.

Every implementation package must:

- refresh the §16.1 impact scan for its symbols/callers and classify new matches;
- update the authoritative `Specs/<Domain>/` owner for any durable behavior discovered or changed;
- add deterministic correctness/race coverage and source-contract guards where appropriate;
- define/reuse instrumentation and archive performance evidence when the package can affect File
  Operations latency, throughput, queueing, provider I/O, cancellation, or memory retention;
- pass focused tests, a Debug x64 build, `git diff --check`, and its phase stop-condition audit;
- record the validation run/evidence and commit ID below before checking the package complete.

### Phase 0 — normative migration and executable skeleton

- [x] Rewrite authoritative domain owners and remove stale REVIEW/FO-D split ownership.
- [x] Land the mandatory capability-v2 ABI and typed plan/result/conflict skeleton across all in-tree
  providers and tests.
- [x] Retire persisted pre-calculation settings, add Verify persistence, and close popup/provider/
  artifact/S3/profile owner gaps.
- [x] Validate Phase 0 and record the final evidence.

Evidence: commits `559d8eb58` and its recorded Phase 0 build/test receipts in §15.

### Phase 1 — P0 identity, containment, and ownership

- [x] **P1.1 — Central typed admission and qualified endpoints.**
  - [x] Require every capability-v2 response to carry a non-empty path-scoped `rootId`; fail closed
    when it is absent, and keep provider instance, profile, and namespace-root identity as distinct
    endpoint axes.
  - [x] Refresh every Copy/Move/Rename/Delete/Recycle/Create Directory ingress and direct-provider
    mutation match in §16.1.
  - [x] Materialize the central plan factory/admission API for `TransferPlan`, `RenamePlan`, and
    `DeletePlan`; reject malformed/incomplete plans before queue publication.
  - [x] Carry qualified source/destination provider, instance, profile, path, intent, immutable
    settings, and comparison-only ingress snapshots through every admitted plan.
  - [x] Route each migrated ingress through exactly one confirmation/admission owner; preserve
    clipboard and external-OLE ownership rules.
  - [x] Instrument plan construction/admission latency, source count, payload bytes, and rejection
    buckets; add deterministic source-contract and plan-validation tests.
  - Root-identity correction evidence (2026-08-22): full Debug x64 rebuild passed with 0 warnings/
    0 errors (attestation `d312af6572ad149aed9254286ce2dddff0a42ff316bd1b909848b9422dc0c2d7`);
    `PluginContractTests.exe` passed; direct filtered Commands
    `file_operations_phase0_typed_contract_source_guard` passed. Commit: this work-package commit.
  - Typed-admission evidence (2026-08-22): Debug x64 build passed with 0 warnings/0 errors
    (attestation `eef545cf39cd7e034de01e5eca79edc4cfb1b3db22dbdab7f58a9708183ce184`);
    focused Commands typed-contract and clipboard guards passed; File Operations
    `FileOps_ProviderCapabilityMatrix` passed its Setup/case/Cleanup sequence. The queue owns a
    `shared_ptr<const FileOperationPlanGroup>` and clipboard Move consumption is a pre-publication
    ownership barrier. Commit: this P1.1 work-package commit.
  - Mixed-root decision and implementation (2026-08-22): one child plan per qualified source root
    under one immutable task/admission group. Confirmation, task identity, cancellation, result
    presentation, and the captured clipboard Move barrier apply once to the group. Child mappings
    are rebased after stable partitioning; no child borrows another root's endpoint authority.
    Debug x64 build passed with 0 warnings/0 errors (attestation
    `6150b9d69c918fc9199c70c21cbd8725f702e70d2fcd93912a9736c964d9dada`); focused Commands
    `file_operations_phase0_typed_contract_source_guard` and File Operations
    `FileOps_ProviderCapabilityMatrix` passed. Commit: this P1.1 closeout work-package commit.
- [x] **P1.2 — No-follow object binding and identity abstraction.**
  - [x] Implement the host identity/binding abstraction over optional
    `IFileSystemObjectBinding`/provider identity.
  - [x] Implement local no-follow `FILE_ID_INFO` binding with volume/object identity, revision facts,
    kind, and explicit unsupported/indeterminate results.
  - [x] Define snapshot lifetime, reopen/revalidate semantics, and provider contract validation for
    null or malformed success outputs.
  - [x] Add exact identity match/change/disappearance/reparse tests without using pathname/basic
    metadata as authority.
  - Evidence (2026-08-22): Debug x64 rebuild passed with 0 warnings/0 errors (attestation
    `1439ca6b8effa7a9b7ce33855c0fdf2b9d7d862a288728f643e5600c922afee6b`); focused Commands
    `file_operations_phase0_typed_contract_source_guard`, File Operations
    `FileOps_ProviderCapabilityMatrix`, and `PluginContractTests.exe` passed. The host copies and
    retains validated provider authority, distinguishes unsupported/missing/indeterminate/contract
    violation, and the local provider proves hard-link equality, replacement, disappearance,
    retained-handle reading, and no-follow junction identity. Commit: this P1.2 work-package commit.
- [x] **P1.3 — Same-object, alias, subtree, and device-envelope guards.**
  - [x] Reject malformed/escaping mappings and local `GLOBALROOT`/DOS-device/NT-device namespace
    envelopes before publication; keep this structural check distinct from object identity.
  - [x] Route local per-item Copy/Move through just-in-time source/destination/ancestor binding and
    final provider-boundary revalidation in both serial and parallel executor paths; remove the old
    top-level canonicalizing filesystem preflight.
  - [x] Implement direct same-path Copy as automatic Keep Both, reject hard-link/object aliases,
    same-folder Move, destination-inside-source, and destination link ancestors before mutation.
  - [x] Cross-check copied object/revision snapshots against provider `IsSameObject`; disagreement is
    a provider-contract violation and every non-Ready mutation-gate state carries a failing HRESULT.
  - [x] Make local link classification name-surrogate-aware and make capability/name/concurrency
    gates query the concrete governed path instead of an instance-root placeholder.
  - [x] Qualify `Native`, `CopyOnly`, and currently unavailable `Managed` execution explicitly:
    Local native Move uses `FILESYSTEM_MOVE_NATIVE_ONLY`; a Move admitted through the Copy pair
    executes the Copy branch, never calls source Delete, and reports `Copied; source kept`; an empty
    Move pair prohibits managed copy-delete rather than the safe downgrade.
  - [x] Extend executable same-object/containment authority beyond the local profile to each
    provider/native strategy that can mutate or truncate; unsupported binding must remain an honest
    strategy restriction rather than falling through to pathname authority.
  - [x] Complete deterministic direct/case/8.3/SUBST/mapped-UNC/junction-mount and qualified-provider
    alias coverage, including every Copy/Move executor ingress and final-boundary swap/indeterminate
    result.
  - Evidence (local-profile iteration, 2026-08-22): Debug x64 build passed with 0 warnings/0 errors
    (attestation `fde11c80dd04257cdbfe49f4a6d26b03df8cd54a4ce811c962d0e683004e14a4`).
    `FileOps_ProviderCapabilityMatrix`, `Fairstream_CopyIntoSelfAliasRejected`, the Commands typed
    contract source guard, and `PluginContractTests.exe` passed. The two focused matrix failures and
    their repairs are preserved under
    `Specs/TestRuns/SINON/Continuation/2026-08-22_163341_fileops_p13_provider_matrix/`.
    At that checkpoint P1.3 remained open for nonlocal strategies and the remaining alias matrix;
    the final evidence below closes those items.
  - Evidence (strategy-qualification iteration, 2026-08-22): Local exposes a tested native-only
    route that fails before generic/reparse copy-delete fallback; guarded per-item execution binds
    admitted Move work to `Native` or `CopyOnly`, rejects unavailable `Managed`, and rejects bulk
    Copy/Move before provider I/O. The executable CopyOnly regression publishes the destination,
    retains the source, and returns `S_FALSE`; the provider matrix also proves that Curl export may admit
    CopyOnly through the Copy pair while its empty Move pair still forbids cleanup. Full Debug x64
    rebuild passed with 0 warnings/0 errors (attestation
    `88a798bc676ae7cdaa55e3503aec585044c5f6c3820e311f6ce2990e2260721a`); focused Commands
    typed-contract guard passed (run
    `20260822T155303Z-58536-45f41c679d73417ca11e55da5f8c0621`); File Operations provider
    matrix/CopyOnly regression passed (run
    `20260822T155335Z-7964-3b2d27623adf4459854a1833fdc4577d`); native-only same-volume folder
    merge passed (run `20260822T155408Z-70120-4cd3eb2b742146cbbae111fc034527f1`); rebuilt
    `PluginContractTests.exe` passed, including local plugin debug selftests 74/74. Final
    performance-contract hardening emits one bounded `fileops.operation.strategy` row per
    non-empty intent/strategy/topology bucket, never per file or descendant. The final full Debug
    x64 rebuild passed with 0 warnings/0 errors (attestation
    `95e061e3bd2621a1340d0f792bff32e54b549774a14c7804d5a28d1692e6ff1a`); the Commands source
    guard passed (run `20260822T160800Z-63664-ff50740ee93940509a72b3ff2d7e8287`); the repeated
    provider matrix passed 6/6 (run `20260822T160830Z-62420-f6138830c9da42f5939577106754f5d0`)
    and archived `move.copy-only.same-root` while proving destination publication plus source
    retention under `Specs/TestRuns/4cb089111a23/FileOps/2026-08-22_180854/`; the repeated native
    folder-merge case passed 6/6 (run
    `20260822T160905Z-14204-c7e07626c44e4511bfa7fb1c42feea82`) and archived
    `move.native.same-root` under
    `Specs/TestRuns/4cb089111a23/FileOps/2026-08-22_180928/`. Both archives pass
    `Test-TestRunArchive.ps1`; no before/after throughput claim is made because this is a new
    safety-routing baseline. The rebuilt `PluginContractTests.exe` also passed, including local
    plugin debug selftests 74/74. Commit: this strategy-qualification work-package commit.
  - Evidence (qualified-provider authority iteration, 2026-08-22): the just-in-time guard now uses
    the stable path-profile contract to stop exact-path and strict destination-inside-source cases
    within one qualified root even when optional object binding is unsupported; this remains
    structural rejection and never source-delete authority. Every current `nativeMove: true`
    provider accepts `FILESYSTEM_MOVE_NATIVE_ONLY`: Dummy proves locked node relocation/merge,
    Microsoft Drive proves one item-ID-preserving Graph `PATCH` with no `DELETE`, and S3 proves its
    pinned-revision publication plus exact conditional source delete. All single/batch boundaries
    reject unknown move modes before provider I/O. Full Debug x64 rebuild passed with 0 warnings/
    0 errors (attestation `bd96596587ec4f76534d4ca23a23bbdeea4cc40dc168cad02e95aad46bd9a589`);
    the Commands source guard passed (run
    `20260822T163920Z-73196-5238bbbfa6274a0ca81b99d5fb601a9c`); the repeated provider matrix
    passed 6/6 (run `20260822T163952Z-71552-f7d7a1441c794a32bc1fa205142996e6`);
    `PluginContractTests.exe` passed, including Local 74/74, Microsoft Drive 218/218, and S3
    155/155 debug assertions. The bounded identity/strategy evidence is archived under
    `Specs/TestRuns/4cb089111a23/FileOps/2026-08-22_184016/` and passes
    `Test-TestRunArchive.ps1`; its 34 deterministic Debug guard samples averaged 214 us and
    remained at or below 480 us. No throughput improvement is claimed. Commit: this
    qualified-provider-authority work-package commit.
  - Final P1.3 evidence (2026-08-22): Copy/Move defaults to guarded per-item execution; bulk
    transfer, unavailable `Managed`, and `Native` routed through a destination bridge fail before
    provider I/O. Serial and parallel scheduling call one strategy dispatcher. CopyOnly uses Copy
    verification, never arms the legacy Move cleanup manifest, retains every source, and reports
    source-kept. Same-folder and explicit-mapping checks use the qualified path profile. Exact object
    comparison crosses root aliases inside one plugin/instance/profile identity domain, while
    `rootId` remains the strategy-topology boundary. Out-of-pane explicit filesystems derive their
    plugin identity through `IInformations`, preventing pane/provider misrouting.
    The final Debug x64 rebuild passed with 0 warnings/0 errors (attestation
    `60dce37d779866df470a9bd98336e8ab1d1c40e52a2fb75696a3ccfdbb3e757d`); the Commands source
    guard passed (run `20260822T183211Z-71024-b1e4ae6c87494fd6a878631171277fe5`); the repeated
    provider/alias/swap matrix passed 6/6 (run
    `20260822T183243Z-3184-a6044c9a5f464e689c78698b77258291`); the CopyOnly verification,
    cleanup-exclusion, and directory-retention cases passed (runs
    `20260822T183319Z-65980-c4aca27e94dc4987a74dcf45831be67e`,
    `20260822T182222Z-80596-9653a03bd8bc45e5b1bf7bfc51021794`, and
    `20260822T182253Z-28052-015d4bd7d76d464aac901c59f1b4a342`); serial/parallel conflict
    cases and native folder merge passed; `PluginContractTests.exe` passed. The repeated identity
    matrix and CopyOnly verification evidence are archived under
    `Specs/TestRuns/4cb089111a23/FileOps/2026-08-22_203307/` and
    `Specs/TestRuns/4cb089111a23/FileOps/2026-08-22_203342/`; both pass
    `Test-TestRunArchive.ps1`. No throughput improvement is claimed. Commit: this P1.3 closeout
    work-package commit.
- [x] **P1.4 — Bound destination ancestry and identity-owned staging/publication.**
  - [x] Bind and revalidate destination root/ancestors no-follow; turn destination links and ancestor
    replacements into typed conflicts/failures.
  - [x] Create stages exclusively and retain their provider identity/creation token.
  - [x] Permit publish, rollback, and cleanup only against the exact retained identity; never remove
    a replacement object by pathname.
  - [x] Diagnose `S_OK` with null/empty reader, writer, binding, stage, publication, or capability
    result and fail closed.
  - [x] Add collision, replacement-after-create, promotion failure, cleanup failure, and crash-point
    deterministic tests plus stage/publication timing and retained-state metrics.
  - Evidence (2026-08-23): the local binding creates every bridge stage with `CREATE_NEW`, retains
    no-follow `FILE_ID_INFO` authority, conditionally publishes through the exact expected
    destination, and uses an exact `.rs_bak_{GUID}` rollback object without pathname deletion.
    Unknown cleanup retains a visible Possible artifact. The host rejects null-success reader,
    writer, stage, binding, and publication outputs; measures aggregate stage creation,
    publication, and retained-stage counts; and never arms managed-Move cleanup for CopyOnly.
    Concurrent pathname replacement after publication is verified through a duplicate of the
    retained content-capable handle, leaving the replacement owner untouched. Full Debug x64
    rebuild passed with 0 warnings/0 errors (attestation
    `962f2fd6b8a33c37187e7e8edc438a560344f69950e79aa9791ced10a51bb1b1`).
    `PluginContractTests.exe` passed, including Local 95/95, Microsoft Drive 218/218, S3 155/155,
    Dummy capability/ABI coverage, and all unload quiet points. Focused replacement, CopyOnly
    cleanup-exclusion, parallel source-failure, and Recycle batching cases passed. The 52-case
    Fairstream/Floodgate family passed 47/47 executable cases with seven reviewed P2/P3 deferrals
    (run `20260822T225853Z-74684-76000aad815d41e397579b3cc6a58e02`). The final complete File
    Operations matrix passed 111, failed 0, skipped 27 (run
    `20260822T230027Z-79788-952f58990c5e42d797b4f4d3674c55a5`); its same-machine stage/
    publication and scheduler evidence is archived under
    `Specs/TestRuns/4cb089111a23/FileOps/2026-08-23_011612/`. No throughput improvement is claimed;
    bounded safety-path timing is the P1.4 baseline. Commit: this P1.4 work-package commit.
- [x] **P1.5 — Inline F2 central-engine cutover.**
  - [x] Route in-pane F2 through worker-executed `RenamePlan(InlineRename)` as soon as the central path
    and identity revalidation exist.
  - [x] Remove the UI-thread provider `RenameItem`/legacy `ReportError` mutation bypass.
  - [x] Bypass global Queue admission waiting while retaining overlapping-root interlocks and exact
    identity/capability revalidation.
  - [x] Keep the ordinary visible task card until Phase 5 lands delayed reveal/silent clean success.
  - [x] Add worker-thread, Queue-bypass, interlock, teardown, and no-direct-call regression coverage.
  - Evidence (2026-08-23): accepting the inline editor now submits a typed one-step
    `RenamePlan(InlineRename)` to the central worker engine. The worker binds and revalidates the source,
    its parent, and any existing destination before using the provider's conditional
    `RenameIfUnchanged` boundary; a committed-but-indeterminate result is never retried. The old
    UI-thread `RenameItem` and pane `ReportError` mutation path is removed. Inline F2 bypasses the
    global Queue wait while an identity-aware mutation interlock serializes exact overlapping roots;
    each task interns common root-to-ancestor bindings once per unique provider path, and cancellation
    wakes a waiting task without provider mutation. The ordinary visible card remains intentionally
    active until the Phase 5 presentation package implements the normative 500-ms delayed reveal.
    Final full Debug x64 rebuild passed with 0 warnings/0 errors (attestation
    `fd95dad5bfc45aded7855ed9f7fc7eea53be3c047dc95e008092a2527d17806f`). The final focused
    worker/Queue/interlock/cancellation case passed 3/3 (run
    `20260823T002807Z-26200-d5c7f0baff1e43068ef7f564fb7b5a38`); the final Commands
    source-contract guard passed (run `20260823T002706Z-43380-6ad63b5f73b543fc8405d41e1d775ae1`);
    `PluginContractTests.exe` passed, including Local 95/95, Microsoft Drive 218/218, and S3
    155/155 debug assertions. The real inline
    rename dialog and live DxUi interaction passed (runs
    `20260823T000635Z-26024-73e3c2184ef84e3888f5f689450ed11c` and
    `20260823T000704Z-44608-cacf4f35346b4b5db9e92fa368425f08`). The complete File Operations
    aggregate passed 112, failed 0, skipped 27 (run
    `20260823T000732Z-42220-f73014eafbbb49bdb51baf71cee20367`). The final focused interlock/
    cancellation timing evidence is archived under
    `Specs/TestRuns/4cb089111a23/FileOps/2026-08-23_022826/`; no throughput improvement is claimed.
    Commit: this P1.5 work-package commit.
- [x] **P1.6 — Phase 1 integration and stop-condition gate.**
  - [x] Route Pack and Unpack delete-after-success through `DeletePlan` with one captured
    `ArchiveDeleteAfter` receipt; remove the Unpack `std::filesystem::remove` bypass.
  - [x] Classify interlock scopes as `ReadSource`, `WriteSource`, or `PublishDestination`; permit
    Read/Read overlap and serialize every overlapping pair containing a writer or publisher.
  - [x] Fail an item before writer creation when cryptographically secure stage-name entropy is
    unavailable; never fall back to PID/TID/time-derived names.
  - [x] Use the admitted provider path profile for resolved destination parent/directory-shell
    creation; leave no host `std::filesystem::parent_path()` mutation authority in that route.
  - [x] Keep Copy and Copy-only on one immutable optional-verification policy. Mandatory byte-count
    and committed-size validation remain shared; do not expand the transitional FNV reread into a
    permanent mandatory verification contract.
  - [x] Make `RenameOrigin::Unspecified` fail validation, use `InlineRename` for the explicit F2
    exception, and prove one-step Batch Rename cannot inherit it by default.
  - [x] Show a localized pane error for Inline Rename rejection before task publication or missing
    host wiring; once published, keep the central task UI as the only failure presentation.
  - [x] Make provider-profile parent derivation canonical across validation, exact guards,
    interlock execution, completion destination/focus, cache notification, and refresh; remove the
    compatible task-local helper.
  - [x] Reuse the canonical snapshot/provider identity cross-check and remove the unreachable
    provider bulk Copy/Move executor without weakening the pre-I/O rejection.
  - [x] Record the executable Inline Rename provider matrix: Local succeeds through binding; a
    rename-advertising provider without binding fails visibly before mutation; `rename: false`
    rejects before task publication.
  - [x] Run the complete Phase 1 race/alias/containment/staging/provider validation matrix.
  - [x] Archive same-machine performance evidence for admission/identity/staging overhead and prove
    no unbounded per-selection/tree retention was introduced.
  - [x] Run focused suites, Debug x64 build, spec inventory, and `git diff --check`.
  - [x] Run applicable test-enabled Release coverage.
  - [x] Audit that no destructive path uses pathname/basic metadata as sole authority and no
    production F2 path directly mutates on the UI thread.
  - Evidence (2026-08-23): full-solution Debug x64 build passed with 0 warnings/0 errors
    (attestation `9520c04671e718251fee6f861a04d9a5f1ffc622a0e707fb866b9de104220b97`). The complete File
    Operations suite passed 112, failed 0, skipped 27; its digest-bound compact retention evidence
    is archived under `Specs/TestRuns/4cb089111a23/FileOps/2026-08-23_125000/` and passes the
    explicit pre-stage archive contract. One paused Inline Rename scope retained 13 exact authority
    nodes against the 1,024 path-depth bound; no throughput improvement is claimed. The focused
    Commands source contract, provider matrix, Floodgate interlock, and provider-parent cases passed;
    the final source-contract, interlock, and provider-parent reruns are
    `20260823T113654Z-77304-57da7e5d983f40a6adbde63dda576146`,
    `20260823T113732Z-63672-f4f1e76ae5f3456992dbbb34b9743301`, and
    `20260823T113819Z-61708-9a25dfc7bcee467a9487b4d21889bd65`. Full tooling Pester passed
    634/634; spec inventory reported zero blocking findings; `git diff --check` passed. A diagnostic
    Fresh Full run exposed six stale tooling assertions and three
    Commands failures; the six assertions and one source guard were repaired, while isolated reruns
    proved the other two failures transient, and all affected suites subsequently passed. Release
    coverage remains pending on the user-owned executable above. Commit: this P1.6 partial-gate
    work-package commit.
  - P1.6 closeout iteration (2026-08-23): Pack and Unpack now carry the archive prompt's captured
    consent into central permanent `DeletePlan` admission; the private Unpack pathname delete was
    removed. Copy/Copy-only register `ReadSource` plus `PublishDestination`, destructive source
    operations register `WriteSource`, and only overlapping Read/Read pairs share the interlock.
    Provider-profile parent derivation now owns directory-shell construction, stage entropy failure
    stops before writer creation, and the transitional mandatory Copy-only-excluding FNV reread was
    retired in favor of the shared byte-count/committed-size floor and later optional BLAKE3 policy.
    The stable full Debug x64 rebuild passed with 0 warnings/0 errors (attestation
    `e2cd4e0f05043ded7b528dc4cce41110f5413f32494f240140143571059f4565`). Focused provider/entropy
    coverage passed 3/3 (run `20260823T124611Z-84152-d931b0a65b4346ebaff65816e4a257fa`), the
    pre-calculation/bridge performance family passed 6/6 (run
    `20260823T124641Z-75384-1b4f9fc07a2e4909a38fd9e23c55b249`), the Commands typed source
    contract and live Unpack delete-after prompt each passed (runs
    `20260823T124950Z-19476-a30d6ad5105049eaa3dca86e865a2056` and
    `20260823T125019Z-64504-74ad657f296543f6ba22b84faf848a3c`). The complete File Operations
    suite passed 112, failed 0, skipped 27 (run
    `20260823T125059Z-79808-3f2a1af6dfca4c7281e7542674703c18`); performance evidence is archived
    under `Specs/TestRuns/4cb089111a23/FileOps/2026-08-23_150659/`. Release remains the sole Phase 1
    blocker while the user-owned Release executable is running. After the normative performance and
    selftest owners were finalized, the full Debug x64 solution rebuilt with 0 warnings/0 errors
    (attestation `652819c1dd4394543d14604724d028350c12c4aadea7bc6e2cec710037de98f4`), and the live Commands
    source/spec contract passed again (run `20260823T131750Z-20904-0df85acc5c19449e82d521706c8369c7`).
  - Final P1.6 integration evidence (2026-08-23): execution and completion now resolve every
    transfer destination from the immutable `TransferPlan`, including Compare explicit mappings and
    distinct source/destination path profiles; cache notification, focus, and refresh use that same
    provider-path result. A deterministic 256-source retention case proves the interlock retains 270
    unique authority nodes at maximum provider depth 14, bounded by unique overlapping paths rather
    than selection-by-ancestor multiplication. The final test-enabled Debug x64 solution built with
    0 warnings/0 errors (attestation
    `0d942406474f0ea51e7a41a8f44bb688c5df3028d7875961046d910c0ccd5d06`). The complete Debug File
    Operations suite passed 112, failed 0, skipped 27 (run
    `20260823T141824Z-85124-233d064bd97d44989667e6bb2204b406`); its digest-bound compact evidence
    is archived under `Specs/TestRuns/4cb089111a23/FileOps/2026-08-23_163451/` and makes no
    throughput-improvement claim. The test-enabled Release x64 solution built with 0 warnings/0
    errors (attestation `f63165461862f4779e7e9aca410fb0c2dd6957789fff987b49b1b1cee5aa1896`).
    Release provider qualification passed 3/3 (run
    `20260823T144257Z-15668-79a0da03c9af4d5faad02b27a0b7bb78`), exact destination mappings
    passed 3/3 (run `20260823T144332Z-63468-8b52b3fe8f124c11a75856fd3086f2d3`), the Cross-FS
    Copy/Copy-only source-preservation family passed 9/9 (run
    `20260823T144357Z-63468-b27da3f8c1fa4936b7e0bc0248fbefca`), and Inline Rename worker/
    Queue/interlock passed 3/3 (run
    `20260823T144421Z-63468-0f253fd520894441a5cac4a425951917`). Successful Release evidence
    is archived under the corresponding `2026-08-23_164318`, `2026-08-23_164352`,
    `2026-08-23_164417`, and `2026-08-23_164439` FileOps folders. Commit: this Phase 1 closeout
    commit.

### Phase 2 — managed Move and bounded traversal

- [x] **P2.1 — Strategy qualification and bound transfer endpoints.**
  - [x] Select native Move, managed Move, or Copy-only from executable path/profile/pair facts.
  - [x] Bind the source reader/revision and destination publication proof required by the selected
    strategy; fail closed on missing or stale proof.
  - [x] Add provider-pair strategy tests and strategy-selection telemetry.
  - Evidence (2026-08-23): the pure selector now requires the concrete Copy route for every
    non-native outcome and selects Managed only with the separate Move pair, bound/conditional
    source Delete, exclusive/conditional destination publication, and both binding interfaces.
    The live bridge reads a Managed item from the retained exact source binding, publishes through
    the identity-owned stage, and can only call that binding's `DeleteIfUnchanged`; failure to
    acquire the stronger per-item authority becomes `Copied; source kept`. Local delete bindings
    exclude concurrent write/delete/rename opens. A test-enabled full Debug x64 build passed with
    0 warnings/0 errors (attestation
    `c99b3dac84617611358d7ef528d34c79b1abee0f7119bbb06fc8c476544f559c`). The provider strategy
    matrix passed 3/3 (run `20260823T151609Z-49040-6358bf411eac4de7ac74a55835c55fd4`), including
    exact Local file/tree Managed cleanup and an external-writer Copy-only downgrade; the Commands
    source-contract guard passed (run
    `20260823T151640Z-28708-0a120d9559b44c37b2b2645404b316a7`); the Cross-FS Copy/Copy-only
    preservation family passed 9/9 (run
    `20260823T151710Z-22256-bf989e69812348319633e56a8dc86abb`); and
    `PluginContractTests.exe` passed, including Local provider selftests 95/95. Bounded
    `fileops.operation.strategy` telemetry is reused; no throughput-improvement claim is made.
    Curated FileOps evidence is retained under the corresponding `2026-08-23_171631` and
    `2026-08-23_171733` machine-profile folders, both validated by
    `Test-TestRunArchive.ps1 -RunPath`. The exact final snapshot also passed the complete Debug
    File Operations suite 112/112 with 27 documented environment/deferred skips (run
    `20260823T152411Z-9800-7552897dd32c49f4af615693b91f7b0c`); its 228-MiB raw metric stream
    remains local rather than violating the curated 2/5-MiB archive limits. Commit: this P2.1
    work-package commit.
- [x] **P2.2 — Managed Move and exact cleanup.**
  - [x] Publish each destination atomically, then conditionally/handle-bound delete only the exact
    unchanged source.
  - [x] Downgrade to `Copied; source kept` whenever exact cleanup authority is absent, changed, or
    indeterminate.
  - [x] Implement per-file cleanup, nonrecursive post-order directory cleanup, sharing/access/path/
    disk-full outcomes, and retryable cleanup records.
  - Evidence/commit: the managed bridge now retains one exact bound cleanup record from source read
    through terminal cleanup, publishes before deletion, and retries only `DeleteIfUnchanged` against
    that same record. Local authority keeps the source object bound strongly enough that an external
    rename loses the tested race. Deterministic injections prove a sharing violation enters the
    central Retry/Skip/Cancel surface without recopying, while an unknown mutation outcome is terminal
    `ERROR_IO_INCOMPLETE`, retains the source, and never prompts or issues a second delete. The final
    full Debug x64 test-enabled build completed with zero warnings/errors (attestation
    `8d26fa75dc440adff81eb6efc49a01592b19d126453e1af8fe13ae40efca43a0`). Focused final runs:
    provider/cleanup matrix `20260823T160505Z-89840-64cda4f7cbc34c02a1908229bb065f6c`, Commands
    source guard `20260823T160542Z-75164-b7f4b0b95fef490496b88ccc894f26b0`, conflict Retry cap
    `20260823T160618Z-61952-fbd0ae8d164e43b39330dbe439ddb08a`, and CrossFS floodgate
    `20260823T160656Z-84232-6a72597a0632414cb236840587da74cf`. Curated archives are under
    `Specs/TestRuns/4cb089111a23/{FileOps/2026-08-23_180531,Commands/2026-08-23_180606,FileOps/2026-08-23_180643,FileOps/2026-08-23_180723}` and each passes
    `Test-TestRunArchive.ps1 -RunPath`. This package makes no throughput-improvement claim; it reuses
    strategy, identity-revalidation, and conflict-wait instrumentation. Commit: `4379129c`.
  - [x] Correct Local Native qualification without weakening volume-scoped endpoint truth: at P2.2
    closeout, equal volume-qualified endpoints remained necessary, a destination-absent same-volume
    directory stayed Native, and a regular directory targeting an existing regular directory was
    admitted Managed so colliding children used the central typed conflict engine. P3.2c supersedes
    the destination-absent-tree rule: Local now advertises no native semantic-link transform, so
    every Local tree/link object uses Managed while regular files remain Native. Qualification probes
    top-level shape only and does not restore recursive preflight. `FileOps_MoveMergeIntoExistingFolderSameVolume`
    proves both immutable plan strategies, the exact child `Exists`/Overwrite decision, destination
    sibling preservation, and final source disposition (3/3, run
    `20260823T163403Z-51732-71a2428c6e99417da9fc73eac908378f`; curated archive
    `Specs/TestRuns/4cb089111a23/FileOps/2026-08-23_183424`). The complete Debug solution build
    passed with zero warnings/errors and attestation
    `4d4768563b4a2ab3f36eef3333ee12c4ff8dbc91f9fa8f0a084eb6a021cace50`. This correction makes no
    throughput claim and emits the existing strategy/conflict evidence plus
    `fileops.plan.native_directory_merge_reclassified`. Commit: this qualification-correction
    work-package commit.
- [x] **P2.3 — Retire weak skip/delete authority.**
  - [x] Remove FNV-1a as silent-identical skip or deletion authority from native resume and bridge
    paths.
  - [x] Remove the mandatory destination reread as managed-Move delete authority while preserving
    optional verification.
  - [x] Route every existing-destination file collision through the typed conflict decision.
  - Evidence/commit: the Local provider no longer content-hashes an existing destination into a
    silent successful Skip; byte-identical reruns enter the typed Exists prompt. The host bridge no
    longer hashes copied bytes for cleanup, rereads destination content as deletion authority, or
    retains the retired pathname-keyed whole-tree cleanup map. Managed cleanup calls only the exact
    source binding's `DeleteIfUnchanged`; source-owned sharing/access/path/disk known non-commits
    expose one Retry plus Skip, while unknown outcomes remain one-shot and indeterminate. Retained
    descendants propagate source retention through only their containing directories without a
    second cleanup prompt. A real alternate-volume Local test, when the environment exposes one,
    proves host admission selects Managed rather than relying on a test-injected strategy. Focused
    repeat evidence: provider/cleanup matrix 9/9 run
    `20260823T174944Z-46084-a4b14020987d40f78da598bf7802f108`; retained-child Keep Both 9/9 run
    `20260823T175023Z-65128-79ea7c204059437b99fdbeb9ef5a89e4`; Local native/managed/cross-volume
    qualification 9/9 run `20260823T175058Z-35604-59d7fa866235411896bf68230690f2e1`;
    identical-collision prompt 9/9 run `20260823T175132Z-8408-34568d95dede4d0aaaa038194329853e`.
    The earlier full-suite latency outliers were not reproduced in isolation: pre-calc cancellation
    9/9 run `20260823T175207Z-89888-7ccbaf939d414d9e9d6604a0f542f287` and bandwidth/cancel 9/9 run
    `20260823T175241Z-91308-e31ac8c573f64934ba60f97ed00acbe4`. The exact final Debug snapshot
    built with zero warnings/errors (attestation
    `d943ea76c35aaa8a5f06912c2ad56de18acf00b0033e2df2642c64538159966e`) and the complete File
    Operations suite passed 112/112 with 27 documented environment/deferred skips, zero failures,
    and zero flaky/isolation classifications (run
    `20260823T180111Z-86700-6f9e646d181a433a95937392fd736a1f`). The Commands source-contract
    guard passed 1/1 (run `20260823T181731Z-63724-5f8f1989e0ad4ce2a53063a61d3d89e2`). Curated
    evidence is retained under
    `Specs/TestRuns/4cb089111a23/FileOps/{2026-08-23_195011,2026-08-23_195122,2026-08-23_195157,2026-08-23_195231,2026-08-23_202100,2026-08-23_202101}` and
    `Specs/TestRuns/4cb089111a23/Commands/2026-08-23_201753`; the explicit archive contract passes
    for all 61 retained files. The full raw metric stream and oversized focused bandwidth stream
    remain local rather than violating the curated 2/5-MiB limits. Commit: this P2.3 work-package
    commit.
- [x] **P2.4 — Bounded traversal and retention.**
  - [x] Remove the retired host pathname-keyed whole-tree cleanup manifest/map (landed with P2.3).
  - [x] Replace remaining created-directory/traversal lists with bounded queues, O(depth) ancestors,
    bounded path/depth/proof state, and retryable terminal records.
  - [x] Enforce quantitative ceilings with partial-prior-success semantics and diagnostics.
  - [x] Add deep/wide/long-path and memory-high-water deterministic/performance coverage.
  - Evidence/commit: the host bridge and Local provider now queue only bounded file/reparse work,
    retain directory/metadata authority in O(depth) frames, restore metadata post-order only for
    directories created/replaced by the operation, and never retain a completed-tree-sized directory
    list. Hard limits are 128 recursive levels, 4,096 retained entries, 16 MiB queued UTF-16 path text,
    8 MiB ancestor/directory metadata, 16 admitted bridge files, and the existing 256-MiB aggregate
    pump budget. Every limit is exact task-terminal even under Continue on error; prior publications
    remain published and no conflict prompt or source cleanup follows. The initially proposed 256-level
    recursive bound was rejected by executable evidence: a 192-level Debug fixture exhausted the
    default stack before a controlled result; its WER/trace/dump pointer and hash are retained under
    `Specs/TestRuns/4cb089111a23/Continuation/20260824_000200_fileops_p24_recursive_depth_crash`.
    The final 96-level extended-path fixture passes both Local `CopyItems` and `CopyItem` below the
    conservative 128-level bound (3/3, run
    `20260823T220922Z-8608-8e21da15c567476fadd3ca58a2c6ad16`, archive
    `Specs/TestRuns/4cb089111a23/FileOps/2026-08-24_000942`). Host wide/limit evidence passes 3/3
    (run `20260823T211829Z-86072-d456d6accca64c90b8c951cb2cca717a`, archive
    `Specs/TestRuns/4cb089111a23/FileOps/2026-08-23_232029`): 48/48 files and two directories,
    admission high-water 16, retained high-water 21 entries/11,634 path bytes/9,674 metadata bytes,
    no natural limit hit; the forced second-item depth limit returns exact `0x800703E9`, records one
    limit hit, retains the first 64-byte publication, publishes nothing below the bound, and spends
    zero time in conflict wait. Directory merge metadata and reparse coverage passes 3/3 (run
    `20260823T213211Z-86916-82d08f3f418645e6bca688224e14565e`, archive
    `Specs/TestRuns/4cb089111a23/FileOps/2026-08-23_233232`). The source-contract guard passes 1/1
    on the final constants (run `20260823T221013Z-82748-2ab06320e68744d1b63b2928211690a5`,
    archive `Specs/TestRuns/4cb089111a23/Commands/2026-08-24_001033`). The complete Debug File
    Operations suite passed 112/112 with 27 documented skips and zero classification failures (run
    `20260823T213305Z-3636-896d5017b436403cb6dc4d9d964b2149`, curated archive
    `Specs/TestRuns/4cb089111a23/FileOps/2026-08-23_234843`). Final full Debug and Release x64 builds
    completed with zero compiler warnings/errors and attestations
    `9c9c97eaad62f8c6075cb28f95c22c7bcdc977dbc456e15b660d5592857f3cc2` and
    `ff63da121cfe3f9e63be98ea8b23fd305f540718df7800891393dc8aa2657e46`; Release required one normal
    rerun because the first post-build attestation observed transient root `sqlite3.dll` staging, and
    the second build staged and attested the dependency without a manual copy. Commit: this P2.4
    work-package commit.
- [x] **P2.5 — Single-pass discovery-ahead scheduler.**
  - [x] Replace runtime pre-calc/`Calculating` bridges with one traversal feeding bounded execution.
  - [x] Implement discovery-ahead reservation/throttling so discovery cannot starve behind transfer.
  - [x] Implement one-way Skip discovery, immediate reservation release, just-in-time enumeration,
    and full-speed transfer without skipping safety checks.
  - [x] Instrument first mutation, queue depth, starvation, bandwidth, traversal closure, and ETA
    knowledge state; archive local/high-latency evidence.
  - Evidence/commit (2026-08-24): the operation-control ABI now carries discovery mode plus
    cumulative per-item discovery progress; every provider callback receives its stable item cookie,
    and the host serializes cumulative-to-delta accounting so concurrent callbacks cannot multiply
    totals. Local recursive Copy, permanent recursive Delete, and the host bridge feed their existing
    bounded traversal directly into execution. Delete retains one 256-child mutation batch, caps
    terminal-failure retention at 4,096 paths/16 MiB, and removes the old 200,000-entry flattening
    pass. Discovery-ahead uses low-water/target/maximum depths 32/128/256, reserves one
    worker above low-water, halves active transfer capacity below low-water, and switches one-way to
    just-in-time/full transfer capacity after Skip. Skip atomically releases the host reservation;
    a provider already inside a call observes just-in-time at its next bounded checkpoint. Every
    task, aggregate, compact, whole-operation, taskbar, and ETA presentation remains indeterminate
    until traversal closure. Selected leaf Copy and qualified native Move report their top-level
    object before mutation; recursive roots remain owned by their existing walker. The final full
    Debug x64 rebuild passed with 0 warnings/0 errors (attestation
    `7a724aeab4b7dfb7051a66523a6e6cb4c2c8e0a0b817edd0eabe06d8f43ef995`). The final
    discovery family passed 6/6 (archive
    `Specs/TestRuns/4cb089111a23/FileOps/2026-08-24_020018`); exact overlap passed 3/3 with
    2,097,152/2,097,152 discovered bytes, first mutation before closure, queue high-water 15/256,
    zero starvation, 125,000 µs open time, and 29 µs callback-lock wait (archive
    `Specs/TestRuns/4cb089111a23/FileOps/2026-08-24_015609`); Queue-paused Skip release passed 3/3
    with a 0 µs host-reservation release and no pre-resume publication (archive
    `Specs/TestRuns/4cb089111a23/FileOps/2026-08-24_015914`). Typed source and popup progress
    contracts each passed 1/1 (archives
    `Specs/TestRuns/4cb089111a23/Commands/{2026-08-24_025734,2026-08-24_020001}`). The curated
    archives pass the explicit TestRuns contract. The final 3,000-file permanent-Delete
    baseline also passed 3/3 with exact 3,000-byte completion, mutation before closure, queue
    high-water 256/256, zero starvation, and the 128/256/4,096/16-MiB traversal ceilings emitted
    (archive `Specs/TestRuns/4cb089111a23/FileOps/2026-08-24_021821`). Exact leaf aggregation,
    popup smoke, and the 256-record bridge discovery bound then passed 3/3 each and were rerun in
    the final governed File Operations suite. That suite passed 110/110 with 27 documented skips,
    zero failures, and zero flaky/regression/isolation classifications (digest-bound compact archive
    `Specs/TestRuns/4cb089111a23/FileOps/2026-08-24_031318`). Commit: this P2.5 work-package commit.
- [x] **P2.6 — Phase 2 integration and stop-condition gate.**
  - [x] Prove every advertised destructive Move strategy has exact conditional cleanup or becomes
    native/Copy-only.
  - [x] Run managed-Move swap/failure/cancel/partial and boundedness matrices, builds, specs, and
    archived performance validation.
  - Evidence/commit: the executable `nativeMove: true` set is exactly Local, Dummy, Microsoft Drive,
    and S3 flat-prefix. Provider capability/debug/source contracts prove NativeOnly behavior for each;
    every other profile remains Managed, Copy-only, or unsupported from executable pair facts. Local
    direct cross-volume Move now returns `ERROR_NOT_SAME_DEVICE` with source retained and no
    destination; its retired provider copy/delete proof/manifest/walker is absent. Host Managed Move
    owns the only Local copy/delete route through retained binding plus `DeleteIfUnchanged`.
    Deterministic changed-membership coverage pauses after selected-child publication, injects an
    unselected source child, and proves `S_FALSE` / `Copied; source kept` without recopy, cleanup
    prompt, or ancestor loss. The four legacy atomic-final verification cases now execute truthfully
    as Copy-only and retain source across convergence, permanent miss, wrong size, and canceled
    backoff. Focused native-link-object, membership-race, collision, provider-matrix, recursive
    parallelism, and Commands source-contract cases passed. The Fairstream family passed 52/52 with
    only semantic retarget explicitly deferred to P3.2. The final governed File Operations suite
    passed 116/116 with 21 documented skips and zero failure classifications (digest-bound compact
    archive `Specs/TestRuns/4cb089111a23/FileOps/2026-08-24_045623`); the final Commands source guard
    passed 1/1 (archive `Specs/TestRuns/4cb089111a23/Commands/2026-08-24_050042`). Full Debug and
    Release x64 builds completed with zero warnings/errors and attested receipts
    `9708651fa12067a852108151f4decb820cd132e5be239b0981e8bb766beff0e5` and
    `dea560e7c335bbd52e8dff25fa7eebd37487e70b74db74aa5f77d80fc18bdfc6`. Both curated archives pass
    explicit pre-stage and repository inventory validation; the specification inventory reports zero
    blocking findings. Commit: `a2c6a5d25`.
- [x] **P2.7 — Atomic Native qualification and scheduler-policy hardening.**
  - [x] Replace Local per-child Native rename-merge with one provider mutation: a raced regular
    destination directory must produce a known non-commit before any child is moved; a destination
    link remains a typed link conflict and is never a merge namespace.
  - [x] Retain a bounded admission-qualified fallback on eligible Native directory items. When the
    regular destination directory appears before the provider mutation, requalify that untouched
    item under the same task and run Managed; if Managed is unavailable, use qualified Copy-only or
    fail source-unchanged without an unproductive Overwrite retry. Never replay an unknown or
    partially committed Native result. Copy-only is a forward-compatible fallback for a future
    native provider whose Copy pair is executable but whose exact destructive proofs are absent;
    Local currently qualifies Managed whenever this race fallback is available.
  - [x] Give ordinary Copy an explicit `Copy` strategy instead of naming it `Managed`, including
    validation, telemetry, typed result defaults, and deterministic source-contract coverage.
  - [x] Split Local same-root selections by executable top-level shape so an existing-directory
    merge item does not force unrelated destination-absent siblings away from Native rename.
  - [x] Extract the shared traversal/discovery ceilings and the identical queue-target/worker-
    reservation calculations into one Common policy owner used by the host bridge and Local walker;
    keep their layer-specific scheduler loops local.
  - [x] Centralize the shared post-mutation terminal classification used by serial and parallel
    item loops without changing prompt/cache/continue-on-error ownership.
  - [x] Add deterministic destination-link, destination-appeared, mixed-selection, strategy-name,
    and shared-policy guards; refresh §16.1; archive governed File Operations performance evidence.
  - Evidence/commit (2026-08-24): Local Native Move now performs one provider mutation and preserves
    provider-level case-only directory rename; regular directory merge is host Managed, while a
    regular destination directory that races into existence requalifies the untouched item under the
    same task without replaying unknown/partial Native results. Missing fallback fails source-
    unchanged without an unproductive Overwrite retry. Local mixed selections retain per-item
    Native/Managed strategies, ordinary Copy emits `copy.copy.*`, and the host/Local walkers consume
    the shared bounded policy owner. Debug and Release x64 builds completed with zero warnings/errors
    and accepted receipts `91723b1dcbfa30eba1a73a28286da46b4c3346e8d36c0638ae15987a51600e8f`
    and `ee7fce3e27cfc95772518839f7c60186dbbbe45cff9408b53a67bc048635df48`.
    Focused merge/race/case-only coverage passed 3/3, destination-junction coverage passed 3/3,
    Commands source contracts passed 1/1, provider contract tests passed, and the final governed File
    Operations suite passed 117/117 with 20 documented skips and zero failure classifications
    (`Specs/TestRuns/4cb089111a23/FileOps/2026-08-24_113000`). Continuation archives retain the
    abandoned pause-hook diagnostic and one earlier timing-sensitive queue assertion that passed in
    isolation and in subsequent complete runs. Post-close review removed the redundant pre-provider
    pathname shape probe, hoisted the Local shape `IFileSystemIO` queries to one pair per plan, and
    added a behavioral no-fallback/source-unchanged case; the provider-known-noncommit sequence in
    the normative specs remains the sole race-requalification route. Focused provider-matrix and
    Commands source-contract runs passed 3/3 and 1/1. The complete Debug x64 File Operations run
    passed 117/117 executable cases with 20 documented environment skips and is archived at
    `Specs/TestRuns/4cb089111a23/FileOps/2026-08-24_120449`. Debug and Release x64 builds completed
    with zero warnings/errors and attested receipts
    `9203dc5252039c1893276a0baebdafeaaa755a437801ba033d0f9fade0c3542d` and
    `b29dc9356b3d8a89a9e28a788d5fed242aab1752e3f85d7730a26970681ab734`. Commit:
    `d6d44bab3` plus this following P2.7 hardening commit.

### Phase 3 — links, conflicts, and normal file-manager semantics

- [x] **P3.1 — Preserve/Skip runtime cutover and silent migration.**
  - [x] Remove Follow from provider UI/runtime/payloads and migrate legacy values to Preserve/Skip.
  - [x] Snapshot the provider default plus per-task confirmation choice; never reread live settings.
  - [x] Add provider capability/configuration/migration tests and remove obsolete Follow assertions.
  - Evidence/commit: the public ABI accepts only Preserve/Skip and reserves numeric zero; every
    shipped provider that consumes non-null `FileSystemOptions` validates the complete header before
    I/O. Local schema, serialization, UI, and runtime expose only Preserve/Skip. Legacy
    `followTargets` silently canonicalizes to Skip and legacy `copyReparse` to Preserve while
    preserving unrelated plugin configuration; unknown/numeric values fail closed. The host captures
    the provider default into the immutable typed plan and copies that value into every provider call,
    so workers never reread live settings; a future confirmation override has the same immutable
    field and does not persist. Direct provider precedence proves per-call Preserve overrides a Skip
    default and per-call Skip overrides a Preserve default without following or mutating the target.
    The obsolete session warning/grant and localized resources are removed. Full Debug x64 built
    with zero warnings/errors and accepted receipt
    `79ab5a7f9c060ddac3d4685d820f2d5c3fd645e35a72e80224c5c140419cf40e`.
    Commands source-contract and configuration/migration cases passed 1/1 each (archives
    `Specs/TestRuns/4cb089111a23/Commands/{2026-08-24_055801,2026-08-24_055831}`), provider contract
    tests passed including Local 95/95 and Microsoft Drive 219/219, and the final governed File
    Operations suite passed 116/116 with 21 documented skips and zero failure classifications
    (digest-bound compact archive
    `Specs/TestRuns/4cb089111a23/FileOps/2026-08-24_061531`). The archive makes no throughput claim;
    semantic in-tree retarget remains the single explicit P3.2 product skip. Commit: this P3.1
    work-package commit.
- [x] **P3.2 — Semantic link preservation and retargeting.**
  - [x] **P3.2a — Exact semantic payload and selected-root Managed retarget.**
    - [x] Read file-symlink, directory-symlink, and junction semantics from the retained no-follow
      binding without opening or mutating the target object.
    - [x] Apply selected-root `S -> D` meaning transformation for Managed tree Move, preserve
      outside-root target spelling, and recompute relative spelling from the destination link parent.
    - [x] Create an exclusive identity-owned destination link stage, conditionally publish it, and
      exact-delete the retained source object only after successful publication.
    - Evidence/commit: the public binding ABI now exposes bounded two-call `ReadBoundLink` and
      exclusive `CreateExclusiveLink`, with a new IID so stale binaries fail QI instead of calling a
      changed vtable. Local reads symlink/junction reparse data through the retained no-follow handle,
      supports sparse component transforms in the provider primitive, and returns an owned stage
      binding. The host Managed bridge reads semantic payload, applies the selected-root transform,
      publishes the exclusive stage with `PublishAs`, validates the published link snapshot, and
      calls `DeleteIfUnchanged` only after publication. The formerly deferred
      `Fairstream_MovedTreeRetargetsInternalLinks` case now passes and proves an in-tree junction
      resolves under the destination merge while the exact source tree is removed. Full Debug x64
      built with zero warnings/errors under receipt
      `ad6b105300bc9e20e0a66df93f3628b84dd0dbcbe1fc7a0513798ebcdc5fd6f7`; provider contracts passed
      including Local 102/102; the focused Commands source guard passed 1/1; and the final governed
      File Operations suite passed 117/117 with 20 documented skips and no failure classifications
      (digest-bound compact archive
      `Specs/TestRuns/4cb089111a23/FileOps/2026-08-24_075837`). The archive makes no throughput or
      Release-runtime claim. Commit: this P3.2a work-package commit.
  - [x] **P3.2b — Sparse Keep Both mappings and bounded dependency deferral.**
    - [x] Feed actual Keep Both component renames into the semantic transform rather than only the
      selected-root mapping.
    - [x] Defer unresolved forward targets within explicit count/depth/path bounds, process backward
      references immediately, and retain the source for every unresolved/failed Move dependency.
    - [x] Cover renamed ancestors, forward/back references, and cycles without constructing an
      unbounded whole-tree manifest.
    - Evidence/commit: the public bound-link payload now carries the bounded source-relative target
      dependency under a fresh binding IID, and `TryGetFileSystemRelativePath` supplies the shared
      provider-profile-aware derivation used by the host. Actual nested Keep Both choices populate a
      sparse component map. Sequential traversal classifies provider-order backward references from
      active directory-buffer views, while forward references/cycles and parallel-traversal links use
      bounded exact deferral. Failed prefixes retain dependent source links and held ancestor cleanup;
      no target object is followed and no whole-tree manifest/graph is retained. Full Debug x64 built
      with zero warnings/errors under receipt
      `ff450bb548c2593cfb1d0dc5d37ff6b4ac733d243e8cef196457335aede0b468`; Local provider contracts
      passed; focused Commands path-identity and ABI/source-contract cases passed 1/1 each; focused
      nested Keep Both and moved-tree semantic cases passed 3/3 each; and the final governed File
      Operations suite passed 117/117 with 20 documented skips and no failure classifications
      (digest-bound compact archive
      `Specs/TestRuns/4cb089111a23/FileOps/2026-08-24_133302`). The archive records sparse-map,
      immediate/deferred-resolution, and fail-closed source-retention counters and makes no throughput
      or Release-runtime claim. Commit: this P3.2b work-package commit.
  - [x] **P3.2c — Native qualification and complete link-kind/profile matrix.**
    - [x] Select Native Move only when the provider proves that the exact operation preserves the
      approved semantic transform; otherwise use Managed or Copy-only.
    - [x] Cover file symlink, directory symlink, junction, relative/absolute, root-relative, outside-
      root, native/managed/copy-only, and unsupported-provider cases.
    - Evidence/commit: Debug x64 Rebuild passed with zero warnings/errors and attestation
      `c77835b7e26ab9727b866343d9cff390b7ea5347f55de311004a8ec11bdc18b1`; Local provider contract
      tests passed 113/113; focused same-volume semantic-route and Phase 12 link-policy cases passed
      3/3 each; the focused Commands ABI/source-contract case passed 1/1; and the final governed File
      Operations suite passed 117/117 with 20 documented skips and no failure classifications
      (digest-bound compact archive
      `Specs/TestRuns/4cb089111a23/FileOps/2026-08-24_150932`). The archive records Local tree/link
      Managed classification, destination-absent absolute in-tree junction retargeting, Copy-only
      source retention, and the complete typed provider/profile matrix without making a throughput or
      Release-runtime claim. Commit: this P3.2c work-package commit.
- [x] **P3.3 — Typed conflicts and folder-manager semantics.**
  - [x] Implement distinct file collision, folder merge, type mismatch, destination link,
    replace-read-only, and target-conflict buckets.
  - [x] Preserve folder-on-folder merge and evaluate conflicts only on colliding children.
  - [x] Implement Keep Both, same-pane Copy duplicate naming, same-folder Move rejection, Skip All,
    and scoped Apply-to-all rules.
  - Evidence/commit: the host now refines provider status into typed regular-file, read-only,
    type-mismatch, exact-destination-link, name, target, retry, Recycle, and unsupported buckets from
    no-follow metadata before showing the prompt. Each cached receipt contains the exact bucket,
    source/destination kinds, exact link kind when known, and both path-profile IDs; plain Skip is
    never cached and explicit Skip all remains bucket-scoped. Folder-on-folder merges continue without
    a folder overwrite prompt. File collisions offer Overwrite/Keep Both, read-only files offer the
    dedicated replace grant, type mismatches withhold Overwrite, and destination links expose only
    link-safe actions. Same-pane Copy chooses a unique sibling; same-folder Move remains rejected.
    Local recursive Copy now defers preserved file/directory symlinks and junctions, feeds actual
    child Keep Both names into the sparse semantic transform, and publishes forward links only after
    all deferred names are resolved under the shared 4,096-entry/16-MiB ceiling. Callback-free direct
    provider Copy retains its explicit overwrite contract, while host operations cannot use an
    operation-wide overwrite flag to bypass a typed link/type-mismatch prompt. Full Debug x64 Fresh
    File Operations validation built with zero warnings/errors under attestation
    `3ff80344402dff244ce699a78b7e4eeb1c88e32235959ea95152272bfda361f2` and passed 117/117 with
    20 documented skips and no failure classifications (digest-bound compact archive
    `Specs/TestRuns/4cb089111a23/FileOps/2026-08-24_190228`). Focused typed-conflict, Fairstream,
    direct-provider reparse, and final-link race cases also passed. The archive records the exact
    conflict, nested Keep Both, and semantic Copy rows without making a throughput or Release-runtime
    claim. Commit: this P3.3 work-package commit.
- [x] **P3.4 — Identity-safe metadata/link mutation and phase gate.**
  - [x] Apply no-follow identity validation to link replacement, metadata operations, and conflict
    revalidation.
  - [x] Prove no tree Move leaves an in-tree relative link targeting a removed source and folder merge
    never becomes overwrite.
  - [x] Run link/conflict/provider matrices, builds, specs, and archived bounded-retarget evidence.
  - Evidence/commit: destructive conflict decisions now carry an exact no-follow destination
    authority into one conditional mutation. Link replacement, read-only replacement, exclusive
    directory/link publication, and post-copy metadata writes never authorize mutation from pathname
    text alone; providers without exact metadata authority emit a diagnosed unsupported result and
    preserve provider defaults. Folder-on-folder remains merge, a source link over an ordinary
    destination kind is a typed mismatch, destination links are never merge namespaces, and semantic
    Move publishes mapped links before exact source cleanup. The staged-overwrite fault hooks execute
    at the live identity-owned `PublishAs` boundary and prove pre-publication failure preserves the
    original destination bytes and attributes. The final governed Debug x64 File Operations run
    `20260824T224940Z-54068-2e418dbae5a14677a21adda6aba2b9dc` passed 117/117 executable cases with
    20 documented environment skips and zero failure classifications under build receipt
    `da7f142e3ecb76f28755a7c9b35f7d569c908c8203b55f4b4f8fa3f94a012ef6`; its digest-bound compact
    archive is `Specs/TestRuns/4cb089111a23/FileOps/2026-08-25_010517`. The test-enabled Release x64
    full solution then rebuilt with zero warnings/errors under receipt
    `837bd0f71b8561b7fb9c47c277f802126ddfc7cb46a652b5d74f100dfb3f0020`. Release does not publish the
    Debug-only File Operations case inventory, so this package makes no Release-runtime, throughput,
    or latency-percentile claim. Commit: this containing Phase 3 work-package commit.

### Phase 4 — confirmation, consent, metadata, results, and clipboard

- [x] **P4.1 — Immutable confirmation and prompt payload.**
  - [x] Implement Links and Verify controls, clipboard-consumption copy, Queue/Parallel/bandwidth
    snapshot, and any required breaking `HostPromptRequest` payload.
  - [x] Keep one confirmation owner per ingress and capture immutable intent before workers start.
  - Evidence/commit: `HostPromptRequest::fileOperationOptions` exposes synchronous caller-owned
    Links, Verify, execution-mode, bandwidth, and cut-list-consumption state; `AlertOverlayWindow`
    renders and exposes the four keyboard/mouse/UIA option rows; accepted values are copied into
    every mutable child plan before final validation, the gated-worker clipboard barrier, and
    immutable task publication. Cancel leaves caller state unchanged. Debug x64 rebuilt cleanly with
    zero warnings/errors under receipt `71cda137085c549e847bf479b4c93daa13a87df106255e0a2222f2794bf187f8`.
    Commands run `20260824T234510Z-82948-015bf77dd94549d98e006b092325eea0` passed the live prompt/UIA
    case; File Operations run `20260824T234545Z-69716-69cfa1890ebe4288a3482b42bf53bc6a` passed the
    Phase 10 confirmation/immutable-plan cases 3/3; `ResourceLocalizationContracts.Tests.ps1`
    passed all six contracts. Commit: this containing P4.1 work-package commit.
- [x] **P4.2 — Shared deferred-consent gate.**
  - [x] Implement task pause/resume around EFS plaintext, placeholder hydration, sparse inflation,
    metadata loss, low-space continuation, and Recycle escalation decisions.
  - [x] Use exact conflict records, one-shot/scoped grants, safe defaults, revalidation, and no
    session-sticky consent.
  - Evidence/commit: the shared task arbiter now publishes seven typed risk layouts with immutable
    item/byte/identity facts, Cancel-last/off-layout rejection, task/risk/root receipts bounded to
    64 entries, and separate `FileOps.Consent.*` metrics. `FileSystemItemMutationResult` carries
    per-item provider mutation truth across the in-tree callback ABI. Recycle escalation requires a
    known non-commit, the retained pre-operation authority, a fresh exact delete bind, item-only
    consent, and `DeleteIfUnchanged`; missing/changed/indeterminate truth never prompts or retries by
    pathname. Debug build receipt `34523eaa7f09355d5baa0159b4515a88f962adfc40c5259780e65c2b6e328aff`
    completed with zero warnings/errors. File Operations run
    `20260825T003510Z-96840-7fb790ce4ac743a6819b87ca7368ad7e` passed the Phase 10 family 4/4;
    archived evidence is `Specs/TestRuns/4cb089111a23/FileOps/2026-08-25_023804`.
    `ResourceLocalizationContracts.Tests.ps1` passed all six contracts. Commit: this containing P4.2
    work-package commit.
- [x] **P4.3 — Metadata and source-retention matrix.**
  - [x] Implement ADS/MOTW, EA, timestamps, attributes, sparse, compression, ACL/owner, EFS, and
    placeholder outcomes with exact warnings/loss reporting.
  - [x] Prevent Move source deletion after ungranted security-significant loss.
  - Evidence/commit: optional exact `IFileSystemBoundMetadata` now transfers allocation/encryption
    state before content and streams/EA/basic/security evidence after content through retained source
    and owned-stage authority. Local actual MOTW/ADS/basic/sparse preservation plus deterministic
    placeholder/EFS/sparse and exact-loss source-retention cases passed 3/3 in Fresh File Operations
    run `20260825T012149Z-3684-678ec6bbb27a4b9da394c7ea8d074acb`; archived evidence is
    `Specs/TestRuns/4cb089111a23/FileOps/2026-08-25_032437`. Debug receipt
    `9395ab91df29ee6899cd7790f656f695d2637e8679f58580bbec92c25b1ffa7f` completed with zero
    warnings/errors. The archive records ten bounded metadata inspection/apply/matrix rows and makes
    no throughput claim. Commit: this containing P4.3 work-package commit.
- [x] **P4.4 — Typed results and every consumer.**
  - [x] Emit per-item Published/Verified/SourceDisposition/Completion axes and canonical details.
  - [x] Map axes to HRESULT, task state, badges, Issues, telemetry, FolderView, Find, Compare,
    watch/index refresh, and removal focus without pathname inference.
  - [x] Preserve/refresh rows for Retained/Unknown; remove only exact `Removed`.
  - Evidence/commit: the engine stores one typed result per selected source, reduces aggregate state
    from those axes, and carries the same truth through completion events, localized popup summaries,
    cache/watch refresh, removal focus, Find, and Compare. Missing destructive truth is
    indeterminate and never receives the legacy `S_OK` removal fallback. The governed Debug x64
    focused result run `20260825T015602Z-78104-0e8da17e2f76413e8868ce49b225136e`
    passed 3/3, the Find exact-removal case
    `20260825T015631Z-62664-d6be412bf4e142fd807317ca14aad8fe` passed 1/1, and the provider/
    indeterminate matrix `20260825T015700Z-94080-4a29ed1f41114dd8a68f7c9307b2a835` passed 3/3
    against full-solution Debug receipt
    `c00998e729b6d1766a5d1cc0580cf8fc2c39102d7f5911af97cb21bbb93398d5`. The result matrix emits
    bounded `FileOps.Result.Items` and `FileOps.SelfTest.TypedResults` rows. Commit: this containing
    P4.4 work-package commit.
- [x] **P4.5 — Clipboard admission and retained-source UX.**
  - [x] Clear a matching Move cut sequence after complete accepted queue admission and before worker
    release; never restore it.
  - [x] Implement retained/unknown actions, explicit Cut retained items again, and duplicate-paste
    race coverage.
  - Evidence/commit: task publication and bounded sequence admission are atomic; a successfully
    created worker waits on a task-owned gate while FolderView consumes the still-matching sequence.
    Thread-admission failure rolls the task/sequence receipt back before clipboard access; clear
    failure is retained as a warning and still releases the accepted task. Completed retained-source
    actions rebind every exact source no-follow and reject one missing/replaced object before Select
    or explicit Cut-again. The final full Debug x64 solution rebuilt with zero warnings/errors under
    receipt `06a86654afa788842c7e355d3d1a94fd6fbb1b78687dc313c52b8f0db0360947`.
    File Operations run `20260825T025226Z-40856-e0fb7567f1474829879bb2aa25f36607` passed 3/3;
    the real clipboard duplicate-paste run
    `20260825T025325Z-89456-9c35023560804aa5ac6ec02e6108a816` and source-contract run
    `20260825T025256Z-30764-a1ce2941e7d7480ca5bb07c2d9da4bad` each passed 1/1 against that exact
    receipt. The bounded evidence archive
    `Specs/TestRuns/4cb089111a23/FileOps/2026-08-25_045246` passes explicit pre-stage validation;
    `ResourceLocalizationContracts.Tests.ps1` passed all six contracts. No throughput improvement
    is claimed. Commit: this containing P4.5 work-package commit.
- [x] **P4.6 — Optional verification UI and phase gate.**
  - [x] Implement verification states, BLAKE3/provider proof, mismatch/cancel outcomes, transfer and
    hatched verification segments, per-file bars, graph series, combined ETA, and bandwidth accounting.
  - [x] Prove verification off adds no reads/work and verification never becomes delete/skip authority.
  - [x] Run metadata/consent/result/clipboard/verification matrices, UI tests, builds, and performance
    evidence; audit the Phase 4 stop conditions.
  - Evidence/commit: immutable Verify intent now selects one exact-authority publication/proof route
    for every non-Native copied payload; Native Move remains `NotApplicable`, Verify Off constructs no
    hasher/readback work, and requested proof never authorizes conflict resolution or source deletion.
    Verification-capable tasks serialize transfer then proof, charge both byte streams to the shared
    bandwidth budget, and expose determinate compact/expanded progress only after discovery closes.
    The hosted segmented progress control owns accessible transfer/verification presentation; popup
    results, Issues, telemetry, and retained-source text consume the typed axes. Provider proof,
    host readback, zero-byte, mismatch, unavailable, cancel, Native-Move, serialized two-file, and UI
    segment cases passed 3/3 in final Debug and test-enabled Release exact runs. Debug x64 rebuilt with
    zero warnings/errors under receipt `c2a99fcd336d44a4927f11ed2cf049af4925ad7a46f29ae96a8ec7b7f9cb8a31`;
    test-enabled Release x64 rebuilt with zero warnings/errors under receipt
    `1e2aab69ed538e8bc27dd2b970810f1ba06de632178a5cb92f98ae0c76da2682`.
    The final governed Debug evidence archive is
    `Specs/TestRuns/4cb089111a23/FileOps/2026-08-25_133449` and passes explicit pre-stage validation.
    DxUi, provider-contract, settings-schema, localization, and the retained bridge/result focused
    matrices passed. A broad Debug File Operations run passed 120 cases with 21 declared-capability
    skips and one unrelated Phase 6 cancellation-latency sample at 266 ms against 250 ms; the exact
    Phase 6 rerun passed 3/3 at 32 ms. No throughput improvement is claimed. Commit: this containing
    P4.6/Phase 4 work-package commit.
- [x] **P4.7 — Stable conflict attention, Recycle default, and exact result closeout.**
  - [x] Reveal `Needs attention — Reading details` immediately, expose no decision buttons while
    no-follow metadata is loading, and publish one stable typed action set only after it resolves.
  - [x] Make Cancel first, default, and Escape for Recycle escalation; keep the general Cancel-last
    rule for every other conflict/consent surface.
  - [x] Remove provider Copy handling that could reinterpret `PermanentDelete` as a Recycle opt-out.
  - [x] Keep selected directories `VerificationState::NotApplicable` while nested file proof still
    drives task failure/cancel/source-retention truth, and preserve exact bridge publication on every
    post-copy failure/cancel path.
  - [x] Update authoritative File Operations, bridge, and popup specs plus deterministic runtime and
    source-contract coverage.
  - [x] Complete final Release/Fresh validation, record receipts below, and commit the scoped package.
  - Evidence/commit: Debug x64 full solution rebuilt with zero warnings/errors under receipt
    `f3a9d9bf615410ff214bc26d47fe34e68a62b8751b97b01b43c177e1fb32bac2`.
    File Operations run `20260825T155822Z-45316-18537a0217b84e5ea9a976dc63ccd981`
    passed the complete Phase 10 family 8/8. Commands runs
    `20260825T155909Z-69096-444659125e284c5f983424036c2f3466` and
    `20260825T155949Z-94676-55ac137ac89e48068626cc13d53d760a` passed the exact conflict/popup
    cases 1/1 each. The focused Pester source/inventory/documentation contracts passed 201/201;
    localization passed 6/6 and `git diff --check` passed. The final Fresh Full run
    `20260825T160714Z-79580-bd13bfcdb4f044e0ae9997f3b35dfaa9` completed 1,830 cases with 1,792
    passes, 34 documented skips, and four broad-run failures. Its File Operations process executed
    53 behavioral cases with 53 passes before the runner's three-minute aggregate timeout interrupted
    the remaining families and reported the coverage inventory missing; the complete Phase 10 family
    remained green 8/8. Three Commands broad-order UI failures were outside the Phase 4 mutation
    contract; the potentially related completed-group case passed its isolated rerun after the Fresh
    run, and both exact conflict/popup cases above remained green. The test-enabled Release solution
    compiled with zero warnings/errors, but post-build attestation could not finish because the shared
    runtime staging step did not place `sqlite3.dll` at the Release output root even though identical
    built/vcpkg copies were present under the provider and vcpkg output directories. This is recorded
    as build/deployment infrastructure debt rather than hidden Phase 4 evidence. Commit: this
    containing P4.7/Phase 4 closeout commit.

- [x] **P4.8 — Cross-phase exact-object, link-kind, and executor hardening. — COMPLETE**
  - [x] Replace the Local semantic-link retarget Boolean with an explicit three-way result:
    `OutsideSourceRoot`, `MappedInsideSourceRoot`, or `InsideSourceRootUnmappable`. Only the first
    preserves stored spelling; the last enters `TargetConflict`, publishes no guessed link, retains
    the exact source link and dependent ancestry, and never authorizes Managed cleanup.
  - [x] Make host mutation strategy, destination-directory ensure, overwrite policy, and conflict
    refinement consume one authoritative object-kind classification. A name-surrogate reparse is a
    Link; a known non-name-surrogate cloud/WOF/dedup reparse retains file/directory kind; unknown or
    unavailable kind fails closed without exposing Replace link. Extend the breaking ABI or consume
    bound snapshot/tag authority where `FileSystemBasicInformation` cannot represent this truth.
  - [x] Make ordinary Permanent Delete preserve confirmation-to-mutation identity. Local consumes
    the retained no-follow source authority through `DeleteIfUnchanged`; a provider without host
    binding may proceed only through a contract-tested provider-native exact-identity mutation, never
    by silently rebinding and deleting a pathname replacement. Recycle's shell-owned bulk carve-out
    remains unchanged.
  - [x] Withhold a destructive destination-link action unless the selected source/destination shape
    can consume the exact destination receipt; cover file-on-directory-link and directory-on-file-link
    without a dead Replace link button.
  - [x] Feed identical traversal-limit generation/status facts into serial and parallel item-terminal
    classification. The working tree now uses one classification input helper. The larger executor-loop
    convergence is deliberately retained as P5.7 architecture work so it is not misreported as part
    of this safety closeout.
  - [x] Replace PID/TID/tick case-only rename temporary names with cryptographic/exclusive ownership;
    collision or entropy failure stops before moving the source and no replacement flag may overwrite
    an unowned temporary object.
  - [x] Preserve `Needs attention — Reading details` until the complete typed scope is known. A cached
    Apply-to-all decision may bypass metadata I/O only when authoritative source kind, destination
    kind, link kind, and both path profiles are already available; a provisional HRESULT/bucket is
    never sufficient cache authority.
  - [x] Add deterministic inside-root-unmappable link, WOF/cloud non-surrogate, exact Permanent Delete
    replacement race, file-on-link action, serial/parallel traversal-limit, and case-temp collision/
    entropy tests. Reuse/add bounded `FileOps.*` counters and archive same-machine File Operations
    evidence; record focused Debug, test-enabled Release, Fresh Full, `git diff --check`, and commit.
  - Working-tree evidence (2026-08-25): the ABI and Local provider now expose tri-state link-target
    mapping; an inside-root unmappable target returns `TargetConflict` and never publishes guessed
    source spelling. Host strategy, destination ensure, overwrite, and conflict refinement consume
    no-follow bound object kind, with name-surrogate tag refinement only for link subtype. Permanent
    Delete snapshots before confirmation, binds after confirmation, cross-checks exact object/revision,
    and calls retained `DeleteIfUnchanged`; the replacement-race and recursive-root cases pass. Case-only
    rename uses the shared CSPRNG unique-sibling generator and exclusive first hop. Serial and parallel
    item completion consume the same traversal-limit fact helper. `Needs attention — Reading details`
    remains non-actionable until metadata resolves.
  - Release-closeout remediation (2026-08-25): the exact Permanent Delete Retry-cap fixture now
    injects two known non-commits at `DeleteIfUnchanged`, proves one Retry followed by a no-Retry
    prompt, and records exactly two attempts with the source retained. Exact regular-file delete
    reports committed-size discovery through the governed operation-control callback and promotes
    those bytes to completed progress only after committed removal. The Local delete TOCTOU swap hook
    now runs in all test-enabled builds rather than Debug only. The Dummy connection-override fixture
    records the authority boundary explicitly: `deleteMaxConcurrency` can clamp admitted work but
    cannot authorize Permanent Delete without exact binding; provider-direct recursive deletion is
    test teardown only. The Dummy cross-pane refresh fixture uses its provider-owned Recycle route.
    Focused governed Release passes: Retry cap
    `20260825T204419Z-79144-a6fae8313b544e10aa8456f1e8ed91c9`, delete TOCTOU
    `20260825T205754Z-94520-ad478884a6fa4b6c96ff1a9afe734db0`, meaningful delete bytes
    `20260825T205822Z-94520-d6d3ff5789fb4a2fa2f18d9685739f06`, and Dummy visible refresh
    `20260825T205848Z-94520-f519309e32eb485080c86a8d71a085fb`. The connection-override rerun and
    containing Fresh Full remain pending after aligning the obsolete Delete clamp assertion with the
    exact-authority contract.
  - Additional closeout evidence (2026-08-25): connection override passed 3/3 under
    `20260825T210719Z-39388-a3a4d6d0db3145969fef2a7a9c1b5b4d`; the complete source-contract
    Pester file passed 176/176 after updating the case-only rename guard to the canonical CSPRNG name
    helper. Isolated Compare run `20260825T211242Z-9016-48d1bf5cbc4b46149a7ad5829f7e8698`
    exposed a Release-only selftest control-flow defect: a nested MTP setup helper returned the truthy
    `Skip` value as successful creation, then dereferenced an empty interface. The helper now records
    Skip and returns false; the direct focused Release repro exits cleanly. This is harness closeout,
    not a File Operations behavior change; governed rerun remains pending the next full attestation.
  - Focused governed test-enabled Release cases passed with no failure classification:
    `Phase10_PermanentDelete` (`20260825T194419Z-63624-e3fbf4d8f99245648b5e62cbb0acdc45`),
    `Fairstream_MovedTreeRetargetsInternalLinks`
    (`20260825T194503Z-63624-663c33c02854429ba3bfed735ad0fd6d`),
    `Phase9_ConflictPrompt_LocalFileOntoDirectory`
    (`20260825T194530Z-63624-a6e5ab00b4ae40f0b9625662a1e2a3aa`), and
    `FileOps_MoveMergeIntoExistingFolderSameVolume`
    (`20260825T194557Z-63624-e7f883b416f348b3aebb334d4a059016`). The Commands typed-plan source
    guard passed under `20260825T194304Z-71988-b7aafcc3b3e04d959eea597ec091cf3b`. The final full-solution,
    test-enabled Release build passed with zero warnings/errors (attestation
    `8e6369a1f8ab15955db6146488920ca29a260e8fe285a51ac7e3f2e5a36334e2`).
    `git diff --check` reports no errors. Debug/Fresh Full and the containing commit remain pending
    because `.build/artifact-operations/x64-debug-aea633ba0ab3.lock` is still owned by the pre-existing
    live PowerShell PID `29548`; it was not terminated or bypassed.
  - Final integration/hardening evidence (2026-08-26): the main implementation is committed in
    `3593bd8fc`. The follow-up applies the authoritative name-surrogate classifier to every Local
    Native Copy root, serial child, parallel child, and destination-ensure branch. Non-surrogate
    WOF/cloud/dedup reparses retain regular file/directory behavior; tag-query failure stops before
    publication. Test-enabled Debug build attestation
    `8eb8d62de2e9bc11506be367b6650a6694d2422679871e441dd58d614f6fc48d`
    completed with zero warnings/errors. Plugin contract/debug selftests passed, including Local
    FileSystem 129/129. Recursive Copy matrix
    `20260826T060654Z-70948-5d0a94b11f474453a62656f3d5fcffca` and typed overwrite/read-only
    `20260826T060721Z-70948-883b30b6a9ba413bbc3b923a0601948e` passed 3/3 each. Test-enabled
    Release build attestation
    `ea06314c17fec130c512a01eecd887672f325752c7060a5420b75cd757118fc4`
    also completed with zero warnings/errors; the Release recursive Copy and typed-conflict reruns
    passed under `20260826T061613Z-115440-b7f3cca0821e493f848b73cbbbe4a3ae` and
    `20260826T061644Z-115440-75153982c1e2407980a185f30f2c7861`.

### Phase 5 — Rename, artifacts, recovery, and S3 directories

- [x] **P5.1 — Final Inline-F2 presentation and rename strategies. — COMPLETE (`3593bd8fc`)**
  - [x] Add immediate non-clean reveal, fixed 500-ms clean-running reveal, silent clean success,
    identity focus retention, fake-clock tests, and no duplicate legacy alert.
  - [x] Finish provider native/managed rename and conflict/result coverage. An existing destination used
    by `RenameIfUnchanged` must be bound with rename/publication authority, not metadata-only authority;
    otherwise an accepted Overwrite receipt is unconsumable and must fail rather than reprompt forever.
  - Working-tree evidence (2026-08-25): the presentation state machine implements immediate non-clean
    reveal, the fixed fake-clock 500-ms deadline, silent clean success without active/completed card,
    once-visible-always-visible behavior, identity focus retention, and production Release linkage.
    A focused run exposed and preserved a two-prompt timeout caused by a metadata-only destination
    receipt; Inline Rename now binds the exact destination with `READ_METADATA | RENAME | PUBLICATION`,
    and the test fails fast if an accepted grant ever produces a second stable prompt. The corrected
    governed `Floodgate_InlineF2WorkerQueueBypassAndInterlock` run passed 3/3 with no failure
    classification under `20260825T194337Z-48068-f69b34950c8e4755a0bf030426e74a4e`. The initial timeout evidence is
    retained locally under `Specs/TestRuns/SINON/Continuation/2026-08-25_2128_fileops_inline_rename_focused_hang/`.
    The implementation and archived focused evidence are committed in `3593bd8fc`.
- [x] **P5.2 — Batch Rename central-engine migration. — COMPLETE (`3593bd8fc`)**
  - [x] Submit every one-step/multi-step/cycle plan as `RenamePlan(BatchRename)` through ordinary
    queue/card admission.
  - [x] Remove the private executor exemption and map per-step results/cancellation truthfully.
  - [x] Capture one path-scoped endpoint and ingress identity snapshot per changed row; reject a
    mixed root/profile or changed source before mutation. Dependency layers and cycles are immutable
    plan data, while every actual rename still rebinds/revalidates exact no-follow authority.
  - [x] Admit one row per physical source object. Preview blocks duplicate provider paths using the
    path profile; central admission compares exact no-follow object identities and rejects the whole
    plan when aliases or hard links name the same object. It never invents an execution order for
    duplicate rows or permits two final mappings to mutate one object.
  - [x] Make the former Batch Rename execution engine a pure dependency/cycle scheduler invoked by
    the File Operations task. It may not call `RenameItems`/`RenameItem` itself; final steps, temporary
    hops, and best-effort restoration all cross the engine-owned conditional rename boundary.
  - [x] Route the Batch Rename window's Cancel command to the admitted task, keep progress on both the
    ordinary task card and the generation-scoped window status, and derive report/undo/success refresh
    only from typed terminal item results.
  - [x] Add deterministic one-step, chain, swap/cycle, partial failure, cancellation, source-replacement,
    queue/card, and stale-window-completion coverage plus execution/admission performance evidence.
  - [x] Preserve the explicit Batch Rename selection-level recheck exception: after admission and
    before the first mutation, rebind every changed source and check every unrelated final namespace
    slot once. Planned-source targets remain dependency/cycle edges; the pass grants no replacement
    authority, performs no recursive discovery, and cannot replace exact per-hop revalidation.
  - [x] Unless cancellation is requested, attempt every independent row admitted to the current
    runnable layer after a sibling fails or is skipped, but never unlock a dependent layer/cycle
    through a retained source.
  - Working-tree implementation (2026-08-25): the window-owned mutation worker is removed. The
    window submits one generation-scoped central request, routes Cancel to its returned task id, and
    receives progress/completion only through tokenized UI messages. `AdmitBatchRename` captures one
    path-scoped endpoint plus an immutable ingress snapshot per row, rejects mixed root/profile plans,
    and records dependency cycles. The former execution engine is now a provider-free scheduler;
    final rows, exclusive random temp hops, and restoration all call the File Operations task's
    retained-authority `RenameIfUnchanged` boundary. Returned bound authority is captured and carried
    across every hop. Typed per-row results distinguish Completed, Skipped, failed-before-publication,
    and indeterminate retained-temp outcomes. The worker now performs the specified one-time
    selection-level source/unrelated-target recheck before any row mutation, and the scheduler finishes
    already-runnable independent siblings without advancing a dependency blocked by a retained source.
    Existing window execution tests now use a test-only central-task host rather than restoring a
    production bypass. Production source-replacement and ordinary Queue/card coverage pass, and Batch
    Rename admission plus the 1,024-row independent scheduler scenario emit explicit performance
    metrics.
  - Focused Release evidence (2026-08-26): the final test-enabled x64 Release build passed with zero
    warnings/errors (attestation
    `e0d1d5c75533dfd637bb82ca4a0de88ab816f46c11ed19b83dbd0316c0f5b651`; log
    `.build/logs/msbuild-20260826_011110_750-pid116548-029d85f1.log`). Direct scheduler
    (`2026-08-26_011845`), swap (`2026-08-26_010447`), chain (`2026-08-26_010516`), partial
    failure (`2026-08-26_010545`), cancellation (`2026-08-26_010612`), no-bypass unsupported
    provider (`2026-08-26_011553`), production source replacement (`2026-08-26_011809`), stale
    generation (`2026-08-26_012210`), and close-while-running (`2026-08-26_012248`) cases each
    passed and are archived under `Specs/TestRuns/4cb089111a23/Commands/`. The source-contract case
    passed under `2026-08-26_012049`; the complete source-contract Pester file passed 176/176.
    The close-while-running case result is green; its runner aggregate remains red only because the
    preserved TestSandbox disk audit reports 110 historical entries, which this package does not
    delete or hide.
    Production pane admission/refresh passed under `2026-08-26_011730` and emitted
    `batchrename.admit.us` in 9,166 us for two rows. The 1,024-row independent scheduler baseline
    passed under `2026-08-26_011642`, emitted `batchrename.execute.us` in 188,143 us, completed every
    row with exactly one conditional mutation each, and makes no throughput-improvement claim. The
    selected archives pass explicit pre-stage validation (108 files) and the complete TestRuns
    inventory passes (1,117 files). Repository integration is committed in `3593bd8fc`.
  - Duplicate-source hardening evidence (2026-08-26): the Commands Batch Rename family passed 76/76
    under `20260826T080926Z-71492-43a286d89aa646ce8db685ffad097ed7`, including the localized
    path-duplicate preview block. The focused File Operations identity suite passed 3/3 under
    `20260826T081141Z-114988-13efc7efd1674391a7f7dd264c52ee44`, including two Local hard-link
    aliases rejected before task publication. The complete source-contract file passes 177/177 and
    the resource-localization contract passes 6/6. These cases distinguish textual duplication from
    exact physical identity and prove that neither form reaches a rename mutation.
  - Evidence/commit: `3593bd8fc`.
- [x] **P5.3 — Durable rename journal and recovery. — COMPLETE (`3593bd8fc`, hardened by this containing commit)**
  - [x] Before the first mutation, atomically create
    `%LOCALAPPDATA%\RedSalamander\State\FileOperations\<task-id>.json` with a versioned schema,
    qualified endpoint, original/current/final provider paths, exact no-follow object/revision
    identities, kind, depth, and final leaf for every admitted row. The record is
    selection-proportional and bounded by the accepted selection; it is not a recursive manifest.
  - [x] For every final rename, exclusive `.rs_ren_` cycle hop, and restoration hop, durably record
    an `inFlight` intent before calling `RenameIfUnchanged`, then durably record `committed` plus the
    provider-returned exact identity before another mutation may begin. A known non-commit removes
    only that intent; an unknown outcome retains it for recovery inspection and is never retried.
  - [x] Keep a failed/canceled/partial journal whenever any committed or indeterminate step may need
    recovery. Remove it only after a clean terminal result, or after a terminal result that proves no
    mutation committed. Journal write/flush/publish failure before a mutation prevents that mutation;
    failure after a committed mutation returns indeterminate and preserves the last durable record.
  - [x] On recovery load, reject unknown schema/kind, malformed or oversized identities, mixed or
    missing endpoint facts, non-contiguous step sequences, and impossible item/step transitions.
    Reconcile the single possible `inFlight` step by reopening both recorded names no-follow: source
    match means known non-commit, destination match means committed, neither/dual/replacement means
    indeterminate and offers no mutation action.
  - [x] Implement `Resume` by rebuilding the remaining dependency/cycle schedule from each row's
    exact current identity and final mapping. Before every resumed hop, rebind the recorded current
    name and require the exact recorded object/revision; require an unrelated destination to remain
    absent and use an exact destination receipt for any approved replacement. Never repeat a step
    whose outcome is unknown.
  - [x] Implement `Roll back` by traversing committed steps in reverse order. Each inverse hop binds
    the committed destination identity and requires the prior name to be vacant or to be the same
    recorded object. A replacement object, missing/ambiguous identity, changed endpoint/profile, or
    failed durable transition stops recovery and leaves the journal. Roll back is not Undo/G2 and is
    offered only for this interrupted task.
  - [x] Keep `Open location` read-only and available when mutation recovery is withheld. P5.4 owns
    startup discovery, Proven/Possible classification, badges, touch warnings, and the recovery UI;
    P5.3 owns the exact journal parser/planner/executor those surfaces call.
  - [x] Add deterministic create/begin/commit/known-noncommit/terminal fault injection, crash after
    every rename step, source/destination replacement, in-flight reconciliation, Resume, reverse
    Roll back, cycle/intermediate-name, cancellation, malformed-record, and clean-removal tests.
    Archive journal write latency/bytes/row-count and large-selection load/plan evidence; do not make
    a throughput-improvement claim.
  - Focused Release evidence (2026-08-26): test-enabled x64 Release build passed with zero
    warnings/errors (attestation
    `f61ddbf4dbb594851b4525a03acccbefe66c1a08c44f2ba6d111c06a320cd52c`; log
    `.build/logs/msbuild-20260826_025451_481-pid41848-a7d7f6d9.log`). The four recovery cases passed
    under `2026-08-26_025844` and its nine-file archive passes explicit pre-stage validation. They
    cover every durable transition fault, provider-committed/inFlight reconciliation, every committed
    checkpoint of a two-name cycle, Resume, reverse Roll back, cancellation, current-name and
    prior-name replacements, malformed schema rejection, exact clean removal, and retry after a
    retained record. The 1,024-row record is 683,213 bytes; `batchrename.journal.write.us` recorded
    3,514 us and `batchrename.journal.load.us` recorded 6,400 us. These are durability/boundedness
    baselines and make no throughput-improvement claim. The implementation is committed in
    `3593bd8fc`. The follow-up restores the live journal to its last durable `inFlight` state when
    commit persistence fails, prevents same-instance `Finalize` from manufacturing a committed
    transition, and stops Roll back on an unexpected current path instead of skipping later inverse
    hops. Test-enabled Debug build attestation
    `8eb8d62de2e9bc11506be367b6650a6694d2422679871e441dd58d614f6fc48d`
    completed with zero warnings/errors; durable-transition run
    `20260826T060810Z-70948-ecf3feb814554b0fac2f17afd30bdd64` passed. Test-enabled Release
    attestation `ea06314c17fec130c512a01eecd887672f325752c7060a5420b75cd757118fc4`
    passed the durable-transition and replacement Roll-back cases under
    `20260826T061739Z-115440-7657c9beafc0444f92ea00e67b96b903` and
    `20260826T061805Z-115440-88e7ebbab7b4418caad46b274c06d770`.
  - Evidence/commit: `3593bd8fc` plus this containing hardening commit.
  - Test-host shutdown hardening (2026-08-26): repeated `0xC0000005` exits in
    `BatchRenameCountingReadDirectoryFileSystem::Release` were traced to completed central-test tasks
    retained in `g_batchRenameCentralTestTasks` until CRT static destruction, after plugin teardown.
    The Batch Rename Commands family now owns a final cancel/move/join/release quiet point. The x64
    Debug build passed with zero warnings/errors (attestation
    `c134efca3739b2fd6371a837d6585a32e4c566213979b04d103ec2ec0c5b63ff`; log
    `.build/logs/msbuild-20260826_140040_222-pid68848-035c6838.log`), the wrapper-backed exact case
    passed in ten separately awaited processes with no new crash sidecar, and source contracts pass
    178/178. The original 24 matching stacks and external dump hashes are retained under
    `Specs/TestRuns/SINON/Continuation/2026-08-26_batchrename_test_task_shutdown_crash/`.
- [x] **P5.4 — Artifact registry, display, warnings, and recovery.**
  - [x] Treat each atomically published, versioned per-task recovery record as the claim store; scan
    only canonical `<16-hex-task-id>.json` regular non-reparse records, require the filename task id
    to match the strict payload, quarantine malformed siblings without discarding other valid claims,
    and never invent authority from a partial parse.
  - [x] Project exact claims only for mutation-bearing durable phases. A committed temporary/final
    Batch Rename row claims its exact current identity; the one permitted `inFlight` step exposes
    source and destination alternatives for classification, but only the name whose current no-follow
      identity matches can become Proven. Successful terminal records grant no artifact authority.
  - [x] Implement one shared endpoint + provider-path-profile + current no-follow identity join:
    matching valid claim and identity is Proven; a claimed-path identity mismatch, unavailable probe,
    malformed/missing claim, or recognized current/legacy generated-name shape is at most Possible;
    everything else is Ordinary. Cached UI/index values are display hints only.
  - [x] Remove generated-name suppression from local Search/index and its obsolete source contract.
    Add runtime visibility coverage for `.rs_tmp_`, `.rs_bak_`, `.rs_ren_`, `.rs_copy_tmp_`, and
    `.~rs-write-`; no pane/Find/Search/Compare filter may hide either artifact class.
  - [x] Define one semantic projection payload: classification/reason, qualified location, task/time,
    intended final name, durable phase, and Inspect/Reveal/Open-location/Resume/Roll-back capability
    bits. Proven alone receives phase-authorized recovery capabilities; Possible is inspection-only;
      the payload is display state and never mutation authority.
    Durable phase authorization is explicit: in-flight or partially committed unfinished records may
    expose Resume and Roll back; fully final retained records expose Roll back only; records with no
    in-flight or committed forward mutation expose neither.
  - [x] Connect that projection to FolderView, Find/Search, Compare, and the File Operations popup,
    including localized accessible badge/status text and refresh/re-query behavior before every
    authority-bearing action.
    - Working-tree progress: Folder enumeration captures the provider instance with the enumeration
      request, loads one registry snapshot, and performs path-scoped capability/no-follow identity
      projection only for recognized candidate names. FolderView renders a persistent warning glyph
      in every mode; Compare inherits it through the same FolderViews. Find loads once per search and
      renders localized Proven/Possible warning badges. Registry/qualification/binding failure keeps
      the row visible as Possible. Popup/status projection, accessible action descriptions, and
    the popup projects active/terminal recovery cards and task-observed artifact warnings. It is not
    a provider-wide artifact scanner: claimless Possible objects are discovered and persistently
    badged by FolderView/Find/Search/Compare without being hidden. Localized status text and menu/card
    action names provide the accessible classification and recovery descriptions.
    The current presentation slice also appends the localized Proven/Possible text to the focused
    pane status bar and adds an artifact-only context submenu with host-internal read-only Inspect
    and Reveal in pane. Phase-eligible Proven rows now add Resume/Roll back only when the strict
    recovery callback is connected; projection capability bits alone never create a live button.
  - [x] Implement the core pre-admission exact-set receipt: only Proven/Possible rows are captured;
    ordinary rows are excluded; the receipt is ephemeral/non-Apply-to-all; any set/path/endpoint/
    identity change invalidates the whole request; unavailable exact identity cannot mint Continue.
  - [x] Route every mutation-capable FolderView, Find/Search, Compare, drag/clipboard, external-open,
    and publication-target ingress through the artifact warning before ordinary File Operations
    admission. Cancel is default, inspection is prompt-free, and each accepted receipt is revalidated
    immediately before use.
    - [x] Direct internal mutation commands have their own explicit closeout checklist; being outside
      `StartOperation` is not an exemption:
      - [x] Guard the selected exact set for non-recursive Change Attributes after its options are
        accepted and immediately before the first mutation; Cancel performs no attribute, timestamp,
        or alternate-stream change.
      - [x] Guard the selected exact set for non-recursive Change Case after its options are accepted
        and before the worker is admitted; Cancel starts no rename worker.
      - [x] Retain the non-recursive Change Case receipt on its worker and revalidate the accepted
        artifact identities immediately before each rename batch; the admission check is no longer
        the final worker-side race proof.
      - [x] For recursive Change Attributes and Change Case, discover candidates in the existing
        one-pass walk, pause before the first guarded mutation, hand the exact set to the UI-thread
        warning surface, and retain the accepted ephemeral receipt on the worker.
      - [x] Reopen/rebind each recursive guarded object no-follow immediately before its mutation.
        A missing, replacement, endpoint/profile, set, or identity mismatch cancels/fails closed; it
        never consumes the earlier grant by path text and never mutates a different object.
      - [x] Add ordinary prompt-free, Cancel-before-mutation, accepted exact-object, replacement-race,
        worker cancellation/teardown, and recursive nested-artifact tests for both commands. Archive
        warning/pause latency and retained-receipt bounds without a throughput-improvement claim.
    - Working-tree progress: the central typed `StartOperation` path now guards selected sources and
      existing publication targets for pane, Find/Compare, drag/clipboard, Batch Rename, and archive
      cleanup admissions. It binds only claim/name candidates no-follow, uses one localized
      Continue/Cancel exact-set warning with Cancel as default, stores an ephemeral receipt, and
      reloads/rebuilds/revalidates the whole set on the worker before I/O. Missing publication names
      are not treated as objects. FolderView Shell Open/Open With and configured external viewer/
      editor launches (including Find-resolved actions), Find default Shell Open, and user-menu
      programs now consume the same exact-set warning and immediately reload/rebind/revalidate before
      process launch. User-menu scope comes from the action's referenced path macros and is resolved
      before any selected-path manifest is created. Internal viewer/VFS inspection is prompt-free.
      The current-directory Shell context menu warns after verb selection and before invocation; the
      Shell Security page warns before it can change ACLs. Non-recursive Change Attributes and
      Change Case guard/revalidate their selected set before mutation/worker admission. The direct
      command worker cutover now retains that receipt: recursive discovery pauses for the shared
      UI-thread warning, Change Attributes revalidates each guarded object before its mutation, and
      Change Case revalidates the guarded members of every provider rename batch. The remaining
      direct-command test/performance matrix stays open under the checklist above, so the parent
      checkbox stays unchecked.
  - [x] Add the Proven recovery adapter over P5.3's strict parser: reload current claims, require the
    candidate to remain Proven, reopen the strict journal, rebuild claims from that reopened payload,
    and reject a replacement or name-only Possible object before exposing the P5.3 executor.
  - [x] Connect Resume/Roll back/Open location to the recovery UI and P5.3 probe/mutation callbacks.
    Resume/Roll back are offered only when the durable phase permits them; Reveal in pane is the
    read-only Open-location presentation; unknown or replacement identity exposes inspection only.
    The Cancel-default action reloads/re-proves on its MTA worker, requires the journal endpoint and
    complete source identity on every hop, requires a vacant destination, and calls only
    `RenameIfUnchanged`. There is no conflict replacement, pathname retry, or automatic post-crash
    cleanup. The active/terminal state is one File Operations informational card and teardown joins
    the cancellable worker before releasing providers.
  - [x] Add crash-before/after-claim, malformed sibling, inFlight source/destination, mismatch/
    replacement, user-created matching name, never-hidden projection, exact-set guarded-touch,
    recovery race, and no-automatic-cleanup tests. Archive registry-load/classification and
    Search/index visibility evidence without a throughput-improvement claim.
  - Focused working-tree evidence (2026-08-26): test-enabled x64 Release build passed with zero
    warnings/errors (attestation
    `dc55dfffc56bf62509bb4035f77091db0ec1b02dd2381072026b29a523332be1`; log
    `.build/logs/msbuild-20260826_033324_758-pid78696-66566aaa.log`). Commands
    `cmd_fileops_artifact_registry_claim_identity_join` passed under `2026-08-26_033728` and its
    nine-file archive passes explicit pre-stage validation. It covers malformed-sibling quarantine,
    committed and inFlight claims, exact Proven identity, replacement/name-only Possible, shared
    projection capabilities, exact-set receipt capture, successful immediate revalidation,
    replacement invalidation, unavailable-identity refusal, strict recovery reopen, and replacement/
    Possible recovery rejection. Compare `local_index_fileops_artifacts_remain_visible` passed under
    `2026-08-26_033728`; all five current/
    legacy generated-name families remained query-visible and
    `fileops.artifact.index.visibility.us` recorded 4,246 us for five indexed/five returned rows.
    Initial registry load recorded 1,873 us for three claims across three scanned records (one
    quarantined); the three recovery reopens recorded 336/165/185 us over the same bounded fixture.
    These are boundedness/visibility baselines and make no throughput-improvement claim. The first
    FolderView/Compare/Find projection slice passes the complete source-contract file 177/177 and a
    test-enabled x64 Release full-solution build with zero warnings/errors. After the central typed-
    operation and external-launch touch guards landed, the final test-enabled x64 Release full-
    solution build again passed with zero warnings/errors. After the remaining immediate Shell and
    user-menu ingress was guarded, a contamination-clearing full-solution rebuild passed with zero
    warnings/errors (attestation `9859e25595f92d82165e13e94db5fb3c116491f76d373b7850a5efd19ba81f24`;
    log `.build/logs/msbuild-20260826_045407_346-pid70268-9f5f8432.log`). The production provider-
    projection and exact-receipt helper case `cmd_fileops_artifact_registry_claim_identity_join`
    passed under `20260826T024433Z-19336-77a5ce482ba5430c965f47cb25a85da4`; Compare/Search visibility
    case `local_index_fileops_artifacts_remain_visible` passed under
    `20260826T024543Z-106432-cdfe45d68c4c45e8a0a0b38eb5ae46b8`. The external-launch production
    guard case `Phase10_ArtifactTouchGuard` passed 3/3 under
    `20260826T025912Z-48020-630ef0d18bef4aaabdc6c28cbe38b166`; its nine-file archive at
    `Specs/TestRuns/4cb089111a23/FileOps/2026-08-26_045938` passes explicit pre-stage validation.
    It proves ordinary paths are prompt-free, artifact-shaped Possible paths default to Cancel,
    Continue performs an immediate exact-set rebind/revalidation, and neither path mutates the test
    object. `fileops.artifact.touch.external.us` recorded 4,457 us for zero guarded objects, 791 us
    for a canceled one-object receipt, and 880 us for an accepted/revalidated one-object receipt.
    These are warning-latency/boundedness baselines and make no throughput-improvement claim. The
    external-action macro-scope case `file_action_external_launch_plan_macros` passed under
    `20260826T025835Z-75044-9480d43f213c4304bef56a4285108a76`, proving exact item/current-folder/
    opposite-pane/selection scope and excluding non-path macros before launch-plan side effects. The
    complete source-contract file passes 177/177 and the resource-localization contract passes 6/6.
    Popup/status projection, the complete direct-command race/performance matrix, recovery UI,
    containing commit, and Fresh Full remain pending.
    The direct-mutation admission slice then added Cancel-default exact-set guards for non-recursive
    Change Attributes and Change Case. The complete source-contract file remains green at 177/177;
    a clean test-enabled x64 Release full-solution rebuild passed with zero warnings/errors
    (attestation `4421a75e371a177392a9bc1d4d5ee494a471013d2a485cc43891373b5a651d10`;
    log `.build/logs/msbuild-20260826_050751_869-pid67120-e32f2696.log`). Focused File Operations
    `Phase10_ArtifactTouchGuard` passed 3/3 under
    `20260826T031223Z-112280-7af6403a6f844572a15c877950edf139`. At that checkpoint recursive
    discovery/pause, Change Case worker-side receipt retention/revalidation, and their mutation race
    matrix were still open; that evidence does not claim the later worker cutover.
  - Ordinary-name and ingress hardening (2026-08-26): claim leaf/path hints now admit only the exact
    no-follow projection probe, so a claim-backed final name without a generated-name marker remains
    visible and can become Proven only by identity. Configured viewer/editor/user-menu guards derive
    their touched set from the same item/current-folder/opposite-pane/selected-paths macro analysis;
    unrelated provider objects are excluded. Default destructive conflict actions are withheld unless
    the exact destination receipt is already bound. Test-enabled Release attestation
    `ea06314c17fec130c512a01eecd887672f325752c7060a5420b75cd757118fc4`
    completed with zero warnings/errors. Release macro scope
    `20260826T061712Z-115440-2df453607f5d4612a67b25c5fe034672` and registry/ordinary-name
    projection `20260826T061832Z-115440-ab90f1003a6c4ea3ba495f8162a4ba93` passed. At that checkpoint
    popup/status projection, recursive Change Attributes/Change Case authority, and recovery actions
    were open. The recursive direct-command authority was closed by the later worker cutover above;
    popup/status projection and recovery actions were subsequently closed by the P5.4 commits named
    in the document status.
  - Integration validation (2026-08-26): the frozen hardening tree completed a clean test-enabled
    Debug full-solution build with zero warnings/errors (attestation
    `addfca3e8e80432a0aeb29b00ce089497aa74bf1714a68922c344d50eaddc181`; log
    `.build/logs/msbuild-20260826_120623_908-pid20740-03b2425a.log`). Fresh Full run
    `20260826T093408Z-70916-426cdd8c29a8497fa5b80fe070df47c6` passed the artifact-mutating
    build-toolchain gate 2/2, re-attested the complete Debug profile, and promoted all completed
    standalone executables. Its in-process tooling-Pester entry then deadlocked after its displayed
    tests passed: no child process, no CPU progress, and all PowerShell threads waiting. The owned
    runner was interrupted and the profile was rebuilt cleanly; therefore this is retained as
    incomplete infrastructure evidence and is not claimed as a green Fresh Full result. The first
    Release rebuild exposed and closed an `ENABLE_TESTS`-scope error in the Change Attributes/Change
    Case artifact-guard failure presenter. The corrected production translation unit passes the x64
    Release `ClCompile` target; two full Release retries were separately blocked by MSVC PDB-server
    RPC failures in unrelated projects, so no new full-Release attestation is claimed. The affected
    source-contract and resource-localization suites remain green at 177/177 and 6/6.
  - Recursive direct-command authority evidence (2026-08-26): governed Commands run
    `20260826T105422Z-112992-6c7ab8b06ad944e697a11a75d7721a38` passed the Change Case worker
    receipt/revalidation cases. Governed Commands run
    `20260826T110540Z-112196-cd2afe53c2eb4ffd9efa51800f79d8d1` passed the recursive nested-
    artifact Change Attributes Cancel-before-mutation case after its long-path fixture was corrected.
    Governed File Operations run
    `20260826T110845Z-112944-b7f6fe028a2244aca67afb01cae3abec` passed
    `Phase10_ArtifactTouchGuard` 3/3, including clean immediate revalidation and replacement-object
    rejection. Its build completed with zero warnings/errors (attestation
    `fe03ed8b35402ddde2ebfe7576a6029cfb9411f9f07fe57d67903efbcca3ed91`; log
    `.build/logs/msbuild-20260826_130854_975-pid112944-d00c72cf.log`). These focused cases prove the
    newly checked implementation bullets; worker teardown/cancellation, accepted recursive command
    mutation for both commands, and the archived warning/retained-receipt performance matrix remain
    open, so the direct-command test checkbox and P5.4 parent remain unchecked.
  - Artifact presentation checkpoint (2026-08-26): recovery capabilities now come from the durable
    journal state, not Proven classification alone; in-flight/unfinished records expose Resume and
    Roll back, while a fully final retained record exposes Roll back only. FolderView adds the
    localized classification to its focused-item status and exposes read-only internal Inspect plus
    Reveal in pane without publishing unconnected recovery buttons. A test-enabled focused executable
    run of `cmd_fileops_artifact_registry_claim_identity_join` returned success; the first governed
    wrapper attempt exited before discovery and is not counted as green evidence. The production x64
    Release build passed with zero warnings/errors (attestation
    `939c7a08066ce3aff0d90139cf8759dcb8fb6fcfae9e5c1c61d29ba1233053b3`; log
    `.build/logs/msbuild-20260826_133155_002-pid44128-2c4e3fd2.log`). Source contracts pass 177/177,
    localization contracts pass 6/6, and `git diff --check` passes.
  - Strict recovery UI checkpoint (2026-08-26): FolderView now exposes Resume/Roll back only when the
    reopened durable phase permits the action and the strict host callback is connected. Confirmation
    is Cancel-default. The MTA worker reloads and re-proves the candidate after confirmation, checks
    the complete endpoint/profile/root and no-follow identity for every hop, requires the destination
    to be vacant, and calls only `RenameIfUnchanged`; active and terminal state is projected through
    one File Operations informational card. FolderWindow teardown requests cancellation and joins the
    worker before releasing providers. The x64 Debug build passed with zero warnings/errors
    (attestation `c134efca3739b2fd6371a837d6585a32e4c566213979b04d103ec2ec0c5b63ff`; log
    `.build/logs/msbuild-20260826_140040_222-pid68848-035c6838.log`). Exact Commands cases
    `cmd_fileops_artifact_registry_claim_identity_join` and
    `cmd_fileops_artifact_recovery_provider_callbacks` passed 2/2 in the 2026-08-26 12:08:49Z Debug
    run; the callback case proves exact probe, successful vacant-destination conditional rename,
    occupied-destination known non-commit, and replacement-source rejection before mutation. Source
    contracts pass 178/178, localization contracts pass 6/6, and `git diff --check` passes. The
    complete cross-surface accessibility audit, direct-command race/performance matrix, recovery
    crash/race archive, containing commit, and Fresh Full remain open.
  - Direct-command closeout checkpoint (2026-08-26): governed Debug Commands run
    `20260826T122927Z-32548-033c9902b26a4b47a80f356a5bc7c971` passed
    `cmd_pane_changeAttributes_recursive_artifact_guard` and `cmd_pane_changeCase` 2/2. Change
    Attributes proves Cancel before mutation, accepted exact-object revalidation/mutation, and worker
    completion teardown for a recursively discovered nested Possible artifact. Change Case proves
    recursive exact-set preparation Cancel, accepted retained-receipt revalidation and rename,
    replacement-race rejection, and cooperative worker cancellation without mutation. The compact
    archive at `Specs/TestRuns/4cb089111a23/Commands/2026-08-26_142952/` passes explicit pre-stage
    validation. Observed warning/receipt bounds were 2,318 us for Cancel, 3,815 us for accepted
    warning, and 27/1,209 us for empty/one-object immediate revalidation; recursive Change Attributes
    completed in 5,351 us canceled and 6,233 us accepted. These are same-machine boundedness
    baselines and make no throughput-improvement claim. Source contracts pass 178/178 and the full
    Debug solution attestation is `3513bbc82bd949faee1858f9248f944801f827096efb5b51699ed397ca1b25b8`
    with zero warnings/errors.
  - Recovery matrix closeout checkpoint (2026-08-26): test-enabled Debug Commands archive
    `Specs/TestRuns/4cb089111a23/Commands/2026-08-26_143441/` passes all six selected cases and the
    explicit pre-stage archive contract. The durable-transition, Resume/Roll-back replacement,
    malformed/large-selection, and crash-step/cancellation cases cover crash before/after claim,
    in-flight source/destination reconciliation, mismatch/replacement, and retained journals. The
    registry case covers malformed-sibling quarantine, user-created matching names, ordinary-name
    claims, never-hidden projection, exact-set guarded touch, and now explicitly proves that registry
    load, projection, warning, and recovery-open perform no automatic cleanup. The provider callback
    case covers vacant-destination recovery plus occupied-destination and replacement-source races as
    known fail-closed outcomes. No new crash sidecar was produced. The focused app build completed
    with zero warnings/errors (attestation
    `aa3448ba8e3ae848125b16a14b0cf03200230428bf73da7de6b39301211a26d7`; log
    `.build/logs/msbuild-20260826_143114_672-pid79384-7d90894a.log`).
- [x] **P5.5 — Move breadcrumb and recovery summary.**
  - [x] Persist one compact versioned Move breadcrumb atomically after typed admission and before
    task publication, worker creation, clipboard consumption, discovery, or provider mutation;
    reject the admission if that initial fail-if-exists write fails.
  - [x] Bound the record to strategy counts, exact qualified-root count, at most 16 representative
    root samples, one qualified destination, pane-role navigation hints, and a 128-KiB parser cap;
    store no object identity, per-item mapping, deletion receipt, or recovery authority.
  - [x] Advance once to Executing immediately before discovery/I/O, roll memory back on persistence
    failure, and retire the exact JSON only after a known terminal transition.
  - [x] Project interrupted records as Indeterminate restart notices with phase/strategy/root facts,
    pane-faithful Open source/destination, no Resume/Delete/Retry, no synthetic Issues actions, and
    Dismiss acknowledgement restricted to the regular non-reparse JSON child.
  - [x] Reject malformed/reparse/out-of-root records without mutation; preserve valid records from a
    partial directory scan; cover initial-persist failure, lifecycle rollback, host projection,
    popup actions, dismissal, bounded retention, and persistence/load timings deterministically.
  - Evidence (2026-08-26): the final Debug x64 full-solution build passed with zero warnings/errors
    (attestation `0d4a9bdd7897d060b311accd2e3365496f404460b096501583f9cba7899531e3`;
    log `.build/logs/msbuild-20260826_155118_192-pid48884-1366af2b.log`). Focused Commands run
    `20260826T135456Z-69056-ad3aac148de444819e687b030e38e1c0` passed
    `cmd_fileops_move_breadcrumb_durable_lifecycle` 1/1. The compact archive at
    `Specs/TestRuns/4cb089111a23/Commands/2026-08-26_155519/` passes explicit pre-stage archive
    validation. Observed bounded Debug timings were 2,160 us for the 17-root create, 471–649 us for
    load/partial-load, 1,797 us for successful phase advance, and 2,451 us for terminal persistence;
    the injected admission create failed in 610 us before task/clipboard/provider publication. These
    are durability/startup baselines and make no Move-throughput or recovery-success claim. Commit:
    this P5.5 work-package commit.
- [x] **P5.6 — Qualified Create Directory and S3 durable directories.**
  - [x] Qualify the concrete parent/candidate path through capability v2 before mutation; require a
    non-empty path-scoped `rootId`, validate provider-profile containment and name feasibility, and
    reject device/`GLOBALROOT`/NT namespace envelopes before any Local `CreateDirectoryW` fallback.
    The command remains outside `StartOperation`, but the fallback is not an admission exemption.
  - [x] Use the provider directory interface or an explicitly qualified Local native path; never
    reinterpret a provider path as Win32 merely because it is textually absolute.
  - [x] Commit trailing-`/` markers for flat-prefix Create Directory and empty-tree Copy/Move.
  - [x] Keep provider-native directory-bucket behavior fail-closed and non-advertised until a real
    native operation exists; never synthesize false durable success and never treat S3 Table as that
    profile.
  - [x] Add empty-`rootId`, device namespace, containment, marker merge, cleanup, refresh,
    cancel, and capability-profile tests plus admission-latency evidence.
  - Evidence (2026-08-26): the final Debug x64 full-solution build passed with zero warnings/errors
    (attestation `0dbab6c6d6183f5132354185c5eadb6e05cb30e28fdd71b71ab66a2441d22a1e`;
    log `.build/logs/msbuild-20260826_163344_805-pid82240-93c1f35c.log`).
    `PluginContractTests.exe` passed, including all required `operations.createDirectory` documents
    and the S3 debug matrix at 163/0. Focused File Operations run
    `20260826T143735Z-63656-35813db17d694a44940a5aadfaf30042` passed 3/3; focused Commands run
    `20260826T143820Z-106636-07d807415f4440fcb04e4d3381446f27` passed the live Create Directory
    prompt/create route 1/1. Compact archives
    `Specs/TestRuns/4cb089111a23/FileOps/2026-08-26_163807/` and
    `Specs/TestRuns/4cb089111a23/Commands/2026-08-26_163848/` pass explicit pre-stage validation.
    Observed bounded Debug admission timings were 1,384 us for qualified Local, 121 us for Dummy,
    and 37–390 us for rejected candidates. These are admission/fail-closed baselines and make no
    directory-creation-throughput claim. Commit: this P5.6 work-package commit.
- [x] **P5.7 — Phase 5 integration and stop-condition gate.**
  - [x] Run rename/recovery/artifact/S3 race and performance matrices, builds, specs, and stop audit.
  - [x] Replace the duplicated serial/parallel item conflict/terminal loops with one policy executor;
    scheduling may differ, but traversal-limit, cancellation, Keep Both, Recycle escalation,
    prompt clearing, and unconsumed-grant behavior must have one owner and symmetric tests.
  - [x] Prove no artifact is hidden, no pattern grants ownership, no touch bypasses warning/
    revalidation, and recovery never mutates a replacement.
  - Integration evidence (2026-08-26): one `processIndex` policy owns mutation,
    Native race fallback, Keep Both, conflict/Recycle decisions, terminal classification, prompt
    cleanup, progress, and typed result publication. Concurrency one invokes it on the operation
    worker; concurrency greater than one invokes it through `PerItemTaskScheduler`. This deleted
    694 lines of the second state machine and closed the prior serial-only skipped-reparse behavior.
    The clean Debug x64 full-solution rebuild used for the full suite passed with zero warnings/errors
    (attestation `78dd4609fd994c4203a20014af870a92d80850a81ec526bf8b93307ba9476178`; log
    `.build/logs/msbuild-20260826_192004_649-pid125316-a2d80176.log`). The focused source-contract
    run `20260826T151925Z-26072-88be128f19724e0ab8fd3b50fb6a270b` passed 1/1. The provider
    contract/debug matrix passed, including Local 129/0 and S3 163/0. A full Commands run passed
    874 cases and reported two unrelated suite-order failures; both passed immediately in isolation
    (`20260826T152004Z-116424-4c8c0b13a17e4c0899458fb1dbb3c85c` and
    `20260826T152042Z-125516-9426609052dc44478cccf5cfb519547f`). During full FileOps validation,
    two access violations were diagnosed as Debug UI-thread stack exhaustion in oversized compiled
    test dispatchers, not executor memory corruption. The dispatcher split and PDB frame evidence are
    archived under `Specs/TestRuns/4cb089111a23/Continuation/2026-08-26_p57_fileops_stack-exhaustion/`;
    the former Phase 11 crash case passed 3/3 in immutable run
    `20260826T170937Z-99416-ef7f90f74ea14583bd14928bb3cba545`. The complete immutable FileOps
    run `20260826T172428Z-24428-b09b4e4a5d2b48dbb9af0cfb5e2b7b0c` then passed 123/0 with 20
    explicit environment/delegation skips. Its compact performance archive is
    `Specs/TestRuns/4cb089111a23/FileOps/2026-08-26_194120/` and passes explicit pre-stage
    validation; it makes no final Release-throughput or live remote-provider claim. The final
    closeout Debug x64 full-solution rebuild also passed with zero warnings/errors (attestation
    `5e56490fe07315f205f575979cc92f3e7749ee672744aaadd82f1c80b1a1e849`; log
    `.build/logs/msbuild-20260826_194743_365-pid522712-fef2d046.log`), and focused source-contract
    run `20260826T175115Z-48000-bf0c8c7f79a243f8bcc47da8a81dd070` passed 1/1 against that receipt.
    Commit: this P5.7 work-package commit.
- [x] **P5.8 — Durable-directory truth and provider-local namespace hardening.**
  - [x] Remove the S3 trailing-`/` attribute shortcut: an absent flat-prefix path remains missing
    until an exact marker or current child prefix is queryable.
  - [x] Make the built-in Local provider enforce the same shared drive/UNC/volume-GUID admission
    envelope as the host before any direct `CreateDirectoryW` call; device, `GLOBALROOT`, NT,
    relative, and incomplete namespaces fail before mutation.
  - [x] Add provider-executable tests for absent/published S3 markers and direct Local namespace
    rejection. The injected empty-`rootId` capability-v2 negative test already exists and remains
    the host parser/admission guard; it was not missing as claimed by the post-Phase-5 review.
  - Evidence (2026-08-26): Debug FileSystem and FileSystemS3 builds passed with zero warnings/errors
    (attestations `c43d23b5303d8c635185073c4a1579748dc7c0cd543b58d51e83c216cfce2ad8` and
    `5bcdd531074c03171cab7ebf5504023686b234e935fd7a570f6989764cb52c6b`).
    `PluginContractTests` passed; Local debug selftests were 146/0, S3 debug selftests were 164/0,
    S3 multipart was 49/0, and S3 directory-transfer probes were 157/0. Commit: this P5.8/P6.1a
    foundation work-package commit.

### Phase 6 — cancellation, provider performance, executable matrix, and closeout

- [x] **P6.1 — Cooperative cancellation/deadline/quiet points.**
  - [x] Add one ABI-owned abort/deadline checkpoint, parse/carry every mandatory capability-v2
    cancellation field, and make malformed or missing fields fail closed.
  - [x] Plumb Local bound Read/Write/Commit and recursive provider work through the checkpoint,
    including callback-free calls and deterministic post-open abort/deadline tests.
  - [x] Plumb the canonical abort/deadline checkpoint through every options-capable built-in provider
    boundary used by enumeration/scheduling, Read, Write, Commit, publication, verification,
    metadata, delete, interlock, and consent waits. Providers whose blocking transport call cannot
    honor the complete contract keep `abort: false` and `deadline: false`; the bounded checkpoints
    do not turn a conservative false claim into true.
  - [x] Implement provider-specific quiet-point diagnostics/quarantine without unsafe thread
    termination or plugin unload.
  - [x] Add cancellation latency, teardown, disconnect, and indeterminate-result tests/metrics.
  - Evidence (2026-08-27): the full Debug solution rebuilt with zero warnings/errors (receipt
    `6d6c0f124f046c6aad9c62e911fc598d34f0723ea17e2e87c79ac2b40e5af801`).
    `PluginContractTests` passed Local 146/0, Microsoft Drive 219/0, S3 164/0, Curl 221/0,
    S3 multipart 50/0, and every S3/Curl unload quiet-point cycle. Local deterministic post-open
    abort/deadline, File Operations discovery/bandwidth cancel latency, MTP watchdog/quarantine,
    remote callback cancellation, disconnect, and indeterminate-result fixtures remain executable.
    Focused Debug File Operations passed the provider capability matrix 3/0, Phase 6 cases 7/0,
    and Local discovery-cancel cases 3/0.
    Remote capability bits remain honestly false until their blocking transports gain complete
    deadline cancellation. Commit: `3a652074`.
- [x] **P6.2 — Provider and bridge performance/memory closure.**
  - [x] Measure and fix MTP, S3, native, bridge, collision-index, proof-state, queue, buffer, and
    retained-record high-water behavior.
  - [x] Make the FBP-10 thread/buffer decision only from archived before/after evidence.
  - Evidence (implementation): MTP public writes stage to a delete-on-close file with zero private
    payload buffer; production WPD upload/device relay buffers are at most 4 MiB, compare proof at
    most 8 MiB, and the reader owns one reusable 8-MiB request state. S3 four-known-4-KiB writers
    reserve at most 16 KiB. Batch Rename retains one provider-canonical collision-key index bounded
    at 64 MiB for 65,536 x 240-UTF-16 names. Exact deterministic cases and metric names are owned by
    the File Operations, MTP, S3, SelfTest, and Performance Validation specs. Focused Debug MTP
    passed 53/0 with the one opt-in live-device case skipped; the 65,536-name Commands memory gate
    passed 1/0 in Debug and test-enabled Release. The same-machine test-enabled Release FBP-10
    archive is `Specs/TestRuns/4cb089111a23/FileOps/2026-08-27_011215/`: four 32-MiB Dummy files
    with 30-ms chunk latency measured the disabled pipeline at 130.484 seconds and the enabled
    pipeline at 0.250 seconds with an 8-MiB resolved buffer. Decision: retain the current overlapped
    pipeline, 16-pump admission bound, and 256-MiB aggregate buffer ceiling; add no second
    worker/buffer layer. Commit: `3a652074` plus this evidence/spec closeout commit.
- [x] **P6.3 — Executable capability matrix.**
  - [x] Execute every advertised operation/profile/pair claim for every production provider mode.
  - [x] Fail dishonest `true`, missing required v2 fields, null success, unsupported cleanup, and
    disconnected/read-only state transitions.
  - Evidence (2026-08-27): Debug `PluginContractTests` passed the Local, 7z, Dummy, Microsoft Drive,
    Google Drive, S3/S3 Table, Curl, MTP, and Terminal capability/profile and debug-selftest matrix;
    MTP fake-backend execution passed 53/0 with only the explicitly opt-in live-device smoke
    skipped. Test-enabled Release `PluginContractTests` passed mandatory capability-v2 shape and
    null-output validation, Microsoft Drive 219/0, Curl 221/0, S3 multipart 50/0, S3 directory-
    transfer 157/0, and every unload quiet-point cycle. Release File Operations provider
    qualification passed 3/0 (run
    `20260826T230736Z-35804-fd7d705d0b4c49eab09e8e687bd65757`). Providers whose fake backends
    are intentionally Debug-only were executed in the Debug matrix rather than reported as Release
    coverage. Commit: `3a652074` plus this closeout commit.
- [x] **P6.4 — Final correctness/performance evidence.**
  - [x] Complete §17 deterministic Debug and test-enabled Release matrices.
  - [x] Archive same-machine performance evidence and validate TestRuns inventory/quality.
  - [x] Run Fresh Full, spec inventory, tooling/archive inventories, and `git diff --check`.
  - Evidence (2026-08-27): Fresh Full run
    `20260827T062921Z-83176-b6c6742db09143c0ad9a8eb673185c62` passed 1,934 total / 1,881 passed /
    0 failed / 53 documented skips in 105m13s, with zero flaky, regression, isolation-suspect,
    unclassified-failure, or quarantine classifications. Its initial Debug build receipt was
    `4030b346e47bc3a27b031a45ffb90a377dc884f69c330496b23f298072d3efdd`; the post-writer full
    reattestation receipt was `9bf85a0ede4133c8197f2386a51b657a587af6a226369bdc402767b2eafc2493`.
    The test inventory completed successfully, TestRuns whole-inventory validation passed 1,190
    files, tool inventory passed with no findings, spec inventory reported 0 blocking findings,
    and `git diff --check` passed. The FBP-10 Release archive remains
    `Specs/TestRuns/4cb089111a23/FileOps/2026-08-27_011215/` and supports the bounded-pipeline
    decision recorded in P6.2.
- [x] **P6.5 — Normative closeout.**
  - [x] Reconcile every checklist item, finding, decision crosswalk, impact match, domain owner, and
    dependent WIP; remove no longer applicable compatibility scaffolding.
  - [x] Update status to Done, move this plan to `Specs/Plans/Done/`, reconcile the WIP index, and
    commit the closeout.
  - Evidence/commit: this containing closeout commit.

## 1. Purpose and implementation authority

This is the single WIP implementation specification for the File Operations redesign. It consolidates the accepted product decisions with the security, performance, provider, ABI, UI, and test findings needed to implement them. It intentionally removes the former question register, market comparison, repeated rationales, and parallel I11 queue.

The required sequence is:

1. Phase 0 moves the contracts in this file into their authoritative specifications and adds the implementation tickets/checklists there or in this file.
2. Phases 1–6 change code, tests, providers, UI, and evidence in the stated order.
3. Closeout moves this file to `Specs/Plans/Done/` only after the normative specifications, tests, archived performance evidence, and Fresh Full validation are complete.

No executor may reinterpret an old normative sentence to defeat a rule in this file. During Phase 0, quote each replaced normative clause and replace it explicitly. If current code or a normative spec exposes a behavior not resolved here, stop and obtain a product decision with the exact file and clause or code location.

### 1.1 Legacy decision-ID crosswalk

The former decision IDs remain only as migration anchors for specifications that still cite them. They do not revive the deleted question register or supersede the clauses in this file.

| Legacy ID | Current owner |
|---|---|
| `FO-D007` | Typed result axes and aggregation: §12 |
| `FO-D009` | Preserve/Skip link policy: §8.1 |
| `FO-D010` | Folder merge and child conflict scope: §§7.1, 10 |
| `FO-D011` | Per-item atomic publication: §§3, 6.2 |
| `FO-D013` | Bounded streaming/retention without a whole-tree manifest: §§3, 6.1–6.4, 13 |
| `FO-D019` | Path-scoped provider capabilities: §5.2 |
| `FO-D021` | Capability-v2 host/provider cutover: §5.1 |
| `FO-D022` | Bound source cleanup without mandatory destination reread: §6.3 |
| `FO-D023` | Optional content verification and digest contract: §9.1 |
| `FO-D025` | Confirmation Verify control and capability behavior: §§4.4, 9.1 |
| `FO-D027` | Recycle escalation gate: §7.4 |
| `FO-D029` | Placeholder hydration consent: §9.3 |
| `FO-D032` | Destination-name feasibility: §10 |
| `FO-D037` | Identity/staging before managed Move cleanup: Phases 1–2 |
| `FO-D038` | Verification progress, graph, and ETA: §12.5 |
| `FO-D039` | Central Rename ingress: §7.3 |

At the current baseline, a repository scan outside this file finds only these 16 legacy IDs in active `Specs/` files. IDs with no remaining citation are retired history and intentionally are not expanded back into a 61-row decision ledger. Phase 0 reruns the scan; any newly found citation must be mapped here before its owning REVIEW banner or normative clause is rewritten.

## 2. Audit verdict

### 2.1 What is good and must survive

- Identity is not path text. Source, destination, endpoint, object, revision, content, and complete-item identity are distinct.
- Move has three honest strategies: native Move, managed copy-then-delete, and Copy-only with source retained.
- No hash, size, timestamp, name, or path comparison may silently authorize skipping user data or deleting a source.
- Links are objects. There is no follow-target transfer mode.
- EFS decryption, sparse inflation, placeholder hydration, lossy Move metadata, low-space continuation, and Recycle escalation are explicit task-scoped decisions.
- Atomicity is per file/object. Cancel or crash may leave both complete copies; the engine reports that state instead of pretending the whole tree rolled back.
- Verification is optional and off by default. It verifies content when selected but is not the destructive-delete authority.
- Existing file-manager behavior—folder merge, Keep Both, queue/parallel execution, speed limits, filters, and direct same-pane workflows—survives unless explicitly changed below.

### 2.2 What was less good and is removed

- The former file called itself complete while retaining stale open questions and contradictory clauses.
- Product decisions and the I11 remediation queue duplicated ownership.
- “Phase 2” mixed identity, links, prompts, metadata, artifacts, timestamps, space, and overlap into an untestable package.
- A whole-tree preflight/manifest conflicted with streaming discovery and bounded memory.
- Outcome prose multiplied display sentences without a typed result model.
- Confirmation controls, ingress payloads, provider contracts, HRESULT aggregation, clipboard consequences, and test gates were not specified together.
- The former `followTargets` migration rationale incorrectly said the setting had no Preferences surface. Silent migration remains approved because the product is not complete, not because the setting was invisible.

### 2.3 Current code gaps this program closes

| Area | Current evidence | Required state |
|---|---|---|
| Same-object safety | `Plugins/FileSystem/FileSystem.FileOps.cpp`: `TryAreSameFile` is effectively reached only for case-only rename; alias routes can reach overwrite/publication | Identity/containment check at every ingress and immediately before destructive or truncating mutation |
| Weak equivalence | Native resume and bridge Move cleanup use FNV-1a as skip/delete evidence | Digest may verify content; it never silently skips and never authorizes deletion |
| Source deletion | `FolderWindow.FileOperations.State.cpp`: bridge cleanup is pathname/basic-info based | Reader-bound identity/revision and handle-bound or conditional delete; otherwise Copy-only |
| Destination containment | Bridge may merge through destination reparse directories; traversal classifications can become stale | No-follow ancestor binding and revalidation; destination links are explicit conflicts |
| Staging | Native/bridge staging cleanup and promotion can be path-owned | Create-exclusive stage plus retained identity/token; cleanup/promotion only by that authority |
| Links | P3.1 removed Follow from the ABI/UI/runtime, canonicalized legacy settings, and made immutable per-call Preserve/Skip authoritative; tree Move semantic retarget remains deferred | P3.2 implements semantic-preserving retarget for tree Move without following or mutating the target object |
| Folder conflicts | Old decision wording could make an existing folder an overwrite conflict | Folder-on-folder merges; conflicts are evaluated on colliding children |
| Rename ingress | In-pane F2 calls `RenameItem` directly; Batch Rename has a separate executor | Both submit typed rename plans through the central operation engine |
| Confirmation | `RedSalamander/FileOperationConfirmation.cpp` builds an OK/Cancel `HostPromptRequest`; there is no Links selector or Verify control | Host-internal confirmation model captures Links and Verify into the immutable task snapshot |
| Recycle | `RedSalamander/FolderWindow.FileOperations.Popup.cpp` adds `PermanentDelete` for `RecycleBinFailed` under the current conflict action table | Fresh one-shot Recycle escalation gate with default Cancel and identity recheck |
| Results | Host primarily has HRESULT plus Completed/Partial/Canceled/Failed | Typed per-item axes mapped to HRESULT, popup, Issues, pane refresh, telemetry, and clipboard |
| Clipboard | `RedSalamander/FolderView.FileOps.cpp` calls `InvalidateMoveClipboardAfterVerifiedCompletion` from the completion callback | Move clipboard clears when the complete accepted request is queued; it is never restored |
| Discovery/progress | P2.5 removed the runtime `preCalcEnabled`/`preCalcMaxWorkers`/`Calculating` bridge. Local and host-bridge traversal now report bounded cumulative discovery into execution; Skip switches one-way to just-in-time | Keep every progress/ETA surface indeterminate until traversal closure and preserve the 32/128/256 reservation, non-starvation, and ≤50 ms Skip-release contracts |
| Recursive memory | Native queue, created-directory list, and bridge Move manifest can grow with the tree | Streaming bounded work; O(depth) ancestors; no whole-tree Move manifest |
| Cancellation | Host checks between provider calls, but Read/Write/Commit/enumeration may block | Abort/deadline contract and known/indeterminate quiet point without unsafe unload |
| Provider truth | Capability v1 is instance-wide and not executable; success/null outputs can crash host | Path-scoped capability v2 plus executable matrix and success-output validation |
| MTP/S3 performance | MTP may retain at least two payload copies; S3 reserves multipart-scale memory for tiny objects | Streaming bounded buffers and size-aware reservation |
| Artifacts | `.rs_tmp_`, `.rs-bak-`, and rename intermediates look like ordinary user files | Proven/Possible/Ordinary provenance with distinct badge and recovery actions |
| S3 directory creation | `FileSystemS3::CreateDirectory` can create only an in-session synthetic prefix | Zero-byte trailing-`/` marker for flat S3 namespaces |

This table is an implementation audit, not a promise that line numbers remain stable. Before each phase, use `rg` and `git diff` to refresh the exact call sites.

## 3. Non-negotiable invariants

1. **Exact authority.** Delete, replace, restore attributes, promote, roll back, and clean up only the exact object or immutable revision authorized by the task.
2. **No-follow safety.** Classification and mutation of a link operate on the link object. Ancestor containment checks never follow a link silently.
3. **Fail-closed Move.** If exact source cleanup is unavailable or publication is indeterminate, keep the source. Do not fall back to pathname deletion.
4. **No silent data choice.** Hashes and metadata heuristics may inform verification or diagnostics; they do not choose Skip, Overwrite, or source deletion.
5. **Per-item atomic publication.** A partially written final object is never reported as published. Directory-wide transactions are not promised.
6. **Bounded streaming.** Memory and retained records are bounded independently of total tree size. Discovery may continue during execution.
7. **Immutable task intent.** Every worker uses the accepted plan snapshot; it never rereads live settings or pane/provider configuration.
8. **Truthful UI and telemetry.** The result describes destination publication, verification, and source disposition separately.
9. **Task-scoped consent.** A grant applies only to the described risk in the current task and never becomes a session or Preference flag.
10. **No target mutation.** Preserving or retargeting a link may rewrite the link object being copied/moved; it never reads, writes, deletes, or changes the target object as a consequence of link traversal.
11. **No silent elevation.** File Operations never relaunches itself or the app with UAC without a separately specified, explicit user action.
12. **No unsafe teardown.** A plugin/module stays alive until its in-flight request reaches a known or indeterminate quiet point.

## 4. Terms and typed model

### 4.1 Identity

| Term | Meaning |
|---|---|
| Qualified endpoint | Provider instance/profile plus namespace/root identity; not a display path |
| Path identity | The name/address requested in a qualified endpoint |
| Object identity | Stable no-follow local file ID or provider object ID |
| Revision identity | Immutable generation/version/ETag only where the provider contract makes it immutable |
| Content identity | Cryptographic digest or exact byte proof; not object identity |
| Complete item | A fully committed destination object, including required publication semantics |
| Bound object | An open handle or provider token retaining no-follow identity and authorized operations |

Local object identity uses volume identity plus `FILE_ID_INFO` from a no-follow handle. A 64-bit `nFileIndex`, filename, basic metadata, or FNV-1a is not sufficient destructive authority. Remote identity uses the provider’s documented object and revision semantics. ETag is accepted only where the provider declares it immutable for that operation; S3 destructive cleanup requires `VersionId`/generation or an equivalent conditional-delete token.

### 4.2 Operation plan

The host creates immutable child plans before task admission. The usual case has one child. A
mixed-root selection is stably partitioned into one child per qualified source endpoint and stored
in one immutable `FileOperationPlanGroup`; the group has one confirmation, task ID, cancellation
surface, result presentation, and gated-worker clipboard barrier.

```cpp
enum class TransferIntent { Copy, Move };
enum class LinkPolicy { Preserve, Skip };
enum class ExecutionMode { Queue, Parallel };
enum class DeleteMode { Recycle, Permanent };
enum class DeleteOrigin { PaneCommand, FindResults, CompareResults, PackCleanup, UnpackCleanup, RecycleEscalation };
enum class RenameOrigin : uint8_t
{
    Unspecified,
    InlineRename,
    BatchRename,
};
enum class ConsentKind { PermanentDelete, ArchiveDeleteAfter, RecycleEscalation };

struct QualifiedEndpoint
{
    std::wstring pluginId;
    std::wstring instanceId;
    std::wstring profileId;
    std::wstring rootId;
};

struct ProviderIdentitySnapshot
{
    std::vector<std::byte> objectId;
    std::vector<std::byte> revisionId;
    std::wstring pathProfileId;
};

struct QualifiedSourceItem
{
    std::wstring providerPath;
    std::optional<ProviderIdentitySnapshot> ingressSnapshot;
};

struct QualifiedDestination
{
    std::wstring providerFolderPath;
};

struct TransferDestinationMapping
{
    size_t sourceIndex;
    std::wstring destinationProviderPath;
};

struct ClipboardSequence
{
    uint32_t windowsSequenceNumber;
};

struct RenameStep
{
    QualifiedSourceItem source;
    std::wstring finalLeafName;
};

struct RenameCycle
{
    std::vector<size_t> stepIndices;
};

struct DestructiveConsentReceipt
{
    ConsentKind kind;
    uint64_t taskNonce;
    std::optional<size_t> itemIndex; // Required for one-item Recycle escalation.
};

struct OperationOptions
{
    LinkPolicy linkPolicy;
    bool verifyAfterCopy;
    ExecutionMode executionMode;
    std::optional<uint64_t> bandwidthLimitBytesPerSecond;
};

struct TransferPlan
{
    TransferIntent intent;
    QualifiedEndpoint sourceEndpoint;
    QualifiedEndpoint destinationEndpoint;
    std::vector<QualifiedSourceItem> selectedItems;
    QualifiedDestination destination;
    std::vector<TransferDestinationMapping> explicitMappings;
    OperationOptions options;
    std::optional<ClipboardSequence> moveClipboardSequence;
};

struct RenamePlan
{
    RenameOrigin origin;
    QualifiedEndpoint endpoint;
    std::vector<RenameStep> finalMappings;
    std::vector<RenameCycle> cycles;
    OperationOptions options;
};

struct DeletePlan
{
    QualifiedEndpoint endpoint;
    std::vector<QualifiedSourceItem> selectedItems;
    DeleteMode mode;
    DeleteOrigin origin;
    bool recursive;
    OperationOptions options;
    std::optional<DestructiveConsentReceipt> initialConsent;
};

using FileOperationPlan = std::variant<TransferPlan, RenamePlan, DeletePlan>;
using FileOperationPlanGroup = std::vector<FileOperationPlan>;
```

Source spelling may follow repository style, but these semantic fields and their separation are normative. Copy/Move intent exists only in `TransferPlan`; Recycle is `DeletePlan::mode == Recycle`, not a separate operation variant. `RenamePlan::origin` is mandatory because a one-step Batch Rename must not inherit the inline-F2 admission/presentation exception merely because its mapping count is one. `ingressSnapshot` is a captured comparison value, never mutation authority: execution obtains a bound object and must match/revalidate any authoritative ingress snapshot before mutation. An absent snapshot does not waive execution-time binding. `DeletePlan::mode == Permanent` requires a matching initial consent receipt except when a fresh per-item Recycle-escalation receipt is created by §7.4. Archive checkboxes produce an archive-delete receipt only after the archive operation reaches its own success/verification contract; the receipt cannot authorize any path not present in `selectedItems`.

Every child plan has exactly one `sourceEndpoint`. A request spanning roots is not rejected and is
not weakened into one synthetic endpoint: the host stably partitions selected items into child
plans keyed by the complete qualified endpoint, rebases any explicit mappings to child-local
indexes, and admits one immutable group. All children share the captured intent/options and, for a
transfer, the same qualified destination. One malformed child rejects the entire group before
publication. The group owns one confirmation, task ID, queue/card, cancellation state, and
gated-worker clipboard barrier; execution results remain per item and may become partial across
children.

`TransferPlan::explicitMappings` is empty for ordinary F5/F6, paste, and drag/drop: each selected top-level item is placed under `destination.providerFolderPath` using the task's collision rules. Compare synchronization supplies exactly one mapping for every selected source index; each mapping names that item's complete destination provider path. Mapping indices must be unique, in range, and collectively cover the selected set. Every mapped path must belong to `destinationEndpoint` and be at or below `destination.providerFolderPath`; the host rejects an incomplete, duplicate, cross-endpoint, or escaping mapping before queue admission. A mapped directory establishes the destination root for its streamed descendants; it does not require a pre-enumerated child manifest.

#### 4.2.1 Ingress ownership and confirmation

Every mutation ingress produces one of the plans above before provider execution. A current direct call listed below is migration evidence, not an exemption.

| User/system ingress | Current production entry | Required plan and confirmation |
|---|---|---|
| F5/F6 Copy/Move to other pane | `FolderWindow.FileOperations.cpp`: `CommandCopyToOtherPane` / `CommandMoveToOtherPane` | `TransferPlan`; normal Copy/Move confirmation and immutable Links/Verify/options snapshot |
| Pane/Finder/Compare permanent Delete or Recycle | `FolderWindow.FileOperations.cpp`; `FindFilesWindow.cpp`; Compare delegates to pane commands | `DeletePlan`; Recycle or Permanent mode and the §7.4 confirmation/escalation rules |
| Clipboard paste, including same-pane duplicate | `FolderView.FileOps.cpp`: `PasteItemsFromClipboard` | `TransferPlan`; captured clipboard sequence for Move; §7.2 same-folder rules; clear cut list at accepted queue |
| Internal drag/drop | `FolderView.DragDrop.cpp`: `PerformDrop` with source context | `TransferPlan`; preserve qualified source endpoint and central options snapshot |
| External OLE/`CF_HDROP` drop | `FolderView.DragDrop.cpp`: `PerformDrop` without internal context | Qualify the shell/local source endpoint before `TransferPlan`; report Copy to the external source until asynchronous Move is proven complete |
| Find Files Copy/Move/Delete | `FindFilesWindow.cpp`: `CopyOrMoveSelectedResultsToOtherPane` / `DeleteSelectedResults` | `TransferPlan` or `DeletePlan`; retain resolved provider instance and per-result correlation |
| Compare Directories synchronization | `CompareDirectoriesWindow.cpp`: `StartCompareSyncToOtherPane` | `TransferPlan` with explicit per-item destination mappings; the existing Compare Move summary remains the ingress confirmation copy |
| Pack “delete sources after success” | `FolderWindow.FileSystem.Commands.cpp`: archive pack completion submits `DeletePlan(Permanent, PackCleanup)` through `StartFileOperationFromFolderView` | Preserve the central route and its captured archive-delete receipt |
| Unpack “delete archives after success” | `FolderWindow.FileSystem.Commands.cpp`: archive unpack completion submits `DeletePlan(Permanent, UnpackCleanup)` through `StartFileOperationFromFolderView` | Preserve the central route and its captured archive-delete receipt; no private pathname deletion |
| F2 inline Rename | `FolderView.FileOps.cpp`: `RenameFocusedItem` submits `RenamePlan(InlineRename)` through the File Operations callback | The inline rename editor/prompt is the only initial confirmation—do **not** show a second File Operations confirmation; §7.3 owns immediate admission and delayed card reveal; later conflicts use the central conflict UI |
| Batch Rename | `BatchRenameExecutionEngine.cpp` / `BatchRenameWindow.cpp` currently execute a separate rename engine | `RenamePlan(BatchRename)`, including when only one mapping remains; the Batch Rename preview/Run action is the initial confirmation; execution, cycles, results, cancel, recovery, queueing, and cards move to the central engine |
| Future plugin/command ingress | Any host callback or command that can mutate an item | Must call the same plan factory; direct provider mutation is forbidden unless this specification names the narrow exemption |

The plan factory is the sole location that applies qualified endpoint capture, same-folder/subtree admission, clipboard start barrier, confirmation receipts, and immutable settings. Provider calls are execution details below that boundary.

### 4.3 Pre-queue validation versus streaming discovery

Before task admission, validate only the immutable request envelope:

- required fields and selected-item list are non-null/non-empty;
- source and destination fields are qualified and structurally valid;
- local paths use supported ordinary drive/UNC or equivalent extended syntax; reject `\\?\GLOBALROOT`, `\\.\`, `\??\`, and unsupported device namespaces;
- the requested command is represented by the correct typed plan;
- a clipboard Move request is a complete captured item list, not a live clipboard reference.

Do **not** enumerate the full selection or preflight all top-level filesystem state before the first mutation. Object kind, existence, identity, containment, capabilities, metadata, space, and recursive children are discovered and revalidated as execution reaches them. A later discovery failure preserves prior successful items and produces a partial result. This rule supersedes I11 `FBC-18`’s proposed all-input filesystem preflight while preserving its valid envelope-consistency requirement.

The old separate pre-calculation model is retired for Copy, Move, Delete, and Recycle:

- remove `fileOperations.preCalcEnabled`, `fileOperations.preCalcMaxWorkers`, and the separate recursive `Calculating` pass/state from their normative contracts;
- use one traversal whose discoveries feed the same bounded execution queue;
- never traverse an item solely to calculate totals and then traverse it again to perform the operation;
- totals and ETA are estimates while discovery remains open and become authoritative only when traversal closes.

Discovery-ahead is enabled by default to keep the transfer queue warm. It is not a safety preflight and it does not delay the first safe mutation. The popup exposes a one-way **Skip discovery** action while discovery-ahead is active. Choosing it:

- stops admitting discovery-ahead work at its next safe checkpoint and discards no work already discovered;
- immediately releases the reserved discovery request capacity and discovery-driven execution throttling;
- switches the task to just-in-time discovery performed by the execution path and permits the full configured execution concurrency and, for byte-copying work, bandwidth;
- does not skip recursive enumeration, name validation, identity/containment checks, conflict classification, capability queries, or any other check required before an item is acted on;
- does not mean Skip files, Cancel, or ignore errors, and cannot be reversed within that task;
- leaves whole-task totals/ETA indeterminate or estimated until traversal naturally closes; current-item byte progress remains exact where the provider reports it.

### 4.4 Confirmation snapshot

Copy/Move confirmation is initially a host-owned model. This program may break and replace `HostPromptRequest` if a shared typed prompt ABI materially simplifies confirmation or provider-driven consent; there is no compatibility requirement because the host and every shipped plugin are updated atomically in this tree.

The confirmation contains:

- operation, source count/summary, source endpoint, destination endpoint/path;
- **Links:** `Preserve links` or `Skip links`, initialized from the source provider setting;
- **Verify copied file contents**, initialized from `verifyAfterCopy` (default off) and enabled according to §9.1;
- queue/parallel and current bandwidth controls where already exposed;
- known up-front EFS, sparse-inflation, and placeholder-hydration facts, without forcing recursive enumeration.

For a clipboard Move, confirmation also states: **“Starting this Move consumes the cut list. If an item cannot be removed, its source will remain and the cut list will not be restored.”**

Accept returns a complete immutable options snapshot. Cancel queues nothing and changes no clipboard. Workers may open deferred gates for facts discovered during streaming, but may not reread Preferences or plugin JSON.

The old `followTargets` value migrates silently to `Skip`; `copyReparse` migrates to `Preserve`; old `skip` migrates to `Skip`. Preserve is the default for new or absent settings. The old Follow numeric value remains reserved and is rejected at ABI/config boundaries so it cannot be reinterpreted accidentally.

## 5. Provider ABI and capability architecture

### 5.1 ABI evolution

Change the file-system plugin ABI in one lockstep host/provider/test change. Compatibility with the unreleased ABI is not a goal:

- add public `FileSystemLinkPolicy::{Preserve, Skip}` and per-call operation options;
- replace instance-wide capability v1 with path-scoped capability v2;
- add optional object-binding and conditional mutation interfaces;
- remove v1 after updating every in-tree caller and provider; do not retain a compatibility discovery path;
- require every production and Dummy provider to return a valid v2 profile, with unsupported abilities reported explicitly as false/unsupported;
- keep binding/conditional-mutation interfaces optional and queried explicitly; a destructive strategy may rely only on interfaces and capabilities it actually obtained;
- increment the ABI version and update every production and Dummy provider, QI table, contract test, comments, and normative provider matrix together.

This compatibility break includes public and host-internal operation payloads. Phase 0 may replace `HostPromptRequest`, `FileSystemOptions`, operation callbacks, task request/result structs, capability JSON, and the §4.2 plan layout when doing so makes the typed contract executable. Do not add v1 adapters, dual parsers, legacy struct-size branches, or old/new prompt paths. Any durable development artifact or recovery record that survives a process restart must carry an explicit format version and fail closed or be quarantined when incompatible; it is not a reason to preserve the old ABI. The absence of released consumers removes backward-compatibility work, but every landed work package must still build and test with the host and all in-tree providers on one coherent ABI.

Candidate interface IDs reserved by this plan:

- `IFileSystemPathCapabilities2`: `9be5ee85-f247-4a95-953b-5f18ed76719d`
- `IFileSystemObjectBinding`: `d8ae290a-b84c-42ec-982c-7c01dedc7603`
- `IFileSystemBoundObject`: `5ed3921d-a486-42cc-a5c3-33d97f75c5b5`

Phase 0 may adjust spelling and layout for ABI conventions, but not remove or weaken this method set:

```cpp
enum FileSystemBindFlags : uint32_t
{
    FILESYSTEM_BIND_NO_FOLLOW       = 0x01,
    FILESYSTEM_BIND_READ_CONTENT    = 0x02,
    FILESYSTEM_BIND_READ_METADATA   = 0x04,
    FILESYSTEM_BIND_DELETE          = 0x08,
    FILESYSTEM_BIND_RENAME          = 0x10,
    FILESYSTEM_BIND_PUBLICATION     = 0x20,
};

enum FileSystemBoundObjectKind : uint32_t
{
    FILESYSTEM_BOUND_REGULAR_FILE,
    FILESYSTEM_BOUND_DIRECTORY,
    FILESYSTEM_BOUND_LINK,
    FILESYSTEM_BOUND_OTHER,
};

struct FileSystemBoundObjectSnapshot
{
    uint32_t sizeBytes;
    uint32_t kind;                 // FileSystemBoundObjectKind
    const void* objectId;          // Immutable for this bound-object lifetime.
    uint32_t objectIdBytes;
    const void* revisionId;        // Empty only when the profile declares no revision token.
    uint32_t revisionIdBytes;
    uint64_t committedSizeBytes;   // UINT64_MAX when not applicable/unknown.
};

struct FileSystemConditionalMutationResult
{
    uint32_t sizeBytes;
    BOOL mutationCommitted;
    BOOL originalStillPresent;
    BOOL outcomeKnown;
};

enum FileSystemLinkKind : uint32_t
{
    FILESYSTEM_LINK_KIND_SYMBOLIC_FILE = 1,
    FILESYSTEM_LINK_KIND_SYMBOLIC_DIRECTORY = 2,
    FILESYSTEM_LINK_KIND_JUNCTION = 3,
};

enum FileSystemLinkTargetMapping : uint32_t
{
    FILESYSTEM_LINK_TARGET_OUTSIDE_SOURCE_ROOT = 1,
    FILESYSTEM_LINK_TARGET_MAPPED_INSIDE_SOURCE_ROOT = 2,
    FILESYSTEM_LINK_TARGET_INSIDE_SOURCE_ROOT_UNMAPPABLE = 3,
};

struct FileSystemLinkInformation
{
    uint32_t sizeBytes;
    FileSystemLinkKind kind;
    BOOL targetIsRelative;
    FileSystemLinkTargetMapping targetMapping;
    uint32_t targetLengthUtf16;
    uint32_t targetCapacityUtf16;
    wchar_t* targetBuffer;
    uint32_t sourceRelativeTargetLengthUtf16;
    uint32_t sourceRelativeTargetCapacityUtf16;
    wchar_t* sourceRelativeTargetBuffer;
};

struct FileSystemLinkComponentMapping
{
    const wchar_t* sourceRelativeComponentPath;
    const wchar_t* destinationComponentName;
};

struct FileSystemLinkTransform
{
    uint32_t sizeBytes;
    const wchar_t* sourceLinkPath;
    const wchar_t* destinationLinkPath;
    const wchar_t* sourceRootPath;
    const wchar_t* destinationRootPath;
    const FileSystemLinkComponentMapping* componentMappings;
    uint32_t componentMappingCount;
};

interface IFileSystemBoundObject;

interface IFileSystemPathCapabilities2 : IUnknown
{
    virtual HRESULT GetPathCapabilities(const wchar_t* path,
                                        FileSystemOperation operation,
                                        const char** jsonUtf8) noexcept = 0;
};

interface IFileSystemObjectBinding : IUnknown
{
    virtual HRESULT BindObject(const wchar_t* path,
                               FileSystemBindFlags flags,
                               IFileSystemBoundObject** bound) noexcept = 0;

    virtual HRESULT CreateExclusiveWriter(const wchar_t* stagePath,
                                          const FileSystemOptions* options,
                                          IFileWriter** writer,
                                          IFileSystemBoundObject** ownedStage) noexcept = 0;

    virtual HRESULT CreateExclusiveDirectory(const wchar_t* stagePath,
                                             const FileSystemOptions* options,
                                             IFileSystemBoundObject** ownedStage) noexcept = 0;

    virtual HRESULT ReadBoundLink(IFileSystemBoundObject* boundLink,
                                  const FileSystemLinkTransform* transform,
                                  const FileSystemOptions* options,
                                  FileSystemLinkInformation* information) noexcept = 0;

    virtual HRESULT CreateExclusiveLink(const wchar_t* stagePath,
                                        const FileSystemLinkInformation* information,
                                        const FileSystemOptions* options,
                                        IFileSystemBoundObject** ownedStage) noexcept = 0;
};

interface IFileSystemBoundObject : IUnknown
{
    virtual HRESULT GetSnapshot(FileSystemBoundObjectSnapshot* snapshot) noexcept = 0;
    virtual HRESULT IsSameObject(IFileSystemBoundObject* other, BOOL* same) noexcept = 0;
    virtual HRESULT OpenReader(IFileReader** reader) noexcept = 0;
    virtual HRESULT GetBasicInformation(FileSystemBasicInformation* info) noexcept = 0;

    virtual HRESULT PublishAs(const wchar_t* finalPath,
                              IFileSystemBoundObject* expectedDestination,
                              FileSystemFlags flags,
                              FileSystemConditionalMutationResult* result,
                              IFileSystemBoundObject** published) noexcept = 0;

    virtual HRESULT RenameIfUnchanged(const wchar_t* destinationPath,
                                      IFileSystemBoundObject* expectedDestination,
                                      FileSystemFlags flags,
                                      FileSystemConditionalMutationResult* result,
                                      IFileSystemBoundObject** renamed) noexcept = 0;

    virtual HRESULT DeleteIfUnchanged(FileSystemFlags flags,
                                      FileSystemConditionalMutationResult* result) noexcept = 0;

    virtual HRESULT AbortOwnedObject(FileSystemConditionalMutationResult* result) noexcept = 0;
};
```

`rootId` is a provider-defined opaque identity for the namespace root containing the queried path,
interpreted only together with `pluginId`, `instanceId`, and `profileId`. It is mandatory and
path-scoped in capability v2. A provider with one logical root per configured instance may return a
stable provider-root token; a multi-root instance must distinguish its roots (for example local
volume, S3 bucket/directory bucket, MTP device storage, or cloud drive). A display path, pane label,
or host-manufactured hash is not root identity. Missing or malformed `rootId` rejects admission.
Exact object authority still comes from the Phase 1 no-follow binding/revalidation contract; the
capability assertion does not replace object binding.

The host turns pane/provider instance context into a non-empty comparison identifier:
`host/default` for an absent context and `host/context/<opaque-context>` for a present one. The
namespace prefix prevents an absent context from colliding with a literal provider context named
`default`. This value is never display copy or mutation authority.

ABI rules:

- `IFileSystemPathCapabilities2` is mandatory for every shipped provider after the breaking cutover. Missing QI, invalid JSON, a profile mismatch, or an advertised operation that fails its executable contract disables the affected strategy.
- `operations.move` is required provider-entry-point admission; `operations.nativeMove` is a separate required strategy claim. The former never proves the latter. Native strategy routing requires the exact qualified source/destination profile pair and every provider-specific pair predicate.
- Object-binding interfaces are optional. Their absence disables managed destructive Move and identity-owned recovery for that profile, but does not disable honest Copy or a separately advertised native Move.
- `BindObject` is no-follow; a provider that cannot guarantee that returns `ERROR_NOT_SUPPORTED`. The returned COM object owns the handle/token and immutable revision guard until release.
- Snapshot pointers are owned by the bound object and remain valid/unchanged until its release. `objectId` must be non-empty for a successful binding; revision may be empty only when the v2 profile says revision binding is unavailable.
- `CreateExclusiveWriter` fails on any existing stage name and returns both the writer and its ownership token atomically. Success with either output null is a provider-contract failure.
- `ReadBoundLink` is a two-call bounded read from the exact no-follow link authority. It resolves
  relative text lexically, applies the selected-root plus bounded sparse component transform without
  opening the target, and preserves outside-root stored text. `CreateExclusiveLink` fails on any
  existing stage name and returns the exact owned link-stage token. Success with a null token is a
  provider-contract failure.
- `expectedDestination == nullptr` means publish/rename only if the final name is absent. Replacing an existing destination requires the exact bound destination object. A mismatch returns `ERROR_REVISION_MISMATCH`/typed conflict without mutation.
- Conditional methods always initialize `outcomeKnown`, including transport failure. Unknown outcome forbids retry, cleanup, or source deletion until reconciliation establishes identity.
- `AbortOwnedObject` is valid only for an object created exclusively by that binding. It is never a general path delete.
- Read, Write, Commit, enumeration, and conditional mutation receive the task deadline/Abort contract through the final `FileSystemOptions`/operation-control shape. Phase 0 must add that shape and lifetime rules beside these methods rather than leaving cancellation as prose.
- Capability v2 declares path-scoped name, case, directory, link, native Move, object-store, metadata preservation/loss, digest, deadline, and quiet-point behavior. Every `true`/supported claim has a provider contract test that calls it.

All successful provider factories and query methods validate required output pointers. Success with a null reader, writer, directory result, bound object, or capability result becomes a diagnosed provider-contract failure (`E_POINTER` or `ERROR_INVALID_DATA`), never a host dereference.

### 5.2 Capability v2 sketch

```json
{
  "version": 2,
  "pathProfile": "stable-provider-defined-id",
  "rootId": "opaque-path-scoped-root-id",
  "operations": {
    "copy": true,
    "move": true,
    "nativeMove": true,
    "rename": true,
    "delete": true,
    "recycle": false
  },
  "identity": {
    "object": "stable",
    "revision": "conditional",
    "boundDelete": true,
    "conditionalDelete": false
  },
  "publication": {
    "exclusiveStage": true,
    "conditionalPublish": true,
    "committedSize": true
  },
  "links": {
    "preserveFileLink": true,
    "preserveDirectoryLink": true,
    "retargetInTree": true,
    "exactLinkRemoval": true
  },
  "metadata": {
    "motw": "preserved",
    "alternateStreams": "reported-loss",
    "extendedAttributes": "reported-loss",
    "sparse": "preserved",
    "efs": "preserved"
  },
  "cancellation": {
    "abort": true,
    "deadline": true,
    "quietPointTimeoutMs": 30000
  },
  "names": {
    "comparison": "ordinalIgnoreCase",
    "caseOnlyRename": true,
    "maxComponentUtf16": 255
  },
  "directories": {
    "model": "native"
  }
}
```

Unknown fields are ignored only where the version contract says they are optional. Missing/non-boolean `operations.move` or `operations.nativeMove`, `S_OK` plus `{}`, and unknown required semantics fail closed. Capabilities are resolved at the bound path/profile, not assumed instance-wide.

### 5.3 Strategy routing

| Strategy | Admission | Completion |
|---|---|---|
| Native Move | One provider/path profile truthfully supports an OS/provider atomic or native move for this source/destination and the call preserves the required semantics | Provider reports final destination identity and source absence/retention truthfully |
| Managed Move | Exact reader-bound source identity/revision, protected copy, identity-owned publication, destination proof, and handle-bound or conditional source delete all exist | Destination published, then exact source removed; any failed safety gate keeps source |
| Copy-only | Copy is supported but any destructive Move prerequisite is missing | Publish destination and report `Source retained` with the reason |

The local provider keeps Native Move only for a path-scoped same-volume regular-file pair that can
complete through one in-place provider rename. Local declares `nativeMoveSemanticTransform: false`,
so every directory tree and link object uses Managed semantic publication even when the destination
name is absent. Existing regular-directory destinations remain Managed folder merges so only
colliding children enter the central typed conflict surface; this qualification must not recursively
enumerate the tree. Local volume-specific `rootId` still prevents cross-volume Native admission. A
native regular-file rename can succeed for an open file when managed copy/delete cannot; the engine
must not convert a native sharing failure into pathname deletion. Across providers or volumes, use
managed Move only when every destructive precondition is real. Otherwise, route to Copy-only before
source cleanup.

The provider matrix must be executable. Expected baseline:

| Provider/mode | Copy | Move strategy |
|---|---|---|
| Local Win32 | Native/managed | Native same-volume; managed cross-volume only with bound source and identity-owned publication |
| 7z | Export | Copy-only |
| Curl FTP/SFTP/SCP | Bridge export Copy; provider-native methods remain non-advertised while conditional publication/rollback is unsafe | `operations.copy/move/rename` remain false despite existing direct method bodies; enable a native strategy only after v2 advertises and contract-tests its exact safe subset; otherwise Copy-only |
| S3 flat bucket | Stream/server-side copy | Conditional generation/VersionId Move when real; otherwise Copy-only |
| S3 directory bucket | Provider-native directory semantics | Path-profile-specific |
| Microsoft Drive | Stream/server-side copy | Native Graph Move for the same drive/profile using item ID plus conditional ETag/reconciliation; cross-profile transfer remains managed or Copy-only by advertised proof |
| MTP writable | Stream copy | Native only for same-device leaf-preserving WPD Move when v2 advertises/tests it; transfer-copy fallback is Copy-only until bound source cleanup exists |
| MTP read-only/disconnected | Export only or disabled | Copy-only/unsupported |
| Google Drive current read-only mode | As currently advertised | Copy-only/unsupported |
| Dummy | Test-declared | Exercises every strategy and failure |

## 6. Transfer execution contract

### 6.1 Streaming walk and containment

Execution uses a bounded work queue and an O(depth) stack of bound ancestors. It does not build a complete tree manifest.

For each discovered item:

1. classify the source no-follow and bind its identity/revision;
2. validate the child component before joining it to the destination;
3. bind/revalidate the destination root and every traversed ancestor no-follow;
4. reject a linked ancestor; an exact final destination link is a conflict under §8.4;
5. determine the path-scoped capability profile and strategy;
6. execute the item and emit its terminal record before releasing item-sized state;
7. after a directory’s children finish, restore directory metadata and, for Move, remove the exact now-empty source directory non-recursively.

Hard limits:

- current recursive-walker depth: 128 descendant directory edges below the selected root (root depth 0); raising it requires iterative traversal plus matching evidence;
- admitted bridge files: 16;
- host pump buffers: 256 MiB total;
- queued work: 4,096 entries and 16 MiB retained UTF-16 path text;
- ancestor/directory metadata: 8 MiB and O(depth);
- retry time and provider deadlines: configured and observable;
- no retained record proportional to the total completed tree.

Reaching a bound stops discovery at that item, keeps prior successes, and reports the exact limit. Hidden, system, and filter-hidden recursive children are part of the selected directory and are processed; pane display filters do not redefine selection contents.

#### 6.1.1 Discovery-ahead scheduler

Discovery-ahead and execution share one task scheduler and one bounded queue; they are not two passes. This scheduler applies to recursive Copy, Move, Delete, and Recycle; “transfer” below means the byte-copying branch of Copy/Move. The default scheduler keeps enumeration responsive without permanently withholding user-configured throughput:

- begin mutation as soon as the first item is safely actionable; do not wait for a directory total;
- start performance validation with queue low-water/target candidates of 32/128 items and an adaptive target no greater than 256 items, always below the §6.1 hard limits;
- while discovery remains open and discovery-ahead is enabled, reserve one provider request slot for enumeration when the provider supports concurrent requests; transfer may use the other configured slots;
- when only one provider request can run, interleave one bounded enumeration batch between completed transfer items instead of attempting impossible concurrency;
- below low-water, prioritize enumeration and do not admit additional transfer items beyond the already active set; active pumps yield at bounded checkpoints of no more than 4 MiB or 50 ms so queued discovery/cancel work can run;
- at or above target, let transfer consume work without further discovery until the queue falls below target;
- when traversal closes, release the reservation and use the full configured concurrency and bandwidth;
- when the user selects **Skip discovery**, stop discovery-ahead and release its reservation/throttling immediately as defined by §4.3. Execution then discovers the next required child just in time and otherwise runs at full configured speed/concurrency.

The configured bandwidth limit counts transfer and verification data, not metadata enumeration bytes. Enumeration still shares provider concurrency and is measured separately. The numeric queue candidates are internal tuning parameters, not product settings or ABI. Phase 0 registers them as perf parameters; Phase 2 chooses final defaults from §14/§17 evidence. Their boundedness, no-starvation behavior, ≤50 ms/4 MiB scheduling checkpoint, and immediate Skip-discovery release are normative even if the thresholds change.

### 6.2 Copy publication

For a regular file/object:

1. bind source identity/revision and open the reader against it;
2. create the destination stage/final object exclusively and retain its identity/token;
3. stream through bounded buffers, validating counts, EOF, zero-progress, overflow, and cancellation;
4. Commit/flush and obtain committed identity/revision/size;
5. publish/promote using retained authority, never a recognizable pathname;
6. revalidate final destination identity;
7. copy/report metadata under §9;
8. optionally verify content under §9.1;
9. emit final per-item axes.

Crash before publication leaves no reported final destination. Crash after publication may leave a complete destination. Cleanup never removes an object based only on `.rs_` naming.

### 6.3 Managed Move source cleanup

Managed Move adds these mandatory protections:

- retain a no-follow source handle or immutable provider token from source binding through cleanup;
- exclude source writes/replacement during the copy where the platform supports it; otherwise rely on immutable revision-bound reads;
- copy the protected object/revision, not a later pathname lookup;
- publish to an identity-owned stage/final destination;
- prove the final destination still names the committed identity/revision;
- delete through the bound handle or provider conditional-delete operation;
- if any step is unavailable, failed, canceled, or indeterminate, keep the source.

There is no mandatory post-copy content reread and no mandatory hash for deletion. Verification is independent and optional. This replaces FNV retirement sequencing and avoids cloud egress or a serial second pass. A provider that cannot meet the contract exposes Copy-only.

### 6.4 Directory Move cleanup

After each file reaches a terminal state, release its proof state. Remove a source directory only after all discovered children for that directory are terminal and all children that must be removed are confirmed removed. Removal is nonrecursive and identity-bound. A non-empty, replaced, inaccessible, or indeterminate directory remains and is reported. This yields bounded memory and honest partial Move behavior without a whole-tree deletion manifest.

### 6.5 Overlapping operations

Tasks whose bound destination roots overlap serialize destructive/publication work in deterministic root order. Disjoint roots remain concurrent. Existing Queue/Parallel selection and bandwidth limits survive. Waiting time is reported separately from transfer throughput. Interlocks release on success, failure, cancel, and teardown.

`RenamePlan(InlineRename)` is the explicit Queue-mode admission exception defined by §7.3: it does not wait merely because the global new-task mode is Queue, but it must wait on this overlapping-root interlock. A scheduler or queue cleanup must not generalize “all tasks wait in Queue mode” across this origin distinction.

## 7. File-manager behavior

### 7.1 Folder merge and collisions

Directory onto existing directory is a normal recursive merge. The directory itself does not raise Overwrite. Each colliding child is classified independently. This narrows former FO-D010 to:

- regular file/object collision;
- source/destination type mismatch;
- exact final destination link;
- colliding children inside a folder merge.

File onto directory, directory onto file, and other type mismatches use `Keep Both / Skip / Cancel`, default Cancel. Overwrite is never offered for a type mismatch.

### 7.2 Same-folder operations and containment

- Copy into the same folder is allowed only through Keep Both and creates a unique sibling name.
- `Ctrl+C`, then `Ctrl+V` in the same pane uses that rule and creates a duplicate.
- Move into the same folder is rejected as no-op/invalid; it is not converted to Copy.
- Copy or Move into the selected directory’s own subtree is rejected when discovered, before writing that subtree item.
- A destination resolving to the exact source object through hard link, SUBST, mapped drive/UNC, junction/mount, 8.3 name, case alias, or provider alias must stop before truncation or replacement.
- Copying one hard-link name to a genuinely new sibling is allowed; overwriting another name for the same object is not.

Copy treats separate hard-link directory entries as separate copied files; it does not promise to reconstruct the source hard-link group. Native Move may preserve the provider/volume hard-link structure as part of its native semantics. Managed Move reports no metadata loss merely because the destination files are independent—the content objects were intentionally copied—but a future preserve-hard-link feature requires a separate explicit contract and capability.

### 7.3 Rename

F2 and Batch Rename submit `RenamePlan` through `StartOperation`. The planner owns preview, macros, validation, and cycle declaration. The engine owns admission, worker execution, identity revalidation, conflicts, progress, cancellation, recovery, and results. No production F2 path may call provider `RenameItem` on the UI thread. A rejection before task publication or missing host wiring uses the localized pane operation-error overlay because no central task exists; after publication the central task UI is the only failure presentation and FolderView must not add the old duplicate pane alert.

F2 keeps Explorer-like inline UX: accepting the edited name submits `RenamePlan(InlineRename)` without a second File Operations confirmation dialog. Escape, empty text, or an unchanged name creates no task. An inline-rename plan is admitted immediately and is not gated by the global Queue/Parallel new-task mode; its captured `OperationOptions::executionMode` does not override this origin rule. It still waits on the overlapping-root interlock and all identity/capability safety gates. Provider work runs on the File Operations workers, never on the UI thread.

Inline F2 uses a delayed-reveal presentation policy with the product constant `kInlineRenameCardRevealDelayMs = 500`; it is not a Preference. The delay applies only while the task is running without a prompt or non-success result:

- waiting on the overlapping-root interlock, conflict/other user attention, failure, partial result, cancellation, or an indeterminate axis reveals the popup and task card immediately, without waiting for the deadline;
- if a clean-running task is still active when 500 ms elapses, reveal its card;
- if it reaches `Completed` before reveal with no conflict, no Issues row, and no indeterminate result axis, create no card or completed-group row and do not activate the popup; still emit normal telemetry, publish result/cache notifications, refresh the pane, and retain focus on the renamed object identity;
- once revealed for any reason, the task remains an ordinary visible File Operations task. A later clean success follows the existing success auto-dismiss/retention setting and completed-card rules; it does not become silent again.

The reveal clock belongs to the task/popup state, not to a sleeping worker. Deterministic tests inject a fake clock or a zero reveal delay and must not sleep for 500 ms.

`RenamePlan(BatchRename)` always uses ordinary Queue/Parallel admission and task-card behavior, even when it contains one step. Choosing **Batch Rename** from the F2 prompt launches that ordinary Batch Rename flow and creates no `InlineRename` task. Batch Rename treats its preview plus explicit Run action as confirmation; central execution does not show a redundant second summary prompt.

S3 F2 Rename remains available. Each rename declares the chosen strategy:

- `native` when the provider has real rename;
- `managedMove` for prefix/object copy plus conditional source deletion;
- rejected/Copy-only when destructive proof is unavailable.

Case-only rename is a supported native special case where the path profile declares it. Rename cycles use create-exclusive, identity-owned same-directory intermediate names. Batch Rename maintains a compact, selection-proportional recovery journal containing exact identities, final mappings, completed steps, and intermediate ownership. Journal updates are atomic/durable at every irreversible step. Recovery offers `Resume`, `Roll back`, and `Open location`. Rollback is identity-bound; it never renames a replacement object by pathname.

### 7.4 Delete and Recycle

| Command/state | Behavior |
|---|---|
| Permanent Delete requested | Normal destructive confirmation; delete exact bound selected object only |
| Recycle available | Revalidate selected object no-follow immediately before the shell/provider operation |
| Recycle unavailable or fails before known recycling | Pause the item and offer `Delete permanently / Skip / Cancel`; default Cancel; no Apply-to-all |
| Delete permanently chosen after Recycle failure | Fresh no-follow identity check, then exact delete; result is `Deleted`, with Issues detail that Recycle failed and escalation was granted |
| Recycle outcome indeterminate | Do not offer/delete permanently in the same task; report Unknown and refresh |

The escalation grant applies only to that item and task. It is not a Preference. Delete/Recycle never follows a selected link target.

### 7.5 Create Directory

Create Directory is in this program for provider capability, validation, containment, S3 durability, and truthful completion, but remains outside `StartOperation` in this implementation slice.

For flat S3 namespaces, creating `photos` writes a zero-byte `photos/` marker and reports success only after Commit/queryable visibility according to the provider contract. Copy/Move of an empty prefix recreates the same marker. Providers with native directories do not use markers. S3 directory buckets use their native path-scoped profile.

### 7.6 Archive cleanup and Undo boundary

Pack/unpack “delete source after success” never performs private pathname cleanup. After the archive operation has reached its own successful/verified contract, it submits an ordinary typed `DeletePlan` through File Operations with the original qualified source identities and the command’s explicit destructive consent. If identity cannot be rebound exactly, cleanup fails closed and the source remains.

Undo remains deferred to the Whim `G2` program. This specification supplies the typed results, exact identities, recovery records, and safe inverse-operation prerequisites that Undo may consume; it does not promise an Undo command or retain source data solely for future Undo.

### 7.7 Existing file-manager contracts that survive

These current contracts are explicitly carried forward; “survives” is not permission for the executor to omit them:

| Contract | Required behavior | Current durable/code owner to migrate or preserve |
|---|---|---|
| `Skip All` | Remains an explicit user action for an eligible typed bucket; never inferred from Skip and never available for Recycle escalation | `Specs/FileSystem/FileSystem_FileOperations.md`; File Operations conflict layout/state/popup |
| More-menu action layout | Rare actions such as Replace read-only and Skip All stay reachable under **More…** instead of expanding every prompt | `Specs/FileSystem/FileSystem_FileOperations.md`; `FolderWindow.FileOperations.Popup.cpp` |
| File-on-directory replacement | Type mismatch never offers Overwrite; only Keep Both/Skip/Cancel under §7.1/§10 | File Operations popup contract and conflict model |
| Queue `Start now` and reorder | A queued card can release only itself; adjacent reorder changes waiting order only and never changes global Queue/Parallel mode | `Specs/UI/UI_FileOperationsPopup.md`; File Operations queue/popup files |
| Storage-adaptive concurrency | The destination path/profile’s storage characteristics cap execution concurrency; HDD remains serial, SSD/NVMe/removable/network/cloud use their truthful tested profiles | `Specs/FileSystem/FileSystem_FileOperations.md`; `FileSystemStorageCharacteristics`; scheduler/runtime tests |
| Bandwidth controls | User changes remain live scheduler controls; they do not mutate the immutable data-choice/consent portion of the plan | File Operations popup/runtime and `FileSystemOptions` |
| Pane filters and recursive selection | The top-level selection is what the user selected; once a directory is selected, hidden/system/filter-hidden descendants remain included unless the command explicitly defines a filter | §§2.1, 6.1; FolderView selection contract |
| Removal-focus behavior | Delete/Move repairs focus only after exact per-source `SourceDisposition::Removed`; Retained keeps the row, Unknown refreshes, and stale epochs cannot steal focus | `Specs/UI/UI_FolderView.md`; `BeginRemovalFocusTracking` / `CompleteRemovalFocusTracking` |
| Find/Compare result refresh | Remove a resolved result only for Removed, preserve it for Retained, refresh on Unknown, and refresh a Published/Unknown destination | §§12.3; Find/Compare completion subscriptions |
| Search/index refresh | Normal published/removed paths notify the existing cache/index/watch consumers; Proven and Possible artifacts remain searchable/visible with classification badges | §§12.3, 13; search/index and directory-cache specs |

## 8. Link contract

### 8.1 Policy

- `Preserve`: copy/move the link object without following its target.
- `Skip`: create nothing, read/mutate no target, retain source on Move, and report Skipped with link kind.
- There is no `FollowTargets` mode, session warning grant, or implicit dereference.

The rule applies to top-level and nested file symlinks, directory symlinks, junctions, supported provider shortcuts, and unknown reparse tags. An unknown kind may be preserved only when the provider can round-trip it as an opaque link object; otherwise it is skipped with a reason.

### 8.2 Semantic preservation and retargeting

“Preserve” means preserve what the link means after the operation:

- If a link target is inside the selected source tree and the corresponding target is included in the operation, rewrite the moved/copied link object so it resolves to the destination counterpart.
- If a link target is outside the selected tree, preserve the same external target meaning. Keep raw relative text only when it still resolves to that external target from the destination; otherwise rewrite to an equivalent target representation supported by the provider.
- A tree Move must not leave a relative in-tree link pointing back to the source tree that is then removed.
- The algorithm is uniform for file symlinks, directory symlinks, junctions, native Move, and managed Move, subject to provider capability.
- Rewriting stored target text is mutation of the copied/moved link object, not following or mutating the target.

If semantic preservation cannot be proven, Skip/fail that link and retain it for Move. An unresolved forward in-tree target also blocks deletion of that target’s not-yet-processed source path as described below. Do not silently create a dangling or retargeted-to-wrong-object link.

### 8.3 Streaming retarget implementation

Whole-tree pre-enumeration is not required. Retargeting uses a root transform plus sparse exceptions:

1. Lexically normalize the stored link target against the source link’s parent without opening or following the target object.
2. Record the final selected-root mapping, including a top-level Keep Both name: `sourceRoot -> destinationRoot`.
3. If the normalized target is inside that selected root, compute its destination counterpart by replacing the root and applying only finalized Keep Both component mappings. If it is outside, preserve the original external target meaning from the new link location.
4. When entering a directory, enumerate and reserve destination names for its immediate children in bounded batches before publishing those children. Final Keep Both choices enter a sparse `source-relative component -> final component` map. Ordinary unchanged components require no record.
5. Choose an absolute or relative stored target representation supported by the destination provider. Keep the original raw text only when resolving it from the destination link location produces the required mapped/external path.
6. Publish the link through the normal identity-owned stage/final contract. Native tree Move must run the same semantic check and rewrite when the provider’s native operation would otherwise preserve the wrong path meaning.

A back-reference normally resolves immediately because the target component mapping is already final. A forward reference whose component mapping is not final becomes a deferred link dependency containing the bound source-link identity, raw link payload, normalized source-relative target, destination link location, and required name decisions. The link is not published and its source is not removed until the dependency resolves. Traversal discovery—not link following—finalizes the target mapping; the engine then publishes the link and may remove the exact source link.

Deferred links and sparse rename exceptions share the §6.1 queue/path-memory bounds. If a new forward dependency cannot fit:

- Copy skips/fails that destination link and continues; the source is untouched;
- Move stops further processing of that selected root at the safe checkpoint, leaves the link and all not-yet-processed descendants (including the forward target) at source, and reports Partial/Source retained;
- already completed per-item moves remain completed; there is no false whole-tree rollback claim.

If a target mapping ultimately cannot be established because of a canceled conflict, provider limitation, or discovery failure, apply the same rule. Never delete a forward target after retaining a source link that still depends on it.

The local provider already contains the simpler root-substitution foundation in `Plugins/FileSystem/FileSystem.FileOps.cpp` (`TryRetargetPathIntoDestination` and its reparse-copy call site). Extend or replace that canonical helper with the typed root/sparse-mapping algorithm; do not create a second retargeting dialect in the host.

Worked cases:

| Source link | Operation | Required destination link |
|---|---|---|
| `Tree\docs\latest -> ..\releases\v2` | Copy `Tree` to `Backup\Tree` with unchanged internal names | `Backup\Tree\docs\latest -> ..\releases\v2`; raw relative text is valid because the root transform preserves the relative relation |
| `Tree\shortcut -> C:\Tree\docs\a.txt` | Move `Tree` to `D:\Archive\Tree` | Rewrite to `D:\Archive\Tree\docs\a.txt` or an equivalent relative form; never leave it targeting the removed `C:\Tree` root |
| `Tree\shortcut -> docs\a.txt` | Copy where `docs` becomes `docs (2)` through Keep Both | Rewrite to `docs (2)\a.txt`; the sparse component map supplies only the renamed component |
| `Tree\docs\external -> ..\..\Shared\policy.txt` | Move `Tree` while `Shared` is outside the selection | Recompute from the destination link location so it still denotes the same external `Shared\policy.txt`; do not apply the selected-root transform |
| `Tree\a\link -> ..\z\target` before `z` is discovered | Streaming Move | Defer the link; once `z` and any Keep Both name are final, publish the retargeted link, then remove the exact source link |

### 8.4 Destination link conflict

An exact final destination link uses `Replace link / Keep Both / Skip / Cancel`, default Cancel.

- Replace removes the exact link object only and never its target.
- Apply-to-all is scoped to the same link kind and conflict class.
- If exact no-follow removal is unavailable, Replace is disabled.
- A linked destination ancestor is a hard stop, not a conflict prompt.
- Revalidate link identity immediately before replacement; a changed kind/object invalidates the grant and re-enters classification.

## 9. Verification, metadata, and consent

### 9.1 Optional verification

Verification defaults off and applies to copied regular file/object contents after final publication.

- The host streaming digest is BLAKE3. A provider digest/proof may replace host reread only when the bound path profile names an equally strong content-proof contract, binds it to the exact committed source/destination revisions, and contract tests prove equivalent bytes.
- Zero-byte files are verified by existence, kind, committed size zero, and the selected content-proof contract.
- Directories report verification NotApplicable; files expanded from a directory selection are verified.
- Native Move reports NotApplicable and performs zero verification reads.
- Host read-back verification for an item begins only after that item is published. It does not overlap transfer of the next item in this slice; the combined ETA is the sum of remaining transfer/publication and verification work. A future overlapping pipeline requires an explicit amendment and new memory/bandwidth evidence.
- Verification failure/cancel on Move keeps the source.
- Verification unavailability is a typed result; it is not silently reported as verified.
- Verification does not authorize overwrite, skip, artifact cleanup, or source deletion.

The confirmation Verify checkbox is visible for Copy/managed-Move intents. Its enabled state comes from the bound destination-root capability profile, not an unqualified provider-wide flag. A known unsupported destination disables it with the reason. Capability lookup has a 300 ms UI budget: on timeout the checkbox remains enabled with “availability checked during operation”; a later per-item unsupported result becomes `VerificationState::Unavailable` rather than silently clearing the user’s selection. The timeout never changes the saved checkbox value. Native-Move-only items remain NotApplicable.

### 9.2 Metadata policy

| Metadata | Copy | Managed Move consequence |
|---|---|---|
| Mark-of-the-Web / `Zone.Identifier`, non-default ADS, extended attributes | Preserve when supported; otherwise publish data and report exact loss | Keep source by default; offer task-scoped `Move anyway; metadata lost`, default Keep Source |
| ACL/owner | Destination inheritance is accepted; warn/report effective change | Does not by itself block source deletion after warning |
| EFS encryption | Preserve when supported; never silently use decrypted-destination behavior | If plaintext would result, gate before affected bytes; default do not decrypt/keep source |
| Sparse allocation | Preserve when supported; otherwise state logical-size inflation | Gate before affected write; free-space estimate uses logical size |
| Compression, timestamps, ordinary attributes | Best effort with exact loss/warning | Failure does not block source deletion after data publication/required gates |
| Directory timestamps | Restore deepest-first after final child write using bound no-follow directory authority | Warning on failure; no whole-tree second walk |

Native Move follows provider/native metadata semantics and reports deviations it can observe.

### 9.3 Shared deferred-consent gate

One host component implements all mid-stream risk gates. It pauses affected work, defaults safe, accepts at most one grant per task/risk/bound destination root where appropriate, records the decision, supports cooperative cancel, writes no Preference, and performs no affected writes while unanswered.

| Risk | Choices and default |
|---|---|
| Insufficient reported destination space | `Continue / Stop`; default Stop |
| Space cannot be reported | Warn once; continue until an actual write error unless the provider contract requires Stop |
| EFS would become plaintext | `Copy/Move as plaintext / Keep source or Stop`; safe option default |
| Sparse file would inflate | `Continue / Stop`; default Stop; display logical destination bytes |
| Placeholder hydration discovered | `Download and continue / Stop`; default Stop; display authoritative known count/bytes |
| Move would lose MOTW/ADS/EA | `Move anyway; metadata lost / Keep source`; default Keep source |
| Recycle failed | Per-item `Delete permanently / Skip / Cancel`; default Cancel; never Apply-to-all |

If confirmation already displayed and captured the same authoritative fact, do not ask again. Facts discovered later use the gate.

Ordinary sharing violation, access denied, path too long, or disk-full-during-write are classified item errors with Retry/Skip/Cancel where the existing conflict framework supports them. They do not trigger UAC or unsafe pathname cleanup.

## 10. Conflict model

Every prompt is generated from a typed conflict record containing source kind/identity, destination kind/identity, risk class, allowed actions, safe default, and Apply-to-all scope. Revalidate the record immediately before mutation.

Every conflict decision is one-shot by default. Apply-to-all occurs only when the prompt exposes an eligible scope from the table and the user explicitly selects it; the engine never infers Apply-to-all from a prior Overwrite/Skip choice. The task snapshot records the exact conflict class, endpoint profile, and kind scope of every grant so it cannot leak to a sibling class.

| Conflict | Actions | Default | Apply-to-all scope |
|---|---|---|---|
| Regular file/object exists | `Overwrite / Keep Both / Skip / Cancel` | Cancel | Same regular-file conflict and compatible endpoint profile |
| Existing destination is read-only regular file | `Replace read-only / Keep Both / Skip / Cancel` | Cancel | Read-only regular files only |
| Read-only known with exists | One combined prompt; do not prompt twice | Cancel | Same combined class |
| Type mismatch | `Keep Both / Skip / Cancel` | Cancel | Same source/destination kind pair |
| Exact destination link | `Replace link / Keep Both / Skip / Cancel` | Cancel | Same link kind only |
| Name not representable/too long | `Keep Both` only if a valid generated name exists, otherwise `Skip / Cancel` | Cancel | Same feasibility class where safe |
| Sharing/access/transient provider failure | Existing `Retry / Skip / Cancel` semantics | Cancel or current safe default | Same exact error class only |
| Explicit `Skip All` for any eligible row above | Execute `Skip` for the current item and record the explicit Apply-to-all grant for that row; expose through **More…** | No implicit default | Exactly the corresponding row's scope; unavailable where no safe scope exists and never available for Recycle escalation |

Folder merge is not a conflict. A digest match does not auto-Skip. Replacing read-only content changes only the exact destination object that was revalidated; attribute clearing/restoration uses retained no-follow authority and never a later pathname lookup.

Keep Both naming is deterministic within the task, collision-safe under concurrent creation, and finalized by create-exclusive publication. Previewed names are advisory until publication succeeds.

## 11. Cancellation and provider lifetime

When Cancel is accepted:

1. task state becomes `Stopping` immediately;
2. no new item or provider call is admitted;
3. host loops stop at bounded checkpoints;
4. active provider operations receive Abort and/or their configured deadline;
5. in-progress stage objects are cleaned only through retained ownership when their outcome is known;
6. source objects are never deleted after cancellation unless the exact source deletion had already completed;
7. the task remains visible until each in-flight call reaches a known or indeterminate quiet point;
8. the host never kills a worker thread or unloads a module under an active callback.

There is no universal 10-second quiet-point rule. Each provider/path profile declares a tested deadline/quiet-point timeout and emits diagnostics when exceeded. The existing fake-provider “Cancel visibly responds within 2 seconds” remains a UI/performance target, not authority to unload unsafe code. An exceeded timeout produces an indeterminate result, preserves possible artifacts/source, quarantines provider use as appropriate, and retains lifetime until the callback really returns.

Pause applies during transfer and verification through explicit checkpoints. App exit treats Verifying/Stopping as live I/O under the existing File Operations exit policy.

## 12. Results, UI, clipboard, and consumers

### 12.1 Per-item axes

```cpp
enum class PublicationState { NotAttempted, NotPublished, Published, Unknown };
enum class VerificationState { NotRequested, NotApplicable, Verified, Failed, Unavailable, Canceled };
enum class SourceDisposition { Removed, Retained, Unknown };
enum class ItemCompletion { Completed, Skipped, Canceled, Failed, Indeterminate };
```

Each item also records strategy, source/destination final names and identities where safe to expose, conflict/grant used, link kind/action, metadata losses, artifact references, HRESULT/provider error, and durations/bytes.

Canonical summaries are composed from axes, not stored as dozens of unrelated outcome enums:

| Axes/example | Summary | Badge |
|---|---|---|
| Published + Verified + Removed | `Moved and verified` | Success + verified mark |
| Published + NotRequested + Removed | `Moved` | Success |
| Published + any non-success verification + Retained | `Copied; source kept` plus verification detail | Warning/partial |
| Published + NotRequested + Retained, planned Copy | `Copied` | Success |
| Published + Retained, Move downgraded | `Copied; source kept` plus reason | Warning/partial |
| NotPublished + Retained + Skipped | `Skipped` plus reason | Neutral |
| Unknown publication or source | `Outcome unknown` | Indeterminate |
| Removed after Recycle escalation | `Deleted` with escalation detail | Success + Issues detail |

Different badges are required because destination success, verification, and Move completion are different facts.

### 12.2 Aggregate HRESULT/status

| Aggregate condition | Result |
|---|---|
| All items fulfilled requested intent | `S_OK`, Completed |
| Only intentional skips or preplanned Copy-only/source-retained outcomes, no hard error | `S_FALSE`, Partial/Completed-with-notes according to UI contract |
| User canceled and no higher-risk indeterminate state | `HRESULT_FROM_WIN32(ERROR_CANCELLED)`, Canceled or Partial if earlier items completed |
| Mixed successes and unexpected failures/source retention | `HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY)`, Partial |
| No mutation and every item failed with one exact error | Preserve the exact error, Failed |
| Any publication/source disposition indeterminate | `HRESULT_FROM_WIN32(ERROR_IO_INCOMPLETE)`, Indeterminate/Partial |

Telemetry and Issues rows use the typed axes even when HRESULT is coarser.

### 12.3 Pane, Find, Compare, and search consumers

- Remove/ghost a source row only for `SourceDisposition::Removed`.
- Retained leaves the row and refreshes metadata as needed.
- Unknown triggers refresh; never optimistically remove it.
- Refresh destination for Published or Unknown publication.
- Find Files and Compare Directories use the same rules for resolved rows.
- Pane, Find, Search, and Compare never exclude an item because of artifact classification. Proven
  and Possible artifacts remain visible and are prominently badged.

### 12.4 Clipboard contract

For a clipboard Move:

- capture the entire accepted source list in the immutable plan;
- after the complete request is accepted and successfully queued, clear the cut clipboard only if its sequence still matches the captured sequence;
- clear it before execution begins so repeated `Ctrl+V` cannot queue duplicate Move tasks;
- never restore it after failure, cancel, Copy-only downgrade, `Copied; source kept`, or indeterminate outcome;
- if confirmation is canceled or queueing fails, do not clear it;
- Copy clipboard contents remain unchanged.

Queue admission therefore has a short start barrier: atomically publish the accepted task and sequence receipt, construct its worker behind a task-owned gate, invalidate the matching clipboard sequence, then release the worker and return successful admission. Thread-admission failure rolls back the task and sequence receipt before clipboard access. If clipboard invalidation itself fails, log/warn and still release the already accepted task; the bounded receipt ledger rejects another admission of the same sequence, and the failure is not permission to discard the captured plan.

The UI makes this consumption visible without implying failure:

- after the start barrier releases, show **“Move queued — cut list cleared.”** in the task/popup acknowledgement;
- when Move retains any source, show **“Some sources remain. The cut list was consumed when the Move started and was not restored.”** with `Open source`, `Open destination`, and `Select retained items` where those locations are known;
- offer **“Cut retained items again”** only as an explicit new user action when the retained set is known exactly; rebind every pathname no-follow at click time and require the captured object identities before publishing the new cut payload;
- for `SourceDisposition::Unknown`, offer Refresh/Open actions but no Cut-again action until refresh establishes which exact sources remain;
- never synthesize a new clipboard list automatically from failures, because that would reintroduce duplicate-paste races and could capture replacement objects.

### 12.5 Discovery and verification presentation

There is no standalone `Calculating` phase. While discovery-ahead remains active, the running task shows a secondary **Discovering** activity and a **Skip discovery** action. Selecting it changes the secondary label to **Discovering as needed**, removes the action, and immediately releases the scheduler reservation defined by §6.1.1. It never changes the selected files or safety checks.

Whole-task progress follows the knowledge state:

- while traversal is open, show indeterminate/estimating total progress plus exact discovered, completed, current-item, and byte counts where known; the UI must not present an estimated denominator as authoritative;
- after traversal closes, freeze the total denominator and compute the combined remaining-time estimate as unfinished transfer/publication time plus requested verification time;
- with verification enabled, use a two-segment whole-task presentation: ordinary `fileOps.progressTotal` transfer/publication fill and hatched `fileOps.progressVerify` verification fill;
- show per-file transfer and verification mini-bars/labels, `StatusKind::Verifying` when every remaining byte of work is verification, and a separate `fileOps.graphVerify` history series;
- verification bytes consume the configured data bandwidth limit but are excluded from pump-only transfer throughput; end-to-end throughput and ETA include verification;
- with verification off, preserve the ordinary single-segment presentation and create no verification reads, graph samples, or fake verification phase.

After **Skip discovery**, totals and ETA remain indeterminate/Estimating until just-in-time traversal naturally closes; full transfer speed does not imply that the engine knows the remaining tree.

## 13. Engine artifacts and recovery

### 13.1 Proven/Possible/Ordinary

| Class | Evidence | Authority and presentation |
|---|---|---|
| Proven engine artifact | A valid atomically persisted recovery claim binds the task/artifact/qualified endpoint/durable phase and the current no-follow provider identity/creation token matches it | Always visible with an obvious Proven badge; exact phase-authorized recovery and guarded touch may be offered |
| Possible engine artifact | Name matches a current or legacy `.rs_*` convention but no valid matching claim plus current identity exists | Never automatic authority; always visible/selectable with “possible interrupted-operation artifact” badge; recovery UI offers Inspect/Reveal only |
| Ordinary | No provenance signal | Normal user file; never treated as engine-owned |

Implementation is intentionally small: exclusive create returns the identity/token; the engine writes one compact record under `%LOCALAPPDATA%\RedSalamander\State\FileOperations\<task-id>.json`; record replacement is atomic; recovery reopens the object no-follow and compares identity before presenting Proven actions. The record is the durable claim source and the reopened provider identity is current object truth; only their conjunction is Proven. The artifact registry/classifier is the sole host join point. A filename, badge, pane row, index field, or in-memory task is only a hint. No filesystem-wide database and no “scan then delete” mechanism is required.

Use a shared recognizable namespace for new artifacts such as `.rs_tmp_`, `.rs_bak_`, and `.rs_ren_`; continue classifying legacy spellings as Possible. Every normal consumer displays both classes; no artifact-specific filter exists. The pane shows artifact name, intended final name when recorded, task/time, and actions such as Resume/Roll back/Open location/Delete after exact validation. There is no automatic post-crash cleanup.

Selection, Properties, Reveal/Open location, and host read-only viewing are inspection. Any user or foreign task that copies/exports, externally opens/edits, moves, renames, deletes/recycles, overwrites, writes, or uses a Proven artifact as a target receives a special warning with Cancel as the safe default. Continue covers only the exact identities in that accepted request, is never persisted/Apply-to-all, and is revalidated before use. The owning task follows only its durable phase and still revalidates. Possible recovery offers Inspect/Reveal only; an explicitly user-initiated ordinary command requires a Possible warning but gains no recovery/cleanup authority.

### 13.2 Move breadcrumb

Each active Move has a compact task record containing task ID, intent, qualified roots, start time, strategy counts, and last durable phase. It is O(1) plus active recovery objects, not a per-file tree manifest. On restart it can say that an interrupted Move may have left both source and destination and offer locations/refresh. It does not authorize deletion.

Batch Rename is the exception: its recovery journal is selection-proportional because rollback/resume requires exact mappings. That cost is explicit and bounded by accepted selection size.

## 14. Performance and observability

### 14.1 Required metrics

Every File Operations item/scenario records, as applicable:

- qualified topology/profile and selected strategy;
- source bytes read, destination bytes written, source bytes hashed, destination bytes reread;
- pump, Commit/publication, verification, metadata, source-delete, interlock wait, consent wait, and end-to-end duration;
- digest algorithm and CPU time;
- publication, verification, source disposition, and completion axes;
- link policy/kind/action and metadata-loss classes;
- discovery mode (`ahead`, `justInTimeAfterSkip`, `closed`), discovered/admitted/completed items and bytes, enumeration calls/latency, queue low-water/target/high-water, reservation state, starvation time, and Skip-discovery release latency;
- queue items/path bytes, walk depth, retained ancestor bytes, deferred links, sparse name mappings, active proof records, buffers, workers, and retry/deadline high-water;
- provider abort/deadline/quiet-point result;
- final metric emitted only after the item’s last required phase.

Pump throughput is named pump-only. End-to-end throughput includes Commit, publication, selected verification, metadata, and source cleanup, but reports user consent/interlock waiting separately.

### 14.2 Quantitative gates

- Host bridge: no more than 16 admitted files and 256 MiB pump buffers.
- MTP writer: streaming memory no greater than 8 MiB plus documented WPD-owned buffers; no whole-payload duplicate.
- MTP reader: reusable/no-copy request path; no per-read allocation proportional to request count.
- S3: known tiny objects use at most 16 KiB reservation for four 4 KiB fixtures; multipart remains bounded by 4 payloads and 256 MiB.
- Native/bridge recursive work: limits in §6.1; total completed tree size does not increase retained memory after steady state.
- Discovery-ahead: no transfer or enumeration starvation under the high-latency/many-small scenarios. A Skip-discovery action releases admission reservation/throttling by the next scheduler checkpoint (target ≤50 ms, excluding an already active provider call) and makes the full configured transfer concurrency/bandwidth available.
- Collision/name index: no more than 64 MiB for the 65,536 × 240 UTF-16 deterministic case while preserving exact and folded collision semantics.
- Destination DACL fallback: lazy; the normal successful overwrite path does not query it unless the fallback needs it.
- FBP-10 buffer/thread restructuring is measure-first: instrument serial/small and parallel utilization before changing the design; accept a change only with lower measured cost and no throughput/cancel regression.
- UI cancel acknowledgement on deterministic fake providers targets ≤2 seconds; provider quiet-point completion uses the declared profile deadline.

No universal throughput number is imposed. Compare test-enabled Release baseline/candidate on the same machine and fail on correctness, quality, memory gates, or statistically credible regression under `Testing_PerformanceValidation.md`.

## 15. Implementation phases

Each phase begins with drift review, updates normative specs/tests in the same change, and stops on the conditions listed. Phase numbers are ordering constraints, not one giant pull request.

### Phase 0 — normative migration and executable skeleton

Deliver:

- apply §16 migration map, use §1.1 to replace every legacy FO-D citation, and remove every stale REVIEW notice/old contradiction;
- create the §16.1 impact manifest, classify every discovered caller/consumer, and update the map before implementation;
- retire the separate pre-calculation settings/state and specify discovery-ahead, Skip-discovery, progress, graph, and ETA behavior in the authoritative owners;
- establish the typed plan/result/conflict models and central-ingress source-contract tests;
- finalize the §5.1 method spelling/layout in `Common/PlugInterfaces/FileSystem.h`, specify the breaking ABI-v2 cutover, and complete the host/production-provider/Dummy lockstep list;
- publish the §8.3 link-retarget cases and dependency/bound behavior as normative executable cases;
- register deterministic selftest/perf cases and metric dictionary before behavior work.

Expected files: authoritative specs in §16, `FolderWindow.FileOperationsInternal.h`, File Operations selftest inventory, provider contract test inventory.

Stop if any normative behavior conflicts with this file and has not been explicitly resolved here.

#### Phase 0 completion evidence — 2026-08-21

Phase 0 is complete. The authoritative File Operations, bridge, virtual-filesystem, provider, UI, settings, and testing owners now state the approved contract; stale REVIEW notices and legacy contradictions in the known scope were replaced. The host and every in-tree provider compile against the mandatory capability-v2 ABI with no v1 adapter, the typed plan/result/conflict skeleton is present, the persisted pre-calculation controls are retired, Verify is persisted, and §16.1 records the complete known caller/consumer impact manifest. Runtime mutations remain unchanged except where required to make this executable skeleton compile and test; Phases 1–6 own the behavior cutover.

Validation:

- command: `.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh`;
- result: **PASSED** — 1,913 total, 1,860 passed, 0 failed, and 53 expected environment/capability skips; no flaky, regression, isolation-suspect, or unclassified results;
- checkpoint: `.build/ValidationEvidence/runs/20260821T201144Z-94052-3ba40111f9d34061936f1c30bd147cd3`;
- build attestation: `cda2ca26f3dbaf5c1b704f4949da364d521c3211e5c68d20343e36d6e0fa10c3`;
- executed-plan hash: `b996b39bea73fde9e70405fa819091a3288b53bbd323c49602eb673261094e7c`;
- build completed with 0 warnings and 0 errors.

Phase 1 may begin from this baseline. This WIP remains active until all later phases, durable spec closeout, archived performance evidence, and final Fresh Full validation are complete.

#### Phase 0 capability-truth correction — 2026-08-22

The post-Phase-0 review found that the first capability-v2 payload reused the legacy `move` value as `nativeMove`. That made the JSON structurally valid but did not distinguish current provider Move admission from proof of the qualified native strategy. The executable contract now requires both fields:

- `operations.move` means that the provider exposes its current Move operation entry point for the queried path/profile; the host may use it only through the strategy and safety gates owned by the implementation phase;
- `operations.nativeMove` means that the provider has a qualified native-Move strategy for the queried endpoint pair; it is never inferred from `move`;
- Local advertises `move: true`, `nativeMove: true` only after P1.3 added and contract-tested
  `FILESYSTEM_MOVE_NATIVE_ONLY`, which fails before every generic/reparse copy-delete fallback;
  writable MTP remains `move: true`, `nativeMove: false` until its native leaf route is similarly
  separated from fallback;
- S3 flat-prefix and Microsoft Drive advertise `nativeMove: true` because their provider-native exact-revision operations are intentional strategies; this fact is independent of whether a managed copy/delete strategy has bound-delete support;
- every required `operations` Boolean and every required v2 section is mandatory. `S_OK` with `{}`, a missing field/section, a null result, or malformed JSON is a diagnosed provider-contract failure and fails closed.

P1.3 replaced the legacy host capability gates that substituted `L"/"` with concrete source and
destination path queries. The plan factory now enforces every same-endpoint operation Boolean or
the complete cross-endpoint Copy export/import pair before task publication. A Move using that Copy
pair is `CopyOnly`: it cannot call source Delete and reports `Copied; source kept`. The distinct Move
pair remains reserved for a future exact managed destructive route. This is top-level admission,
not recursive discovery, and `nativeMove: true` remains insufficient strategy authority until the
qualified endpoint pair and provider-native implementation satisfy the execution gate.

The correction also materializes optional File Operations test settings immediately before a legacy Phase 2 runtime-bridge mutation. Persisted pre-calculation authority remains retired; a selftest may not assume the otherwise-default `fileOperations` optional is engaged.

Correction validation:

- full Debug x64 build: **PASSED**, 0 warnings and 0 errors; receipt `57f94cb7e3b49d86bc206b868ee8088ad0272fa651305c1665e9d95039119289`;
- focused assertion repro `Causeway_BridgeRejectsHostileChildNames`: **PASSED**, run `20260822T094741Z-73380-46bd1baed619483988144c869059cccb`;
- full File Operations family: **PASSED**, 138 total, 118 passed, 0 failed, 20 expected environment/capability skips; run `20260822T095423Z-76116-c75303f30ea6462ca291ec68bdc2e1fb`;
- MTP instance-honest capability matrix: **PASSED**, run `20260822T101124Z-82104-1ee0f3169e144c54ba00ab6c6262b788`;
- Batch Rename missing-identity and duplicate-source provider fixtures: **PASSED**, runs `20260822T101208Z-49108-29efcecad86e4911b65f4ebc0a17984c` and `20260822T101244Z-75900-6d7fd61ef1ae45089d92322e73e9cce5`;
- rebuilt `PluginContractTests`: **PASSED**.

The assertion and dump analysis are preserved under `Specs/TestRuns/SINON/Continuation/2026-08-22_fileops_capability_v2_empty_optional_assert/`. The earlier unmanaged direct-launch result from that session is invalid evidence because runner cleanup raced its sandbox; it is explicitly excluded in the continuation record.

#### Phase 0 normative-owner closure pass — 2026-08-22

- [x] `UI_FileOperationsPopup.md`: confirmation Links/Verify/clipboard snapshot, streaming
  discovery and Skip-discovery, Inline-F2 reveal, typed results/Issues, verification presentation,
  and debug-snapshot ownership.
- [x] Preferences/SettingsStore: provider-owned Preserve/Skip default and silent legacy migration
  are specified without inventing a duplicate global `fileOperations.linkPolicy` key. Runtime
  parser/schema/engine removal remains in its already-owned Phase 3 package.
- [x] Search/Compare/Find/FolderView/Popup artifact projection: no artifact is hidden. Proven and
  Possible remain visible with prominent classification; Proven truth is the atomic recovery claim
  plus current matching no-follow identity; guarded touches use an exact one-request grant; Possible
  recovery is Inspect/Reveal only.
- [x] `FileSystem_S3.md`: flat-prefix marker, directory-bucket/native distinction, durable Create
  Directory/empty-tree transfer, conflict, cleanup, and deterministic test contract.
- [x] `Plugins_VirtualFileSystem.md`: complete built-in capability-v2 profile matrix, with Local,
  Dummy, and 7-Zip owned in the shared matrix rather than empty wrapper specs.
- [x] Whim G2 dependent wording now permits only exact typed receipts and Proven Engine artifact
  tokens; pathname sets are not an inverse journal.

Phase 0 normative-owner closure is complete and Phase 1 is authorized. Current closure validation:

- full Debug x64 project/solution-graph build: **PASSED**, 0 warnings and 0 errors; artifact receipt
  `56d70793a9bed86fc977e6c5b2a13ebd167953a367d3538c1d8e8125fd18ebce`;
- direct filtered Commands entrypoint
  `file_operations_phase0_typed_contract_source_guard`: **PASSED** with exit code 0;
- `git diff --check`: **PASSED**.

The runtime behavior packages remain assigned to Phases 1–6; this normative closure pass does not
authorize out-of-order mutation-engine changes.

### Phase 1 — P0 identity, containment, and ownership

Deliver:

- qualified endpoints on every ingress;
- materialize the central typed plan factory/admission path and route in-pane F2 through worker-executed `RenamePlan(InlineRename)` as soon as that path exists; remove the UI-thread `RenameFocusedItem` → provider `RenameItem`/legacy `ReportError` mutation bypass in this phase, not Phase 5;
- enforce the §7.3 immediate-admission/global-Queue exception and §6.5 overlapping-root wait on that Phase 1 F2 route, with exact identity/capability revalidation before rename;
- local no-follow `FILE_ID_INFO` binding and provider identity abstraction;
- same-object/alias and subtree stops before truncation;
- bound destination root/ancestor checks;
- create-exclusive identity-owned staging/publication/cleanup;
- null-success provider validation and device-namespace envelope rejection.

The 500-ms delayed-reveal/silent-clean-success presentation may land with the Phase 5 popup work. Until then, an in-development `InlineRename` task may use the ordinary visible-card presentation, but it may not execute directly, run on the UI thread, or inherit global Queue waiting. This temporary presentation is not release/Done behavior.

Closes: `FBC-01`–`FBC-06`, `FBC-11`, `FBC-16`, `FBC-21` foundations.

Stop if a destructive path still uses pathname/basic metadata as sole authority, an ABI cannot land atomically across host/providers/tests, or any production F2 path can still mutate through direct UI-thread provider `RenameItem`.

### Phase 2 — managed Move and bounded traversal

Deliver:

- reader-bound source, destination proof, conditional/handle-bound deletion;
- Copy-only downgrade where any prerequisite is absent;
- remove FNV silent skip/delete authority and mandatory destination reread;
- per-file source cleanup, nonrecursive directory cleanup, O(depth) metadata;
- bounded queue/path/depth/proof state; no whole-tree manifest;
- single-pass discovery-ahead scheduler, just-in-time Skip-discovery mode, and starvation/bandwidth instrumentation;
- retryable cleanup records retained until terminal.

Closes: `FBC-02`, `FBC-04`, `FBC-10`, `FBC-13`, `FBC-14`, `FBC-18`–`FBC-20`, `FBP-02`, `FBP-03`, `FBP-05`.

Stop if a provider is expected to advertise destructive Move without exact conditional cleanup.

### Phase 3 — links, conflicts, and normal file-manager semantics

Deliver:

- Preserve/Skip ABI and silent migration;
- semantic retargeting and bounded deferred links;
- selected-root transform, sparse Keep Both component mappings, and forward-target deletion dependencies;
- destination-link/type/read-only conflict classes;
- folder merge, Keep Both, same-pane duplicate, same-folder Move rejection;
- identity-safe metadata operations and no-follow link replacement.

Closes: `FBC-07`–`FBC-09`, remaining `FBC-13`, `FBC-17`.

Stop if tree Move can leave an in-tree relative link targeting a source object it removes, or if folder merge becomes folder overwrite.

### Phase 4 — confirmation, consent, metadata, results, and clipboard

Deliver:

- immutable confirmation controls and settings persistence;
- shared deferred-consent gate;
- metadata matrix and source-retention rule;
- typed axes, HRESULT/UI/Issues/consumer mapping;
- accepted-queue clipboard clearing;
- clipboard-consumption/result actions;
- optional verification states, two-segment progress, graph series, and ETA contract.

Closes: `FBC-12`, `FBC-19`, result/consent/product gaps.

Stop if a Move deletes a source after ungranted security-significant metadata loss, or if any consumer removes a row without `SourceDisposition::Removed`.

### Phase 5 — Rename, artifacts, recovery, and S3 directories

Deliver:

- complete the Phase 1 F2 route with the §7.3 delayed-reveal/silent-clean-success popup policy and final rename strategy coverage;
- migrate Batch Rename to ordinary-admission `RenamePlan(BatchRename)`, including one-step plans, cycles, results, and recovery;
- durable rename journal and recovery actions;
- Proven/Possible/Ordinary artifact classifier, badge, grouping, and exact recovery;
- never-hidden projection across pane/Find/Search/Compare plus the exact-identity Proven/Possible
  touch-warning gate;
- Move task breadcrumb;
- S3 flat-prefix trailing-`/` Create Directory/empty-directory transfer behavior.

Stop if F2 regresses to a direct mutation path, an `InlineRename` plan waits only because global Queue mode is selected, a one-step Batch Rename inherits the F2 presentation exception, an artifact is hidden because of its classification, a pattern match becomes cleanup authority, a Proven/Possible touch bypasses its warning/revalidation gate, recovery can mutate a replacement object, or S3 reports a durable folder without committing its marker/native directory.

### Phase 6 — cancellation, provider performance, executable matrix, evidence

Deliver:

- Abort/deadline/quiet-point plumbing for enumeration/Read/Write/Commit/mutation;
- MTP/S3/native/bridge memory and allocation fixes;
- truthful end-to-end metrics and measurement-first FBP-10 decision;
- executable capability matrix across every production mode;
- deterministic Debug and test-enabled Release runs, archived evidence, Fresh Full;
- normative closeout and move this plan to Done.

Closes: `FBC-15`, `FBP-01`, `FBP-04`, `FBP-06`–`FBP-10`, `FBCON-01`–`FBCON-03`.

Stop if cancellation requires unsafe thread termination/module unload, or performance evidence reports success before the final required phase.

## 16. Normative specification migration map

Phase 0 must update these owners before corresponding production code lands:

| Normative owner | Clauses to merge/replace |
|---|---|
| `Specs/FileSystem/FileSystem_FileOperations.md` | Entire user-visible contract: strategies, single-pass discovery and Skip-discovery, merge, conflicts, same-folder behavior, Move safety, hard-link copies, archive cleanup, results, delete/recycle, metadata, clipboard, cancellation, recovery, and Undo boundary |
| `Specs/Core/Core_FileSystemBridge.md` | Bound source/publication/delete flow, bounded streaming, provider validation, metrics, no mandatory reread, Copy-only downgrade |
| `Specs/Plugins/Plugins_VirtualFileSystem.md` | Capability v2, qualified endpoints, per-call link policy, optional binding interfaces, cancellation/deadline, truthful provider matrix |
| `Specs/SettingsStore.schema.json` and `Specs/Core/Core_SettingsStore.md` | Remove `preCalcEnabled`/`preCalcMaxWorkers`; `verifyAfterCopy` default false; link-policy migration/default; progress/graph theme keys; no task grants or Skip-discovery choice persisted |
| `Specs/UI/UI_PreferencesDialog.md` | Remove pre-calculation toggle/worker controls; add File Operations verify setting and provider link-policy control without Follow |
| `Specs/UI/UI_FileOperationsPopup.md` | Confirmation controls; Discovering/Skip-discovery/Discovering-as-needed; Waiting/Verifying/Stopping; estimating/final totals and ETA; two-segment verification progress/graph; axes/badges; clipboard-consumption copy/actions; Issues rows; deferred gates; inline-F2 500-ms delayed reveal/immediate non-clean reveal/silent clean success; artifact recovery entry points |
| `Specs/UI/UI_FolderView.md` | same-pane duplicate; accepted-queue cut-list consumption replacing verified-success invalidation; worker-executed `RenamePlan(InlineRename)` with no second confirmation, Queue bypass, interlock wait, delayed card reveal, identity-based focus retention, and no post-publication duplicate pane alert; source disposition refresh; artifact presentation/grouping |
| `Specs/UI/UI_BatchRenameWindow.md` | typed plan submission, engine ownership, journal/recovery, cycle intermediates |
| `Specs/Core/Core_Search.md`, Compare/Find specs | Proven/Possible artifact visibility and source/destination refresh semantics |
| `Specs/FileSystem/FileSystem_S3.md` | flat-prefix marker and directory-bucket distinction; conditional generation cleanup |
| Local, Curl, MTP, Microsoft Drive, 7z, Google Drive, Dummy provider specs | Accurate path profiles, identity/publication/delete/cancel/metadata capabilities and Copy-only cases |
| `Specs/Testing/Testing_SelfTests.md` | Deterministic functional/race/cancel/provider-matrix inventory; replace pre-calc/`Calculating` fixtures with discovery-ahead/Skip-discovery scheduler and popup cases |
| `Specs/Testing/Testing_PerformanceValidation.md` | metric fields, bounds, baseline/candidate quality, end-to-end phase semantics |

### 16.1 Known code and test impact map

This is the minimum known implementation set, not a closed allowlist:

| Area | Known files/components | Required change |
|---|---|---|
| Public plugin ABI | `Common/PlugInterfaces/FileSystem.h`; `FileSystemPluginManager.*`; capability parsing in `FolderWindow.FileOperations.cpp`/state | Breaking ABI version, v2 path capability parser, binding/publication interfaces, size/output validation, QI tables, diagnostics |
| Host plan/engine | `FolderWindow.FileOperationsInternal.h`; `FolderWindow.FileOperations.State*.cpp`; `FolderWindow.FileOperations.cpp` | Materialize §4.2 plan/result/conflict types, plan factory, identity-bound strategy router, discovery scheduler, source disposition, central ingress |
| Confirmation/popup/issues | `FileOperationConfirmation.*`; `FolderWindow.FileOperations.Dialog.cpp`; `FolderWindow.FileOperations.Popup.*`; `FolderWindow.FileOperations.IssuesPane.*` | Typed confirmation, Links/Verify/clipboard copy, deferred gates, Skip discovery, progress/ETA, scoped Apply-to-all/Skip All, badges/actions |
| Settings/Preferences/resources | `Common/SettingsStore.h`; `Common/Common/SettingsStore.cpp`; `Specs/SettingsStore.schema.json`; `Preferences.FileOperations.*`; `Preferences.Dialog.cpp`; main/provider `.rc` and satellite resources | Remove pre-calc controls/settings; migrate link policy; add Verify/theme/resource strings; preserve localization placeholder rules |
| Local provider | `Plugins/FileSystem/FileSystem.h`; `FileSystem.cpp`; `FileSystem.FileOps.cpp`; local QI/factory/tests | No-follow `FILE_ID_INFO`, exact alias/containment, owned stage/publication, Preserve/Skip, semantic retarget, Recycle contract, remove FNV authority/FollowTargets |
| Curl provider | `Plugins/FileSystemCurl/FileSystemCurl.*`; Curl specs/selftests | Capability-v2 false/true claims matching safe executable paths; preserve non-advertised direct code until it can meet conditional publication/cleanup; bridge export truth |
| MTP provider | `Plugins/FileSystemMtp/FileSystemMtp.*`; MTP spec/selftests | Distinguish native same-device leaf-preserving Move from unsafe transfer-delete fallback; binding/cancel/memory truth; Copy-only fallback |
| Microsoft Drive provider | `Plugins/FileSystemMicrosoftDrive/FileSystemMicrosoftDrive.*`; provider spec/selftests | Advertise/test Graph ID+ETag native Move accurately; add v2 path profile/binding/digest/placeholder facts; preserve conditional reconciliation |
| S3 provider | `Plugins/FileSystemS3/*`; `Specs/FileSystem/FileSystem_S3.md` | Flat-prefix marker, directory-bucket profile, conditional generation/VersionId cleanup, v2 capability/tests |
| Other providers | `Plugins/FileSystem7z/*`, `FileSystemGoogleDrive/*`, `FileSystemDummy/*`, every other `IFileSystem` implementation and factory/QI table | Compile-time ABI cutover, explicit false unsupported fields, executable contract rows, Copy-only routing where proof is absent |
| Clipboard/pane/drag-drop | `FolderView.FileOps.cpp`; `FolderView.DragDrop.cpp`; `FolderView.h`; `FolderWindow.cpp` | §4.2.1 plan creation, same-folder policy, qualified OLE/internal endpoints, accepted-queue cut clearing, removal-focus result mapping |
| Find/Compare | `FindFilesWindow.*`; `CompareDirectoriesWindow*.cpp`; Compare engine/specs | Preserve resolved identities/mappings, central plans, per-item results and truthful refresh/removal |
| Rename | `FolderView.FileOps.cpp`; `BatchRenameExecutionEngine.*`; `BatchRenameWindow.*`; File Operations popup/task state; Batch Rename specs/tests | Typed `RenameOrigin`; F2 worker execution, immediate admission, 500-ms reveal state, no-second-confirmation/no-legacy-alert path; Batch normal admission/card and multi-step/cycle plan; eliminate direct execution exemptions; journal/recovery/results |
| Pack/unpack cleanup | `FolderWindow.FileSystem.Commands.cpp`; archive prompt/resources/tests | Preserve the completed central Pack/Unpack `DeletePlan` routes, captured receipt binding, and no-private-pathname-delete guard |
| Create Directory | `FolderWindow.FileSystem.Commands.cpp`; `IFileSystemDirectoryOperations`; S3/local/provider implementations and tests | Keep the command outside `StartOperation`; apply path capability/containment and make flat-S3 success commit a durable trailing-`/` marker |
| Cache/search/watch consumers | `DirectoryInfoCache.*`; `Common/SqliteIndexStore.*`; `Common/LocalSearchIndexCore.*` including `IsRedSalamanderStagedTempName`; `FindFilesWindow.*`; `CompareDirectoriesWindow*`; search/Compare/Find specs/tests including `CompareDirectoriesEngine.SelfTest.Cases.SearchAndIndex.cpp` | Consume Published/Removed/Retained/Unknown axes; remove generated-name search suppression; project never-hidden Proven/Possible badges; route guarded touches through File Operations; no pathname inference |
| File Operations tests/perf | `RedSalamander/SelfTest/FileOperations/*`; `SelfTest/Commands/Commands.SelfTest.FileOps.cpp`; provider tests; `Specs/TestRuns/` | Replace obsolete pre-calc/Follow/FNV assertions, add §17 races/matrix/perf evidence, archive test-enabled Release results |

#### Unknown-impact closure rule

Unlisted code is **unknown**, not out of scope. Before each phase changes code, the executor must refresh an impact manifest using `rg`/`rg --files` for at least:

```text
StartOperation|CopyItem|MoveItem|DeleteItem|RenameItem|CopyItems|MoveItems|DeleteItems|RenameItems
GetCapabilities|CanSameFileSystemOperation|CanCrossFileSystemCopyMove|QueryInterface
preCalc|Calculating|followTargets|FollowTargets|FNV|PermanentDelete|SkipAll|Apply.to.all
FileOperationRequest|FileOperationCompleted|InvalidateMoveClipboard|NotifyPath|NotifyFolderContentsChanged
IsRedSalamanderStagedTempName|.rs_tmp_|.rs_copy_tmp_|.~rs-write-|.rs_bak_|.rs_ren_
DeleteUnpackedArchives|deleteSourcesAfterPack|deleteArchiveAfterUnpack
```

The phase ticket records every match as **Update**, **Test-only update**, or **Reviewed unrelated**, with a short reason. It also runs `rg -n "FO-D[0-9]{3}|REVIEW REQUIRED" Specs` and maps every remaining File Operations citation/banner through §1.1 or a current section. Any newly discovered mutating ingress, provider implementation, result consumer, setting/resource, or test is added to this table before implementation continues. This prevents the present inventory from becoming a false completeness claim as the tree changes.

#### Phase 0 impact manifest — 2026-08-21 baseline

The §16.1 scan was rerun against `Common/`, `Plugins/`, `RedSalamander/`, `Tests/`, and active
`Specs/`. Matches are exhaustively classified by the following directory/file groups; a file matched
through generic `QueryInterface` text alone is included in the reviewed-unrelated rows.

| Classification | Current matches | Reason / owning phase |
|---|---|---|
| **Update — Phase 0 complete** | `Common/PlugInterfaces/FileSystem.h`; `Common/FileSystemPathIdentity.cpp`; every shipped `Plugins/FileSystem*` capability implementation/header/QI table; `FolderWindow.FileOperations.cpp`; `FolderWindow.FileOperationsInternal.h`; `SettingsStore.*`; Preferences File Operations/dialog/resources and satellites; `Tests/PluginContractTests`; capability/parser test doubles under Commands, Compare, Viewer, and Performance tests | Breaking capability-v2 and typed plan/result skeleton; retired persisted pre-calc controls; verification default; compile-time provider cutover and contract guards. |
| **Update — Phase 1** | `FileSystemPluginManager.*`; `FolderWindow.FileOperations*.cpp`; `FolderView.FileOps.cpp`; `FolderView.DragDrop.cpp`; `FolderWindow.cpp`; Find/Compare ingress; `DirectoryInfoCache.*`; local `FileSystem.FileOps.cpp`/`DirectoryOps.cpp`; factories and optional binding QI tables | Materialize plan factories, qualified bindings, no-follow identity/containment, exclusive stage ownership, central ingress, clipboard acceptance, and Inline-F2 admission. No runtime behavior is authorized before this phase. |
| **Complete — Phase 2** | `FolderWindow.FileOperations.State*.cpp` and private/internal state; popup discovery fields/actions; Local provider traversal; File Operations Fairstream and Phase 05–09 tests | P2.1–P2.6 implemented and closed Managed Move cleanup, bounded traversal, one-pass discovery/execution, Skip-discovery, truthful open/closed totals, and the integrated destructive-strategy stop-condition gate. Later audits may add explicit post-close hardening packages without reopening provider copy/delete. |
| **Update — Phase 3** | local/Curl/MTP/Microsoft Drive/S3 mutation bodies; `Preferences.Plugins.cpp`; link-policy resources; File Operations Phase 10–13 tests | Remove Follow behavior, implement Preserve/Skip and semantic retarget, conflict/name gates, conditional cleanup, and Recycle escalation. All remaining `followTargets`, FNV mutation-authority, and `PermanentDelete` conflict matches belong here unless the test row explicitly asserts retirement. |
| **Update — Phases 4–5** | confirmation/dialog/popup/issues files; `BatchRenameExecutionEngine.*`; `BatchRenameWindow.*`; Find/Compare/Search result consumers | Typed results/verification/metadata/consent/artifact presentation, remaining central Rename/DeletePlan execution, and exact consumer refresh. Archive cleanup entered central `DeletePlan` in P1.6. |
| **Update — Phase 6** | MTP/S3 allocation and I/O files; cancellation plumbing in every provider; File Operations diagnostics/perf; `Specs/TestRuns/` evidence | Abort/deadline/quiet-point execution, bounded provider memory, final metric emission, and archived evidence. |
| **Test-only update** | `RedSalamander/SelfTest/FileOperations/*`; relevant Commands/Compare/provider selftests; `Tests/PluginContractTests`; Viewer/Performance filesystem stubs | Replace old capability/pre-calc/Follow/FNV assumptions and add the §17 behavioral/race/performance inventory. Test-only direct provider calls are fixtures, not production ingress. |
| **Reviewed unrelated** | `Common/DxUi/*`, `Common/Helpers.h`, `Common/WindowMessages.h`, generic `FactoryImpl.h`; Terminal and Viewer plugin implementation/factory files; DxUi tests; UI theme/accessibility files; `Specs/Reviews/**`, `Specs/Plans/Done/**`, and `Specs/TestRuns/**` historical captures | Matches arise from generic COM `QueryInterface`, unrelated clipboard/accessibility text, message names, or immutable historical evidence. They do not implement an `IFileSystem` mutation ingress or current File Operations authority. |
| **Reviewed adjacent** | `Common/LocalSearchIndexCore.cpp`; Search/Compare cache code; connection specs/code; `Plugins/Plugins_PluginAPI.md`; `Specs/NormativeConsistency.json`; FolderView performance plans | These are consumers/dependencies. They retain their own authority but must consume typed File Operations results or preserve v2 ABI facts when their owning implementation phase lands. |

The active-spec citation scan now finds legacy `FO-D` anchors only in §1.1 and explanatory
crosswalk text in this file. File Operations REVIEW banners in authoritative and dependent active
specifications have been replaced with direct clauses. Historical Done plans, archived TestRuns, and
review reports remain evidence and are not rewritten as current contracts.

#### Phase 1.1 impact refresh — 2026-08-22

The mutation/admission scan was rerun after central typed admission and mixed-root grouping. Every
current production-tree match is classified below; provider-internal implementation calls remain
in scope through their owning later packages even though they are not separate user ingress.

| Classification | Exact current matches | Reason / next owner |
|---|---|---|
| **Update — P1.1 complete** | `FolderWindow.FileOperations.cpp`: every production Copy/Move/Delete/Recycle/Find/Compare request calls `AdmitOperation`; `FolderWindow.FileOperations.State.Runtime.cpp`: the sole typed `StartOperation(OperationAdmission)` publication boundary; `FolderView.FileOps.cpp` and `FolderView.DragDrop.cpp`: clipboard/internal/external request callbacks | One user request now builds qualified immutable child plans. Mixed roots become one child per endpoint under one confirmation/task/clipboard barrier. No production caller publishes a legacy argument envelope directly. |
| **Update — P1.2 complete** | `FolderWindow.FileOperations.cpp` host binding/revalidation abstraction; `Plugins/FileSystem/FileSystem.cpp` local `IFileSystemObjectBinding` QI, retained `LocalBoundObject`, and `FILE_ID_INFO`; `Plugins/FileSystem/FileSystem.FileOps.cpp` provider debug matrix; File Operations and Commands selftests | Exact local no-follow identity and typed host outcomes are executable. Null/malformed success is a provider contract violation; snapshot pointers are copied immediately and retained handles/tokens own lifetime. Conditional mutation flags and exclusive publication remain honestly unsupported for P1.4. |
| **Update — P1.3 complete** | `FolderWindow.FileOperations.State.cpp` rejects bulk Copy/Move, unavailable `Managed`, and `Native` plus destination bridge before provider I/O; serial and parallel scheduling call one shared per-item strategy dispatcher. `Native` passes `FILESYSTEM_MOVE_NATIVE_ONLY`; `CopyOnly` uses the same mandatory byte-count/committed-size publication validation as Copy, never arms source cleanup, and retains the source. Optional BLAKE3/provider-proof verification remains Phase 4 work and the transitional mandatory FNV destination reread was retired rather than expanded to Copy-only. `FolderWindow.FileOperations.cpp` applies path-profile mapping/same-folder/descendant rules, retains and revalidates source/destination/all destination ancestors for every source kind, compares exact object IDs across qualified root aliases, and derives out-of-pane explicit-provider identity from `IInformations`. Dummy, Microsoft Drive, and S3 contract-test their NativeOnly routes; `FolderWindow.FileOperations.State.Runtime.cpp` no longer performs the canonicalizing top-level filesystem overlap preflight. The executable Move-cleanup entry point was removed at P1.3; the then-dormant bridge manifest internals were removed by P2.3. Local `Plugins/FileSystem/FileSystem.FileOps.cpp` legacy `TryAreSameFile` remains provider-internal and must be reassessed when Phase 2/3 retires the remaining fallback implementation. | Qualified-provider structural/exact/native authority is executable immediately before mutation. Same-path Copy selects Keep Both, native Move cannot silently fall through to copy-delete, CopyOnly produces truthful source-kept completion, and alias/swap/indeterminate cases are deterministic. |
| **Update — P1.4 complete** | `Common/PlugInterfaces/FileSystem.h`; local `Plugins/FileSystem/FileSystem.cpp`; Dummy atomic writer; `FolderWindow.FileOperations.State*.cpp`; File Operations Fairstream/Phase 05–13 selftests; bridge/File Operations/VFS specs | Exclusive provider-created stages and exact expected-destination publication replace pathname staging authority. Local publication, rollback, cleanup, metadata, and content verification retain exact handles; a concurrent pathname replacement is preserved. Null/malformed success fails closed, CopyOnly never reaches managed cleanup, and stage/publication/retained-artifact metrics are bounded aggregates. |
| **Update — P1.5 complete / Phase 5 presentation and Batch Rename remain** | `FolderView.FileOps.cpp` Inline-F2 request creation; `FolderWindow.FileOperations.cpp` typed rename admission; `FolderWindow.FileOperations.State*.cpp` conditional rename and mutation interlock; File Operations/Commands selftests; `BatchRenameExecutionEngine.cpp`; `FileSystemRenameBatch.cpp` | Inline F2 now mutates only on a File Operations worker through `RenamePlan(InlineRename)`, exact authority, and conditional provider rename. It bypasses global Queue waiting but not the identity-aware overlapping-root interlock. Phase 5 still owns the normative delayed-reveal/silent-clean-success presentation and Batch Rename's typed execution, rollback, and journal migration. |
| **Update — P1.6 safety integration complete** | `Common/FileSystemPathIdentity.*`; `FolderWindow.FileOperationsInternal.h`; `FolderWindow.FileOperations.cpp`; `FolderWindow.FileOperations.State*.cpp`; `FolderView.h`; `FolderWindow.FileSystem.Commands.cpp`; File Operations/Commands selftests; File Operations and testing specs | Archive delete-after carries captured consent into central `DeletePlan`; interlock roles distinguish shared source reads from destructive/publication overlap; provider-profile parents govern bridge shells and immutable plan-owned destinations across execution/completion/refresh; CSPRNG failure stops before writer creation; Copy and Copy-only share byte-count/committed-size validation and immutable optional-verification intent. Debug and test-enabled Release closeout coverage is recorded in the P1.6 checklist. |
| **Update — P1.6 archive complete / Create Directory remains** | `FolderWindow.FileSystem.Commands.cpp` Pack/Unpack cleanup; provider/CreateDirectoryW command | Pack and Unpack cleanup submit central consent-carrying permanent `DeletePlan`. Create Directory remains outside `StartOperation` but must use path capability/containment and durable provider creation rules. |
| **Test-only fallback/update** | `FolderView.FileOps.cpp:600,682,738,882-883,1358`; `FolderView.DragDrop.cpp:781,785`; File Operations/Commands test sources | Direct FolderView/provider fallbacks are reachable only behind explicit selftest gates and are forbidden when the production callback is absent. Tests must migrate or remain named fixtures as the relevant package lands. |
| **Reviewed provider-internal** | local shell `IFileOperation` delete calls; MTP backend Copy/Move/Rename/Delete/CreateDirectory; Microsoft Drive provider selftest paths | These implement a provider entry point or deterministic provider selftest. They do not bypass host admission, but their identity/strategy truth remains governed by the provider and Phase 2/3/6 contracts. |
| **Reviewed adapter** | `CompareDirectoriesEngine.cpp:4553-4621` virtual filesystem pass-through methods | This read/comparison adapter implements the mandatory `IFileSystem` surface and forwards to its base provider; synchronization ingress itself already submits typed explicit mappings. It gains optional binding forwarding only if the wrapper is used on a path requiring that interface. |
| **Reviewed unrelated filesystem housekeeping** | crash/session/settings/perf/index diagnostic directory and marker operations; ViewerSqlite temporary local-file cleanup | These paths own application state, logs, indexes, or viewer materialization and are not user File Operations. They do not operate on selected pane/provider items. |

No unclassified production mutation match remains from this refresh. Future packages must rerun the
same scan because this table is a dated inventory, not an allowlist.

#### P2.3 impact refresh — 2026-08-23

The complete §16.1 mutation/legacy scan was rerun across `Common/`, `Plugins/`,
`RedSalamander/`, `Tests/`, and active `Specs/`: 80 mutation/capability files, 75 legacy-symbol
files, and 122 unique files. The P2.3 matches are classified as follows:

| Classification | Exact current matches | Result / next owner |
|---|---|---|
| **Update — P2.3 complete** | `Plugins/FileSystem/FileSystem.FileOps.cpp`; `FolderWindow.FileOperations.State.cpp`; state private/queue/internal declarations; `Specs/FileSystem/FileSystem_FileOperations.md`; `Specs/Core/Core_FileSystemBridge.md` | Removed content-derived identical Skip, host copy-time cleanup hashing, mandatory managed-cleanup reread/path revalidation, pathname-keyed whole-tree cleanup records, and destination-owned cleanup classification. Exact bound cleanup and typed collision decisions remain. |
| **Test-only update — P2.3 complete** | File Operations selftest state plus Fairstream, Phase 05–06, and Phase 07–09; Commands plugin source contracts | Added identical-collision prompts, four cleanup buckets × Retry/Skip, real alternate-volume admission, affected-ancestor retention, and guards against the retired host walker/hash names. |
| **Update — P2.6/P2.7** | Local provider Move fallback and its copy-time snapshot/delete walker | Removed. Local provider Move is one native rename only; cross-volume, folder-merge, and semantic-copy routes must use host Managed publication plus exact `DeleteIfUnchanged` or fail source-kept. No provider-owned copy/delete or per-child rename manifest remains to revive. |
| **Historical P2.5 audit boundary** | Follow/link-policy code; Delete/Recycle conflict actions; remaining Phase 10–16 tests and UI popup clauses | P2.5 removed the active runtime pre-calc/`Calculating` bridge. Phase 3 subsequently delivered links, conflicts, and Recycle; Phases 4–6 own the remaining presentation, typed results, provider completion, and final evidence. No P2.3 mutation authority depends on these matches. |
| **Reviewed unrelated/historical** | Generic COM/capability matches and excluded `Specs/Plans/Done/**`, `Specs/Reviews/**`, `Specs/TestRuns/**` evidence | No user File Operations mutation ingress or current cleanup authority. |

#### P2.4 impact refresh — 2026-08-24

The traversal/retention scan was rerun over the host executor, Local provider, File Operations tests,
Commands source contracts, and current domain/testing specs. Current matches are classified as follows:

| Classification | Exact current matches | Result / next owner |
|---|---|---|
| **Update — P2.4 complete** | `Plugins/FileSystem/FileSystem.FileOps.cpp`; `FolderWindow.FileOperations.State.cpp`; state private/queue/internal declarations | Bounded file/reparse admission, O(depth) directory authority, post-order created-directory metadata, exact quantitative ceilings, terminal limit routing, and emitted high-water/limit evidence. No completed-tree-sized directory or cleanup list remains. |
| **Test/spec update — P2.4 complete** | File Operations Phase 07–09 and Phase 10–13 tests plus shared fixture helpers; Commands plugin source contracts; `Specs/FileSystem/FileSystem_FileOperations.md`; `Specs/Core/Core_FileSystemBridge.md`; `Specs/Plugins/Plugins_VirtualFileSystem.md`; `Specs/Testing/Testing_SelfTests.md` | Adds 96-level extended-path provider coverage, wide/forced-limit bridge coverage, merge-metadata preservation, quantitative source guards, and the recursive-depth safety contract. |
| **Diagnostic evidence retained** | `Specs/TestRuns/4cb089111a23/Continuation/20260824_000200_fileops_p24_recursive_depth_crash` | The rejected 256-level recursive proposal failed at 192 levels on the default Debug stack. The archive preserves runner/WER evidence and the external dump pointer/hash; the executable contract is 128 until traversal is iterative. |
| **Closed by P2.5** | The previously inventoried `preCalc*`/`Calculating` executor/popup fields and Phase 05–09 fixtures | The separate runtime pass/state is gone. Single-pass discovery-ahead reports cumulative totals from the mutation traversal and the UI remains indeterminate until closure. |

#### P2.5 impact refresh — 2026-08-24

The discovery/progress scan was rerun across the ABI, host executor and popup, Local/Dummy providers,
File Operations and Commands selftests, resources/satellites, settings state, and current
File Operations/VFS/UI/testing specs. Current matches are classified as follows:

| Classification | Exact current matches | Result / next owner |
|---|---|---|
| **Update — P2.5 complete** | `Common/PlugInterfaces/FileSystem.h`; `Common/SettingsStore.h`; `Plugins/FileSystem/FileSystem.FileOps.cpp`; `FolderWindow.FileOperations.State*.cpp`; `FolderWindow.FileOperations.Popup.*`; `FolderWindow.FileOperationsInternal.h`; File Operations resources and four satellites | One traversal reports discovery and feeds bounded mutation. The host owns exact per-item cumulative accounting, open/closed totals, one-way Skip, immediate reservation release, and indeterminate-until-closed presentation. Local permanent Delete uses the same no-follow walk with one 256-child batch and bounded terminal-failure retention; the prior 200,000-entry flattening path is removed. No persisted or runtime pre-calculation control/state remains. |
| **Test/spec update — P2.5 complete** | File Operations Fairstream, Phase 05–10 and shared fixtures; Commands File Operations/plugin/settings source contracts; `Specs/FileSystem/FileSystem_FileOperations.md`; `Specs/Plugins/Plugins_VirtualFileSystem.md`; `Specs/UI/UI_FileOperationsPopup.md`; `Specs/Testing/Testing_{SelfTests,PerformanceValidation}.md`; Operation perf WIP | Replaced obsolete pre-calc cases/wording with single-traversal accounting, mutation-before-close, Skip-continuation/release, cancellation, scheduler, popup, source-contract, and archived metric evidence. |
| **Reviewed legacy API, not discovery authority** | Local/Dummy directory-size APIs and their Phase 10–13 file-root tests; folder-size `IDS_STATUS_CALCULATING_SIZE` resources | These serve explicit folder-size/status consumers. They do not run before File Operations mutation, populate File Operations totals, or restore a `Calculating` task state. |
| **Reviewed Phase 3+** | Follow/link-policy code; Delete/Recycle conflict actions; typed-result/verification/provider completion work | P2.6 removed the remaining Local provider-owned copy/delete Move fallback and closed the Phase 2 destructive-strategy gate. Later phases own links, conflicts, results, verification, and provider-specific completion without reopening a totals pre-pass. |

The active citation scan finds `FO-D` only in this file's §1.1 crosswalk, current explanatory text,
and the scan instructions themselves; no `REVIEW REQUIRED` banner remains. No newly discovered
production ingress or unclassified P2.4 authority remains.

#### P2.6 impact refresh — 2026-08-24

The destructive-strategy and cleanup scan was rerun across every shipped capability-v2 document,
provider Move entry point, the host selector/executor, File Operations tests, Commands source
contracts, and current File Operations/VFS/testing specs. Current matches are classified as follows:

| Classification | Exact current matches | Result / next owner |
|---|---|---|
| **Update — P2.6/P2.7** | Local `Plugins/FileSystem/FileSystem.FileOps.cpp`; host `FolderWindow.FileOperations.State.cpp`; File Operations Fairstream and Phase 05–13 fixtures; Commands provider/engine source contracts | Local provider Move is one native rename only and has no copy/delete or per-child merge fallback. Bound Managed cleanup alone may delete after publication. A regular destination directory that races into existence is freshly requalified under the same task; directory membership change after discovery becomes automatic `Copied; source kept`, and successful nested retention propagates to the task result. |
| **Executable native-strategy proof** | Local, Dummy, Microsoft Drive, and S3 flat-prefix profiles advertising `nativeMove: true`; `FileOps_ProviderCapabilityMatrix` plus provider debug/source contracts | These are the exact shipped `nativeMove: true` profiles. Each accepts NativeOnly and proves a provider-native object relocation/conditional implementation; all other profiles remain native false and therefore Managed, Copy-only, or unsupported from pair facts. |
| **Compatibility publication, never destructive authority** | Legacy atomic-final writer verifier and `Cinderstar_LegacyWriter*` fixtures | Bound Local source to unbound Dummy destination is Copy-only. Bounded final-path revalidation/cancel remains tested, but every terminal outcome retains source and no Managed cleanup is armed. |
| **Reviewed Phase 3** | Follow/link-policy code and semantic Preserve/retarget fixtures | The final semantic contract remains normative. `Fairstream_MovedTreeRetargetsInternalLinks` is explicitly deferred to P3.2 because Phase 2 does not yet preserve/retarget a link object on the host Managed route. |
| **Reviewed historical/dependent** | `Specs/Plans/Done/**`, `Specs/Reviews/**`, `Specs/TestRuns/**`, and unrelated active WIPs containing retired symbol names | Historical evidence is not current authority. Future owning packages must refresh active dependent WIPs when their implementation symbols are touched. |

#### P2.7 impact refresh — 2026-08-24

The Native-shape, strategy, traversal-policy, and scheduler scan was rerun across the Local provider,
host admission/executor, shared Common helpers, File Operations and Commands tests, project manifests,
and current File Operations/bridge/VFS/shared-helper specs. Current matches are classified as follows:

| Classification | Exact current matches | Result / next owner |
|---|---|---|
| **Update — P2.7 complete** | Local `Plugins/FileSystem/FileSystem.FileOps.cpp`; host `FolderWindow.FileOperations.cpp`, `FolderWindow.FileOperations.State.cpp`, and typed plan declarations | Local Native is one provider mutation. Regular directory merge and a raced regular-directory destination requalify to Managed under the same task; destination links remain typed conflicts; case-only directory rename is not misclassified as merge. Ordinary Copy uses the explicit `Copy` strategy. |
| **Shared bounded policy — P2.7 complete** | `Common/FileOperationTraversalPolicy.h`; Local and host project manifests/consumers; `Specs/Core/Core_SharedHelpers.md` | One dependency-neutral owner defines traversal ceilings plus overflow-safe discovery queue/worker calculations. Provider and host scheduling loops remain local because their ownership and work-item types differ. |
| **Test/spec update — P2.7 complete** | File Operations merge/race/junction/discovery fixtures; Commands source contracts; `Specs/FileSystem/FileSystem_FileOperations.md`; `Specs/Core/Core_FileSystemBridge.md`; `Specs/Plugins/Plugins_VirtualFileSystem.md` | Deterministic destination-appeared, destination-link, mixed-selection, case-only directory rename, explicit strategy, terminal classification, and shared-policy evidence is executable and archived. |
| **Reviewed compatibility input** | Local `FILESYSTEM_MOVE_DEFAULT` handling | Default/omitted Local Move options remain an explicit Native-only compatibility input. They cannot restore provider copy/delete or per-child rename merge. |
| **Retired by P3.2c** | `nativeDirectoryRaceFallback == CopyOnly` validation/executor arm | A directory merge fallback must preserve Move semantics through exact Managed cleanup. Copy-only is never a raced-directory merge fallback; validation rejects it. |
| **Reviewed historical/dependent** | `Specs/Plans/Done/**`, `Specs/Reviews/**`, `Specs/TestRuns/**`, and unrelated active WIPs | Historical evidence remains non-authoritative. Current source/spec scans find no live provider child-merge symbol or `copy.managed.*` strategy name. |

#### P3.1 impact refresh — 2026-08-24

The link-policy scan was rerun across the public VFS ABI, every shipped provider operation entry
point, Local configuration/schema/runtime, host plan admission/execution, resources and satellites,
File Operations and Commands selftests, and current File Operations/VFS specs. Current matches are
classified as follows:

| Classification | Exact current matches | Result / next owner |
|---|---|---|
| **Update — P3.1 complete** | `Common/PlugInterfaces/FileSystem.h`; Local `FileSystem.{h,cpp,FileOps.cpp}`; host `FolderWindow.FileOperations{,Internal,State,State.Private}.h/.cpp`; main resources and cs/fr/ja/sk satellites | Only Preserve/Skip are executable. Numeric zero remains reserved and fails validation. The host snapshots immutable task intent and the Local provider honors the per-call value over its configured default. No Follow UI, runtime branch, session grant, or warning resource remains. |
| **Provider ABI boundary sweep — P3.1 complete** | Local, Curl, Dummy, Microsoft Drive, MTP, and S3 mutation entry points; Curl shared progress helper | Every provider that consumes non-null options validates the common complete header, including link policy, before I/O. 7z and Google Drive remain unsupported/non-consuming for these entry points and do not form a hidden policy path. |
| **Test/spec update — P3.1 complete** | Commands plugin-config/source-contract tests; File Operations shared, Fairstream, Phase 05–13 fixtures; `Specs/FileSystem/FileSystem_FileOperations.md`; `Specs/Plugins/Plugins_VirtualFileSystem.md` | Canonical schema/round-trip, silent legacy migration, invalid ABI values, immutable propagation, per-call precedence, Skip target safety, Native link relocation, and bridge reparse behavior are executable. The governed archive retains the complete per-case authority and compact claim-bearing metrics. |
| **Intentional compatibility text** | Legacy parser aliases and their migration/source-contract tests; explicit historical/migration clauses in current specs and this WIP | `followTargets` and `copyReparse` may appear only as accepted legacy input or text proving their retirement. They are never emitted, presented, or executed as policy values. |
| **Delivered — P3.2a** | `Fairstream_MovedTreeRetargetsInternalLinks`; public bound-link payload/stage ABI; Local exact no-follow link methods; host Managed bridge link publication | Preserve now copies the exact link object without following its target, applies selected-root transformation, conditionally publishes an identity-owned stage, and only then exact-deletes the retained source. |
| **Delivered — P3.2b** | Host Keep Both mapping feed; bounded forward-target dependency queue; failed-prefix and held-cleanup records; provider-relative dependency payload | Actual child conflict choices now feed the semantic transform. Forward references and cycles defer under the shared traversal ceilings; failed dependencies keep the exact source and ancestors. No whole-tree manifest or target-object traversal is introduced. |
| **Delivered — P3.2c** | Native semantic qualification; complete link-kind/profile matrix | Native tree/link Move requires explicit path-profile semantic-transform proof. Local lacks that proof and routes trees/link objects to Managed while regular files remain Native; the executable link-kind/profile matrix and compact governed archive close the package. |

#### P3.2a impact refresh — 2026-08-24

The semantic-link scan was rerun across the public binding ABI, Local no-follow binding and reparse
implementation, host bridge copy/publication/cleanup paths, File Operations fixtures, provider
capability parsing, and current File Operations/bridge/VFS specifications. Current matches are
classified as follows:

| Classification | Exact current matches | Result / next owner |
|---|---|---|
| **Update — P3.2a complete** | `Common/PlugInterfaces/FileSystem.h`; Local `FileSystem.{h,cpp,FileOps.cpp,Internal.h}`; host `FolderWindow.FileOperations{,State}.cpp` | The host receives a bounded semantic payload from the exact retained source object, creates an exclusive owned link stage, conditionally publishes it, validates the published identity/kind, and only then permits exact cleanup. No target object is opened, followed, or mutated. |
| **Test/spec update — P3.2a complete** | Commands source-contract guard; Local provider debug contract; `Fairstream_MovedTreeRetargetsInternalLinks`; `Causeway_BridgeFileReparsePolicy`; `Phase12_ReparsePointPolicy`; authoritative File Operations/bridge/VFS specs | The formerly skipped moved-tree case is executable. Typed unsupported-provider behavior remains intact, and the compact governed archive binds the complete 117/117 behavioral result plus claim-bearing link metrics. |
| **Reviewed hardening boundary — P3.4** | Local directory-link stage creation before opening its owned no-follow handle | Exclusive name creation rejects collisions, and all subsequent publication/cleanup is identity-owned. P3.4 owns a stronger atomic directory-object create-and-bind primitive and the adversarial name-swap proof; P3.2a does not claim that later hardening gate. |
| **Delivered — P3.2b** | Host conflict outcomes and `FileSystemLinkComponentMapping` population; forward/back/cycle fixtures | Actual Keep Both outcomes now populate the bounded sparse map; provider-order backward references publish immediately, forward references/cycles defer, and a failed dependency retains its exact source. See the P3.2b impact refresh below. |
| **Delivered — P3.2c** | Native Move admission/qualification; file-symlink, directory-symlink, junction, relative/absolute and provider-profile matrix | The explicit `links.nativeMoveSemanticTransform` proof gates Native tree/link admission. Local routes those shapes to Managed; profiles unable to expose links may prove the requirement vacuously. The executable matrix is archived with the package. |

#### P3.2b impact refresh — 2026-08-24

The sparse-retarget scan was rerun across the public binding ABI, shared provider-path helpers, Local
no-follow link transformation, host bridge conflict/traversal/publication/cleanup paths, File
Operations and Commands fixtures, performance instrumentation, and current File Operations,
bridge, VFS, shared-helper, and testing specifications. Current matches are classified as follows:

| Classification | Exact current matches | Result / next owner |
|---|---|---|
| **Update — P3.2b complete** | `Common/PlugInterfaces/FileSystem.h`; `Common/FileSystemPathIdentity.{h,cpp}`; Local `FileSystem.FileOps.cpp`; host `FolderWindow.FileOperations{Internal.h,State.cpp}` | The provider reports the exact bounded source-relative dependency without following its target. The host retains only sparse actual renames, active directory-buffer views, deferred exact links, failed prefixes, and held exact cleanup records under the shared traversal limits. |
| **Test/spec update — P3.2b complete** | Commands path-identity/source-contract cases; Local provider debug contract; `Fairstream_MovedTreeRetargetsInternalLinks`; `Phase9_ConflictPrompt_KeepBothNestedCacheEligibility`; authoritative File Operations/bridge/VFS/shared-helper/testing specs | Drive/slash-root relative derivation, backward/forward/cycle resolution, nested Keep Both cache isolation, failed-target source retention, and bounded-state metrics are executable. The compact governed archive binds the complete 117/117 result and claim-bearing rows. |
| **Reviewed hardening boundary — P3.4** | Local directory-link stage creation before opening its owned no-follow handle | Exclusive name creation rejects collisions, but P3.4 still owns the stronger atomic directory-object create-and-bind primitive and adversarial name-swap proof. |
| **Delivered — P3.2c** | Native Move admission/qualification; file-symlink, directory-symlink, junction, relative/absolute/root-relative/outside-root and provider-profile matrix | Exact path/profile proof gates Native tree/link admission; Local uses Managed and unsupported destinations remain Copy-only or typed unsupported. Complete matrix evidence is archived with the package. |

#### P3.2c impact refresh — 2026-08-24

The Native-semantic qualification scan was rerun across capability-v2 producers and parsers, Local
no-follow link capture, host plan validation/admission/execution, provider contract tests, and File
Operations semantic-link fixtures.

| Classification | Exact current matches | Result / next owner |
|---|---|---|
| **Update — P3.2c complete** | `links.nativeMoveSemanticTransform` in every shipped capability profile and strict parser; Local shape qualification in `FolderWindow.FileOperations.cpp`; Managed semantic-link bridge in `FolderWindow.FileOperations.State.cpp` | Native admission for a tree or top-level link requires executable provider proof of the approved semantic transform. Local advertises false, so its regular files remain Native while trees and link objects enter Managed publication and exact cleanup. |
| **Test/spec update — P3.2c complete** | Local provider binding/link debug contract; File Operations capability matrix, same-volume semantic route, moved-tree, Copy-only, unsupported-provider, and Phase 12 cases; Commands source guard; authoritative File Operations/bridge/VFS/testing specs | File/directory symlinks, junctions, relative/absolute/root-relative/outside-root targets, Native/Managed/Copy-only, and unsupported profiles are executable. `ACCESS_DENIED` cannot trigger Native merge requalification, and Copy-only is rejected as a Native directory-race fallback. |
| **Delivered — P3.3** | Public issue actions; host conflict metadata/scope/cache; popup action layout; Local recursive Copy/link deferral; same-pane and nested Keep Both | Typed file/read-only/type/link/name/target/retry buckets now own the action set. Plain Skip is one-shot, Skip all and Apply-to-all use the exact scope, folder merge survives, and Local child Keep Both decisions feed the bounded semantic Copy transform before deferred links publish. |
| **Test/spec update — P3.3 complete** | Phase 9 typed-conflict matrix; Fairstream/Riptide/Causeway link and nested-collision cases; Local direct-provider reparse compatibility; authoritative File Operations/bridge/VFS/popup/testing specs | Existing-folder merge, exact destination junction classification, non-destructive type mismatch, same-pane duplicate naming, same-folder Move rejection, child-level Keep Both, forward-link retarget, and truthful Skip/Partial reporting are executable. The compact governed archive binds all 117 passing cases and the claim-bearing metrics. |
| **Delivered — P3.4** | Prompt-to-mutation destination replacement, exact link replacement, metadata writes, adversarial swaps, and Phase 3 Release gate | P3.4 retains and revalidates the exact destination object across consent, performs Replace link and metadata mutation through bound no-follow authority, and fails or re-prompts on removal/replacement without mutating the substitute. Phase 3 closeout evidence is recorded above. |

#### P4.7 impact refresh — 2026-08-25

The closeout scan was rerun across host conflict publication/rendering, Local provider issue-action
switches, bridge publication/verification result reduction, confirmation resources and satellites,
Phase 10/Commands fixtures, Pester source contracts, and the authoritative File Operations, bridge,
and popup specifications.

| Classification | Exact current matches | Result / next owner |
|---|---|---|
| **Update — stable attention and Recycle escalation** | `FolderWindow.FileOperations{Internal.h,State.cpp,Popup.h,Popup.cpp}`; source/satellite resources | Metadata loading is visible but non-actionable. Recycle alone renders Cancel first and binds it as both default and cancel control; cached or stale submissions cannot bypass metadata completion. |
| **Update — exact publication and directory verification** | `FolderWindow.FileOperations.State.cpp`; Phase 10 verification/metadata cases; File Operations and bridge specs | The bridge's exact publication state survives cancellation/failure. Nested file proof controls completion and source retention, while a selected directory's displayed verification axis remains `NotApplicable`. |
| **Update — no provider Recycle opt-out during Copy** | Local `FileSystem.FileOps.cpp`; Pester source-contract guard | Copy issue handling treats `PermanentDelete` as cancellation and cannot clear a shared Recycle flag or create a pathname-delete retry. Recycle escalation remains host-owned and identity-bound. |
| **Known later architecture work** | duplicated serial/parallel scheduling policy; hosted progress ownership; retained-source localization breadth | These are classified Phase 5/6 maintainability/presentation work. They do not reopen the Phase 4 destructive-safety or typed-result stop conditions. |

### 16.2 Clauses that must be replaced explicitly

Required explicit replacements in the authoritative text include:

- any silent identical-file FNV skip;
- any folder-level Exists/Overwrite wording that defeats merge;
- `FollowTargets`, session-sticky grant, and copy-reparse target-follow behavior;
- copy-reparse raw-text preservation that would break Move semantic retargeting;
- Permanent delete offered without the fresh Recycle escalation contract;
- F2 direct `RenameItem` as a behavioral exemption;
- whole-tree Move manifest, separate recursive pre-calculation, separate `Calculating` phase, or full filesystem preflight requirements;
- `preCalcEnabled`/`preCalcMaxWorkers` settings and totals/ETA language that presents an open traversal as complete;
- Preferences controls, runtime fields, popup status, resource strings, and selftests that exist only for the retired pre-calculation model;
- “clipboard survives source kept” wording;
- mandatory destination reread as Move delete authority;
- provider-wide capabilities where behavior varies by bound path/profile;
- artifact filename pattern as ownership;
- Create Directory text that permits non-durable S3 synthetic folders.
- provider ABI-v1 compatibility language; every in-tree provider cuts over to capability v2 atomically.

### 16.3 Stale File Operations notices and dependent WIP

Phase 0 must replace these known stale notices/rows explicitly; deleting a banner without updating its owned normative clause or dependent row is not closure:

| Owner and current location | Stale claim | Required replacement |
|---|---|---|
| `Specs/UI/UI_FolderView.md:3,498` | F2 confirmation is open; current F2 executes directly/silently; a Move clipboard clears only after verified success | F2's inline editor is the only initial confirmation. Replace direct UI-thread rename with the §7.3 worker `RenamePlan(InlineRename)`: immediate admission outside global Queue, interlock safety, immediate non-clean reveal, 500-ms clean-running reveal, and silent clean success with identity focus retention. The exact cut payload clears after accepted queue admission and is never restored (§12.4). |
| `Specs/Plans/WIP/Operation_Atlas_RemainingSpecificationDecisions_2026-08-04.md` | Former `ATLAS-DEC-CLIP-01`/`ATLAS-DEC-BR-01` open rows | **Replaced in Phase 0:** both rows are now accepted decisions. §12.4 owns cut-payload consumption after accepted queue admission; the central `RenamePlan` contract owns Batch Rename. |
| `Specs/Plans/WIP/Product_WhimFilesGapAnalysisAndImprovementPlan_2026-07-08.md:3,172-215` | G2 depends on FO-D013 as a recovery-journal decision and may invert operations from pathname sets | Keep G2 deferred, but bind it to §12 exact result/identity axes and §13 proven recovery records. An inverse Copy/Delete may mutate only an object proven created by the original task; FO-D013 maps to bounded streaming, not a recovery-journal product decision. |
| `Specs/FileSystem/FileSystem_Mtp.md:3` | MTP is blanket Copy-only unless it gains persistent ID plus exact delete | Replace with the §5.3 path-profile matrix: a tested same-device, leaf-preserving WPD Move may be native; file transfer-copy fallback is Copy-only until bound exact cleanup exists; unsupported directory/non-leaf cases remain rejected. |

Any line numbers above are discovery anchors, not durable identifiers. Phase 0 relocates the row by matching the quoted claim when the owner has drifted.

## 17. Verification matrix

### 17.1 Deterministic correctness cases

At minimum cover:

1. same-object aliases: direct, hard link, case, 8.3, SUBST, mapped/UNC, junction/mount, cross-profile provider alias;
2. destination ancestor and final-link swaps at every classification/mutation boundary;
3. stage-name collision, replacement after create, promotion failure, cleanup failure, crash after publication;
4. folder merge with file collisions, type mismatch, read-only, links, Keep Both, hidden/system children;
5. Copy and Move of file symlink, directory symlink, junction, external relative target, root-mapped absolute target, unchanged internal relative target, in-tree forward/back target, nested/top-level Keep Both retarget, deferred-link bound and target-conflict retention;
6. same-pane `Ctrl+C`/`Ctrl+V`, same-folder Move reject, subtree discovery during streaming, discovery-ahead default, Skip-discovery one-way switch, immediate reservation release, and mandatory just-in-time checks after the switch;
7. clipboard sequence changed before accepted queue, queue failure, cancel, Copy-only, partial, indeterminate;
8. managed Move source swap, destination swap, sharing violation, conditional-delete conflict, directory non-empty cleanup;
9. verification off/on, BLAKE3/provider proof, zero-byte, mismatch, capability timeout/unavailable, cancel, native Move NotApplicable, progress segments, graph, and ETA before/after traversal closes;
10. MOTW/ADS/EA loss, ACL inheritance, EFS refusal/grant, sparse inflation, hydration, space drift/failure;
11. Recycle success, unavailable, failure, escalation, identity change, indeterminate outcome;
12. F2 and Batch Rename native/managed/cycles, crash after each step, Resume/Roll back/replacement object; inline-F2 worker-only execution, global-Queue bypass, overlapping-root wait with immediate card, immediate conflict/failure/partial/canceled/indeterminate reveal, fake-clock 499-ms silent running versus 500-ms reveal, silent clean success with no card/popup activation and identity focus retention, visible-late success following auto-dismiss, no legacy `ReportError`, and one-step Batch Rename retaining ordinary queue/card behavior;
13. Proven claim-plus-identity match/mismatch, Possible legacy/current names, user-created matching
    name, no artifact-specific hiding, exact-set touch warning/revalidation, and no automatic cleanup;
14. flat S3 marker create/refresh/copy/move and native directory-bucket behavior;
15. null-success outputs and capability-advertised operation execution for every production provider mode;
16. cancel/pause/exit during enumeration, Read, Write, Commit, verify, metadata, source delete, and unanswered gate;
17. queue/depth/path/memory limits with partial prior success and no unbounded retained manifest.
18. Pack/unpack cleanup submits an exact `DeletePlan`; identity replacement retains the source and no private pathname delete executes.

Every race test uses test-owned roots and deterministic injection. Never run destructive tests against ambiguous live cloud/device paths.

### 17.2 Performance cases

- same-volume native Move with verification off/on (zero verify reads);
- local cross-volume Copy/Move, many-small and large files;
- discovery-ahead versus Skip-discovery on local and high-latency providers: first-mutation latency, discovery starvation, queue depth, transfer ramp/full-speed release, and ETA knowledge state;
- high-latency provider, Commit delay, conditional delete, reconnect/indeterminate;
- depth/width/path-byte high-water;
- MTP streaming writer/reader memory;
- S3 tiny and multipart reservation;
- collision index at 65,536 long names;
- bridge concurrency 1/4/16, pause/resume, bandwidth changes;
- overlapping versus disjoint roots;
- artifact/recovery journal update cost and Batch Rename large selection;
- FBP-10 thread/buffer utilization before deciding a change.

Focused iteration may use affected/filtered suites. Final authority requires:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh
.\Tools\Test-TestRunArchive.ps1 -Inventory
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
git diff --check
```

Archived test-enabled Release evidence belongs under `Specs/TestRuns/` and must pass the repository performance-quality contract.

## 18. Finding closure ledger

The removed I11 file contained 34 findings. They remain traceable here:

| Findings | Owning clauses/phases |
|---|---|
| `FBC-01` destination reparse merge; `FBC-05` top-level containment; `FBC-06` stale traversal | §§3, 6.1; Phase 1 |
| `FBC-02` path-based source delete; `FBC-04` detached native proof; `FBC-10` FNV destructive proof; `FBC-14` cleanup eligibility | §§5.3, 6.3–6.4; Phase 2 |
| `FBC-03` unowned staging collision | §§3, 6.2, 13; Phase 1 |
| `FBC-07` destructive directory-link replace; `FBC-08` followed metadata; `FBC-09` unbound reparse cleanup | §§8–10; Phase 3 |
| `FBC-11` cross-context aliases | §§4.1–4.3, 7.2; Phase 1 |
| `FBC-12` silent metadata loss | §9.2; Phase 4 |
| `FBC-13` unbounded/cyclic traversal | §§6.1, 8; Phases 2–3; Follow removed |
| `FBC-15` uninterruptible providers | §11; Phase 6 |
| `FBC-16` null-success outputs | §5.1; Phase 1 |
| `FBC-17` path-raced attributes | §§3, 10; Phases 1/3 |
| `FBC-18` sequential late validation | §4.3; envelope validation retained, whole-filesystem preflight rejected |
| `FBC-19` lost aggregate `S_FALSE` | §12.2; Phase 4 |
| `FBC-20` retry record removed too early | §§6.4, 12–13; Phase 2 |
| `FBC-21` device namespaces | §4.3; Phase 1 |
| `FBP-01` MTP whole payload; `FBP-06` per-read allocation | §14.2; Phase 6 |
| `FBP-02` recursive queue; `FBP-03` Move manifest; `FBP-05` cleanup reread | §§6.1, 6.3–6.4; Phase 2 |
| `FBP-04` premature metrics | §14.1; Phase 6 |
| `FBP-07` S3 tiny reservation; `FBP-08` name indexes; `FBP-09` eager DACL; `FBP-10` thread/buffer cost | §14.2; Phase 6 |
| `FBCON-01` missing executable provider matrix | §§5.2–5.3, 17; Phases 0/6 |
| `FBCON-02` inaccurate normative contract | §16; Phase 0 |
| `FBCON-03` incomplete Release evidence | §§14, 17; Phase 6 |

No I11 finding is deferred to an unnamed owner.

## 19. Done criteria and stop conditions

Done means all of the following are true:

- [x] Every normative owner in §16 contains the durable rule and no stale REVIEW notice contradicts it.
- [x] The §16.1 impact manifest classifies every current mutation/capability/result/settings/resource/test match; no discovered caller is silently outside scope.
- [x] Every production ingress creates the typed immutable plan and retains qualified endpoints.
- [x] Same-object, containment, staging, link, replacement, and source cleanup races are fail-closed and identity-bound.
- [x] Folder merge, conflicts, same-pane duplicate, discovery modes, clipboard clearing/UX, Rename, Recycle, and result consumers match this spec.
- [x] Provider capability v2 and optional binding/cancel interfaces execute truthfully for every production mode.
- [x] Recursive work, MTP, S3, collision indexes, proof state, and recovery records pass quantitative memory gates.
- [x] Every item emits final typed axes and final-phase telemetry; no FNV auto-skip/delete authority remains.
- [x] Deterministic Debug and test-enabled Release cases are green; performance evidence is archived and quality-valid.
- [x] Fresh Full, archive inventory, spec inventory, and `git diff --check` are green.
- [x] This plan is moved to `Specs/Plans/Done/` and the WIP index is reconciled.

Stop and request a new product/architecture decision, citing the exact code/spec conflict, if:

- product wants destructive Move from a provider that cannot bind the read source and conditionally delete it;
- a platform cannot preserve link semantics and product does not accept Skip/source retention;
- product wants silent loss of MOTW/ADS/EA or silent EFS plaintext on Move;
- an ABI change cannot land atomically across host and all shipped providers;
- safe replacement would require deleting an unowned object or following a destination link;
- a proposed optimization violates a measured bound or removes a proof without an equal/stronger replacement;
- a new user-visible behavior is missing from this specification.

Otherwise, Phase 0 migration is authorized. Production implementation is authorized in the remaining phase order only after Phase 0 has made the normative corpus coherent and its stop condition is clear.
