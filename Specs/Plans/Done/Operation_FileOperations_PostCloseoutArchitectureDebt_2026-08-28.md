# File Operations post-closeout architecture and validation debt

> **NON-NORMATIVE RETIRED WORKING-PLAN RECORD.** Moved from `Specs/Plans/WIP/`
> on 2026-08-29 after every unfinished A0-A8 requirement, invariant,
> test/performance obligation, done criterion, and STOP condition was transferred
> to `Specs/Plans/WIP/Operation_FileOperations_ReliabilityFirstSuccessorPlan_2026-08-29.md`,
> section 8 E0-E8. **No remaining implementation, validation, or decision action
> belongs to this file.**
>
> The successor is the sole live owner. Its E0-E8 lane retains the prior
> behavior-preserving authorization; its D2-A01-D2-A20 and R0a-R9 lanes remain
> product-arbitration-only until explicitly activated. Retirement approves none of
> that gated scope.
>
> **Frozen history: do not resume or edit.** Unchecked/active markers below preserve
> state at transfer; they do not form a live queue and do not claim completion.
> Durable behavior remains owned by authoritative specifications under
> `Specs/<Domain>/`. The completed Phase 0-6 program remains frozen at
> `Specs/Plans/Done/Operation_FileOperations_GlobalBehaviorDecisionReview_2026-08-16.md`.

## Status

| Field | Value |
|---|---|
| State | **RETIRED — all unfinished A0-A8 scope transferred; no action remains here** |
| Index owner | **None — historical I14 identity retained for retrieval** |
| Priority | **P1** current-HEAD validation and behavior-preserving ownership extraction; **P2** remaining rename/admission/index cleanup |
| Planned at | `6094a966dc0450f75a93341e8364035212610491` (`6094a966d`) |
| Updated | 2026-08-29 |
| Scope | File Operations post-closeout validation, maintainability, ownership consolidation, and bounded responsiveness debt |
| Product behavior | Preserve the completed File Operations contract unless a package explicitly updates an authoritative spec before code |
| Drift check | `git diff 6094a966dc0450f75a93341e8364035212610491 -- RedSalamander/FolderWindow.FileOperations.State.cpp RedSalamander/FolderWindow.FileOperations.Popup.cpp RedSalamander/FolderWindow.FileOperations.cpp RedSalamander/FolderWindow.FileOperations.State.Runtime.cpp RedSalamander/ChangeCase.cpp RedSalamander/FileSystemRenameBatch.cpp RedSalamander/BatchRenameRecoveryJournal.cpp RedSalamander/FileOperationArtifactRegistry.cpp RedSalamander/FileOperationMoveBreadcrumb.cpp RedSalamander/FolderView.Enumeration.cpp RedSalamander/FindFilesWindow.cpp Plugins/FileSystem/FileSystem.FileOps.cpp Specs/FileSystem/FileSystem_FileOperations.md Specs/UI/UI_FileOperationsPopup.md Specs/UI/UI_BatchRenameWindow.md Specs/UI/UI_CommandMenuKeyboard.md Specs/Core/Core_SharedHelpers.md Specs/Testing/Testing_PerformanceValidation.md` |

## Progress checklist

**Historical transfer snapshot only.** `[x]` meant implemented, validated,
committed, and documented; `[+]` meant active when this file was retired. Open
markers intentionally remain truthful. Their complete executable equivalents are
E0-E8 in the sole live successor; do not act from this snapshot.

- [+] **A0. Establish a current-HEAD validation baseline**
  - [ ] Build the merge result at `6094a966d` or its drift-reviewed successor.
  - [ ] Re-run the 96-level recursive Local Copy stress that previously exposed stack/lifetime failure.
  - [ ] Re-run focused File Operations and popup/Commands cases affected by the merge.
  - [ ] Run and archive a current-HEAD Fresh Full; the earlier Phase 6 archive predates later teardown, progress, and merge commits.
  - [ ] Validate test and spec inventories and record the exact run IDs here.
- [ ] **A1. Extract the bridge and per-item policy without changing behavior**
  - [ ] Characterize the existing bridge and `processIndex` behavior before moving code.
  - [ ] Extract `CrossFileSystemBridge` from the body of `Task::ExecuteOperation` into a named implementation boundary.
  - [ ] Extract the per-item retry/conflict/terminal policy from the `processIndex` lambda into a named type or functions.
  - [ ] Keep serial versus parallel scheduling separate from the single item policy.
  - [ ] Prove byte/result/conflict parity and no throughput or memory regression.
- [ ] **A2. Make the engine snapshot the only conflict-action policy owner**
  - [ ] Store the final ordered action set, primary/More placement, default action, and Cancel/Escape action in the prompt snapshot.
  - [ ] Remove popup-side bucket-to-action derivation; the popup renders the snapshot only.
  - [ ] Cover every bucket, especially Recycle Cancel-first/default, destination links, Skip All, and destructive-action withholding.
- [ ] **A3. Share safe durable-file I/O while keeping schemas separate**
  - [ ] Inventory existing helpers before introducing a new one.
  - [ ] Extract only the bounded, no-follow, regular-file, atomic-persist, scan, acknowledge/remove, and rollback-on-persist-failure mechanics shared by the three stores.
  - [ ] Keep Batch Rename journal, artifact claims, and Move breadcrumb schemas and authority semantics distinct.
  - [ ] Add fault-injection parity tests for create, replace, malformed input, reparse rejection, size limits, and in-memory rollback.
- [ ] **A4. Retire Change Case as a second rename mutation engine**
  - [ ] Characterize current recursive ordering, batching, cancellation, progress, and artifact-guard behavior.
  - [ ] Route Change Case mutations through a typed `RenamePlan` and the central File Operations rename authority.
  - [ ] Preserve case-only rename handling, cycle handling, identity revalidation, and recursive artifact warnings.
  - [ ] Remove or explicitly quarantine `FileSystemRenameBatch` after all production mutation callers are gone.
- [ ] **A5. Remove selection-proportional Batch Rename admission work from the UI thread**
  - [ ] Measure current synchronous capability/bind cost for 1, 64, 1,024, and provider-delayed rows.
  - [ ] Split immutable request capture from worker-side capability/binding qualification.
  - [ ] Preserve reject-before-first-mutation, one identity domain per plan, duplicate-object rejection, cancellation, and teardown safety.
  - [ ] Define and specify the Preparing/admission-failure presentation before implementation if the user-visible state changes.
  - [ ] Archive responsiveness and retention evidence.
- [ ] **A6. Scope artifact candidate hints by endpoint and parent**
  - [ ] Replace the global leaf-name hint with an endpoint/identity-domain and parent-qualified candidate index.
  - [ ] Preserve always-visible Proven artifacts, Possible warnings, and no-follow identity proof.
  - [ ] Keep all hints non-authoritative: no hint may authorize mutation, recovery, or deletion.
  - [ ] Measure enumeration/Find cost for common leaf names and a large claim registry.
- [ ] **A7. Use one stable identity for interrupted-Move notice cards**
  - [ ] Remove the displayed durable task ID versus in-session `summary.taskId` ambiguity.
  - [ ] Preserve notice-only semantics: no Resume, no automatic cleanup, no source deletion, and Acknowledge only inside the breadcrumb store.
  - [ ] Cover restart projection, open-source/open-destination, dismissal, and multiple breadcrumbs.
- [ ] **A8. Close out this plan**
  - [ ] Update every affected authoritative domain spec and `Core_SharedHelpers.md` when applicable.
  - [ ] Run focused Debug and test-enabled Release validation after each package.
  - [ ] Run `git diff --check`, spec inventory, test inventory, and archive validation.
  - [ ] Run a final Fresh Full against the exact closeout commit.
  - [ ] Move this file to `Specs/Plans/Done/` and remove I14 from the WIP index.

## 1. Purpose

The completed File Operations program established the product contract: exact
identity authority, Native/Managed/Copy-only Move, no-follow links, folder
merge, typed conflicts and outcomes, identity-owned destructive actions,
streaming bounded discovery, immutable task intent, honest clipboard behavior,
verification, recovery artifacts, qualified Create Directory, and cancellation
checkpoints.

This follow-up does **not** redesign those behaviors. It prevents the remaining
implementation shape from becoming the next source of safety drift. The work is
split so each package can land, test, and revert independently.

The immediate prerequisite is validation, not refactoring. The current HEAD is
a resolved merge. The Local recursive-copy implementation retains heap-owned
traversal frames and `ClassifyLocalCopyPathKind`; there are no unresolved merge
markers at the planned-at commit. However, the last recorded Fresh Full predates
the later shutdown, progress, and merge commits. A0 establishes trustworthy
evidence before architecture work begins.

## 2. Ownership and coordination

### 2.1 This plan owns

- behavior-preserving extraction of the cross-filesystem bridge and per-item policy;
- one conflict-action policy owner;
- shared durable-store I/O mechanics, not shared schemas;
- Change Case migration to central typed rename authority;
- asynchronous/responsive Batch Rename admission qualification;
- artifact candidate-index scoping;
- interrupted-Move notice-card identity cleanup;
- current-HEAD validation needed before these extractions.

### 2.2 Existing owners remain authoritative

- **I3 — File Operations stress validation** owns its five already-defined
  behavior/evidence areas. I14 may consume or extend its archives but must not
  duplicate or silently close I3 rows.
- **I12 — Terminal/FileOps review follow-up** owns completion affinity, reaper
  lifetime, quiet-point, SwapPanes insertion identity, and removal-focus work.
  Any A1 change touching task shutdown, completion posting, or reaper ownership
  waits for I12 closeout or is coordinated in the same commit with its tests.
- **Retired I17 implementation baseline** — archived at
  `../Done/UI_FileOperationsPopupParallelProgressPendingPrompt_2026-08-28.md` —
  delivered Queue→Parallel dest-leaf admission, determinate vs marquee
  presentation, Waiting/Needs-attention/failed-card layout, UNC owned-stage
  identity, and failed-before-publication axes. I14 must not restore destination-
  folder-only `PublishDestination` when extracting interlock/bridge code, relitigate
  the shipped prompt baseline while extracting conflict-action policy, or treat raw
  provider `ERROR_IO_INCOMPLETE` as the unknown-outcome sentinel for a Move/Copy that
  never published. E8 in the sole live reliability-first successor owns the
  inherited exact merged qualification gate; this retired record does not resume I17.
- **I4 — Astrolabe** owns repository-wide replacement of behavioral source-shape
  tests. I14 may add focused runtime tests and retain narrow structural guards,
  but it must not create a second global test-migration queue.
- **I5 — Performance measurement contract** owns reusable repository-wide
  measurement infrastructure. I14 uses that contract and adds only File
  Operations-specific scenarios/counters.
- **I7 — HRESULT/status formatting** owns general status localization cleanup.
  I14 localizes only new or materially changed user-visible strings required by
  these packages.

### 2.3 Authority order

1. Public ABI headers and machine-consumed schemas/manifests.
2. Authoritative specs under `Specs/<Domain>/`.
3. Current implementation and executable tests.
4. This WIP plan.
5. Frozen historical plans and review text.

If current authoritative behavior conflicts with a proposed simplification,
stop that package and amend the owning domain spec only after an explicit
product decision. Do not use this plan to override settled behavior.

## 3. Current-state evidence and findings

Line numbers are anchors at `6094a966d`; executors must re-run the drift command
before editing.

| ID | Priority | Evidence | Risk / opportunity | Owning package |
|---|---:|---|---|---|
| FOA-01 | P1 | `FolderWindow.FileOperations.State.cpp` is about 17.6k lines; `CrossFileSystemBridge` is a local type near line 10798 inside `Task::ExecuteOperation`. | Copy, verification, links, metadata consent, publication, and cleanup are difficult to test or change independently. | A1 |
| FOA-02 | P1 | The shared per-item behavior is still a large `processIndex` lambda near line 16719. | One policy exists conceptually, but it has no named ownership boundary and remains coupled to scheduling/task state. | A1 |
| FOA-03 | P1 | `BuildConflictActionLayout` exists in the engine near `State.cpp:4359` and in the popup near `Popup.cpp:563`. | Recycle default/Escape, Skip All, and destructive-action eligibility can drift between execution and presentation. | A2 |
| FOA-04 | P2 | `BatchRenameRecoveryJournal.cpp`, `FileOperationArtifactRegistry.cpp`, and `FileOperationMoveBreadcrumb.cpp` separately implement bounded durable JSON-file lifecycle mechanics. | Reparse rejection, atomic persistence, malformed-file handling, and rollback rules can diverge. | A3 |
| FOA-05 | P1 | `ChangeCase.cpp:16` includes `FileSystemRenameBatch.h`; mutation batches execute through `FileSystemRenameBatch::Execute` near lines 472-473. | Change Case remains a production rename authority separate from `RenamePlan`, with weaker central conflict/result/journal integration. | A4 |
| FOA-06 | P2 | `AdmitBatchRename` loops capabilities and `BindObjectAuthority` synchronously near `FolderWindow.FileOperations.cpp:3390-3444`; the modeless window calls the host callback synchronously. | Large or delayed providers can block the UI before the task is admitted. | A5 |
| FOA-07 | P2 | `Registry::HasClaimLeafHint` is a global leaf set; FolderView consults it for every enumerated item near `FolderView.Enumeration.cpp:576-578`. | One common claimed leaf can trigger no-follow projection work in unrelated folders. Classification remains safe, but responsiveness can degrade. | A6 |
| FOA-08 | P2 | Restart projection allocates `summary.taskId = _nextTaskId++` while the message displays durable `record.taskId` near `State.Runtime.cpp:439-450`. | A single notice card can expose two task identities to logs, diagnostics, and actions. | A7 |
| FOA-09 | P1 validation | The last completed-program Fresh Full predates `fa2ca6a4e`, `a00b24dd1`, and merge `6094a966d`. | The resolved post-closeout tree lacks one exact Fresh Full receipt. | A0 |

### 3.1 Confirmed non-findings

Do not turn the following into work items without new evidence:

- The merge at `6094a966d` is resolved; the Local recursive-copy walker keeps
  heap-owned frames and the name-surrogate-aware kind classifier.
- Native and Managed Copy/Move are deliberate strategy boundaries, not
  accidental duplicate engines.
- Create Directory deliberately remains outside `StartOperation`; its qualified
  behavior is already owned by the current specs.
- Recycle bulk `DeleteItems` is an explicit shell-owned exception.
- Optional object binding is an explicit provider capability; destructive paths
  still fail closed when exact authority is required.
- Providers advertising `abort: false` or `deadline: false` are honest. Do not
  flip those flags until the provider transport can actually honor the contract.
- Move breadcrumbs are notices, not recovery journals and not deletion authority.

## 4. Architecture target

The target is smaller ownership, not a new universal framework.

```text
Task scheduling (serial / parallel)
        |
        v
Named per-item policy
  - strategy dispatch
  - retry/conflict flow
  - terminal/result mapping
        |
        +--> Native provider mutation
        |
        +--> Named cross-filesystem bridge
               - bounded discovery/streaming
               - publication/verification
               - semantic links/metadata
               - Managed cleanup authority

Engine conflict snapshot
  - ordered actions
  - primary / More placement
  - default and Cancel/Escape actions
        |
        v
Popup renderer (no bucket policy)

Schema-specific durable owners
  - Batch Rename recovery journal
  - artifact registry
  - Move breadcrumb notice
        |
        v
Shared app-local safe-file mechanics
  - bounded read/parse envelope
  - no-follow regular-file checks
  - atomic/exclusive persistence
  - failure rollback and safe acknowledgement
```

The design intentionally does not combine Native and Managed provider paths,
does not invent a general operation framework, and does not merge the three
durable schemas or their authority rules.

## 5. Package A0 — current-HEAD validation baseline

### 5.1 Required work

1. Confirm the worktree is free of conflict markers and that the recursive Local
   Copy implementation still uses heap-owned traversal frames and
   `ClassifyLocalCopyPathKind`.
2. Build Debug and test-enabled Release from the same commit.
3. Run the exact deep-recursion stress that previously failed at 96 directory
   levels, plus the popup/compact progress cases touched by the merge.
4. Run focused File Operations, Commands, provider contract, and archive
   validation appropriate to the changed merge surface.
5. Run Fresh Full with `-ValidationMode Fresh`; archive the exact run and record
   its commit, build receipts, counts, and zero-failure result.

### 5.2 Files and evidence

- `Plugins/FileSystem/FileSystem.FileOps.cpp`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.FileOps.cpp`
- relevant `RedSalamander/SelfTest/FileOperations/*`
- `Specs/TestRuns/<commit>/...` archive generated by the standard runner
- this plan's A0 checklist/evidence row only; no behavior spec change unless a
  failure reveals a real contract gap

### 5.3 Done criteria

- current planned-at or drift-reviewed successor builds Debug and Release;
- the deep recursive-copy scenario completes without access violation, timeout,
  leaked task, or stack growth proportional to tree depth;
- focused cases are green;
- a current-HEAD Fresh Full archive validates.

### 5.4 STOP conditions

- If A0 finds a correctness or lifetime regression, stop architecture extraction
  and create/fix the bounded defect first.
- Never resolve a regression by restoring stack-recursive traversal, weakening
  link classification, killing an unrelated running app, or reusing stale build
  receipts.

## 6. Package A1 — named bridge and item-policy boundaries

### 6.1 Design

Extract in two commits:

1. **Bridge extraction only.** Move `CrossFileSystemBridge` and its private helper
   state from the local function body into an app-local named implementation,
   proposed as `FolderWindow.FileOperations.Bridge.h/.cpp`. The initial move must
   preserve signatures, counters, cancellation polling, traversal limits,
   publication order, verification serialization, metadata consent, semantic
   link handling, and cleanup authority byte-for-byte where practical.
2. **Item-policy extraction only.** Replace the `std::function`/lambda ownership
   of `processIndex` with a named policy/context boundary, proposed as
   `FolderWindow.FileOperations.ItemPolicy.h/.cpp` or private `Task` methods.
   Scheduling may call the same policy serially or through the worker pool.

Prefer private `Task` methods when a proposed type would merely expose most of
`Task` as a back-reference. The goal is testable ownership and smaller functions,
not dependency injection for its own sake.

### 6.2 Invariants to preserve

- one conflict prompt per task;
- Copy-only never deletes the source;
- Managed cleanup uses the retained bound authority only;
- unknown mutation outcome is never retried as if non-commit were known;
- Native directory race requalification remains one-way and within the same task;
- Verify On does not overlap verification with the next transfer;
- traversal/resource limits remain task-terminal where specified;
- per-item publication and source-disposition axes remain truthful;
- task cancellation, completion posting, and shutdown lifetime remain I12-safe.

### 6.3 Tests and performance evidence

- existing File Operations focused suite, including Native/Managed/Copy-only,
  folder merge, links, Keep Both, Recycle, verification, and traversal limits;
- behavioral parity for concurrency budget 1 and greater than 1;
- cancellation at discovery, read, write, verification, prompt wait, and cleanup;
- counters proving no new copies, bridge re-creation, or retained whole-tree map;
- archived before/after throughput and retained-memory evidence for a large tree
  and many independent files.

### 6.4 STOP conditions

- Do not change product behavior in an extraction commit.
- Do not merge Native and Managed copy implementations.
- Do not expose owning raw COM pointers or weaken WIL ownership.
- Do not make bridge lifetime outlive the task or callback mutex it references.

## 7. Package A2 — one conflict-action layout owner

### 7.1 Design

The engine already decides which actions are allowed. Extend the immutable
conflict prompt snapshot so it carries the complete presentation contract:

- ordered action IDs;
- primary versus More grouping;
- default action;
- Cancel/Escape action;
- Apply-to-all and Skip All eligibility;
- metadata-loading state and whether buttons are publishable.

The popup must not switch on `ConflictBucket` to reconstruct those choices. It
may measure, wrap, and paint the supplied actions, but it cannot add, remove,
reorder, or choose a different default.

### 7.2 Files

- `RedSalamander/FolderWindow.FileOperationsInternal.h`
- `RedSalamander/FolderWindow.FileOperations.State.cpp`
- `RedSalamander/FolderWindow.FileOperations.Popup.cpp`
- `RedSalamander/SelfTest/Commands/Commands.SelfTest.FileOps.cpp`
- `Specs/FileSystem/FileSystem_FileOperations.md`
- `Specs/UI/UI_FileOperationsPopup.md`

### 7.3 Required matrix

Cover all typed buckets and at least these high-risk cases:

- regular/read-only Exists with destructive action both allowed and withheld;
- type mismatch;
- destination semantic link;
- TargetConflict;
- Retry-only transport/access buckets;
- Recycle escalation with Cancel first, default, and Escape;
- deferred consent Proceed/Retain Source cases;
- Skip versus Skip All and Apply-to-all scope;
- metadata-loading prompt with no decision buttons;
- cached action without a second unsafe action derivation.

### 7.4 Done criteria

- one production `BuildConflictActionLayout`/equivalent policy owner remains;
- popup tests assert rendering of the supplied snapshot, not duplicate bucket
  policy;
- keyboard, accessibility, default-button, and Escape behavior match the domain
  specs.

## 8. Package A3 — shared durable-file mechanics

### 8.1 Discovery first

Before adding code, search `Specs/Core/Core_SharedHelpers.md`, `Common/`,
`LocalFileTransaction`, and `Tests/TestSupport/`. Reuse or extend an existing
helper if its ownership and security semantics match. If the helper stays
RedSalamander-only, keep it app-local instead of pushing it into `Common`.

### 8.2 Shared responsibilities

The shared layer may own:

- configured byte/count limits;
- no-follow open and regular-file validation;
- exclusive create and atomic replacement mechanics;
- durable write error propagation;
- rollback of in-memory phase/state when persistence fails;
- safe directory enumeration with malformed/reparse rejection;
- acknowledgement/removal constrained to the configured store root.

### 8.3 Responsibilities that must remain schema-specific

- JSON fields, versioning, and parsing;
- Batch Rename intent/commit/reconciliation semantics;
- artifact claim identity and Proven/Possible classification;
- Move breadcrumb phase/count/pane projection;
- which files can be acknowledged, recovered, resumed, rolled back, or merely
  dismissed;
- user-facing projection and action capabilities.

Do not build a generic journal abstraction. A Move breadcrumb must never acquire
recovery or deletion authority by sharing I/O code with the rename journal.

### 8.4 Tests

Run each schema's existing lifecycle tests against the shared mechanics, plus
fault injection for:

- parent creation and exclusive create;
- persist failure before/after in-memory transition;
- atomic replace failure;
- reparse, directory, oversized, truncated, invalid UTF-8/JSON, and wrong version;
- partial directory scan with a rejected sibling;
- acknowledgement outside the store root;
- concurrent readers and one writer under the documented lock model.

Update `Specs/Core/Core_SharedHelpers.md` and each owning domain spec in the same
commit as a new shared helper.

## 9. Package A4 — Change Case through typed RenamePlan

### 9.1 Current behavior to characterize

`ChangeCase.cpp` currently discovers paths, sorts deepest-first, groups equal
depth, revalidates artifact guards in batches of 64, and calls
`FileSystemRenameBatch::Execute`. Capture tests for:

- file and directory case conversion;
- recursive deepest-first order;
- case-only local rename;
- siblings, parent/child selections, collisions, and cycles;
- cancellation during discovery and between batches;
- artifact warning, rejection, and revalidation before mutation;
- progress totals and partial failure reporting;
- unbound or non-rename-capable provider behavior.

### 9.2 Target

Discovery may remain command-specific. Mutation may not. Produce immutable
`RenameStep` mappings and submit them to the same central RenamePlan admission
and execution authority used by Batch Rename, with an explicit origin such as
`ChangeCase` if presentation or journaling differs.

Any new origin must define:

- queue/admission behavior;
- card/popup presentation;
- recovery journal participation;
- artifact guard ownership;
- conflict and Keep Both policy;
- completion refresh/focus behavior.

If these differ from Batch Rename, document the difference in
`UI_CommandMenuKeyboard.md` and `FileSystem_FileOperations.md`; do not bypass the
central identity and mutation gates.

### 9.3 Retirement rule

Delete `FileSystemRenameBatch` only when `rg` proves zero production mutation
callers. If a test-only or ABI-specific caller legitimately remains, rename and
document it as a narrow adapter rather than leaving a misleading shared engine.

## 10. Package A5 — responsive Batch Rename admission

### 10.1 Problem

The modeless Batch Rename window invokes the host callback synchronously. The
host then performs path-scoped capability queries and no-follow object binds for
every row before `StartOperation`. This is correct for identity-domain admission
but selection-proportional and provider-latency-sensitive.

### 10.2 Recommended design

1. On the UI thread, capture only immutable pane/provider identity, the proposed
   source-to-leaf mappings, and the completion/progress endpoints.
2. Transfer that request to a cancellable worker-owned admission stage.
3. On that worker, perform path capabilities, provider-parent derivation,
   no-follow binding, duplicate-object detection, identity-domain grouping, and
   dependency/cycle construction.
4. Publish the executable File Operations task only after qualification succeeds,
   or publish an explicit non-mutating **Preparing** card if product UX requires
   observable long admission. Choose and specify one model before code.
5. Post success/failure back through registered payload ownership; never capture a
   modeless HWND without the existing generation/drain discipline.

The simplest acceptable model is worker qualification followed by task publish;
the Batch Rename window remains in its executing/preparing state and receives a
single failure result if qualification fails. Do not create a second mutation
executor.

### 10.3 Required guarantees

- no mutation before all mappings are qualified;
- one executable identity domain per plan;
- duplicate hard-link/object inputs rejected deterministically;
- pane navigation after capture cannot retarget the plan;
- cancel/window close stops admission and drains posted payloads;
- provider unload observes the normal quiet point;
- no UI-thread capability or binding loop proportional to row count.

### 10.4 Performance scenario

Measure UI-thread time, wall time, retained bindings, and cancellation latency for
1, 64, and 1,024 mappings using Local and a deterministic delayed test provider.
The UI-thread capture budget must be effectively constant with row count; record
the chosen numeric guardrail in `Testing_PerformanceValidation.md` or the owning
Batch Rename spec after baseline evidence, not by guess in this plan.

## 11. Package A6 — scoped artifact candidate index

### 11.1 Design

Replace `_claimLeafHints` as a repository-wide leaf set with a bounded candidate
index keyed by at least:

- plugin ID;
- instance ID;
- path identity/profile or root ID where available;
- provider parent path under that identity contract;
- leaf name under the provider's case rules.

The index may answer only “worth probing.” Final projection still requires the
full claim path and no-follow object identity. If the parent cannot be represented
safely for a provider, fall back to the conservative current probe behavior for
that provider rather than hiding a claim-backed artifact.

### 11.2 Consumers

- `FolderView.Enumeration.cpp`
- `FindFilesWindow.cpp`
- Compare/search projection if it consumes registry hints after drift review
- `FileOperationArtifactRegistry.h/.cpp`

### 11.3 Tests and performance

- ordinary-looking Final artifact remains Proven in its exact parent;
- same leaf in an unrelated folder/endpoint does not cause a no-follow probe;
- Possible name-shape artifacts remain visible;
- claim identity mismatch remains Possible, never Proven;
- registry reload/update invalidates the candidate index safely;
- measure 10k ordinary rows named from a common corpus with 1, 100, and 10k
  claims; archive bind-count and enumeration/Find latency evidence.

## 12. Package A7 — interrupted-Move notice identity

### 12.1 Design

Keep the durable `record.taskId` as historical context, but allocate one explicit
session card identity and expose both only through named fields if both are
actually needed. User-visible text, commands, telemetry, and dismissal must not
ambiguously call both values “task ID.”

Recommended model:

- `summary.taskId` remains the in-session action/card key;
- add or retain a separately named `interruptedOperationId`/`durableTaskId` for
  diagnostic text;
- localize the notice so it does not imply that the historical ID is an active
  resumable task;
- actions address the session card, while Acknowledge uses the exact breadcrumb
  file already associated with that card.

### 12.2 Tests

- one and multiple breadcrumb projection;
- stable source/destination pane and plugin actions;
- displayed IDs have unambiguous labels;
- dismissal acknowledges only the selected breadcrumb;
- malformed or outside-root records remain untouched/rejected;
- no Resume, Roll back, retry, Delete, Cut-again, or automatic recovery action.

## 13. Verification matrix

Use the repository's build/test coordination rules. Do not kill an independently
launched executable that owns an output artifact.

### 13.1 Every package

```powershell
git diff --check
.\build.ps1 -ProjectName RedSalamander
.\Tools\Get-TestInventory.ps1 -Format Json
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
```

Run the focused Debug selftests and source-contract tests for the changed package.
For hot-path, queueing, rendering, startup, provider-I/O, or retention changes,
read `.github/skills/perf-validation/SKILL.md` and
`Specs/Testing/Testing_PerformanceValidation.md` before implementation, add
instrumentation and deterministic tests from the beginning, and archive the
test-enabled Release evidence.

### 13.2 Package focus

| Package | Minimum focused validation |
|---|---|
| A0 | deep recursive Local Copy; affected File Operations/Commands cases; Fresh Full |
| A1 | File Operations strategy/conflict/link/verification/cancel matrix at concurrency 1 and N; memory/throughput archive |
| A2 | exhaustive conflict snapshot/layout/default/Escape/accessibility cases |
| A3 | all three durable lifecycle suites plus persist/reparse/malformed fault injection |
| A4 | Change Case characterization and central RenamePlan mutation/recovery/artifact cases |
| A5 | 1/64/1,024-row delayed-provider admission, cancel, window close, plugin quiet point; UI-thread perf archive |
| A6 | FolderView/Find exact-parent projection and large-registry bind/latency archive |
| A7 | restart notice projection/actions/dismissal/multi-record tests |

### 13.3 Closeout

```powershell
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh
.\Tools\Test-TestRunArchive.ps1 -Inventory
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
git diff --check
```

Record the exact Fresh Full run ID, commit, totals, and archive-validation result
in this plan before moving it to Done.

## 14. Commit and integration strategy

Do not implement this as one branch-sized rewrite. Preferred sequence:

1. A0 evidence-only checkpoint.
2. A1 bridge extraction commit, then A1 item-policy extraction commit.
3. A2 single conflict snapshot/layout owner.
4. A3 shared durable I/O mechanics.
5. A5 responsive admission foundation.
6. A4 Change Case migration using the central/worker admission foundation.
7. A6 artifact candidate index.
8. A7 breadcrumb identity cleanup.
9. A8 final spec/evidence closeout.

A6 and A7 may run after A0 in either order if they do not touch files currently
owned by I12 or an active A1-A5 package. Only one active implementation package
may modify `FolderWindow.FileOperations.State.cpp` at a time.

Each behavior-changing commit must update its authoritative spec and tests in the
same commit. Pure extraction commits must state “no normative behavior change”
and prove parity.

## 15. Program done criteria

This plan is complete only when all are true:

1. A current closeout commit has a validated Fresh Full archive.
2. The bridge and item policy have named, testable ownership outside the giant
   `ExecuteOperation` body.
3. The engine snapshot is the sole conflict-action/default/cancel policy owner.
4. Durable stores share safe file mechanics without sharing authority semantics.
5. Change Case no longer mutates through a separate rename engine.
6. Batch Rename admission performs no row-proportional provider I/O on the UI thread.
7. Artifact hints are scoped enough to avoid global common-leaf probing while
   preserving Proven/Possible visibility.
8. Interrupted-Move notice identities are unambiguous and remain notice-only.
9. Focused Debug and Release evidence is archived for every perf/lifetime package.
10. All affected domain specs describe the resulting durable behavior and
    architecture boundaries.
11. I3 and I12 ownership is not duplicated or silently closed.
12. This plan is moved to `Specs/Plans/Done/` and the WIP index is reconciled.

## 16. Global STOP conditions

- Stop if an extraction changes a result axis, conflict action, destructive
  authority, clipboard rule, folder merge, or link behavior without an explicit
  domain-spec amendment and product decision.
- Stop if any path reintroduces pathname delete after Managed publication,
  FollowTargets, FNV/digest mutation authority, or silent identical-file Skip.
- Stop if UI-thread work becomes proportional to enumerated descendants or
  selected rows without measured and approved bounds.
- Stop if a shared durable helper makes schemas mutually readable or grants a
  breadcrumb/journal/claim authority it did not previously own.
- Stop if serial and parallel scheduling acquire separate conflict/result policy.
- Stop if a cross-thread UI payload bypasses `PostMessagePayload` ownership or its
  target window lacks init/drain lifecycle.
- Stop if an independently launched application blocks an artifact; report the
  exact owner and wait for user action rather than killing it.
- Stop closeout if Fresh Full, spec inventory, test inventory, archive validation,
  or `git diff --check` is not green at the exact closeout commit.

## 17. Considered and rejected

- **Reopen the completed Phase 0-6 plan:** rejected. It is frozen history; this
  plan is a bounded post-closeout owner.
- **Treat the current merge as unresolved:** rejected at the planned-at commit.
  The conflict report was stale; A0 still revalidates the merged behavior.
- **Merge Native and Managed transfer engines:** rejected. They have intentionally
  different provider, sharing, publication, and cleanup semantics.
- **Move Create Directory into `StartOperation`:** rejected. Qualified F7 remains
  a deliberate outside-engine command contract.
- **Advertise abort/deadline true to make Phase 6 look more complete:** rejected.
  Capabilities remain false until the provider can execute them honestly.
- **Make one generic durable JSON journal:** rejected. Only safe-file mechanics
  are shared; schemas and authority remain separate.
- **Delete `FileSystemRenameBatch` before migrating Change Case:** rejected. First
  characterize and move the final production caller.
- **Use artifact hints as proof:** rejected. Hints only select candidates for
  no-follow claim-and-identity classification.
- **Add a second File Operations admission framework:** rejected. Batch Rename
  qualification becomes a worker stage feeding the existing typed plan and task
  engine.
