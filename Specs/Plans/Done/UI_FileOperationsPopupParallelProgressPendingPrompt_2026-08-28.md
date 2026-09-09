# File Operations P0 release blocker: Parallel admission, truthful progress/prompts, and safe SMB publication

> **NON-NORMATIVE RETIRED IMPLEMENTATION RECORD.** Moved from
> `Specs/Plans/WIP/` on 2026-08-29 after the I17 implementation, focused
> correctness/performance evidence, localization/resource work, and authoritative-
> spec updates completed. **No remaining implementation action belongs to this
> file.**
>
> The implementation is **not yet Fresh-Full qualified**. By product-owner
> decision, all residual I17 qualification work—final merged FileOps/Commands/
> resource-contract confirmation, spec-inventory/patch hygiene on the closeout
> commit, and one exact Fresh Full—is transferred to
> `Specs/Plans/WIP/Operation_FileOperations_ReliabilityFirstSuccessorPlan_2026-08-29.md`
> section 9.4 and its E8 exact merged closeout. That successor must record one
> receipt qualifying both its own scope and this
> archived I17 scope. Until that receipt is green, release qualification remains
> pending. Decision-gated rows do not retroactively own or implement I17 behavior.
>
> **Frozen history: do not resume.** The historical executor instructions, gates,
> and STOP conditions below document how the delivered work was produced; they are
> not a live queue. Historical live-SMB execution remains unavailable environmental
> coverage, while deterministic retained-authority coverage remains the baseline.
>
> **Historical implementation-plan preamble.** Durable behavior remains owned by
> `Specs/UI/UI_FileOperationsPopup.md`,
> `Specs/FileSystem/FileSystem_FileOperations.md`,
> `Specs/Plugins/Plugins_VirtualFileSystem.md`, and the applicable testing
> specifications. Merge every lasting contract into those owners before this
> plan moves to `Specs/Plans/Done/`.
>
> **Release rule:** this is a **CRITICAL / P0 / SHOW-STOPPER RELEASE
> BLOCKER**. Do not ship while Local→SMB/UNC Copy or Move can lose mutation
> truth, report a known pre-publication failure as Outcome unknown, discard
> owned-stage cleanup truth, or delete a source without exact proof that the
> destination was published and verified.
>
> **Executor rule:** execute the safety gates before enabling the Parallel/UI
> feature work. Each gate has explicit invariants, RED tests, verification,
> and STOP conditions. Do not replace an exact-object guarantee with pathname
> optimism, a byte-count heuristic, a zero/empty file identifier, or a weak
> file identifier.

## Status

| Field | Value |
|---|---|
| State | **RETIRED — implemented scope and focused evidence complete; all residual aggregate qualification transferred to the live unified successor E8 gate** |
| Severity | **CRITICAL / P0 / RELEASE BLOCKER** |
| Index owner | **I17** |
| Planned at | `25314bea5350e246f6ff4b9b2bc020ffabcc8794` (`25314bea5`) |
| Updated | 2026-08-29 |
| Primary failure | Resolved: Local→SMB/UNC owned stages retain exact handle-bound authority when persistent file identity is unsupported, while every ambiguous mutation remains explicitly Unknown |
| Scope | File Operations typed results, Local owned-stage identity/publication, mutation interlock admission, popup progress/prompt/accessibility, focused tests, perf evidence, localization, authoritative specs |
| Non-goals | intra-file chunk concurrency; conflict-action policy redesign; I12 completion/reaper work; I14 bridge extraction; global File Operations redesign |
| Closeout gate | Run one exact Fresh Full on the merged E8 closeout commit of `Operation_FileOperations_ReliabilityFirstSuccessorPlan_2026-08-29.md`; the earlier detached-tree Fresh run remains focused evidence, not the final repository receipt |

> **Product-owner validation consolidation (2026-08-29):** Do not schedule a
> standalone Fresh Full solely for this retired record. Its final whole-repository
> qualification comes from the successor gate named above. Until that run is green,
> I17 is retired but not Fresh-Full qualified. Its completed focused correctness,
> performance, resource, and authoritative-spec evidence remains the scoped
> baseline; the successor's undecided behavior is not retroactively owned or
> implemented here.

### Working-tree preflight

The saved checkout containing this plan had unrelated user changes and this
plan was untracked when reviewed. Implementation must start in a dedicated
clean worktree at `25314bea5` or a deliberately rebased descendant. Never
overwrite, reset, or absorb unrelated saved-checkout changes.

Before implementation:

```powershell
git status --short
git rev-parse HEAD
git diff --check
```

Record the output in the implementation task. If the implementation worktree
is dirty in any in-scope file, STOP and reconcile ownership before editing.

Drift check:

```powershell
$driftPaths = @(
  'RedSalamander/FolderWindow.FileOperations.State.cpp'
  'RedSalamander/FolderWindow.FileOperations.State.Queue.cpp'
  'RedSalamander/FolderWindow.FileOperations.State.Runtime.cpp'
  'RedSalamander/FolderWindow.FileOperations.Popup.cpp'
  'RedSalamander/FolderWindow.FileOperations.Popup.h'
  'RedSalamander/FolderWindow.FileOperationsInternal.h'
  'RedSalamander/FolderWindow.FileOperations.cpp'
  'Plugins/FileSystem/FileSystem.cpp'
  'Plugins/FileSystem/FileSystem.FileOps.cpp'
  'Common/PlugInterfaces'
  'RedSalamander/RedSalamander.rc'
  'RedSalamander/Resource.h'
  'RedSalamander/Lang'
  'RedSalamander/SelfTest/Commands/Commands.SelfTest.FileOps.cpp'
  'RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.cpp'
  'RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.Phases05_06.cpp'
  'Specs/UI/UI_FileOperationsPopup.md'
  'Specs/FileSystem/FileSystem_FileOperations.md'
  'Specs/Plugins/Plugins_VirtualFileSystem.md'
  'Specs/Testing/Testing_SelfTests.md'
  'Specs/Testing/Testing_PerformanceValidation.md'
  'Specs/Plans/WIP/README.md'
  'Specs/Plans/WIP/UI_FileOperationsPopupParallelProgressPendingPrompt_2026-08-28.md'
)
git diff 25314bea5350e246f6ff4b9b2bc020ffabcc8794 -- $driftPaths
```

If drift invalidates an anchor or an invariant is already implemented, update
this plan and its checklist before continuing. Do not duplicate behavior.

## Progress checklist

`[x]` means implemented, tested, documented, and reviewed. `[+]` means the
only active package.

- [x] **G0 — Establish provenance and deterministic RED tests**
  - [x] Prove deterministic raw Win32 996 phase attribution; retain the live
        historical endpoint as environmental coverage rather than guessing.
  - [x] Add content-free bounded failure-phase telemetry.
  - [x] Add injections for post-`CREATE_NEW` identity failure, cleanup failure,
        zero/empty `FILE_ID_128`, ambiguous rename/delete reconciliation,
        replacement rollback failure, and non-bridge stage-abort failure.
  - [x] Add RED typed-result tests before changing classification.
- [x] **G1 — Make mutation truth authoritative end to end**
  - [x] Store exactly one typed item result on every terminal path, including
        retry-loop Cancel, KeepBoth failure, requalification failure, and
        TraversalLimit.
  - [x] Preserve bridge `Published / NotPublished / Unknown` through
        failure, Skip, Cancel, Continue-on-error, and final aggregation.
  - [x] Keep final-destination publication separate from owned-stage
        cleanup/artifact truth.
  - [x] Remove HRESULT- and byte-count-based inference of mutation truth.
  - [x] Treat ambiguous conditional rename/delete/abort/rollback as Unknown.
- [x] **G2 — Repair Local SMB/UNC owned-stage authority**
  - [x] Keep exact local `FILE_ID_INFO` authority unchanged.
  - [x] Add retained-handle-owned stage authority for a just-created object
        when a remote volume cannot supply `FileIdInfo`.
  - [x] Treat a successful all-zero `FILE_ID_128` as Unsupported, never as a
        persistent exact identity.
  - [x] Make `exclusiveStage` qualification truthful for the destination
        profile and preserve cleanup truth in bridge and non-bridge paths.
  - [x] Never use a weak/reopened-path token for overwrite, source deletion,
        publication reconciliation, or pre-existing-object binding.
  - [x] Preserve exact stage/artifact truth on every cleanup failure.
- [x] **G3 — Fail current alias admission closed, then make destination-leaf Parallel safe**
  - [x] Reject current Indeterminate/contract-violating source or destination
        binding before mutation; do not retain a path-only scope.
  - [x] Register resolved destination leaves, not one destination folder.
  - [x] Represent an absent leaf as exact existing ancestor anchor(s) plus
        normalized relative suffix.
  - [x] Fail closed on indeterminate/contract-violating binding.
  - [x] Preserve same-leaf/ancestor serialization across drive, UNC, SUBST,
        junction, mount, hard-link, and case aliases.
  - [x] Measure provider bind count and overlap-scan cost; add an index only if
        the measured admission budget requires one.
- [x] **G4 — Unify progress state and aggregate semantics**
  - [x] Use one predicate for paint, hosted controls, debug snapshots, UIA,
        taskbar, and footer.
  - [x] Never animate Waiting or Needs-attention as transfer activity.
  - [x] Keep provisional open-discovery progress card-only.
  - [x] Define the footer/taskbar active cohort explicitly.
- [x] **G5 — Rebuild prompt/diagnostic layout and accessibility**
  - [x] Show a stable localized `Waiting for your decision` status for every
        actionable conflict; show `Loading decision details...` while facts
        are incomplete.
  - [x] Question first, then object context/facts, then actions.
  - [x] Handle `metadataLoading` without stale graph/progress.
  - [x] Measure wrapped question and last-note height.
  - [x] Expose question/context and correct reading order through UIA.
  - [x] Validate minimum width, compact density, long localization, and DPI.
- [x] **G6 — Validate, document, and transfer final qualification**
  - [x] Focused Debug correctness tests pass.
  - [x] Test-enabled x64 Release perf evidence is archived and validated.
  - [x] Resource and spec contracts pass.
  - [x] Consolidated Fresh Full obligation transferred to section 9.4/E8 of the sole
        live reliability-first successor; this does not claim that run has passed.
  - [x] Authoritative specs match shipped behavior.
  - [x] I17 is removed from the WIP index and this record moves to Done.

## Release slicing and dependency order

The plan remains one tracked work item, but its gates are not allowed to hold
the emergency safety repair hostage to the larger feature/UI package:

1. **P0 safety hotfix:** G0, G1, G2, and G3A (current interlock
   Indeterminate/contract failure fails closed). This is the Local→SMB/UNC
   release blocker and must be independently buildable, testable, and safe to
   cherry-pick.
2. **Parallel feature gate:** G3B destination-leaf admission. Do not enable
   leaf-level Parallel until anchor+suffix alias tests pass. The hotfix may
   retain conservative folder serialization.
3. **Popup correctness package:** G4 and G5. These remain required to close
   I17 and to ship the popup changes, but they are not prerequisites for
   deploying the mutation-truth hotfix.
4. **Closeout:** G6 applies to every shipped slice; authoritative specs and
   receipts must identify which slice is present.

If release management cannot split these safely, keep the entire item
blocking. Never defer G0–G3A behind visual polish.

## 1. Release acceptance contract

Clauses 1–4 and current fail-closed alias admission are blocking for the P0
hotfix. Clause 5 blocks enabling destination-leaf Parallel. Clauses 6–8 block
shipping the popup correctness package and closing I17.

1. A Local→SMB/UNC Copy/Move either:
   - publishes and returns exact destination truth; or
   - fails with the original/mapped provider error and exact
     `NotPublished / Retained / Failed` truth; or
   - reports genuinely Unknown axes only when a mutation was attempted and
     exact reconciliation could not decide the outcome.
2. Move source deletion occurs only after exact destination publication and
   required verification are proven. Source identity alone is insufficient.
3. Raw provider `ERROR_IO_INCOMPLETE` is not itself authority for Unknown.
   Unknown is derived from typed mutation evidence; aggregate Unknown may
   still map to the existing 996 sentinel.
4. Zero bytes and absence of a provider receipt never prove non-mutation.
   Zero-length files, stage creation, metadata writes, rename, and cleanup can
   mutate state without transferring payload bytes.
5. Queue→Parallel admits distinct non-overlapping destination leaves
   immediately, including Copy and Move, without admitting the same object
   through aliases.
6. Waiting and Needs-attention are idle/decision states. They expose no
   transfer marquee, discovery animation, or throughput graph. An actionable
   conflict shows a stable localized **Waiting for your decision** status;
   metadata loading shows **Loading decision details...** and does not present
   incomplete destructive choices as actionable.
7. The conflict question and its source/destination/fact context are visible,
   unclipped, keyboard reachable, and announced before the action buttons.
8. Last-note failure text wraps and remains readable at supported minimum
   width, density, locale, and DPI.

## 2. Confirmed code findings

Line anchors are from `25314bea5`.

| Severity | Finding | Evidence | Consequence |
|---|---|---|---|
| P0 | Conditional rename can report known non-commit after mutation/rollback becomes ambiguous | `Plugins/FileSystem/FileSystem.cpp:3236`, `:3382-3386`, `:3409-3416` | The original destination may remain under a backup name while the bridge believes it is still present; a published object can be aborted after false non-publication |
| P0 | Failed rename reconciliation preserves false-known truth | `Plugins/FileSystem/FileSystem.cpp:3059-3078` | A remote rename that committed but cannot be re-identified may be treated as non-commit |
| P0 | Exclusive stage cleanup failure is discarded after identity failure | `Plugins/FileSystem/FileSystem.cpp:4380-4399`, `:4470-4482`, `:4570-4582` | A file/directory/link stage may remain visible while the host receives neither ownership nor artifact truth |
| P0 | Bridge publication truth is computed but can be lost | `RedSalamander/FolderWindow.FileOperations.State.cpp:16623-16637`, `:16874-16880`, `:17200-17285` | Fallback result synthesis can turn exact `NotPublished` into Unknown or lose `Published` on Skip/Cancel |
| P0 | Skip ignores supplied bridge publication truth | `RedSalamander/FolderWindow.FileOperations.State.cpp:9969-9974` | A post-publication failure followed by Skip can suppress destination refresh |
| P0 | `processIndex` has terminal returns before its common typed-result store | `RedSalamander/FolderWindow.FileOperations.State.cpp:16763-16769`, `:16806-16814`, `:16919-16927`, `:16978-16989`, `:17021-17025`, `:17245-17253` | Retry-loop Cancel, KeepBoth path failure, native-directory requalification failure, and TraversalLimit can be synthesized later as Unknown/996 instead of preserving the real terminal truth |
| P0 | Continue-on-error can write the same item result twice | `RedSalamander/FolderWindow.FileOperations.State.cpp:17042-17070`, `:17327-17339` | The second store can alter status and, for selected directory roots where verification is not applicable, can rewrite `Published` to `NotPublished` |
| P0 | Owned-stage abort failures are diagnostic-only in the bridge | `RedSalamander/FolderWindow.FileOperations.State.cpp:12064-12093` | A final destination can be exactly NotPublished while a temporary stage remains Possible/Unknown; retry and source deletion lack this artifact axis |
| P0 | Conditional delete/abort preserves initialized known-noncommit after an ambiguous failure | `Plugins/FileSystem/FileSystem.cpp:2931-2993`, `:2996-3024`, `:3083-3098`, `:3132-3148` | A deletion that may have committed can be reported as a known non-commit, allowing unsafe retry, cleanup, or aggregation |
| P0 | A successful all-zero `FILE_ID_128` is accepted as persistent identity | `Plugins/FileSystem/FileSystem.cpp:1917-1939` | Redirectors that return zero for unsupported identity can make unrelated SMB objects compare equal |
| P0 | Local advertises `exclusiveStage:true` for every root and Managed selection consumes the static fact | `Plugins/FileSystem/FileSystem.cpp:5115-5133`; `RedSalamander/FolderWindow.FileOperations.cpp:1192-1212`, `:3075-3090` | A destination may enter a retained-stage strategy before the destination profile has demonstrated the required authority |
| P0 | Non-bridge staged copy/link/directory cleanup discards abort results | `Plugins/FileSystem/FileSystem.FileOps.cpp:3364-3373`, `:4198-4207`, `:4665-4674` | Local Copy/overwrite can orphan an owned stage independently of whether the live 996 came from `CopyFileExW` |
| P0 | Optional interlock binding swallows indeterminate/contract failure | `RedSalamander/FolderWindow.FileOperations.State.cpp:7167-7244`, `:7273` | Cross-root aliases can bypass read/write/publication exclusion after an SMB identity failure |
| P0 | Existing mapped-drive/UNC aliases are already unsafe before leaf Parallel | `RedSalamander/FolderWindow.FileOperations.cpp:1306-1369`; `RedSalamander/FolderWindow.FileOperations.State.cpp:7167-7244`, `:7258-7279` | An existing SMB folder whose bind is Indeterminate becomes a path-only scope, so mapped and UNC spellings of the same object can run concurrently today |
| P0 plan hazard | A missing leaf currently has no retained anchor | `RedSalamander/FolderWindow.FileOperations.State.cpp:7167-7244` and `State.Queue.cpp:542-584` | Naively switching to leaf scopes would admit the same absent target through drive/UNC or other aliases |
| P1 perf | Interlock overlap is Cartesian while holding `_queueMutex` | `RedSalamander/FolderWindow.FileOperations.State.Queue.cpp:587-610` | 512×512 scope pairs per active task can delay admission, cancellation, and Queue→Parallel |
| P0 UI | Paint creates open-discovery hosted progress before Waiting/Conflict suppression | `RedSalamander/FolderWindow.FileOperations.Popup.cpp:8344-8375` | The user sees a fake marquee while the debug snapshot says it is hidden |
| P0 accessibility | Conflict question/context is paint-only | `RedSalamander/FolderWindow.FileOperations.Popup.cpp:5435-5700`, `:8132-8276` | Assistive technology reaches destructive choices without the question or object context |
| P1 UI | Prompt/facts/paths and final diagnostics use fixed/clipped rows | `RedSalamander/FolderWindow.FileOperations.Popup.cpp:6084-6088`, `:7651-7656`, `:8188-8231` | Long localized text, high DPI, or minimum width can omit decision-critical content |

These findings define this plan’s minimum scope. They are not optional cleanup.

## 3. Resolved product and safety decisions

### 3.1 Result semantics

- Typed axes are primary. HRESULT is a status, not mutation evidence.
- Only explicit provider/engine mutation evidence can produce
  `Published`, `NotPublished`, or `Unknown`.
- Bridge-owned stage paths start as `NotPublished`, become `Published` only
  after exact publication commits, and become `Unknown` only after a
  mutation whose exact outcome cannot be reconciled.
- Final-destination publication and temporary-stage disposition are separate
  axes. A failed-before-publication operation can be exactly `NotPublished`
  while stage cleanup is `Unknown`; do not weaken the final-destination axis
  merely because a temporary artifact may remain.
- Add an explicit owned-stage/artifact state with at least
  `NotCreated / Removed / Published / Retained / Unknown`. `Retained` and
  `Unknown` block automatic retry and Move source deletion and surface
  `Possible artifact` diagnostics even when final publication is
  `NotPublished`.
- A provider route without mutation evidence remains Unknown on a failed
  destructive operation. Do not guess from bytes or filesystem existence.
- A true post-publication source-cleanup ambiguity remains
  `Published / Unknown source / Indeterminate` and aggregate 996.
- A failed-before-publication Move with exact bridge truth is
  `NotPublished / Retained / Failed` with the real failure HRESULT.
- A typed item result is single-assignment. A second store for the same item
  is a test failure and debug invariant violation, not last-writer-wins.
- Delete and abort use the same mutation truth model as rename: once a
  potentially committing call begins, known non-commit is legal only after
  exact reconciliation proves the original object/stage still exists.

### 3.2 Owned-stage identity

- `FILE_ID_INFO` remains the persistent exact identity for ordinary Local
  binding, pre-existing destination replacement, revalidation, and source
  deletion.
- A successful query whose `FILE_ID_128` is entirely zero means persistent
  identity is unsupported and must be ignored. It never enters identity
  comparison, caching, interlock admission, or overwrite authority.
- A just-created stage may use a distinct lifetime-scoped opaque capability
  only because the task still owns an exact retained handle to that object.
- The capability is not a filesystem identifier and cannot be reconstructed
  by reopening a pathname.
- `BY_HANDLE_FILE_INFORMATION` volume serial + 64-bit index is forbidden as
  exact authority for this repair.
- If a method cannot prove its result through the retained handle or an exact
  expected-destination binding, it returns Unknown and retains artifacts for
  reconciliation. It does not claim success by weak pathname comparison.

### 3.3 Interlock

- Copy/Move publication scopes are the resolved final destination leaves.
- The same missing leaf through aliases serializes.
- Different sibling leaves under one exact parent may run concurrently.
- Existing hard-link aliases and ancestor/descendant envelopes serialize.
- `Missing`, `Unsupported`, `Indeterminate`, and provider contract failure
  are distinct states:
  - missing destination leaf: anchor to exact existing ancestor plus suffix;
  - unsupported exact identity: conservative conflicting scope within the
    identity domain;
  - indeterminate or contract violation: fail admission before mutation with
    the original HRESULT.
- This fail-closed rule applies to the current folder/source scopes before
  any destination-leaf Parallel change. `bindingRequired=false` must not turn
  Indeterminate or provider-contract failure into path-only admission.

### 3.4 Progress and prompt presentation

- Provisional open-discovery progress is card-only. A started task whose
  traversal is open keeps aggregate/taskbar progress indeterminate.
- Footer/taskbar progress is the **currently executable cohort**: started,
  unfinished tasks that are neither waiting for admission nor blocked on a
  decision. Unstarted Waiting and decision-blocked tasks are shown in counts
  but excluded from the cohort denominator. When cohort membership changes,
  the aggregate may rebase; the authoritative UI spec must state this
  explicitly.
- A Waiting or conflict-active task never exposes a transfer bar, discovery
  activity, graph, or marquee, including while conflict metadata loads.
- Every actionable conflict has a persistent localized status line or chip
  reading **Waiting for your decision**. It is visible without animation,
  included in UIA, and remains present until the decision is submitted or
  canceled. While facts are incomplete, show **Loading decision details...**,
  keep incomplete destructive actions disabled/absent, and never substitute
  a marquee.
- Conflict layout order is:
  **question → From/To → facts/metadata → actions**.
  The question is the primary first block; it is not required to be
  geometrically adjacent to buttons through intervening metadata.
- Visual order, keyboard traversal, UIA sibling order, and announcement order
  match.

## 4. Architecture guardrails

- Preserve WIL RAII for every handle/window/COM resource.
- No `catch (...)`. Do not swallow `std::bad_alloc`.
- Reuse current path-identity and destination-resolution helpers. Search
  `Specs/Core/Core_SharedHelpers.md` and `Common/` before adding a helper.
- Do not move I12 completion/reaper/quiet-point paths.
- Do not extract I14 bridge/`processIndex`/conflict-action policy.
- No new raw cross-thread `PostMessageW` payload ownership.
- No new hardcoded UI strings; use resources and positional placeholders.
- Performance metrics are aggregate and content-free: no paths, names,
  identity bytes, credentials, or file contents.
- Selftests are deterministic; no sleeps and no live-share credentials.

## 5. G0 — Provenance, telemetry, and RED tests

### 5.1 Add bounded failure provenance

Add a small internal enum owned by File Operations:

- `None`, `DestinationParent`, `StageCreate`, `StageIdentity`,
  `StageWrite`, `StageCommit`, `FinalPublish`,
  `FinalPublishReconcile`, `ReplacementRollback`, `StageAbort`,
  `Verification`, `SourceCleanup`, `ProviderNative`.

Carry it in `QualifiedItemMutationResult` and the Local owned-stage
implementation. Emit one aggregate row per terminal item:

- metric `FileOps.Bridge.FailurePhase`;
- bounded enum detail only;
- `value0` bit field for stage-created, cleanup-attempted,
  final-publish-attempted, and source-delete-attempted;
- `value1` typed publication state;
- `hr` original phase HRESULT.

Record owned-stage disposition separately from final publication. Add a
bounded `FileOps.Bridge.StageDisposition` value or an equivalent field in the
terminal item row. Never encode a stage artifact by changing an exactly known
final publication from `NotPublished` to `Unknown`.

Add `FileOps.Local.ExclusiveIdentityQuery` with bounded detail
`file`/`directory`/`link` and the raw query HRESULT. Never emit a path or
identity payload. Update `Specs/Testing/Testing_PerformanceValidation.md`.

### 5.2 Deterministic fault injection

Follow existing `ENABLE_TESTS` hooks and reset each with RAII. Add injections:

1. `FileIdInfo` returns 996 after successful `CREATE_NEW`.
2. Exact stage deletion succeeds after identity failure.
3. Exact stage deletion fails/has unknown outcome after identity failure.
4. Handle rename applies but reports failure and exact reconciliation fails.
5. Existing destination moves to backup, stage publication fails, and backup
   rollback fails.
6. Bridge returns each publication state with the same failure HRESULT.
7. `GetFileInformationByHandleEx(FileIdInfo)` succeeds with an all-zero
   `FILE_ID_128`.
8. `DeleteExactHandle` applies or may apply deletion, reports failure, and
   exact reconciliation succeeds/fails independently.
9. Bridge `AbortOwnedStage` returns known failure and Unknown outcome.
10. Non-bridge file/link/directory scope-exit abort returns known failure and
    Unknown outcome.
11. Each named `processIndex` early exit fires after qualification, and a
    verification failure takes Continue-on-error.

Before GREEN changes, tests must prove:

- the observed/injected 996 phase is named;
- ambiguous rename/rollback is not known non-commit;
- post-create cleanup failure is not discarded;
- bridge `NotPublished` survives failure;
- bridge `Published` survives Skip/Cancel;
- retry-loop Cancel, KeepBoth failure, native-directory requalification
  failure, TraversalLimit, and destination-resolution/guard failures each
  produce one typed terminal result;
- Continue-on-error cannot store an item result twice;
- final publication and owned-stage artifact state remain independently
  observable;
- all-zero `FILE_ID_128` is Unsupported and cannot compare equal;
- direct and non-bridge stage abort failures are retained in typed truth;
- zero-byte fixtures do not alter publication classification;
- metrics contain no path/content.

If `\\rubygloom` is available, run one sanitized smoke after deterministic
tests. Record only profile, phase, HRESULT, typed axes, and artifact truth. If
unavailable, mark **[blocked: live SMB environment]**; deterministic coverage
remains mandatory.

### G0 verification

```powershell
.\build.ps1 -ProjectName RedSalamander
.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter FileOps_ProviderCapabilityMatrix -TimeoutMultiplier 2.0
```

Expected before GREEN work: new targeted assertions fail for the intended
reason while pre-existing assertions remain green.

### G0 STOP

- Root-cause text names `FileIdInfo` without phase evidence.
- A fixture contains live credentials or user file paths.
- Fault injection can escape `ENABLE_TESTS`.
- Mutation is inferred from bytes or path absence.

## 6. G1 — Typed mutation truth for rename, delete, abort, and item finalization

### 6.1 Structured local rename truth

Refactor private `ILocalBoundObjectControl::RenameExactTo` to return:

- status HRESULT;
- `outcomeKnown`;
- `mutationCommitted`;
- `originalStillPresent`;
- reconciliation phase.

Required transitions:

1. Handle rename succeeds → known committed.
2. Handle rename reports failure; exact reconciliation proves final owner →
   known committed.
3. Exact reconciliation proves no mutation → known non-commit.
4. Reconciliation fails/lacks exact authority → Unknown.
5. Existing destination moved to backup; publish fails; exact rollback
   succeeds → known non-commit with original restored.
6. Backup rollback fails/ambiguous → Unknown; retain stage and backup, report
   Possible artifacts, and do not auto-retry.

`FileSystemConditionalMutationResult` must reflect this before every return.
An initialized known-noncommit default must not survive after mutation begins
unless exact reconciliation proves it.

### 6.2 Structured delete and abort truth

Apply the same four-state mutation contract to `DeleteIfUnchanged`,
`AbortOwnedObject`, and their internal `DeleteExact`/`DeleteExactHandle` path:

1. delete succeeds → known committed, original absent;
2. delete reports failure and exact retained-handle reconciliation proves the
   original/stage remains → known non-commit;
3. delete reports failure and exact reconciliation proves delete-pending or
   absence for the exact object → known committed;
4. reconciliation is unavailable or fails after the committing call begins →
   Unknown.

Do not allow `InitializeMutationResult`'s known-noncommit default to leak out
of cases 3 or 4. If the current conditional result cannot express the needed
phase, add a size-versioned result extension; do not overload HRESULT.

### 6.3 Exactly one typed result per terminal item

Refactor `processIndex` so every terminal return passes through one
single-assignment item-finalization block consuming strategy, HRESULT,
`bridgePublication`, source cleanup, verification, action, and retained source
identity.

The executor must explicitly eliminate or route these known bypasses through
the block:

- retry-loop Cancel at `State.cpp:16763-16769`;
- destination resolution/guard/parent preparation failures at
  `:16780-16832`;
- same-path and provider-requested KeepBoth sibling lookup failures at
  `:16806-16814`, `:16919-16927`, and `:17245-17253`;
- native-directory requalification failure at `:16978-16989`;
- TraversalLimit at `:17021-17025`;
- unsupported conflict-action and final Cancel returns at `:17200-17285`;
- progress arithmetic overflow after mutation at `:17303-17313`.

Setup failures that occur before mutation still receive exact typed terminal
truth (`NotAttempted`/`NotPublished` as appropriate); they are not left for
`FinalizeTypedItemResults` synthesis.

- Skip changes completion, not publication.
- Cancel changes completion, not publication.
- Continue-on-error stores exact axes before advancing.
- VerificationFailure and CleanupIndeterminate do not store inside their
  branch and then fall through to the loop-end store. Replace the current
  last-writer-wins behavior with one final value and an assertion/test-only
  counter that rejects a second write.
- `ERROR_IO_INCOMPLETE` means Indeterminate only when typed evidence has an
  Unknown axis.
- `FinalizeTypedItemResults` remains a fail-safe for truly absent provider
  evidence; it never infers `NotPublished` from bytes.

### 6.4 Publication, stage, and source result matrix

| Event | Final publication | Owned stage | Source | Completion | Status |
|---|---|---|---|---|---|
| Stage create fails before object exists, proven | NotPublished | NotCreated | Retained | Failed | original HRESULT |
| Owned stage exists; abort succeeds | NotPublished | Removed | Retained | Failed/Canceled | triggering HRESULT |
| Owned-stage cleanup outcome unknown before publication | NotPublished | Unknown | Retained for Copy and Move | Indeterminate | aggregate 996; preserve triggering HRESULT in item diagnostics |
| Owned-stage cleanup known failed and stage remains | NotPublished | Retained | Retained for Copy and Move | Failed | original/cleanup HRESULT; no automatic retry |
| Publish commits; Copy later fails | Published | Published | Retained | Failed | later HRESULT |
| Publish commits; Move cleanup known non-commit | Published | Published | Retained | qualified CopyOnly/retained result | existing qualified status |
| Publish commits; Move cleanup unknown | Published | Published | Unknown | Indeterminate | aggregate 996 |
| Publish/rollback ambiguous | Unknown | Unknown/Retained | Retained | Indeterminate | aggregate 996 |
| Zero-length copy succeeds | Published | Published | Retained | Completed | `S_OK` |

Destination refresh and Compare Directories invalidation key from typed
publication, not completion label or HRESULT.

### G1 tests

- zero-byte success;
- post-create identity failure with cleanup known and unknown;
- failure before publication after zero bytes;
- prior success plus later NotPublished failure;
- Published then Skip; Published then Cancel;
- Continue-on-error first failure then success;
- Continue-on-error verification failure for an ordinary file and a selected
  directory root; assert one store and unchanged `Published` truth;
- every named early return stores exactly once and preserves its HRESULT;
- abort/delete applies-but-reports-failure with reconcile success/failure;
- final destination NotPublished plus stage Unknown remains two distinct
  axes and blocks retry/source deletion;
- replacement rollback failure;
- rename applied with unavailable reconciliation;
- existing managed post-publish cleanup-unknown remains
  `Published / Unknown / Indeterminate / 996`.

### G1 STOP

- Byte totals determine publication.
- Conditional result says known non-commit after failed reconciliation or
  rollback.
- Skip/Cancel changes `Published` to `NotPublished`.
- The same item result is stored more than once.
- Stage cleanup state is encoded by corrupting final publication truth.
- Delete/abort reports known non-commit after reconciliation is unavailable.
- Move deletes a source from Unknown destination truth.

## 7. G2 — Retained-handle-owned SMB stage authority

Implement the retained-handle identity mode only after G0 proves remote
`FileIdInfo` is unavailable at the owned-stage boundary. Zero-ID rejection,
capability honesty, and cleanup-result preservation are mandatory regardless
of which phase produced the live 996. G1 remains mandatory regardless of
cause.

### 7.1 Identity modes

Add distinct Local bound-object namespace tags:

- persistent `FILE_ID_INFO`;
- retained-handle-owned capability.

The latter is an opaque per-object nonce created after exclusive creation,
valid only while the exact handle is retained, and never recreated by
`BindObject`. `BY_HANDLE_FILE_INFORMATION` volume/index is forbidden as exact
authority.

Eligibility requires Local provider, exclusive creation by this call, exact
handle still owned, remote persistent identity unavailable, and no attempt to
authorize a pre-existing object with the nonce.

Before constructing persistent identity, scan all 16 bytes of `FILE_ID_128`.
An all-zero value returns `ERROR_NOT_SUPPORTED` (or the repository's canonical
Unsupported mapping), even when the Win32 query returned TRUE. Add a source
contract/test that prevents zero identity from entering equality or cache
keys.

### 7.2 Method contract

For retained-handle-owned stages:

- `GetSnapshot`/`IsSameObject` compare the lifetime capability.
- `OpenReader` duplicates the retained handle without `FileIdInfo` re-query.
- metadata/basic-info and `AbortOwnedObject` use the retained handle.
- `PublishAs` renames the retained handle.
- absent-final publication is supported.
- replacement requires independent persistent exact `expectedDestination`;
  otherwise fail closed.
- failed publication uses retained-handle proof only; no weak path reopen.
- unprovable final state returns `outcomeKnown=FALSE` and retains artifacts.
- published authority is the same retained exact handle.

Pre-existing SMB `BindObject` stays unsupported without persistent identity.
Source deletion still requires persistent exact authority.

### 7.3 Capability and strategy honesty

`publication.exclusiveStage=true` must mean the selected destination profile
can either:

- create a stage and return exact persistent authority; or
- create a stage and return the retained-handle-owned authority defined here,
  with structured cleanup truth on every failure.

Do not advertise this solely because the plugin is Local. Qualify the
destination root/profile before Managed selection, or make exclusive creation
fail with `ERROR_NOT_SUPPORTED` before mutation and requalify to the already
proven CopyOnly/provider route. A raw 996 after stage creation is not a
capability probe and must not silently retry through another strategy.

If retained-handle publication is fully implemented and tested on the remote
profile, the capability may remain true. Otherwise it must fail closed for
that profile. Record the selected strategy and capability evidence in bounded
G0 telemetry.

### 7.4 Bridge and non-bridge cleanup ownership

Make stage cleanup result-bearing in both execution families:

- bridge `AbortOwnedStage` feeds the terminal owned-stage/artifact axis;
- `FileSystem.FileOps.cpp` staged file, link, and directory scope exits consume
  the abort HRESULT plus `FileSystemConditionalMutationResult` instead of
  `static_cast<void>`;
- known removed, known retained, and Unknown cleanup outcomes propagate to the
  caller without throwing from a scope exit;
- a retained/Unknown stage is reported once as a Possible artifact, blocks
  automatic retry and Move source deletion, and remains available for human
  reconciliation.

This work is in scope even if G0 proves the observed 996 did not originate in
the direct `CopyFileExW` path.

Apply honest ownership/cleanup to writer, directory, and link exclusive
creation. File Copy/Move is mandatory; directory/link may fail closed when the
server lacks the underlying operation but cannot discard truth.

If current ABI cannot return ownership plus cleanup truth after post-create
failure, STOP and add a size-versioned optional result-bearing binding
interface. Do not encode truth in an overloaded HRESULT.

### G2 method matrix

| Method | Persistent identity | Retained-handle identity |
|---|---|---|
| Bind pre-existing | supported | forbidden |
| Snapshot/same object | persistent exact | exact lifetime capability |
| OpenReader | duplicate + persistent revalidation | duplicate retained handle |
| Publish absent final | supported | supported |
| Replace final | exact expected binding | only with separate exact expected binding |
| Abort stage | exact handle | exact retained handle |
| Rename reconcile | persistent exact comparison | retained-handle proof or Unknown |
| Source delete | supported with grant | forbidden |

### G2 tests

- injected 996 for writer/directory/link identity query;
- successful identity query with all-zero `FILE_ID_128` maps to Unsupported;
- zero/nonzero file publication;
- absent destination;
- pre-existing destination without exact identity remains unsupported;
- OpenReader begins at byte zero;
- abort removes only owned stage; abort failure retains artifact truth;
- non-bridge file/link/directory abort success, known retained, and Unknown
  outcomes reach the result contract;
- a remote profile without retained-handle authority selects/fails over to a
  proven non-destructive strategy before stage creation;
- failed rename without exact reconcile is Unknown;
- verification uses retained published authority;
- managed Move deletes Local source only after publish + verification;
- local NTFS replacement/source-delete tests remain unchanged.

### G2 STOP

- Weak volume/index identity is exact authority.
- `PathRefersToThisObject` reopens for retained-handle identity.
- Lifetime capability is accepted by `BindObject`.
- Pre-existing destination is replaced without independent exact binding.
- Source deletion consumes stage capability.
- Cleanup failure loses authority or Unknown artifact truth.
- All-zero persistent identity is cached or compared.
- Static `exclusiveStage:true` selects Managed for a destination profile that
  cannot satisfy the owned-stage contract.
- Bridge or non-bridge cleanup discards `AbortOwnedObject` results.

## 8. G3 — Current alias fail-closed safety and destination-leaf Parallel

### 8.1 G3A: repair current admission before leaf Parallel

This is part of the P0 hotfix. Today `BindObjectAuthority` correctly classifies
an existing SMB object that cannot provide exact identity as Indeterminate,
but `addScope(..., bindingRequired=false, ...)` can discard that state and
retain only path text. Fix current source and destination scope construction
so:

- Indeterminate and provider-contract violation fail preparation before any
  mutation and preserve the original HRESULT;
- Missing remains distinct and is handled only by an exact ancestor+suffix
  rule when that rule exists;
- Unsupported uses a documented conservative identity-domain envelope;
- mapped-drive and UNC spellings of the same existing folder/object cannot
  both enter on path-only scopes.

Add the mapped-drive/UNC existing-folder regression before changing the scope
granularity. The hotfix is allowed to keep destination-folder serialization.

### 8.2 G3B: destination-leaf scope construction

Resolve destinations with `TryResolveTransferDestinationProviderPath`. Add one
`PublishDestination` scope per unique final leaf and one source
`ReadSource`/`WriteSource` scope per selected item.

Extend `MutationInterlockScope` with per-scope anchor suffixes:

- shared exact authority node;
- normalized relative component sequence from that exact object to target;
- resolution state.

Walk upward even when the target is Missing. Every exact bound
target/ancestor contributes an anchor with target-relative suffix. Intern
common nodes and component storage.

Comparison:

1. preserve same-root lexical equality/ancestor checks;
2. for any common exact anchor identity, compare normalized suffixes;
3. equal suffix or prefix means overlap;
4. different sibling suffixes do not overlap;
5. exact leaf identity catches hard-link/alternate-name aliases;
6. Unsupported identity uses a conservative domain envelope;
7. Indeterminate/contract violation fails preparation.

Use provider path-profile rules; do not lowercase or impose Win32 rules on
provider-native paths.

### 8.3 Cost controls and conditional indexing

Memoize bind attempts by qualified endpoint + normalized path and reuse
destination-parent chains. Emit aggregate, content-free:

- `FileOps.Interlock.BindCount` per task;
- `FileOps.Interlock.CheckUs` per admission attempt;
- `FileOps.Interlock.ScopeComparisons` per attempt;
- `FileOps.Interlock.ActiveCandidates` per attempt.

No per-path/pair rows. The first alias-safe Cartesian anchor+suffix comparison
was measured before indexing. Its archived 256×256 Release result was 65,536
comparisons and 2,600,470 microseconds, which met the STOP-and-review trigger.
The implemented immutable index is keyed by verified identity domain, exact
anchor identity, and normalized suffix; conservative/unkeyable scopes retain
the exact Cartesian fallback.

For 256-vs-256 disjoint siblings:

- record bind, comparison, active-candidate, and wall-time distributions;
- admission completes within the deterministic selftest timeout without UI
  starvation or uncancelable provider work;
- candidate p95 `PrepareUs` and `CheckUs` do not regress by more than 10%
  against the recorded same-machine baseline without an approved explanation;
- retained nodes stay within unique governed paths plus 1,024-depth bound.

### 8.4 Runtime matrix

- Queue→Parallel distinct Copy leaves: both enter.
- Same for Move.
- same/case-equivalent leaf: wait.
- folder vs descendant: wait.
- same missing leaf via drive/UNC: wait.
- distinct missing siblings via drive/UNC: both enter.
- available SUBST/junction/mount aliases.
- existing hard-link aliases: wait.
- injected bind Indeterminate: fail before mutation.
- existing mapped-drive/UNC alias with Indeterminate bind: fail before
  mutation even with destination-folder serialization.
- identity Unsupported: conservative serialization.
- 256-item vs 256-item with multiple active tasks: cost bounds.

Queue wait and overlap wait use distinct localized status copy; neither is a
graph overlay.

### G3 performance evidence

Use test-enabled x64 Release. Archive two-file Queue/Parallel, 256×256
disjoint admission, same-leaf early exit, and alias serialization under:

`Specs/TestRuns/<ComputerHashName>/FileOps/<RunId>/`

Keep `results.json`, `trace.txt`, and `perf_metrics.jsonl`. Use identical
machine/profile/scenario parameters. Bind/comparison counts are required
evidence, not a prescribed data structure. An unexplained p95 regression
above 10% in `PrepareUs` or `CheckUs`, selftest timeout, UI starvation, or
uncancelable provider probing triggers STOP-and-review and may require the
index described above.

Recorded evidence on machine `4cb089111a23`:

- trigger: `FileOps/2026-08-29_022850`, 65,536 comparisons,
  2,600,470 microseconds;
- accepted indexed baseline: `FileOps/2026-08-29_024402`, 8,704 probes,
  2,689 microseconds, three provider-matrix cases passed, zero disk-audit
  findings;
- both archives pass explicit and whole-inventory archive validation.

```powershell
.\Tools\Test-TestRunArchive.ps1 -RunPath <run-folder>
.\Tools\Test-TestRunArchive.ps1 -Inventory
```

### G3 STOP

- Missing leaf has no exact anchor and is unrelated across aliases.
- Indeterminate binding becomes path-only.
- Same alias target enters concurrently.
- Sibling concurrency drops publication conflicts globally.
- Measured admission cost violates the approved budget and no bounded index
  or other correction is supplied.
- Remote probing is unbounded or uncancelable.

## 9. G4 — Progress and aggregate state

Create one pure presentation helper used by D2D paint, hosted descriptors,
debug snapshots, UIA, footer, and taskbar. It returns:

- `showDiscoveryActivity`;
- `showTransferProgress`;
- `showTransferMarquee`;
- `showThroughputGraph`;
- `includeInAggregateCohort`;
- `progressDeterminate`;
- `provisional`.

| State | Discovery | Transfer bar | Graph | Aggregate |
|---|---|---|---|---|
| Unstarted Waiting | hidden | hidden | hidden | excluded |
| Conflict metadata loading | hidden | hidden | hidden | excluded |
| Actionable conflict | hidden | hidden | hidden | excluded |
| Started/open traversal/usable denominator | secondary | provisional determinate card | real throughput only | included, aggregate indeterminate |
| Started/open traversal/no denominator | secondary | marquee only if transfer active | real throughput only | included, indeterminate |
| Started/closed traversal | hidden | determinate when usable | allowed | included |
| Finished | hidden | completion only | hidden | excluded |

Open-discovery card denominator is
`max(discoveredTotalBytes, completedBytes)` or item equivalent. Provisional
percent/ETA remains card-only. Close single-file traversal when the provider
reports `traversalClosed`; test it rather than special-case the popup.

Update `Specs/UI/UI_FileOperationsPopup.md` with active-cohort/rebase
semantics before changing footer/taskbar.

### G4 tests

- Waiting with closed/open-total running sibling;
- conflict and `metadataLoading` with running sibling;
- completed bytes greater than discovered-so-far;
- one-file traversal closes before copy completion;
- paint/hosted/debug/UIA predicates agree;
- cohort entry/exit/rebase is deterministic.

### G4 STOP

- Paint/debug/UIA use separate visibility rules.
- Waiting/conflict creates hosted progress.
- Provisional open traversal becomes final footer/taskbar percent.
- UI spec and implementation describe different aggregate denominator.

## 10. G5 — Prompt layout, diagnostics, UIA, localization

For every `conflict.active`:

1. dedicated prompt block;
2. stable localized state indicator:
   - actionable: **Waiting for your decision**;
   - incomplete facts: **Loading decision details...**;
3. decision question (when actionable) or localized metadata-loading
   explanation;
4. From/To and facts needed to make the decision;
5. actions, enabled only when their required facts/authority are ready;
6. no discovery/progress/graph/marquee.

The state indicator is not an animated progress affordance. It uses the
existing Needs-attention visual language (icon/color/text), remains visible
without hover or focus, and changes only on a real state transition. Do not
use ellipsis animation, pulsing, marquee, or a fake transfer percentage.

Measure question, facts, From/To, and last note with DirectWrite. Card height
comes from content; remove fixed line allowances. Preserve scroll/footer
separation.

Add debug fields:

- `conflictPromptPrecedesActions`;
- `conflictWaitingForDecisionVisible`;
- `conflictDecisionDetailsLoadingVisible`;
- `conflictPromptVisibleWithoutClipping`;
- `conflictContextVisibleWithoutClipping`;
- `conflictDiscoveryIndicatorVisible`;
- `conflictTransferProgressVisible`;
- `lastNoteVisibleWithoutClipping`.

Do not use the contradictory 12-DIP prompt-adjacency rule.

Add semantic UIA text regions for question, From, To, facts, and last note.
UIA sibling/announcement order matches visual order. Hidden
discovery/progress/graph is absent from UIA. Needs-attention announces the
localized waiting state and the question, not only generic status. Metadata
loading announces once on entry and the actionable question announces once
when the state changes; repaint must not repeat announcements.

Validate metadata loading/actionable conflict, visible waiting/loading state,
Copy/Move, minimum/ordinary
width, compact/comfortable density, 100/150/200% DPI, long English and longest
satellite string, long paths, fact-row variants, last-note wrapping,
scroll/footer/action non-overlap, UIA names/order/hidden elements/Invoke.

Add/modify English, `cs-CZ`, `fr-FR`, `ja-JP`, and `sk-SK` resources with
positional placeholders.

```powershell
.\Tools\Tests\ResourceLocalizationContracts.Tests.ps1
.\Tools\Run-AllTests.ps1 -Suite Commands -CommandsFamily file-operations -SkipBuild -CaseFilter cmd_pane_fileops_popup_progress_contracts
```

### G5 STOP

- Question/context remains paint-only.
- Required text clips or disappears.
- Visual, keyboard, and UIA order differ.
- Metadata loading exposes stale progress or enabled destructive action
  before facts are ready.
- An actionable card lacks a persistent visible/UIA
  `Waiting for your decision` indication.
- Repaint causes repeated UIA live-region announcements.

## 11. Implementation file map

| Area | Primary files | Required work |
|---|---|---|
| Typed results | `RedSalamander/FolderWindow.FileOperations.State.cpp` | G0 phase; G1 single-assignment terminal store; publication/source/stage axes |
| Interlock | `State.cpp`, `State.Queue.cpp`, `FolderWindow.FileOperationsInternal.h` | G3A current fail-closed binding; G3B anchor-suffix scopes; conditional index/metrics |
| Queue | `State.Runtime.cpp` | preserve Queue→Parallel release; no second queue |
| Local provider | `Plugins/FileSystem/FileSystem.cpp` | structured rename/delete/abort truth; zero-ID rejection; identity modes; capability/cleanup/publication |
| Provider copy | `Plugins/FileSystem/FileSystem.FileOps.cpp` | always in scope for result-bearing file/link/directory stage cleanup; direct 996 cause still requires G0 proof |
| Popup | `FolderWindow.FileOperations.Popup.cpp/.h` | single presentation rule; measured layout/UIA |
| Resources | `Resource.h`, main/satellite `.rc` | localized wait/loading/failure copy |
| FileOps tests | `FolderWindow.FileOperations.SelfTest.Phases05_06.cpp` plus registration | result, identity, alias, runtime, perf |
| Commands tests | `Commands.SelfTest.FileOps.cpp` | snapshots, layout, UIA, DPI/localization |
| Specs | UI/FileSystem/Plugins/Testing | durable contracts, metrics, evidence |

Do not claim `CopyFileExW` caused the live 996 without G0 evidence. Regardless
of that provenance, fix the independently confirmed `FileSystem.FileOps.cpp`
scope-exit paths that discard `AbortOwnedObject` results. If G0 also proves a
direct transfer branch is implicated, give that branch explicit mutation
evidence; otherwise record it as not implicated in the live failure.

## 12. Consolidated test matrix

| ID | Scenario | Blocking assertion |
|---|---|---|
| R1 | Identity 996 after `CREATE_NEW` | exact phase |
| R2 | Cleanup succeeds | NotPublished / Retained / Failed |
| R3 | Cleanup unknown | artifact retained; Unknown; no source delete |
| R4 | Rename applied/reconcile unavailable | Unknown |
| R5 | Replacement rollback fails | stage + backup retained/reported |
| R6 | Bridge NotPublished failure | real HRESULT; destination truth |
| R7 | Published then Skip/Cancel | Published survives; refresh occurs |
| R8 | Zero-byte file | evidence, not bytes, classifies |
| R9 | Prior success + later failure | aggregate partial/failed, not whole-task Unknown |
| R10 | Managed cleanup unknown | Published / Unknown / Indeterminate / 996 |
| R11 | Retry Cancel/KeepBoth/requalify/TraversalLimit early exits | exactly one typed item result; real HRESULT/axes preserved |
| R12 | Verification failure + Continue-on-error | exactly one store; Published survives for file and selected directory root |
| R13 | Abort/delete applies but reports failure | reconcile decides committed/non-commit or returns Unknown |
| R14 | All-zero `FILE_ID_128` | Unsupported; never equal/cached |
| R15 | Bridge and non-bridge abort failure | final NotPublished remains exact; stage Retained/Unknown reported; no retry/source delete |
| R16 | Destination profile lacks safe exclusive stage | Managed not selected after mutation; proven fallback or fail-closed before create |
| I0 | existing mapped-drive/UNC alias; bind Indeterminate | fail before mutation under current folder scopes |
| I1 | distinct Copy leaves | concurrent |
| I2 | distinct Move leaves | concurrent |
| I3 | same/case-equivalent leaf | serialized |
| I4 | drive/UNC same missing leaf | serialized |
| I5 | drive/UNC siblings | concurrent |
| I6 | SUBST/junction/mount/hard-link | exact overlap |
| I7 | bind Indeterminate | fail before mutation |
| I8 | 256×256 | comparison/bind/node bounds |
| P1 | Waiting | no activity UI |
| P2 | metadata-loading conflict | visible/UIA `Loading decision details...`; no stale progress or enabled incomplete action |
| P3 | actionable conflict | visible/UIA `Waiting for your decision`; question/context before actions; no marquee |
| P4 | open discovery | card provisional; aggregate indeterminate |
| P5 | sibling + waiting/conflict | active cohort |
| P6 | long failed note | wraps without overlap |
| A1 | conflict UIA | waiting state + question/context before actions; transition announced once |
| A2 | hidden UIA | no conflict progress/graph |
| L1 | locale/DPI/density | no clipping/overlap |

## 13. Verification sequence

### Focused

```powershell
.\build.ps1 -ProjectName RedSalamander
.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter FileOps_ProviderCapabilityMatrix -TimeoutMultiplier 2.0
.\Tools\Run-AllTests.ps1 -Suite Commands -CommandsFamily file-operations -SkipBuild -CaseFilter cmd_pane_fileops_popup_progress_contracts
```

Register stable names for new cases and run them explicitly during RED/GREEN.

### Owned suites and contracts

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild
.\Tools\Run-AllTests.ps1 -Suite Commands -CommandsFamily file-operations -SkipBuild
.\Tools\Tests\ResourceLocalizationContracts.Tests.ps1
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
git diff --check
```

### Performance

```powershell
.\build.ps1 -Configuration Release -ProjectName RedSalamander
# Run registered FileOps perf scenarios with capture enabled.
.\Tools\Test-TestRunArchive.ps1 -RunPath <candidate-run-folder>
.\Tools\Test-TestRunArchive.ps1 -Inventory
```

### Release closeout

Do not run this command solely for I17. The sole live reliability-first successor
owns the exact merged E8 closeout run, and that one receipt qualifies both its
change and I17:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh
```

Expected exit 0 with receipt bound to that successor closeout commit. Affected,
Resume, Commands-only, or FileOps-only does not replace Fresh Full.

## 14. Authoritative spec updates

- `Specs/FileSystem/FileSystem_FileOperations.md`: typed evidence precedence,
  failed-before-publication vs Unknown, separate owned-stage artifact truth,
  single-assignment item results, current fail-closed binding, alias-safe
  leaves, Move ordering.
- `Specs/Plugins/Plugins_VirtualFileSystem.md`: conditional outcome truth,
  rename/delete/abort reconciliation, zero-ID Unsupported rule, lifetime
  handle capability, capability honesty, forbidden weak reopens, creation and
  non-bridge cleanup truth.
- `Specs/UI/UI_FileOperationsPopup.md`: Waiting/conflict visibility,
  explicit waiting/loading indicator, no marquee, active-cohort/rebase,
  prompt/UIA order and transition announcements, wrapping.
- `Specs/Testing/Testing_SelfTests.md`: injections and case ownership.
- `Specs/Testing/Testing_PerformanceValidation.md`: metrics, admission budgets,
  conditional-index trigger, scenarios, archive links.

Do not say `FileIdInfo` caused the live failure until G0 proves it. Specs can
state behavior for any equivalent remote identity failure.

## 15. Implementation evidence

- Implemented in source commit
  `0baaede4a8fc222e00d296b29e57f3c463cab16b`, descended from
  `9d375606399c109487188b32569e47c3a62e9b04`, then merged onto the menu-correct
  `master` baseline.
- Debug x64 build attestation before the master merge:
  `39128875daeb04178300e62d51e26003b49ab970945c8049aab62e4de237f1aa`
  with zero warnings and zero errors.
- Test-enabled Release x64 build attestation before the master merge:
  `1f5741dffff05f58648429f983c646184ba47f64353e4327b17a327abdd1347e`
  with zero warnings and zero errors.
- Broad Debug FileOps run
  `20260829T014349Z-106564-d45629cebe5a44dbaeab64a763c4895c`:
  123 passed, 0 failed, 20 declared/environmental skips, disk audit clean.
- Cleanup-unknown overwrite fixture repeat
  `20260829T014306Z-23624-fb14b557f4b242d2adab92c37a0ec675`:
  60/60 repeated assertions passed, disk audit clean.
- Commands-family run
  `20260828T233319Z-111812-d8d11f252a894932a2cded972df10c56`:
  32/32 passed.
- The exact Cartesian Release admission scan crossed the optimization trigger
  at 65,536 comparisons / 2,600,470 microseconds in
  `Specs/TestRuns/4cb089111a23/FileOps/2026-08-29_022850/`. The accepted exact
  retained-anchor/suffix index completed with 8,704 probes / 2,689
  microseconds in
  `Specs/TestRuns/4cb089111a23/FileOps/2026-08-29_024402/`; archive validation
  and the whole-inventory validation passed.
- Deterministic fault injection attributes the representative post-create 996
  to `StageIdentity` and separately exercises stage-abort/reconciliation
  uncertainty. The historical live `\\rubygloom` endpoint/profile was not
  supplied, so this closeout does not claim that its original 996 was observed
  live or came from `FileIdInfo`:
  **[environmental coverage: live SMB endpoint unavailable]**.
- Detached-tree Fresh Full
  `20260829T035127Z-69832-68ff000af1414494a77c9d6f1abb8878`
  passed 17/18 entries, including File Operations 123/123 and Tools Pester
  711/711, but was not a valid final merge receipt because it predated the
  menu-correct master Commands shutdown/isolation fixes. Exact Fresh Full remains
  intentionally pending for the consolidated successor E8 closeout gate stated at
  the top of this record.

## 16. Done criteria

- [x] G0 names the failing phase with deterministic evidence.
- [x] Every post-mutation terminal path stores typed result.
- [x] Every terminal item stores exactly once; named early exits and
      Continue-on-error double-store regressions pass.
- [x] No byte/HRESULT publication heuristic remains.
- [x] Rename/delete/abort/rollback cannot return false-known non-commit.
- [x] Final publication and owned-stage artifact truth are separate.
- [x] Bridge and non-bridge exclusive creation cannot discard cleanup/artifact
      truth.
- [x] All-zero `FILE_ID_128` is Unsupported and never compared/cached.
- [x] Destination-profile capability cannot select Managed without safe exact
      stage authority or a pre-mutation fail-closed fallback.
- [x] SMB stage uses retained-handle exact authority or fails closed.
- [x] Existing mapped-drive/UNC Indeterminate binding fails before mutation
      even if leaf-level Parallel remains disabled.
- [x] Same absent alias target serializes; distinct siblings run concurrently.
- [x] Bind Indeterminate fails before mutation.
- [x] Large-selection cost evidence and admission budgets pass; indexing is
      added only if the recorded trigger is met.
- [x] Waiting/conflict exposes no transfer/discovery animation.
- [x] Actionable conflict visibly and accessibly says
      `Waiting for your decision`; metadata loading says
      `Loading decision details...`; neither uses a marquee.
- [x] Provisional/aggregate rules match UI spec.
- [x] Question/context and last note are wrapped and accessible.
- [x] Five locales pass resource contracts.
- [x] Pre-merge FileOps is 123/123 and focused Commands File Operations is 32/32;
      final merged FileOps/Commands/resource confirmation is explicitly transferred
      to the consolidated successor gate, not claimed complete here.
- [x] Release perf archives validate under the canonical path.
- [x] Spec inventory and `git diff --check` pass before merge; exact-commit rerun is
      transferred to the consolidated successor gate.
- [x] Consolidated Fresh Full obligation is transferred to the successor closeout
      commit; no green receipt is claimed by this archive.
- [x] Authoritative specs contain all lasting behavior.
- [x] I17 row is removed; plan moves to Done as a retired implementation record.

## 17. Global STOP conditions

Stop and report if:

- destination publication cannot be proven before Move source deletion;
- only identity fallback is weak/reopenable;
- an all-zero file ID is proposed as persistent identity;
- cleanup/rollback truth cannot fit current ABI and no size-versioned
  result-bearing extension is added;
- a terminal item can be stored twice or can reach fallback synthesis through
  a known early exit;
- stage cleanup truth is represented by corrupting an exact final-publication
  state;
- `exclusiveStage` remains true for a destination profile that cannot return
  exact authority and cleanup truth, with no proven pre-mutation fallback;
- alias safety treats identical absent targets as unrelated;
- anyone proposes “0 bytes means no mutation”;
- raw 996 is globally redefined/suppressed without phase provenance;
- managed post-publication Unknown is weakened;
- rendered/hosted/debug/UIA states cannot share one decision;
- an actionable human-in-the-loop card has no stable visible/UIA waiting
  indicator or uses animation to imply waiting;
- in-scope files have conflicting unowned edits;
- focused verification fails twice after one reasonable correction;
- the scheduled successor closeout Fresh Full fails twice after one reasonable
  correction. Leave both owners ACTIVE and record the exact failing receipt; the
  intentional consolidation interval by itself is not a blocker.

## 18. Review focus

Reviewers explicitly sign off:

1. rename/delete/abort/rollback truth at every exit;
2. zero-ID rejection, destination capability honesty, and retained-handle
   authority boundaries;
3. separate final-publication and stage ownership/artifact truth after
   post-create failure in bridge and non-bridge paths;
4. exactly-one item-result store across all named early exits and
   Skip/Cancel/Continue/finalization;
5. source-delete dependency on publication + verification + safe artifact
   disposition;
6. current mapped-drive/UNC fail-closed policy and missing-leaf anchor policy;
7. interlock bind/comparison evidence and whether indexing is actually needed;
8. paint/hosted/debug/UIA parity;
9. persistent waiting/loading indication, prompt order, one-time UIA
   announcement, and locale/DPI wrapping;
10. authoritative-spec merge and perf archive.
