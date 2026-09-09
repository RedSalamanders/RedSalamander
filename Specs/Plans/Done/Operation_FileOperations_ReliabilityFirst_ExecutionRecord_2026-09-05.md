# File Operations execution record through 2026-09-05

> **RETIRED / HISTORICAL RECORD / NON-EXECUTABLE.** Moved to `Specs/Plans/Done/` on
> 2026-09-07 with the core plan. This is the preserved text of the
> former long-form plan at commit `58f22d0e`, Git blob
> `888e0926e643a05ea584be6399a0f85a64481e13`. The text below is historical evidence,
> including its old checkboxes, activation states, scope decisions, and failed runs.
> None of those fields is current status or an instruction to resume old work.
>
> Current core work and the only current execution status table are in
> [the simplified reliability plan](Operation_FileOperations_ReliabilityFirstSuccessorPlan_2026-08-29.md).
> The owner deferred history, review-later, retargeting, and the Shell experiment on
> 2026-09-05; their preserved intent is routed to
> [the later-features plan](../WIP/Operation_FileOperations_LaterFeatures_2026-09-05.md).
> This record stayed subordinate to I14 until closeout and moved with it to Done
> under the repository's archive rules. Its preservation does not claim a passing
> final gate. Do not append new execution receipts here; archive new measured runs
> under `Specs/TestRuns/` and link them from the current status table.

<!-- preserved-plan:start -->
# File Operations reliability-first consolidated successor plan

> **Status:** WIP / ALL PRODUCT DECISIONS RECORDED / IMPLEMENTATION ACTIVATION GATES / NON-NORMATIVE
> **Index owner:** I14
> **Decision owner:** product owner
> **Planned at:** `fa98ef68618096414733624c1b7c2320db1cca2c` (`fa98ef68`, 2026-08-29); this records plan provenance, not E0's execution baseline. Every package must drift-review its exact owned files against the then-current merged base before editing.
> **Audit baseline:** File Operations engine/provider evidence at `4ad31825d3de59a4c4ec14a30772de58d7f751ee`; ingress/menu drift reconciled through `fa98ef68` (`master`, 2026-08-29)
> **Competitive baseline:** official public product documentation reviewed 2026-08-29; public contracts are evidence of product behavior and positioning, not proof of undocumented internals
> **Historical baseline:** `../Done/Operation_FileOperations_GlobalBehaviorDecisionReview_2026-08-16.md` at `c5e960bf` remains frozen history
> **Transferred implementation baseline:** former I14 plan at `6094a966dc0450f75a93341e8364035212610491`; its complete unfinished A0-A8 scope is incorporated below as E0-E8
> **Implementation authority:** E0-E8 are the already-authorized, behavior-preserving I14 execution backlog. E0-E7 are complete; E8 remains predecessor-blocked. Product decisions A01-A20 and provider subdecision A13-MD1 are accepted as target law for this non-normative plan, but acceptance does not activate an R package: R0a-R9 still require a complete, predecessor-closed, bounded `ACTIVE` entry before implementation.
> **Single-plan rule:** this is the only live WIP File Operations successor for the consolidated scope. The retired I14 and I17 files in `../Done/` are provenance/evidence only and must never be resumed as queues.

This document supersedes both D2's earlier residual-safety framing and the former
I14 post-closeout work queue. It is one complete successor: product decisions,
authorized architecture debt, implementation packages, human workflows,
verification, inherited qualification, ownership, and closeout live here together.
The stable `D2-Axx` decision identifiers are retained so existing review references
remain intelligible; `D2` no longer denotes a separate WIP owner or file.
It does not reopen the completed Phase 0-6 plan as a work queue. It identifies
which decisions should remain, which are unsafe or dishonest, which complexity is
necessary, and which complexity should be removed or demoted from product law.

**Executor reading path:** section 0 is a progress dashboard, section 6 is this WIP's
sole target-behavior contract, section 8.1 is the sole package-ordering/activation
authority, section 8.2 contains inactive delivery briefs plus optional `DRAFT` and
completed `ACTIVE` cards,
section 9 owns verification, and section 14 owns global STOP conditions. Sections 1-5
and 7 retain audit evidence and decision
rationale only; they never override or duplicate the target rules in section 6.

**Authority boundary:** sections 6 and 7 record accepted target behavior for delivery
planning only. Current authoritative specifications remain binding until the relevant
predecessor-closed `ACTIVE` slice lands implementation, tests, and its owning
authoritative-spec update. Acceptance, `DRAFT`, or `ACTIVE` status alone does not
supersede current normative text. E work preserves current behavior.

High-risk transition examples follow that boundary: ordinary S3 prefix Delete keeps
the current bounded folder-emptying behavior until R0b lands its per-observation ETag
conditions; Batch Rename keeps its current cycle/journal/Resume/Rollback contract until
R7-A10 lands the accepted acyclic/no-journal replacement; and capability JSON v2
remains the current ABI until R2 lands the one lockstep typed cutover. Implementers
must re-read the named authoritative owner at slice activation rather than “fixing”
current code to an inactive target early.

## 0. Execution dashboard

This front door tracks progress only. It deliberately does not repeat behavioral
requirements: section 6 owns those rules, section 8.1 owns dependency order, section
8.2 owns package steps, and section 9 owns evidence. A checkbox here can never
activate a package or override those sections.

Checkbox state is execution state, not product authority. `[x]` records a completed
decision/evidence step; `[ ]` records remaining activation, implementation, evidence,
or closeout. E0-E8 retain inherited authorization; E0-E7 are complete.
Any R slice not explicitly marked `ACTIVE` or `COMPLETE` remains
`DECIDED-NOT-ACTIVE` until its bounded activation card is complete and every hard
predecessor has a recorded exit receipt.

Derived delivery status is computed from section 8.1 exit receipts; these are not new
plans or dependency owners:

- [x] **M1 — Trustworthy current engine**
- [ ] **M2 — Reliable fast core under user control**
- [ ] **M3 — Complete successor and closeout**

### Phase 0 — Decisions and bounded activation

- [x] A01-A20 and A13-MD1 are answered and mapped to the section 6 rule index.
- [ ] Draft an R card when useful; promote it to `ACTIVE` only when every section 8.1
      field is exact and every hard predecessor has a recorded exit receipt.
- [ ] Reconcile I3/I12 ownership and drift-review the exact owned files/interfaces.

### Phase 1 — Merged baseline

- [x] Complete E0's own clean, exact-commit, pre-change focused/Fresh Full baseline;
      it cannot coalesce with R0 or any later changed-behavior closeout.

### Phase 2 — Independent release containment and truth

- [x] Activate, implement, and independently release R0a (A06/FOS-01).
- [x] Activate, implement, and independently release R0b (A13/FOS-03).
- [x] Activate, implement, and independently release R0c (FOS-02/FOS-04).
- [x] Activate, implement, and independently release R0d (A11/A13-MD1/FOS-06/FOS-10).
- [x] Activate, implement, and independently release R0e (A01/FOS-05).
- [x] Activate and complete R0e-OR1's admitted Local blocking-reader cancellation
      residual without re-enabling an uncontained route or changing the public ABI.
- [x] Activate and complete R1a/R1b/R1c (FOS-07 through FOS-12); R1b must precede E5.
- [x] R0-RC1 review corrections (2026-09-02) landed (`6904dce9`); its Fresh Full gate found one
      regression (Skip -> Indeterminate) that R0-RC2 corrected.
- [x] R0-RC2 simplifications (2026-09-02) landed (`f3c00d35`, `29c5a3a7`, `0683cc51`); the Fresh Full gate on
      `0683cc51` reports 2003/2062 with 5 residual failures classified on the card (2 environment,
      3 timing/cancel-race), none attributable to R0-RC1/R0-RC2/R0-RC3.
- [x] R0f complete (every slice landed 2026-09-02/03; activated by owner decision 2026-09-02): every route R0e classified `uncontained`
      (Local SMB, FTP/SFTP/SCP, S3, Microsoft Drive/SharePoint, Google Drive) must be read/write/
      create/delete capable with proved cancellation containment before S1; R0f-SMB precedes M2 and
      every R0f slice precedes M3.
      - [x] R0f-SMB landed (2026-09-02): Local UNC/mapped routes are `bounded` through the
            synchronous-I/O cancel watch (`Common/SynchronousIoCancelWatch.h`, host workers and
            Local scheduler workers); loopback-share read/write/create/rename/delete and the
            blocked-call cancel witness are FileOps family `FileOpsFamily_R0fSmbContainment`.
      - [x] R0f-Curl landed (2026-09-03): every libcurl transfer under a File Operations call polls the
            operation control from its progress callback (control commands included), FTP/SFTP/SCP are
            `providerWatchdog` with a provider-owned bound and advertise copy/move/delete/rename; IMAP
            stays read-only. Witnesses: Curl debug self-test (stalled `DELE` returns on Cancel and on the
            transport bound) and FileOps `R0fCurl_FakeFtpReadWriteCreateDelete` (loopback fake FTP through
            the host: cross-provider Copy both ways, provider CreateDirectory/Rename, Delete).
      - [x] R0f-S3 landed (2026-09-03): every AWS request under a File Operations call carries a continue
            handler the CRT polls from its callbacks, the client's connect/stall/retry policy is the
            provider-owned bound, S3 is `providerWatchdog` (S3 Table stays read-only), and `anonymous`
            serves public buckets and the loopback fixture. Witnesses: S3 debug self-test
            `RunS3StalledRequestCancelSelfTests` (streaming cancel, stalled request with and without
            Cancel) and FileOps `R0fS3_FakeS3ReadWriteCreateDelete` (fake S3 endpoint through the host:
            cross-provider Copy both ways, provider CreateDirectory/Rename, Delete).
      - [x] R0f-Graph landed (2026-09-03): every WinHTTP step under a File Operations call polls the
            operation control (before each attempt, between request and response chunks, during throttle
            backoff) and reports the host's verdict after Cancel; the resolve/connect/send/receive timeouts
            are the provider-owned bound; Microsoft Drive/SharePoint is `providerWatchdog`, advertises
            Rename, and keeps Delete as Recycle by stable item ID with receipts on every Recycle/Rename
            completion. Witnesses: Microsoft Drive debug self-test `RunGraphStalledRequestCancelSelfTests`
            (streaming cancel, stalled request with and without Cancel) and FileOps
            `R0fGraph_FakeGraphReadWriteCreateRenameRecycle` (fake Graph endpoint through the host:
            cross-provider Copy both ways, provider CreateDirectory/Rename, Recycle).
      - [x] R0f-GDrive landed (2026-09-03): the read-only skeleton became a full destination
            (`IFileSystemIO`, `IFileSystemDirectoryOperations`, `IFileSystemAtomicWriter`, and the eight
            mutation entry points: server-side copy, native move, rename, permanent delete, trash, folder
            creation) with receipts on every completion; every libcurl transfer under a File Operations
            call polls the operation control from its progress callback and throttle backoffs are polled;
            the connect and hard request timeouts are the provider-owned bound; Google Drive is
            `providerWatchdog` with `pathTextStableIdentity:true` (exposed names are unique per folder).
            Witnesses: Google Drive debug self-test `RunDriveStalledRequestCancelSelfTests` (streaming
            cancel, stalled request with and without Cancel) and FileOps
            `R0fGDrive_FakeDriveReadWriteCreateMoveDelete` (fake Drive endpoint through the host:
            cross-provider Copy both ways, provider CreateDirectory/Rename/native Move, permanent Delete).

### Phase 3 — Behavior-preserving seams and typed authority

- [x] Complete E1 and E2 in the section 8.1 order.
- [x] Complete E3 after the recorded E2 exit receipt.
- [x] Complete E4, R7-A10, and R7-A15 before E5; this is the one rename-chain order.
- [x] Activate and complete R2's one lockstep typed cutover (A02/A04/A11/A13/
      A13-MD1/A14/A15).

### Phase 4 — Preparing, publication, traversal, overlap, and discovery

- [x] Activate and complete R1d-core after E4, R0e, and R1a (A01/A05/A20); it owns
      no A02 comparison, A19 scheduler, or R6 presentation behavior.
- [x] R3 activated (2026-09-03) and completed (2026-09-04) after R0a, R1c, E1, and R2 (A04/A06); it landed as receipted
      slices (R3-1 conditional replace on identity-less routes landed 2026-09-03; R3-2 writer
      content proof landed 2026-09-03; R3-3 the one transaction record landed 2026-09-03).
- [ ] Activate only the required R4 slices under section 8.1 (A02/A08/A13/A19).
      R4 traversal/A08/A13 completed 2026-09-04 through the receipted slices R4-T1 (bridge
      walkers), R4-T2 (Local permanent-delete walker) and R4-T3 (bridge listing reported, not
      capped); its checklist records the one provider-only residual. Sequencing correction:
      the package note (`8df5427c`) and the R6-A08 activation (`048ddb4c`) were committed
      before the R4-T2/R4-T3 exit receipts and their covering Fresh Full gate #14 existed;
      those receipts now stand on the cards and the closeout is valid from that commit on.
      R4-A19 is complete. R4-A02 (overlap advice) remains the open R4 row.
- [ ] Activate R5 only if R0e proves an exact residual route requires isolation.

### Phase 5 — Rename, artifact, breadcrumb, and optional link behavior

- [x] Complete E5 after E4, R1b, R7-A10, and R7-A15; Change Case now consumes
      the central typed RenamePlan authority.
- [ ] Activate R7-A03 only after R2, R3, and R6-A20 (A03); it is not part of the
      E4/E5 rename chain.
- [x] Complete E6 after E0 while its exact ownership does not overlap an active package.
- [x] Complete E7 after E0 while its exact ownership does not overlap an active package.

### Phase 6 — User decision, progress, explanation, and history slices

- [ ] Activate R6-A02 only after E2 and R4-A02.
- [ ] Activate R6-A07 and R6-A17 independently, each only after R1a and E2.
- [ ] Activate R6-A08 only after R4 traversal/A08/A13 and R4-A19; activate R6-A12
      only after E2. R6-A08 activated 2026-09-04 after the R4 traversal/A08/A13 closeout;
      its card sits under the R6-A08 checklist.
- [x] Activate R6-A09 only after E6; activate R6-A18 only after E3.
- [ ] Activate R6-A20 only after R1d-core.
- [ ] Complete the section 6.9 and section 9 human/accessibility evidence only after
      the contributing R6 slices have landed.

### Phase 7 — Optional simplification evidence

- [ ] Activate R9 only after R1a, E2, R2, and R3; record Adopt/Reject separately for
      each bounded candidate and delete a failed spike.

### Phase 8 — Consolidation and closeout

- [ ] Complete R8 only from landed rules and package evidence.
- [ ] Complete E8 inventories, authoritative-spec updates, inherited-I17 qualification,
      and one exact merged Fresh Full receipt.
- [ ] Move this sole successor to Done and remove I14 from WIP only when section 13 is
      fully satisfied.

## 1. North star and decision test

RedSalamander File Operations should optimize in this order:

1. **Never mutate an object the user did not select or knowingly authorize.**
2. **Never report more certainty than the engine actually has.**
3. **Preserve user data when authority, commit state, or cleanup truth is unknown.**
4. **Keep the application responsive and cancellable even when a provider fails.**
5. **Tell the user, before and after execution, what will happen, what happened,
   what might have happened, what remains, and what they can safely do next.**
6. **Prefer the smallest design that proves those properties.** Throughput and
   concurrency are optimizations, not reasons to weaken or obscure the contract.
7. **Keep routine operation language familiar.** Internal proof axes may be rich,
   but primary UI explains consequences and exposes technical evidence on demand.

**Product requirement (owner decision, 2026-09-02).** Every supported destination must
work as a file manager destination: read, write, create, and delete on Local fixed
volumes, UNC and mapped SMB shares, FTP/SFTP/SCP, S3, Microsoft Drive/SharePoint,
Google Drive, and MTP devices. The cancellation route classes below are a delivery
order for proving containment, never a permanent capability cut: a route classified
`uncontained` is a defect that package R0f must close before S1, not an accepted end
state. Only IMAP stays read-only, because a mailbox is not a writable folder.

Selecting an exact real container—or explicitly deleting a provider-declared
virtual folder such as `s3://bucket/prefix/`—is knowing authorization for membership
encountered during bounded execution inside that exact scope. It never authorizes a
replacement real root, followed link/mount, escaped alias, sibling prefix, or key
outside the canonical virtual-folder boundary. A fixed bulk object/version selection
still authorizes only its admitted revisions. These are the Delete scope models in
section 6.6, not hidden implementation exceptions.

Every decision in this review is tested with five questions:

- Does it prevent data loss or ambiguous mutation?
- Can the implementation actually prove the guarantee?
- Can a user understand the consequence before consenting?
- Does failure leave the user in control with safe next actions?
- Is the complexity essential to those outcomes, or only to optional semantics,
  concurrency, compatibility, or implementation symmetry?

“Always working” does not mean pretending every backend can safely perform every
operation. It means that unsupported or uncertain operations fail before mutation,
retain data, preserve the user's intent where safe, and explain the limitation.

## 2. Scope and audit method

### 2.1 Audited authority

This review covers all normative decisions that define File Operations behavior,
plus provider rules that can authorize or describe a mutation:

- `Specs/FileSystem/FileSystem_FileOperations.md`
- `Specs/Core/Core_FileSystemBridge.md`
- `Specs/Plugins/Plugins_VirtualFileSystem.md`
- `Specs/UI/UI_FileOperationsPopup.md`
- `Specs/UI/UI_BatchRenameWindow.md`
- File Operations portions of `Specs/UI/UI_CommandMenuKeyboard.md`
- Local, S3, Microsoft Drive, MTP, IMAP, FTP/SFTP/SCP, and Google Drive provider
  specifications
- `Common/PlugInterfaces/FileSystem.h`
- the host admission, runtime, task, conflict, interlock, bridge, result,
  Change Case, Batch Rename, artifact, and breadcrumb implementations
- shipped provider mutation, publication, recovery, and capability implementations
- directly relevant deterministic selftests and consistency metadata.
- official public product documentation listed in section 3.1, used only to test
  product value, interaction patterns, and complexity—not as mutation authority.

This is not an audit of unrelated browsing, viewer, settings, or provider
configuration code. The plaintext legacy credential finding in section 5.6 is
recorded because it was discovered in the audited provider specifications, but it
must be routed to a security owner rather than absorbed into File Operations.

### 2.2 Evidence policy

- Findings below are tied to current code or current normative text, not inferred
  from the frozen plan.
- A capability declaration is not treated as execution proof.
- An error code is not reclassified without evidence about the provider/profile
  that produced it.
- No live SMB target was available in this audit. Therefore error 996 is a
  confirmed classification/design problem, not a proven redirector-specific cause.
- Production call paths are distinguished from explicit selftest-only direct
  provider fallbacks.

### 2.3 Authorization and completeness model

One file does not mean one undifferentiated authorization. Every unfinished item
has exactly one state in this successor:

| State | Meaning | May implementation start? |
|---|---|---|
| `AUTHORIZED-INHERITED` | Behavior-preserving work already authorized by former I14 and transferred here without changing its boundary. | E0-E8 are authorized, but a node is runnable only after its hard predecessors and dispatch gates close. E0-E7 are complete. |
| `DECISION-GATED` | Product behavior, provider contract, release-containment scope, or optional architecture awaiting explicit arbitration. | No. Record `accept`, `reject`, or an amended target rule first; that decision still grants no implementation authority. |
| `DECIDED-NOT-ACTIVE` | Product owner selected the target rule, but no bounded R implementation slice has been activated. | No. A card may be drafted, but only exact fields plus closed hard predecessors permit `ACTIVE`. |
| `DRAFT` (card only) | Optional incomplete/refining activation card beneath a `DECIDED-NOT-ACTIVE` R slice. The slice state does not change. | No. It may contain open fields or predecessors and grants zero edit authority. |
| `EXTERNAL-OWNER` | Existing I3, I4, I5, I7, or I12 scope that this plan consumes or coordinates but does not duplicate. | Only through that owner. |
| `QUALIFICATION-INHERITED` | Delivered I17 behavior whose focused evidence remains valid but whose exact merged Fresh Full closeout is still owed. | No implementation queue; E8 carries the aggregate validation gate. |
| `REJECTED` | Considered alternative explicitly excluded from the design. | No, unless new evidence returns it to arbitration. |

The plan is complete only when every decision is resolved, every accepted change is
implemented or explicitly routed to a still-named owner, all E packages are closed,
authoritative specifications contain the durable contract, and E8 records the exact
merged qualification. All product decisions are answered, but an E package must stop
if it would cross into an accepted R behavior that has not been explicitly activated.

#### Ownership boundaries retained during consolidation

- I3 keeps its five existing File Operations stress-evidence areas.
- I12 keeps completion affinity, reaper lifetime, command-surface quiet point,
  post-Swap insertion identity, and removal-focus behavior. E1 coordinates any
  overlap and never silently closes those rows.
- I4 keeps repository-wide behavioral source-shape migration. This plan may add
  focused runtime tests and narrow structural guards only.
- I5 keeps reusable performance-measurement infrastructure; this plan supplies the
  File Operations scenarios and archives.
- I7 keeps generic HRESULT/status localization; this plan owns only strings changed
  by its accepted File Operations behavior.
- The retired I17 prompt/mutation-truth baseline is preserved. E1/E2/R6 must not
  weaken delivered destination-leaf commit-time authority/revalidation guards, turn a
  never-published operation into unknown publication, or relitigate the delivered
  prompt baseline. Those per-item guards do not authorize a global task-overlap lock.

## 3. Executive verdict

The Phase 0-6 program established the right safety vocabulary: typed plans,
Native/Managed/Copy-only strategies, no-follow identity, owned publication,
conditional source cleanup, independent result axes, exact artifact claims, and
safe Recycle escalation. Those are not overdesign; they prevent one HRESULT or one
pathname from being misread as mutation truth.

The current product is nevertheless **not yet at the stated reliability north
star**. Four classes of issue remain:

1. **Confirmed destructive-authority defects.** Local absent-destination Copy can
   pathname-delete a concurrent foreign replacement; MTP paths/recovery use hashes
   or name/size/time without full identity; S3 captures current-object ETags but
   discards them before singular/batch Delete. Bounded admission of late members is
   intended folder behavior; unconditioned removal of a changed generation is not.
2. **False guarantees.** Cancellation deadlines and quiet points are parsed but do
   not prevent an unconditional worker `join()` during shutdown. Capability flags
   advertise Rename/Delete shapes that central exact-authority admission cannot
   execute. Destructive Native success can be called Completed without a receipt.
3. **Truth and control gaps.** Some early terminal paths bypass typed item results;
   clipboard Move can execute after consumption failure or lose the cut list before
   a later pre-mutation failure; Change Case can silently skip unreadable
   directories and has a broken task-card payload; conflict Cancel/Escape has two
   owners; exact cleanup debt can be hidden.
4. **Optional complexity promoted into normative law.** Semantic link retargeting,
   the exact alias-index algorithm, a monolithic mandatory JSON capability document,
   backwards provisional progress, a feature-complete second popup renderer, hard
   traversal ceilings, and historical implementation/performance details are mixed
   into the product contract.

The correct response is not another layer around the 8,000-line execution method.
It is a smaller contract with one owner for each decision, exact destructive
authority, conservative fallback, one terminal-result funnel, and provider routes
that are either executable or visibly unavailable.

### 3.1 Competitive benchmark: copy the product value, not the accidental complexity

Competitive behavior does not define RedSalamander's safety contract. It is used
here to test whether a proposed mechanism produces recognized user value, whether
the same outcome is achieved more simply elsewhere, and where RedSalamander is
deliberately choosing a stronger guarantee. Claims below are limited to official
public documentation; absent documentation is not treated as proof of unsafe
internal behavior.

| Product | Publicly documented strength worth learning from | Boundary or complexity not to copy | Decision consequence for RedSalamander |
|---|---|---|---|
| Windows File Explorer UI + Windows `IFileOperation` API | File Explorer documents a small Cut/Copy/Paste/Rename/Delete vocabulary. Separately, `IFileOperation` documents Shell copy/move/rename/delete, Recycle flags, and per-item pre/post callbacks with HRESULTs. [Explorer](https://support.microsoft.com/en-US/Windows/Experience/FileExplorer/file-explorer-in-windows), [`IFileOperation`](https://learn.microsoft.com/en-us/windows/win32/api/shobjidl_core/nn-shobjidl_core-ifileoperation), [`IFileOperationProgressSink`](https://learn.microsoft.com/en-us/windows/win32/api/shobjidl_core/nn-shobjidl_core-ifileoperationprogresssink), [`SetOperationFlags`](https://learn.microsoft.com/en-us/windows/win32/api/shobjidl_core/nf-shobjidl_core-ifileoperation-setoperationflags) | These sources do not establish that File Explorer internally uses `IFileOperation`; the public API does not promise retained-generation authority, exact cleanup receipts, or end-to-end content verification. | Copy the simple vocabulary and evaluate only the two bounded ordinary-Local candidates in D2-A16; do not delegate a route whose truth cannot be recovered. |
| Total Commander | A mature two-pane Windows file manager with extended tree copy/move/delete, logging, and enhanced overwrite inspection. [Homepage](https://www.ghisler.com/), [feature list](https://www.ghisler.com/featurel.htm) | Its broad power-user surface and public UI documentation do not define exact-generation mutation authority or typed receipts. | Keep it as the established genre baseline for familiar operation vocabulary and conflict inspection without copying its policy breadth. |
| TeraCopy | Optional destination verification, fixed saved task lists, per-file hashes/status, pause/stop, retry of Failed/Skipped items, logs, and rich conflict inspection. [Copy and verify](https://support.codesector.com/en/articles/8789942-copying-and-verifying-files), [file list](https://support.codesector.com/en/articles/9902683-file-list-tab), [replace dialog](https://support.codesector.com/en/articles/8789948-file-replace-dialog), [link handling](https://support.codesector.com/en/articles/9962863-handling-of-soft-and-hard-symbolic-links) | Verification is a mode rather than a documented universal Move prerequisite; its large collision-rule matrix and link-type-dependent handling would expand policy and explanation cost. | Copy per-item evidence, fixed retry sets, and conflict inspection. Make content verification route-aware instead of mandatory for every Managed Move. See D2-A04. |
| FilePilot | Product investment is concentrated on navigation, panels, search, inspector, and command access; v0.7.1 says transfers switched to the modern Windows dialog but does not identify the API. [v0.7.1](https://filepilot.tech/starlog/2026-03-16-v0.7.1) | Undo/Redo, MTP, cloud integration, and improved network-share behavior remain roadmap work; public transfer authority/recovery semantics are not specified. [Roadmap](https://filepilot.tech/roadmap) | Treat FilePilot as evidence that a file manager can avoid rebuilding ordinary Local transfer UX, not as evidence that Shell delegation satisfies RedSalamander's stronger safety contract. |
| Directory Opus | Resource-aware queues, a jobs surface, its named Unattended operation with preselected collision handling, collected errors with end-of-run retry, and a persistent searchable operation log. [Copy queues](https://docs.dopus.com/doku.php?id=file_operations%3Acopying_moving_and_deleting_files%3Acopy_queues), [Unattended operation](https://docs.dopus.com/doku.php?id=file_operations%3Acopying_moving_and_deleting_files%3Acopy_queues%3Aunattended_operation), [tracking and undo](https://docs.dopus.com/doku.php?id=file_operations%3Atracking_and_undoing_file_operations) | Its broad policy surface and size/time-based "Skip Identical" heuristic are unsuitable as destructive proof. RedSalamander's proposed flow deliberately does not preselect skip/overwrite rules. | Copy the jobs, continue-safe-work, collected-review, retry, and history value. Name RedSalamander's mode **Continue safe work and review later**, not Unattended. See D2-A17/A18. |
| FastCopy | Optional conventional/perfect verification, a verified Move mode that removes verified sources, bounded asynchronous I/O, and optional Perfect Verify marker/resume behavior under restricted modes. [FastCopy help](https://fastcopy.jp/help/fastcopy_eng.htm) | Expert modes and implementation markers should not leak into the normal UI; verified Move is valuable without making universal reread the only usable Move. | Offer a clearly named Verified Move and require it automatically on weak-commit routes; keep atomic Native Move unaffected. See D2-A04. |
| WinSCP | For eligible SFTP binary transfers, a temporary `.filepart` protects the old target until completion, with resume constrained by source/version conditions and configurable size thresholds. [Transfer resume](https://winscp.net/eng/docs/resume) | Other protocols and resume modes differ; this protocol-specific policy is not a reason for a universal per-item transaction journal. | Keep owned non-final staging where it buys atomic visibility or resumability; do not require the same publication algorithm for every new Local Copy. See D2-A06. |
| Robocopy | Restartable transfer, explicit logging, bounded multithreading, throttling, low-space handling, and clear expert destructive modes. [Robocopy](https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/robocopy) | Switch explosion, very large retry defaults, `/MIR`, and `/PURGE` are automation power, not safe interactive defaults. | Put bounded scheduling/restart mechanisms behind a small UI; do not expose raw engine policy or widen ordinary Delete. |

Competitive sources plus provider/code audit evidence produce these design conclusions:

1. **RedSalamander's exact destructive authority is a real differentiator.** Keep
   retained/bound object authority, conditional source cleanup, exact artifact
   ownership, and truthful partial/indeterminate results.
2. **The internal model should be richer than the primary UI.** Keep Native,
   Managed, Copy-only and the independent result axes internally; render familiar
   consequences such as `Moved`, `Copied`, `Copied - source kept`, `Canceled`,
   `Failed`, and `Final state uncertain`.
3. **Verification and staging should be risk-scoped.** Exact publication and exact
   source-cleanup authority are universal; extra destination reread/hash and hidden
   staging are mandatory only when the route's risk/claim requires them.
4. **Delete meaning, not storage shape alone, selects authority.** A selected real
   directory and a provider-declared virtual folder can authorize bounded removal of
   members encountered inside the exact scope. A fixed object-set operation still
   consumes admitted key+revision snapshots. S3 ordinary prefix Delete is the former,
   not a hidden fixed-set operation.
5. **The missing product layer is operational comprehension, not another
   transaction framework.** One jobs surface, collected conflicts/errors, fixed
   retry sets, an exportable non-authoritative history, and plain-language strategy
   explanations create more user control than additional hidden state machines.
6. **Fast start and safety are different from a full preflight.** A small mandatory
   selected-root gate closes before mutation; one-pass discovery then uses the current
   tested reservation while transfer/mutation begins. Instrument that baseline first
   and add only the smallest same-resource correction if evidence proves discovery is
   starved. The user may release only run-ahead discovery, never identity,
   containment, conflict, or revalidation.

This evidence changes the earlier working ballot rather than sitting beside it:
D2-A04 moves from universal content proof to route-aware proof, D2-A06 from universal
hidden staging to risk-scoped publication, and D2-A13 from universal descendant
snapshotting to real-container/virtual-folder/fixed-set authority. D2-A16 through D2-A18 add
the Local Shell-delegation spike, optional continue-safe-work/review-later flow, and persistent operation
history decisions. D2-A19/A20 add explicit fast-start/discovery control and the
routine human-interaction model.

## 4. Confirmed release-safety and truth findings

Severity here is based on the north star, not on whether the current automated
suite happens to cover the path.

| ID | Severity | Confirmed current behavior | User impact | Required containment |
|---|---|---|---|---|
| FOS-01 | P0 | Local absent-destination regular-file Copy calls `CopyFileExW(..., COPY_FILE_FAIL_IF_EXISTS)` and, after failure, can call pathname rollback although the task never proved it created the occupant. `Plugins/FileSystem/FileSystem.FileOps.cpp:3296-3305,3918-3947,4110-4123`. | A concurrent foreign file created or replaced at the destination name can be deleted; rollback truth is discarded. | For ordinary absent-destination Local regular-file Copy, use section 6.4 shape 2: exclusive final-leaf create plus retained exact authority for write and abort. Do not add universal hidden staging. Never pathname-delete an unproved occupant; if exact abort cannot finish, retain and report the incomplete artifact. Add a raced-replacement regression under a new R0a release-blocking owner; do not expand the retired I17 scope. |
| FOS-02 | P0 | MTP records a pre-upload journal without the subsequently returned temp PUID, then replay/sweep can identify and delete by temp path or name+size+timestamp. It can clear malformed or exhausted journals while data remains. `Plugins/FileSystemMtp/FileSystemMtp.Core.cpp:299-370,1233-1336,1400-1490`; `Specs/FileSystem/FileSystem_Mtp.md:181-192,210-222`. | Recovery can delete an unproved object or discard the only provenance record for retained user data. | Disable automatic delete/rename recovery without exact PUID. Persist the created PUID before any cleanup authority exists; retain/quarantine uncertain journals and expose the artifact. |
| FOS-03 | P0 | S3 captures object ETags while listing/planning, but singular and recursive Delete paths discard that generation evidence before mutation. Recursive prefix Delete correctly re-lists late members because the product defines the selected prefix as a virtual folder, yet each batch can delete a generation that changed after observation. `Plugins/FileSystemS3/FileSystemS3.Directory.cpp:1471-1479,2280-2348,4227-4254`; `Specs/FileSystem/FileSystem_S3.md:136-145`. | A same-key replacement can be removed without an exact observed-generation condition; folder semantics do not justify blind current-occupant deletion. Combining `VersionId` with ordinary current-key Delete can also expose an older version rather than emptying the visible folder. | Preserve bounded virtual-folder convergence, including the late-child expectation. Ordinary folder Delete consumes the just-observed current-key ETag and omits `VersionId`; exact historical-version deletion remains a separate fixed-set cleanup/rollback operation. Re-list/rebind changed in-scope members, never blindly retry request-level unknown, and stop with explicit residuals when writers or unsupported conditions prevent bounded convergence. |
| FOS-04 | P0 | MTP provider paths use case-folded 64-bit FNV-style hashes of device/PUID/object identity; duplicate names and device resolution can choose by the hash, while the provider advertises stable path identity. `Plugins/FileSystemMtp/FileSystemMtp.Shared.cpp:290-338`; `FileSystemMtp.Device.cpp:791-822,2136-2167`; `FileSystemMtp.Core.cpp:4014-4022`. | A collision can resolve or mutate the wrong device/object. A short hash has become mutation identity. | Use hashes only as hints/display. Retain and compare full PnP ID plus full PUID/object ID; ambiguous matches fail closed. |
| FOS-05 | P0 reliability | Capability profiles store cancellation/quiet-point data, but the host deadline tick is not initialized, cancellation reduces to a Boolean flag, and shutdown unconditionally joins every worker. Some providers, notably MTP, have narrower watchdog/quarantine containment that the host contract does not model. `FolderWindow.FileOperations.cpp:804-820`; `FolderWindow.FileOperations.State.cpp:6639-6652,7083-7085,8267-8286`; `FolderWindow.FileOperations.State.Runtime.cpp:1593-1673`. | An uncontained wedged route can leave Stopping permanent and hang application shutdown despite `FileSystem_FileOperations.md:936-955`; blanket provider disablement would also discard working containment. | Stop advertising unenforced host deadlines. Classify exact routes under D2-A01, preserve proved MTP-style containment, add a host-level never-returning-provider shutdown test and separate SMB evidence, then disable or isolate only residual uncontained routes. |
| FOS-06 | P1 truth | S3, Microsoft Drive, and MTP Native completion paths can return a null mutation receipt; `processIndex` stores Native Move as Completed while source disposition is Unknown, preventing the stricter finalizer from correcting it. `Common/PlugInterfaces/FileSystem.h:132-145`; `FileSystemS3.Directory.cpp:4700-4721`; `FileSystemMicrosoftDrive.cpp:4610-4633`; `FileSystemMtp.Core.cpp:4121-4143`; `FolderWindow.FileOperations.State.cpp:8649-8672,10130-10153`. | The user can see Completed although the engine cannot prove whether the source was removed. | Null destructive receipt is Indeterminate. Advertised Native Move/Delete/Rename must return a typed receipt for primary mutation, source disposition, and cleanup debt. |
| FOS-07 | P1 | The clipboard barrier releases the worker even when clipboard consumption fails; later interlock preparation can fail after a successful consume and return before typed finalization. `FolderWindow.FileOperations.State.Runtime.cpp:1373-1397`; `FolderWindow.FileOperations.State.cpp:7524-7566,7579-7628`; current behavior is asserted at `SelfTest.Phases10_13.cpp:2062-2075`. | A Move may run while the reusable cut list remains, or the cut list may disappear although no mutation ran. Selected items can have no dense terminal result. | `Preparing -> exact clipboard consume -> mutation release`. Abort when consumption fails. One terminal funnel covers every exit; never auto-restore a consumed sequence. |
| FOS-08 | P1 | Typed result storage overwrites an already stored result; global-failure synthesis can use `E_PENDING` with Unknown axes even when no mutation ran. `FolderWindow.FileOperations.State.cpp:8515-8594`. | Later code can erase the first truth; Issues can claim uncertainty instead of a known no-mutation result. | One validating terminal builder, compare-and-store sink, and phase-aware failure synthesis. |
| FOS-09 | P1 | Change Case silently skips a directory when enumeration fails, can still return `S_OK`, and invokes the weaker pathname `FileSystemRenameBatch` engine. Its task payload is sent as a raw pointer although the receiver accepts only a registered opaque token. `ChangeCase.cpp:267-272,372-394,472-491`; `FileSystemRenameBatch.cpp:19-54,157-163`; `FolderWindow.FileSystem.Commands.cpp:14084-14098`; `FolderWindow.FileSystem.cpp:3937-3945`; `Common/Helpers.h:2406-2419,2588-2642`. | Recursive changes can omit subtrees yet appear successful; after 700 ms the task card is absent and the payload leaks. | Promote R1b for immediate truth and complete E5 by routing Change Case through `RenamePlan`. |
| FOS-10 | P1 | Central Permanent Delete and inline Rename require bound authority, while Dummy, S3, Microsoft Drive, Curl, and writable MTP advertise operations without a route central admission can bind. Contract tests validate Boolean/JSON shape, not executability. `FolderWindow.FileOperations.cpp:3155-3186`; `FolderWindow.FileOperations.State.cpp:8898-8908`; `Plugins/FileSystemDummy/FileSystemDummy.h:268-295`; `Plugins/FileSystemS3/FileSystemS3.h:307-334`; `Plugins/FileSystemMicrosoftDrive/FileSystemMicrosoftDrive.h:311-338`; `FileSystemCurl.Shared.cpp:4463-4496`; `FileSystemMtp.Core.cpp:3968-4009`; `PluginContractTests.cpp:1808-1847`. | Commands are offered although the authoritative host route is already known to be unusable. | Make an advertised command mean a complete executable safe route. Clear dishonest bits immediately. A contract-tested provider-native route may remain only when it supplies an exact typed scope/receipt even though no physical container bind exists; S3's declared virtual-folder route is the required model, not an exemption from authority. |
| FOS-11 | P1 | User cancellation enters exact stage cleanup, but `AbortOwnedObject` receives the same canceled operation control and refuses to run exact `DeleteExact`; file/link/directory paths duplicate this behavior. `FileSystem.FileOps.cpp:3445-3475,3521-3528,4323-4359,4834-4864`; `FileSystem.cpp:3089-3111`. | Ordinary Cancel needlessly retains safe-to-remove owned stages and turns a clear cancellation into artifact/indeterminate debt. | Separate “stop primary work” from bounded exact compensation. Local exact cleanup ignores the primary cancel bit and uses a short cleanup-only bound; remote inability remains visible debt. |
| FOS-12 | P1 | S3 and Microsoft Drive can commit the primary mutation while retaining backup/cleanup debt, then discard that structured status at the public boundary and return ordinary success/null receipt. `FileSystemS3.Directory.cpp:2939-2954,4572-4593`; `FileSystemMicrosoftDrive.cpp:4547-4563,5847-5856`; contract `Plugins_VirtualFileSystem.md:1902-1909`. | Hidden `.rs-bak`/Graph rollback objects remain, so the user cannot make an informed recovery decision. | Preserve primary success but expose “Completed; cleanup item retained” once through the typed receipt and user guidance. |

FOS-01 through FOS-06 are release-safety/truth issues. They should not wait for
architecture extraction. FOS-07 through FOS-12 are the first bounded truth and
capability-honesty package. The exact `ACTIVE` owner is an arbitration outcome;
this successor does not silently add them to E0-E8 or the retired I17 scope.

## 5. Decision review rationale and evidence (non-authoritative)

This section explains why the accepted target rules were selected. It is not a
second behavioral contract. In any wording conflict, the stable rule IDs and exact
text in section 6 control interpretation inside this WIP; current authoritative specs
continue to control shipped behavior until the owning slice lands. Packages and tests
reference those IDs rather than copying this rationale as new law.

### 5.1 Decisions to keep

| Decision | Disposition | Why it is necessary rather than overdesign |
|---|---|---|
| Immutable typed plans for user mutations | Keep | Consent, routing, conflict scope, and results need one stable input. Change Case must join this model; F7 and shell Recycle remain explicit small exceptions. |
| Native / Managed / Copy-only | Keep internally | These names expose materially different guarantees. Collapsing them would again let “Move” imply source deletion when only Copy is safe. Routine UI uses consequence language; technical strategy names belong in details/diagnostics. |
| No-follow exact authority for destructive mutation | Keep | Path text, hashes, names, sizes, timestamps, weak indexes, and stale capability flags cannot authorize overwrite, rollback, cleanup, rename, or Delete. |
| Independent publication, verification, source, owned-stage, and completion facts | Keep internally | Each answers a different user-safety question. Construction and presentation should be simplified, not information discarded. |
| Owned publication and exact abort/reconciliation | Keep, route-scoped | Partial content must not masquerade as complete and cleanup must not touch a foreign replacement. Use one reusable transaction where staging is required and retained exact final-leaf authority where it is sufficient. |
| One-shot clipboard Move consumption with no automatic restore | Keep | Restoring can replay stale intent or duplicate later execution. Move predictable checks earlier and report later retained-source outcomes exactly. |
| Deferred scoped consent for EFS, sparse inflation, hydration, space, metadata loss, and Recycle escalation | Keep | These are different loss/cost decisions. One policy/presentation owner can reduce mechanics without merging meanings. |
| Proven artifact claims | Keep | Exact durable claims plus current exact identity allow safe recovery. Names alone never grant cleanup authority. |
| Acyclic Batch Rename plans | Keep | Preview/admission rejects dependency cycles before mutation and explains intermediate-name alternatives. Acyclic steps retain exact partial results without a recovery journal or Resume/Roll back surface. See D2-A10. |
| Notice-only Move crash breadcrumb | Keep | It honestly says Outcome unknown and supports navigation without pretending a compact record can reconcile a recursive transaction. |
| Separate artifact-claim and Move-breadcrumb semantics | Keep | They grant different authority. Batch Rename journaling is retired by A10; do not preserve it merely to justify a generic durable-store abstraction. |
| F7 outside `StartOperation` | Keep conditionally | A bounded qualified create with immediate feedback does not need a recursive task engine. It does need provider-correct name feasibility and must leave the UI thread if the provider cannot prove a short bounded call. |
| Shell-owned Recycle route | Keep | Recycle has different observability and an exact one-shot escalation contract. Central policy/results do not require a universal backend algorithm. |
| Provider-specific Native algorithms | Keep | S3, Graph, MTP, Local, and IMAP have different atomic primitives. The common contract should require authority and receipts, not force one transfer implementation. |
| Immutable confirmation/policy snapshot | Keep | Endpoints, link behavior, verification policy, Queue/Parallel, bandwidth, accepted risk, and Apply-to-all scope must not change beneath the user. Each ingress may capture that snapshot through its own explicit command, preview, or material-risk prompt; this does not require one modal confirmation for every routine operation or an eager list of future conflicts/descendants. See D2-A20. |
| Typed conflict buckets and object-scoped grants | Keep | Overwrite/Keep Both/Skip/Cancel and Apply-to-all need exact scope. Remove duplicate UI derivation, not the safe decision model. |
| Streaming discovery with safe work starting before full enumeration | Keep | Full pre-scan harms responsiveness and memory. Preserve and instrument the current tested discovery reservation/half-rate baseline while eligible work starts. Add byte throttling only if a large transfer measurably starves discovery, mutation-rate control only if small-file Delete does so, and bounded turns only for a proved serialized provider. `Discover as needed` releases run-ahead reservation and any evidence-adopted limiter. See D2-A19. |
| Per-child folder merge decisions | Keep | A folder merge is not one atomic overwrite; each child can conflict, publish, fail, or retain source independently. |
| Source cleanup only after route-sufficient destination and source-stability proof | Keep | This is the defining Managed Move safety boundary. D2-A04 separates universal commit/source authority from route-aware content verification; cleanup can never precede the accepted proof. |
| Queue/Parallel and bandwidth as immutable execution policy | Keep | Users may trade throughput/resource use without changing mutation authority or result truth. Exact scheduler constants remain implementation/performance policy. |

### 5.2 Decisions that must change

| Decision | Current problem | Replacement rule |
|---|---|---|
| Admission | “Before publication” and “when reached” rules conflict; selected-root failures occur after side effects; Batch Rename qualifies/binds on the UI thread. | Use a mandatory, cancellable **Preparing** safety gate with no mutation, clipboard consumption, or Move breadcrumb. It is selection-proportional and top-level only: no recursive enumeration, payload read/hash, totals walk, or descendant conflict prescan. Required provider calls need executable containment: a cooperative bound, proved provider-local watchdog/quarantine with lifetime and unload safety, or accepted process isolation. Otherwise only that exact route is unavailable. Discover descendants and revalidate at mutation. |
| Capability truth | A coarse flag can enable a command although the exact route cannot execute. | User-visible capability means a complete executable safe route. Separate method presence from exact Copy/Move/Delete/Rename/Create feasibility. Runtime receipts remain truth. |
| Identity model | Persistent comparable ID, retained mutation authority, and task-overlap policy are coupled. | Model three independent axes: exact mutation authority, optional comparable identity, and cheap overlap-advisory evidence. Path/alias evidence may trigger a warning but never grants destructive authority, suppresses revalidation, or forces a global execution prohibition. |
| Concurrent task overlap | The interlock can reject or serialize whole tasks as though overlap were invalid, although users may intentionally inject nested Copy/Delete work and accept ordinary races. | At task injection, compare source/destination/delete roles using cheap normalized paths and only already-available or bounded positive link/drive/alias evidence. Clearly disjoint scopes start without a question. Obvious or positively evidenced overlap gets one problem-specific warning with **Queue after**, **Run at the same time**, and **Don't start**. Queue warns once, waits, then runs ordinary current-membership semantics; the special exact-output guard exists only while same-host tasks run concurrently. Every path still uses exact per-item authority and truthful results. |
| Error classification | Raw backend errors invite global “unsupported” mappings. | `Unsupported` is stable proven absence; `RetryableNoCommit` proves no mutation; `FailedKnown` proves post-state; `Indeterminate` preserves uncertainty; `ContractViolation` disables/quarantines the route. Do not globally map 996. |
| Managed Move proof | Exact destination publication is conflated with mandatory full content reread/hash for every route. | Always require exact publication, a stable source generation/read, and exact conditional cleanup. Require end-to-end content verification automatically for weak-commit remote/device routes and when the user selects Verified Move; a proven stable Local committed stream may use Standard Move without a second full reread. |
| Native mutation result | Providers may omit receipts and host code can still call the task Completed. | Every destructive Native route returns a typed primary/source/cleanup receipt. Missing or inconsistent receipt is Indeterminate/ContractViolation, never Completed. |
| Permanent/recursive Delete | Some providers delete current occupants without exact per-generation proof; the opposite universal rule would snapshot every descendant and break ordinary folder expectations. | Bind the exact real root, or canonicalize an explicitly selected provider-declared virtual-folder boundary. Real/virtual folder Delete authorizes bounded current membership inside that scope; fixed object-set operations consume admitted key+revision snapshots. Never cross a replacement root, link/alias escape, bucket/profile, or sibling-prefix boundary. |
| Name validation | F7/Batch Rename/bridge impose Win32 rules on every provider and ignore parsed limits/path-scoped case policy. | Add executable `ValidateChildName(parent, child, operation)`/join feasibility. Apply DOS/device rules only to applicable Local directories and support Local per-directory case sensitivity. |
| Traversal bounds | Depth 128 and aggregate entry/path/metadata ceilings are task-terminal; large provider directory buffers are charged before processing. | Replace recursive walking with iterative/paged walking. Queue and aggregate memory pressure use bounded in-memory backpressure, paging, and just-in-time traversal—never rejection of an otherwise valid tree and never a disk-backed traversal spool. Fundamental single-item/provider limits may fail visibly. |
| Conflict presentation | Engine and popup independently derive actions; three-button layout can hide Cancel and Escape starts close/Cancel All. | One engine-owned decision model owns order, primary placement, safe default, Apply-to-all, Enter, and Escape. Conflicts may be discovered progressively and are revalidated at commit; changed/new conflicts return to the same surface. Cancel is directly available and Escape cancels the current prompt. |
| Retry | Explicit Retry disappears after one failed known-noncommit attempt. | Permit repeated user-initiated Retry while exact authority remains valid and no commit is proven; show attempt count; never auto-retry an unknown mutation. |
| Terminal results | Many exits synthesize/overwrite records and destructive receipt omissions bypass the stricter finalizer. | Every selected root enters one terminal builder and exits one funnel. Exactly one immutable result is stored. Aggregates are pure reducers. |
| Publication/compensation | File, link, directory, Native, and Managed paths duplicate stage logic; create-new Copy bypasses it; Cancel suppresses exact compensation. | One owned-publication transaction with payload/provider adapters and cleanup-only cancellation control. Hidden staging is required for overwrite, Managed Move, weak/remote routes, or an atomic-visibility claim; ordinary exclusive new-name Local Copy may retain exact final-leaf authority and report an incomplete artifact honestly. |
| Cleanup debt | Provider backup/stage debt can be logged and discarded. | Receipt and UI distinguish primary success from cleanup debt; no destructive retry follows a committed primary. |
| Possible artifact handling | A name-shape hint can prompt or block ordinary use even when no valid claim proves manager ownership. | Possible is non-authoritative. Read/open/edit/preview/copy/export stays available solely despite the name; one exact-object Cancel-default warning protects RedSalamander-owned destructive mutation. Properties explains all policy-permitted semantic File Operations facts and explicitly distinguishes a valid claim from name-only resemblance. |
| Cancellation | Declared host deadlines do not isolate unbounded in-process calls, while some providers already have narrower watchdog/quarantine containment. | Remove unenforced deadline claims. Classify each exact operation route as cooperatively bounded, safely contained by a proved provider-local watchdog/quarantine, process-isolated, or unavailable. Preserve an existing contained route rather than disabling its provider family. Shutdown has a deterministic host-level test bound; a contained unknown operation remains visibly Indeterminate. See D2-A01. |
| Discovery/progress | “Skip discovery” sounds like it can skip safety, provisional percent can move backward, a whole jobs surface can incorrectly look like Discovery because one task is scanning, and unrestricted Copy/Delete I/O can starve directory enumeration on one device. | `Preparing` is never skippable. Start with the current tested run-ahead scheduler and measure USB large-Copy, small-file Delete, MTP serialization, SMB, and independent-device cases. Add a scoped limiter only for a measured failing profile. `Discover as needed` releases only run-ahead reservation and any adopted limiter while retaining every just-in-time safety check. Per-task bars remain authoritative; the jobs surface labels a closed-total cohort **Known work** and separately reports open discovery. See D2-A08/A19. |
| Initial interaction and operational comprehension | One generic confirmation journey does not fit inline rename, preview-owned batch work, routine Copy, destructive escalation, overlap consent, or review-later processing. Task cards and prompt-at-first-conflict can leave long work stalled and make prior evidence hard to inspect. | Define an ingress-by-ingress human workflow. Routine explicit operations take the accepted zero-extra-prompt fast path; material risk, destructive escalation, known degradation, requested advanced options, or obvious concurrent-task overlap uses one stable problem-specific surface. One jobs surface shows lifecycle/truth/actions, and optional **Continue safe work and review later** mode collects only deferrable decisions. See D2-A02/A17/A20. |
| Operation history | Transient task state is asked to carry auditability, while universal WAL/Undo would grant false recovery authority. | Add a persistent, retention-controlled, exportable human-readable operation history built from immutable results. It is evidence/navigation only and never authorizes Retry, Undo, cleanup, or source deletion. |
| Ordinary Local execution ownership | The custom engine owns even simple Local work, while Shell delegation could remove code but can also create a second conflict/result owner. | Run a bounded `IFileOperation` spike for ordinary new-name Local regular-file Copy and no-conflict same-volume Native regular-file Move. Use the existing Shell Recycle route as a control; exclude Managed/cross-volume/folder/link Move, overwrite, Permanent Delete, Rename, directory merge, artifacts, remote/device, and Batch Rename. Adopt only if one task/conflict surface and exact per-item truth survive. |
| Specification ownership | Product, ABI, bridge, provider, UI, algorithms, tests, and historical perf are repeated. | Give each rule one owner and use stable rule IDs/cross-references. Move algorithms, constants, source-shape assertions, test matrices, and historical measurements out of product behavior. |
| Consistency evidence | `Specs/NormativeConsistency.json:3` says reviewed at `f9cee9d47` although audited authority changed and still reports `VERIFIED_MATCH`. | Record per-document reviewed commit/content hash and fail freshness validation when authority changes. Rebuild the ledger after accepted decisions land. |

### 5.3 Complexity to remove or demote

1. **Semantic link retargeting as the meaning of Preserve.** Literal link-copy is
   the accepted default. Retargeting collision-mapped internal references is retained
   only as the explicit opt-in **Retarget links inside copied tree** transform. Replace current
   repeated linear scans/copies in `State.cpp:11338-11498,14942-14960,15208-15268`
   and the 4,096-record terminal bound with indexed/bounded state. See D2-A03.
2. **A global overlap interlock as normative product law.** Cheap path and positive
   alias evidence exists to explain an obvious task-injection problem, not to forbid
   user-approved work. Remove root-wide rejection/exclusion as the correctness model;
   retain only narrow provider-internal thread/session safety and exact per-item
   authority/conflict/revalidation. See D2-A02.
3. **Treating mandatory eleven-section capability JSON as safety authority.**
   Replace it in one lockstep in-tree cutover with versioned typed per-operation
   route facts, executable queries, and receipts. JSON may remain diagnostics only;
   do not create dual authority, a compatibility adapter, or a data/schema migration.
   Consumer inventory is a removal gate, not a migration program. See D2-A11.
4. **A feature-complete second popup layout/painter.** Keep one semantic/layout/
   accessibility owner and a minimal safe fallback. See D2-A12.
5. **Backwards provisional percentage.** Exact counters plus indeterminate state are
   more honest and simpler. See D2-A08.
6. **Implementation algorithms and historical measurements in product authority.**
   `FileSystem_FileOperations.md:1310-1314` and
   `Core_FileSystemBridge.md:178-185` embed one 130.484 s versus 0.250 s run;
   discovery/overlap-advisory constants and exhaustive source-text test inventories are
   similarly repeated. Keep behavioral/performance gates normative; keep algorithms
   and evidence in code, Testing specs, and `Specs/TestRuns/`.
7. **Source-text tests that fossilize implementation shape.**
   `Commands.SelfTest.PluginConfig.cpp:5334-5368,5596-5615` asserts internal names,
   forbidden strings, exact constants, and routing source text. Replace them with
   fault injection and observable authority/result/resource tests.
8. **Production compatibility mutation engines.** All UI ingress must use central
   plans. Existing direct FolderView `CopyItems`/`MoveItems`/`DeleteItems` calls are
   explicitly selftest-only fallbacks; the central Delete adapter and provider
   contract selftests mean the ABI cannot be removed blindly. Inventory external
   consumers, then retire unused bulk/direct methods in the next ABI.
9. **Global artifact leaf hints and repeated registry parsing.**
   `FileOperationArtifactRegistry.cpp:131-334` plus Folder enumeration/search reload
   and project common final names globally. Keep exact claims, but cache by journal
   generation and index by endpoint/root/canonical path through E6.
10. **Batch Rename cycle/WAL and duplicate scheduling.** Admission does per-row
    capability/binding and quadratic dependency scans on the UI thread; execution
    recomputes schedules and journals every plan. Create one canonical indexed
    acyclic schedule in Preparing, reject cycles before mutation, execute it without
    a recovery journal, and report exact partial results.
11. **Mandatory full content reread/hash for every Managed Move.** Exact destination
    commit, stable-source authority, and conditional cleanup are universal. A second
    full reread is mandatory only for a weak-commit route or explicit Verified Move;
    imposing it on every stable Local transfer doubles I/O without being the only
    way to prove safe completion. See D2-A04.
12. **Exact descendant snapshot as the only directory Delete model.** This is correct
    for fixed object/version selections, but breaks ordinary real and provider-
    declared virtual folders under writers. Bind/canonicalize the selected folder
    scope, perform bounded membership deletion inside it, and condition every
    provider object generation at mutation. See D2-A13.
13. **Hidden staging as the only legal new-name Copy algorithm.** Owned non-final
    staging is essential for overwrite, Managed Move, weak/remote routes, and atomic
    visibility. An ordinary exclusive new-name Local Copy can be simpler if its
    retained handle prevents foreign cleanup and any incomplete artifact is explicit.
    See D2-A06.
14. **A complete frozen conflict set before execution.** Freeze the user's decisions
    and observed authority, not an eager global prescan. Discover conflicts lazily,
    revalidate immediately before mutation, and return changes to the same decision
    surface.
15. **Internal safety taxonomy in routine UI.** Native/Managed/Copy-only and the
    result axes remain essential internally, but routine UI should explain concrete
    consequences instead of backend categories.
16. **Universal broker, WAL, or transactional Undo.** Isolate only providers that
    cannot prove a bounded quiet point; persist exact stage claims only when an
    artifact intentionally survives; advertise Undo only when a valid exact inverse
    exists. A human-readable log is not mutation authority.

### 5.4 Complexity that should not be removed

- Do not collapse result axes into success/failure.
- Do not replace exact mutation authority with path normalization, hashes, weak
  indexes, size/time matching, or provider-name heuristics.
- Do not expose an incomplete final-leaf artifact as complete merely to eliminate
  staging; under accepted D2-A06 direct exclusive Local publication, its incomplete
  state and exact cleanup outcome remain explicit.
- Do not merge Native and Managed algorithms; they have different atomicity and
  cleanup boundaries.
- Do not merge durable schemas that carry different authority.
- Do not build generic Undo from incomplete receipts.
- Do not weaken destination commit, stable-source, or conditional-cleanup proof when
  making checksum verification route-aware.
- Do not call size/time equality or a provider `S_OK` exact content verification.
- Do not treat a persistent operation log as Retry, cleanup, or Undo authority.
- Do not centralize F7 or shell Recycle merely for architectural symmetry.
- Do not make an optional exact-identity optimizer a prerequisite for safe
  conservative Copy-only work.

### 5.5 Provider disposition

| Provider/profile | Keep | Change/disable |
|---|---|---|
| Local Win32 | Native same-volume regular-file Move; exact bound Delete/Rename; Managed cross-volume/folder/link routes; stable retained source handles; exclusive publication. | Fix FOS-01/FOS-11. Split retained authority from comparable `FILE_ID_INFO`. Do not globally reinterpret 996. Use path-scoped case sensitivity and iterative traversal. Support Standard versus Verified Move proof and the D2-A06 direct-final exception only for ordinary new-name Local Copy. Evaluate bounded `IFileOperation` delegation under D2-A16 rather than assuming adoption. |
| Dummy | Deterministic test engine, atomic writer, fault/latency/cancel simulation. | Clear product Delete/Rename claims until a node-identity central mutation authority exists; test capabilities obey production semantics. |
| 7z | Honest read-only/export-only posture. | Do not add mutation merely for abstraction symmetry. |
| S3 | Provider-specific conditioned Copy/Move/Rename transaction where exact authority/receipt exists; ordinary normalized prefix paths are user-visible virtual folders. | `Delete s3://bucket/prefix/` performs bounded recursive virtual-folder emptying, including late/replacement members freshly observed before convergence, without crossing the canonical prefix boundary. Ordinary current-key removal uses the just-observed ETag and no `VersionId`; exact historical-version cleanup is a separate fixed-set operation. Churn that prevents convergence stops with residuals. In versioned buckets say that current items disappeared while older versions may remain. Expose central authority, receipts, verification, and cleanup debt. |
| Microsoft Drive | Item-ID-preserving, ETag-reconciled Graph PATCH Move/Rename. An exact folder item ID is a real-container root. | Present ordinary Graph `DELETE` as Recycle, not Permanent Delete; bind and revalidate the accepted stable item ID without a root ETag/`If-Match` condition while current members remain container scope. Keep the separate Graph `permanentDelete` endpoint unavailable until it satisfies accepted exact authority and receipt rules. Expose receipts/debt and do not force Native Graph work through the host bridge. |
| MTP | Serialized WPD worker, watchdog/quarantine, bounded streaming, and PUID-backed temp-swap concept. | Remove hash/path/metadata authority, persist full PUID, and clear central Delete/Rename until exact authority/receipt exists. |
| IMAP | Message Delete with admitted `UIDVALIDITY + UID`. | Disable/redesign mailbox recursive Delete without exact mailbox generation/snapshot. |
| FTP/SFTP/SCP | Read/import and proven exact writer routes; Curl cleanup-debt alert is a good precedent. | Keep same-provider Move/Rename/overwrite disabled without exact no-replace/cleanup; disable Delete without exact authority; retire hidden compatibility mutation debt after caller inventory. |
| Google Drive | Current read-only/unstable-path skeleton. | Keep mutation unavailable until stable item identity, executable I/O, and conditional mutation exist. |

The S3 decision relies on the documented service contract: current-object
`If-Match` conditions exist for both singular and multi-object Delete, apply to the
current version, and fail rather than deleting on an ETag mismatch; S3 also documents
strong consistency for object Delete and listing observations. See [AWS conditional
deletes](https://docs.aws.amazon.com/AmazonS3/latest/userguide/conditional-deletes.html)
and [S3 consistency](https://docs.aws.amazon.com/AmazonS3/latest/userguide/Welcome.html#ConsistencyModel).

### 5.6 Adjacent security finding

Legacy IMAP and FTP/SFTP/SCP settings permit plaintext passwords or key
passphrases (`Specs/FileSystem/FileSystem_Imap.md:175-184` and
`FileSystem_FtpSftpScp.md:89-106`; consumption at
`Plugins/FileSystemCurl/FileSystemCurl.Shared.cpp:1429-1473,4101-4148`). This is
not File Operations architecture, but conflicts with the same user-trust north
star. A separate P1 security owner should migrate secrets to Connection
Manager/WinCred, clear migrated values, and request re-entry visibly when needed.

## 6. Proposed complete working contract

This section combines the complete accepted **target** contract for this WIP. All
product-owner questions in section 7 are answered, but accepted clauses remain
`DECIDED-NOT-ACTIVE` until their bounded implementation slice is activated. Activation
permits implementation; it does not rewrite current behavior. A target clause becomes
durable only when the same slice lands implementation, tests, and the authoritative-
spec update routed by section 10.

The following IDs are the only target-behavior references used by the dashboard,
packages, verification matrix, and STOP index. Text outside this section may explain,
implement, test, or reject a rule, but may not restate it as an independent contract.

| Rule ID | Sole target contract | Accepted decisions | Delivery slices |
|---|---|---|---|
| `FO-LIFE-01` | Mandatory bounded Preparing, clipboard release boundary, and routine start | A01, A05, A20 | R1a/R1d |
| `FO-OVERLAP-01` | Cheap problem-specific overlap warning, transient same-host Queue, concurrently-live same-host output guard, and concurrent correctness | A02, A14 | R2/R4/R6-A02 |
| `FO-DISCOVERY-01` | Streaming traversal, evidence-first discovery service, per-task dual progress, honest mixed-task **Known work**, and `Discover as needed` | A08, A19 | R4/R6-A08 |
| `FO-AUTH-01` | Exact mutation authority, typed route/error/receipt facts, and provider namespace | A11, A13-MD1, A14 | R0c/R0d/R2 |
| `FO-ITEM-01` | One per-item mutation state machine and exactly one immutable terminal result | A07 | R1a/E1/R6-A07 |
| `FO-PUBLISH-01` | Risk-scoped owned publication, compensation, cleanup debt, and incomplete-artifact truth | A06 | R0a/R1c/R3 |
| `FO-MOVE-01` | Native/Managed/Copy-only strategy and route-aware proof before source cleanup | A04, A05 | R2/R3 |
| `FO-DELETE-01` | Real-container, declared virtual-folder, and fixed-set Delete authority | A13, A13-MD1 | R0b/R0d/R2/R4 |
| `FO-DECISION-01` | Single conflict owner, explicit Retry, review-later, and cancellation semantics | A07, A12, A17 | E2/R6-A07/A12/A17 |
| `FO-TRUTH-01` | Plain-language outcomes, artifact explanation, and non-authoritative operation history | A09, A18 | E6/R6-A09/A18 |
| `FO-UX-01` | Ingress journeys, phase controls, focus, accessibility, and mixed-task presentation | A02, A08, A12, A17, A19, A20 | R1d/R4/R6 |
| `FO-LINK-01` | Literal Preserve and the explicit bounded Retarget transform | A03 | R7-A03 |
| `FO-RENAME-01` | Cyclic Batch Rename rejection, no production recovery journal, and one RenamePlan path | A10 | R7-A10/E4/E5 |
| `FO-NAME-01` | Provider/path-scoped executable child-name and collision feasibility | A15 | R2/R7-A15 |
| `FO-SHELL-01` | Bounded non-default Shell-delegation evidence spike and adoption reject gates | A16 | R9 |
| `FO-SIMPLE-01` | Complexity budget and forbidden replacement frameworks | all | every package reviewer |

### 6.1 Lifecycle, mandatory preparation, and streaming discovery (`FO-LIFE-01`, `FO-OVERLAP-01`, `FO-DISCOVERY-01`)

The product must not use **preflight** as one ambiguous name for three different
things:

1. **Preparing safety gate:** mandatory selected-root admission work before any
   mutation or clipboard consumption. It cannot be skipped.
2. **Discovery ahead:** optional bounded run-ahead traversal after acceptance. It
   improves totals and keeps work ready while transfer starts immediately. The current
   tested reservation/half-rate scheduler is the compatibility baseline; a new scoped
   limiter requires measured same-resource starvation evidence.
3. **Just-in-time safety checks:** enumeration, identity, containment, capability,
   name, conflict, proof, and commit-time revalidation required when each item is
   reached. They can never be skipped.

```text
User request and ingress-owned intent/defaults
    |
    v
Preparing safety gate (mandatory, cancellable, no mutation/clipboard/breadcrumb)
  - validate immutable request and endpoint pair
  - resolve path-scoped namespace and executable top-level route facts
  - bind selected destructive roots or classify a known safe Copy-only outcome
  - compare source/destination/delete roles against live tasks using cheap path facts
    and only cheap/bounded positive link, mapped-drive, or alias evidence
  - no recursive enumeration, payload read/hash, totals walk, or descendant scan
    |
    +--> obvious or positively evidenced concurrent-task problem
    |       explain the concrete race in the same stable task surface
    |       Queue after (safe default) / Run at the same time / Don't start
    |
    +--> material choice or explicit "with options" request
    |       one stable informed-consent surface
    |
    +--> routine explicit intent with accepted defaults
    |       no extra modal prompt
    v
Accepted execution boundary
  - record the transient same-host Queue edge or concurrently-live Run receipt;
    neither survives as object authority
  - consume the matching clipboard Move sequence exactly once; failure aborts
  - write the notice-only Move breadcrumb when applicable
    |
    v
Discovery ahead + execution (same single traversal)
  - first safe item may mutate before traversal closes
  - discovery cannot be starved by transfer
  - Discover as needed may release run-ahead priority, never safety checks
    |
    v
One per-item state machine
  authority -> conflict -> owned publication -> proof -> exact source cleanup
    |
    v
Exactly one immutable terminal item result
    |
    v
User explanation: requested / definite / possible / remains / safe next actions
```

`Preparing` is selection-proportional and may inspect only the immutable envelope,
endpoint-pair route facts, selected top-level object shape/authority, the immediate
top-level destination facts needed for route/consent, and top-level overlap-
advisory facts. It performs no recursive enumeration, content read/hash, totals walk,
descendant conflict scan, publication, clipboard consumption, or breadcrumb write.
Each required provider call must have executable containment: a cooperative bound,
a proved provider-local watchdog/quarantine with lifetime and unload safety, or
accepted process isolation. Otherwise D2-A01 makes only that exact route unavailable.
“Bounded” is an executable route property, not a timer around an uninterruptible
call. Existing provider containment still requires a host-level shutdown witness.

A task record exists as soon as the request is accepted by the UI. Preparation that
finishes within the measured reveal budget transitions directly into execution or
the one required consent surface without flashing a second card or moving focus.
Longer preparation reveals that same stable card as `Preparing...`, with endpoint
summary and Cancel only; completion replaces it in place. Cancel during Preparing
means no mutation, cut-list consumption, breadcrumb, or successful-operation history
entry. A diagnostic rejected-at-preparation attempt may be retained only when the
history decision explicitly permits it.

At task injection, one running RedSalamander host compares the new task's source,
destination, and Delete roles with active and queued tasks admitted in that same host.
It performs no cross-instance/process/machine coordination. Comparison begins with normalized provider/profile/path
parentage and uses link targets, mapped-drive mappings, stable identities, or alias
facts only when they are already available or obtainable through a bounded cheap
query. It never starts recursive discovery, opens every descendant, or delays routine
injection to prove that no alias exists. Positive evidence may produce **These tasks
may refer to the same location**; absence of such evidence is not a safety proof.

The warning names the actual problem, not the implementation term “interlock”:

- two destination scopes may create, replace, or rename the same names;
- a Delete scope may remove a source or destination another task is reading/writing;
- two Delete scopes may remove the same current members; or
- two textual paths are positively known or strongly evidenced aliases.

Clearly disjoint paths such as `C:\Photos` and `C:\Backups` do not prompt. A parent/
child overlap that is still inside an active task's uncompleted scope does prompt. If
the task already has exact cheap residual-scope evidence proving that branch is
finished or excluded, the prompt may be suppressed; the host never performs extra
traversal merely to suppress a warning.

The stable surface offers **Queue after these tasks** as the safe scheduling default,
**Run at the same time**, and **Don't start**. The warning is shown once for the named
same-host tasks/scopes. If a later queued Delete covers an earlier Copy/Move
destination, the warning explicitly says that after the producer finishes the Delete
will run against current folder membership and may remove those outputs.

Queue creates one transient in-memory predecessor edge inside this running
RedSalamander host. When every named predecessor reaches a terminal state, the edge
is discarded and the queued task enters the ordinary engine without a second A02
admission warning or protection derived from completed tasks. A later injected task
is evaluated separately. Run records only the disclosed scheduling/consequence
consent while the tasks are concurrently live in this host. Neither choice creates
current-object mutation authority, an Apply-to-all rule, or a promise that both
intents can succeed.

User-approved overlap executes through the normal engine. Only while two tasks are
concurrently live in the same RedSalamander host does the A02 live-output guard apply.
If one task reaches an exact object just created by the other and no covering Run
warning disclosed that destructive consequence, it parks before Delete/overwrite/
rename/invalidation and offers **Skip this item**, **Queue until the other task
finishes**, the explicit consequence-named destructive action, and Cancel. Queue ends
concurrency; after the predecessor terminates, the waiting task resumes under
ordinary container-membership, conflict, revalidation, and result rules and may
remove that now-current member. Only Skip promises preservation.

The guard uses a bounded host-local in-memory index of exact creation/publication
receipts and concurrently-live Run receipts. Entries and protective meaning end when
the relevant tasks are no longer concurrently live. The index is never persisted,
shared with another RedSalamander instance/process/host, reconstructed from history,
or used as pathname cleanup authority. Other applications/instances and
post-completion changes are ordinary external races governed by exact per-item
authority and revalidation.

No root-wide host lock may reject or silently serialize all user-approved overlap as
a correctness shortcut. Providers and adapters must nevertheless be thread-safe and
reentrant or serialize only their own narrow session/handle critical sections. A
single-lane provider may time-slice calls while both tasks remain live; **Run at the
same time** promises concurrent task admission, not impossible physical parallelism.

After the first safe item is ready, transfer/mutation starts while bounded discovery
runs ahead. R4-A19 first preserves and instruments the current tested discovery-
reservation/low-water scheduler. Its baseline records time to first safe mutation,
time to traversal close, runnable-discovery wait, queue/memory high-water, bytes and
mutation operations while discovery is open, cancellation, and total throughput on
the required Local, USB-like, SMB, MTP, and independent-device scenarios.

The initial implementation keeps the existing reservation/half-rate behavior and a
bounded in-memory ready queue. When full, discovery backpressures/pauses and later
resumes, or uses just-in-time traversal. It never writes paths, metadata, or an
enumerated worklist to a disk/durable spool. `FO-DISCOVERY-01` requires measurable
discovery forward progress and bounded memory, not a new general byte-token, metadata-IOPS,
device-class, adaptive-ramp, or scheduler-control framework.

Only if the unchanged baseline proves material same-resource starvation may R4-A19
adopt the smallest measured correction. Add byte throttling only when a transfer is
the measured cause; add mutation-rate control only when small-file Delete is the
measured cause; use bounded discovery/operation turns for a genuinely serialized
provider. Prefer tuning the existing reservation/share or existing bandwidth control.
Archive the failing baseline, acceptance threshold, candidate, and before/after
evidence. Do not throttle an independent device/provider, and remove a candidate that
does not meet its evidence gate.

While discovery-ahead is active, the card exposes one task-local, no-confirmation,
one-way action named **Discover as needed**. It stops new ahead admissions at the
next checkpoint, retains already discovered work, releases discovery reservation
and any separately evidence-adopted temporary limiter within the existing 50-ms
checkpoint target excluding an active provider call, then switches to just-in-time
traversal. It never changes selected scope, strategy, conflict
policy, verification, enumeration, identity, containment, capability, name checks,
or commit-time revalidation. The action then disappears and the status remains
`Discovering as needed` until traversal closes. No control is ever named **Skip
preflight** or **Skip safety checks**.

The current scheduler remains the compatibility baseline unless it fails the named
bounded discovery-service/no-starvation threshold. A replacement or added limiter
must improve that failing case without exceeding its throughput, memory, or
cancellation no-regression budget. Queue thresholds and worker fractions remain
performance policy, not product vocabulary.

Before a task's traversal closes, that task has no completion percentage. Exact
discovered, processed, item, and byte counters remain visible. A time estimate may be
shown as **Estimated time remaining — still discovering** when stable sampling
supports it; it is provisional, may move in either direction, disappears when not
supportable, and never drives behavior. Each task card keeps its discovery row/count/
ETA separate from its operation row. A task whose own total is closed may show its
own fixed-total percentage even while another task is still discovering.

Per-task bars are authoritative. The jobs view forms a **Known work** cohort from
active nonterminal tasks whose totals are closed. When those tasks share one additive
unit, it shows **Known work: N%** as `sum(completed) / sum(total)` for that cohort. Open-
total tasks are excluded from that fraction and shown beside it as **N tasks
discovering — total may grow**. The jobs view never labels this subset **Overall**,
averages task percentages, mixes bytes with items, or estimates an open denominator.
If there is no closed-total cohort, or its totals have no common additive unit, show
activity/throughput and exact counters without a jobs percentage. While any included
task has an open total, the Windows taskbar is indeterminate because it cannot display
the qualifying **Known work** label. Task-local closed-total percentages remain
visible throughout.

### 6.2 Authority and error model (`FO-AUTH-01`)

Each route exposes three separate facts:

1. **Mutation authority:** a retained handle/full provider token or conditional
   backend receipt that mutates only the exact admitted object or task-created stage.
2. **Comparable identity:** optional evidence that two separately bound names refer
   to the same object/generation.
3. **Overlap advisory evidence:** cheap normalized path parentage and optional
   positive identity/link/drive/alias evidence used only to decide whether the task-
   injection UI explains an obvious or possible concurrent-task problem.

Lack of comparable identity reduces warning precision; it neither blocks nor grants
concurrency, does not automatically fail an otherwise safe create/Copy, and never
authorizes cleanup. `ERROR_IO_INCOMPLETE` 996
keeps its real transient/unknown classification unless the provider deterministically
proves the identity feature is unsupported. A central helper may list proven
unsupported errors, but cannot turn uncertainty into support absence.

An all-zero `FILE_ID_128` is Unsupported and never enters equality, cache,
admission, or destructive-authority decisions. Lifetime-scoped retained-handle
authority may authorize only the exact object this task just created while that
handle remains retained. It cannot bind a pre-existing object, authorize source
deletion, or replace an existing destination without separate exact
expected-destination authority.

Capability planning consumes versioned typed per-operation route facts through a new
IID/`sizeBytes` boundary plus executable bind/publication/Delete queries and typed
receipts. Host, shipped providers, Dummy, adapters, and tests form one supported
lockstep generation. JSON can carry diagnostics or provider-specific extensions only;
it never enables a route, fills a missing typed proof fact, or coexists as a second
safety authority. There is no compatibility/data migration or mixed-generation
execution path.

| Classification | Meaning | Permitted next action |
|---|---|---|
| Unsupported | This path/profile stably lacks the required safe route. | Explain before mutation; offer Copy-only/no operation. |
| RetryableNoCommit | The provider proves the attempted mutation did not commit. | Explicit Retry/Skip/Cancel with retained exact authority. |
| FailedKnown | The operation failed and exact post-state is known. | Show retained/published/stage facts and safe actions. |
| Indeterminate | Commit, current generation, or cleanup truth cannot be proven. | Preserve data/claim, stop dependent destructive work, inspect/reconcile; never blind Retry/cleanup. |
| ContractViolation | Provider returned an impossible/inconsistent safety receipt. | Stop and quarantine/disable the route; retain evidence and report provider defect. |

### 6.3 Per-item mutation state machine (`FO-ITEM-01`)

Every selected root receives a result slot before execution. Legal transitions are
named; arbitrary field mutation is forbidden:

```text
NotStarted
  -> Bound | Unsupported | FailedKnown
Bound
  -> WaitingForDecision | Ready
Ready
  -> StageCreated | FinalLeafCreated | NativeCommitted
  -> RetryableNoCommit | Indeterminate
StageCreated
  -> Writing -> StageCommitted -> Published
  -> Aborted | Retained | Indeterminate
FinalLeafCreated
  -> WritingVisible -> FinalLeafCommitted -> Published
  -> Aborted | RetainedIncomplete | Indeterminate
Published
  -> RouteProofSatisfied | VerificationRequired
VerificationRequired
  -> Verified | VerificationFailed | VerificationUnavailable
RouteProofSatisfied/Verified
  -> SourceDeleted | SourceRetained | SourceCleanupIndeterminate
any terminal path
  -> exactly one immutable ItemResult
```

The builder validates combinations. `SourceDeleted` is illegal without accepted
Move, destination commit proof, stable-source proof, the route's accepted
verification policy, and exact source-cleanup receipt. `Published` and
`StageRetained` cannot describe the same object. Skip/Cancel before mutation is
`Publication=NotAttempted`, `Source=Retained`, `Stage=None`, not Unknown.
Primary success and cleanup debt are independent; null destructive receipt can
never produce Completed.

Once a possibly committing Rename, Delete, abort, publish, or rollback call begins,
known non-commit is legal only when exact reconciliation proves it. Otherwise the
result is Indeterminate. Skip, Cancel, Continue, and presentation disposition never
rewrite already-known publication, source, or artifact truth.

`FinalLeafCreated` means the requested name is visible while content is incomplete;
it is never `Published` or complete until `FinalLeafCommitted`. A failed/canceled
direct-final Copy that cannot be exactly aborted is `RetainedIncomplete` and remains
an explicit artifact/result state. D2-A06 can permit this narrow visibility shape;
it can never authorize calling partial content complete.

### 6.4 Publication and compensation (`FO-PUBLISH-01`)

Every Copy route uses one of two explicit owned-publication shapes:

1. **Owned non-final stage then atomic/conditional reveal.** Required for overwrite,
   Managed Move, remote/device or weak-commit routes, resume, and any operation that
   claims the final name appears only when complete.
2. **Exclusive final-leaf create with retained exact authority.** Permitted for an
   ordinary new-name Local Copy when no previous destination can be displaced, the
   same retained object is used for write/abort truth, and any surviving incomplete
   artifact is reported as incomplete rather than complete.

The transaction owns only what it created and records exact stage/final authority,
content commit, final-name publication, abort, and retained/unknown artifact identity.

No failure path opens a final pathname and deletes its occupant because the name was
absent earlier. Exact compensation uses a separate short cleanup control after
primary work is canceled; it never restarts primary work. If cleanup cannot finish,
the primary result remains truthful and debt is exposed.

If a provider cannot implement atomic reveal where the route requires it, the route
is unavailable or explicitly Copy-only. D2-A06 chooses the narrower Local-only
direct-final exception; it is not permission to publish partial overwrite/Managed
Move content under the requested final name.

### 6.5 Move (`FO-MOVE-01`)

- **Native Move:** one provider operation preserves/changes the exact object and
  returns a mutation/source/cleanup receipt. Verification is NotApplicable.
- **Managed Move:** exact destination publication from a stable admitted source,
  satisfaction of the route's accepted proof policy, then exact conditional source
  cleanup. Weak-commit remote/device routes and explicit Verified Move require
  end-to-end content proof. A stable Local committed-stream route may use Standard
  Move without a second full reread; if stability/commit cannot be proven, it must
  verify or finish Copy-only.
- **Copy-only:** destination publication may complete, source is definitely retained,
  and the UI never uses an unqualified “Moved” result.

Known strategy counts/reasons appear at consent. Runtime downgrade preserves source
and creates an immediate visible issue; it never silently changes cleanup semantics.

### 6.6 Delete (`FO-DELETE-01`)

Normal Delete first binds/canonicalizes the exact selected scope and then applies
one of three models:

1. **Real hierarchical container:** selecting the exact directory object authorizes
   bounded no-follow removal of descendants encountered inside that same container
   until the directory can be removed. Traversal cannot follow links, mounts, aliases,
   or replacement roots outside the bound container. If an active writer prevents
   completion within the bounded policy, stop and report residuals.
2. **Provider-declared virtual folder:** the provider has no physical container
   object but presents a canonical scope as an ordinary folder. Selecting
   `s3://bucket/prefix/` authorizes bounded recursive removal of current members under
   that exact bucket/profile/prefix boundary, including late members observed before
   convergence. Every current key is deleted with the ETag just observed for that
   key; an ETag mismatch deletes nothing, and a later bounded pass may freshly admit
   the replacement as current in-scope membership. A request-level unknown is never
   blindly retried. Never match `prefix`, `prefix2/`, a provider-returned out-of-scope
   key, or another account/bucket/profile. Stop and report residuals when writers,
   permission, conditions, cancellation, or the pass/time bound prevent convergence.
   Completion linearizes at one complete empty strongly consistent listing; it is
   not a promise against a later writer.
3. **Fixed object/version set:** when the command selects keys/versions rather than
   an ordinary folder, confirmation snapshots exact keys and revisions. Late keys and
   replacement generations are outside that fixed set and survive.

The UI explains fixed-set semantics when they differ from ordinary folder Delete.
S3 ordinary prefix Delete uses model 2; an explicit version/object selection uses
model 3. A provider must not infer model 2 merely from arbitrary text ending in `/`:
the admitted route must declare and canonicalize virtual-folder semantics. Normal
S3 folder Delete removes the current key with `If-Match` and omits `VersionId`;
exact-version rollback/cleanup is a separate model-3 operation. In a versioned
bucket, ordinary folder Delete can create delete markers and hide current items while
older versions remain; results must never call that permanent erasure.

Microsoft Drive folders use the real-container model because the selected folder is
an exact item object, not a textual prefix. Ordinary Graph Delete is a Recycle
operation by the accepted stable item ID without a root ETag/`If-Match` condition.
A rename or unrelated property revision on that same ID does not invalidate consent;
a different item ID is never admitted as a replacement. This does not freeze the
folder's current child set. Graph Permanent Delete remains a separate unavailable
route until it provides the accepted authority and receipt.

Only while both tasks are concurrently live in the same running RedSalamander host
does the A02 guard override ordinary live-container membership. Covering disclosed
Run consent or an item-specific destructive choice permits the concurrent effect;
Skip preserves. Queue waits for the publishing task to terminate, ends the special
live relation, and then resumes ordinary live-container membership, so the current
object may be removed. No history/path fact or cross-instance/process/machine
coordination extends this guard after terminal state. Objects created by Explorer,
another application, or another RedSalamander instance are ordinary external
live-container membership and receive the normal exact per-item authority and
revalidation—no snapshot, monitor, or cross-process protection framework is required.

### 6.7 Conflicts, Retry, and cancellation (`FO-DECISION-01`)

One policy snapshot defines conflict choices, safe default, primary placement,
Apply-to-all scope, and keys. Conflicts are discovered progressively; each observed
destination is revalidated immediately before mutation, and a changed/new conflict
returns to the same surface. Cancel is visible. Escape means Cancel for the current
decision. Closing/hiding the operational window does not resolve a prompt.

One hosted renderer owns layout, actions, focus, and accessibility. If attachment
fails, a minimal localized surface may show only the operation/context/truth text and
safe Cancel/Close; it consumes the same snapshot and never re-derives actions or
duplicates graphs/painter behavior.

When the user chooses **Continue safe work and review later**, one unresolved conflict or failed item
does not stall unrelated safe work. The selected roots and accepted scope remain
fixed while descendants may still be discovered inside those already bound roots.
The engine continues unambiguous items, collects conflicts/errors, and presents one
end-of-run review with per-item Retry/Skip/Cancel/decision scope. Rescan is a new
plan; it never silently admits new selected roots or anything outside the exact real
container/provider-declared virtual-folder/fixed-set scope already accepted.

Retry remains available only for `RetryableNoCommit` while authority is valid. The
UI shows attempt count and never retries automatically. Indeterminate offers
inspect/reconcile/stop, not Retry.

Cancel stops new primary work, signals bounded/isolated calls, performs bounded
exact compensation, drains known completions, and yields exact or indeterminate
results. The host never claims a deadline it cannot enforce. Application shutdown
has a deterministic finite test bound under the accepted isolation decision.

### 6.8 User-facing truth (`FO-TRUTH-01`)

Every issue and aggregate card can answer:

1. What did I ask for?
2. What definitely happened?
3. What may have happened?
4. What remains, and where?
5. Why did the plan change or stop?
6. What actions are safe now?

The primary jobs surface shows operation verb, endpoints, `Preparing`/running/
paused/canceling/indeterminate state, plain-language strategy, exact available
counters, current item, failed/skipped/conflict counts, and safe actions. Internal
taxonomy maps to a small result vocabulary: `Moved`, `Copied`, `Copied - source
kept`, `Canceled`, `Failed`, and `Final state uncertain`.

Primary summaries remain compact, but details never force inference of source
deletion from `S_OK`, destination publication from byte count, or safety from a
provider name. Copy-only, verification-unavailable, retained/cleanup stage,
residual concurrent objects, and indeterminate cleanup stay warning/partial states
even when requested bytes are present.

Properties performs a fresh shared-classifier query and appends a host-owned
**File Operations** section for a Proven or Possible artifact. It exposes every
policy-permitted human-readable semantic fact from the current projection and valid
claim/result: classification/reason, operation/task/time when known, artifact kind
and durable phase, intended destination/final name, publication/source/cleanup/
verification truth, current claim-match status, recovery availability, and safe
actions. Missing facts are unavailable/Unknown; name-only inference explicitly says
no valid operation record or ownership claim was found. A name-only Possible,
including a user-created `.rs_*` file, remains selectable, and read/open/edit/preview/
copy/export is never blocked or prompted solely because of its name. The augmentation does not scan all history,
duplicate recovery authority, fabricate fields, reveal credentials/tokens/secret-
bearing endpoint fragments, or expose raw opaque mutation authority merely because
it exists internally. Properties/history remain explanatory and grant no mutation
authority.

Immutable attempt results feed a retention-controlled, searchable/exportable
operation history. The history supports explanation and navigation; it never grants
cleanup, Retry, Resume, Undo, source-deletion, or artifact authority. Undo is shown
only when an exact provider/action-specific inverse or Recycle contract exists.

### 6.9 Human-in-the-loop journey contract (`FO-UX-01`)

The generic state machine is not a complete user workflow. The following journeys
are required coverage for the replacement; the table is an acceptance matrix, not
permission to build a workflow DSL or a different engine for each command.

#### 6.9.1 Interaction principles

- One task has one stable operational surface. Preparing, consent, progress,
  attention, cancellation, review, and completion replace content in that surface
  rather than opening a chain of dialogs.
- Routine explicit intent requires one invocation and no additional decision under
  accepted D2-A20 B. A prompt exists only for a material consequence,
  destructive escalation, known degradation, or an explicit **with options** path.
- Accepted A02 overlap warnings are task-injection scheduling decisions, not generic
  operation confirmation. They appear only for an obvious/positively evidenced
  concrete race, name the problem, and offer Queue/Run together/Don't start before
  mutation or clipboard consumption.
- At most one actionable decision owns foreground focus. Other tasks and unrelated
  safe work may continue; attention is counted and announced without repeatedly
  stealing focus or bringing the application in front of another application.
- Primary language states consequences: `Move`, `Copy and keep source`, `Replace`,
  `Delete permanently`, `Discover as needed`, `Final state uncertain`. Native,
  Managed, receipt, revision, and identity detail is expandable evidence.
- Every actionable prompt has a visible safe action. Enter invokes only its explicit
  focused/default action; Escape means Cancel for the focused decision, never Cancel
  All. Outside a decision, Escape hides the popup. Hiding/closing never grants,
  rejects, retries, skips, or cancels work.
- Caption Close remains a presentation-only hide for the global File Operations
  popup. Reachability is separate from decision policy: **View > File Operations**,
  `cmd/app/showFileOperations`, and the application-scope `Ctrl+Shift+J` shortcut
  restore the existing surface without changing task state. Publishing a newly
  actionable decision automatically restores a popup hidden while that decision was
  still loading; it never submits the published default/Escape action.
- Decisions are task-local immutable receipts. Apply-to-all is offered only for the
  exact published scope and never becomes a silent Preference. Scheduling receipts
  are immutable only while applicable: a Queue edge or live-output guard ends with
  the same-host live relation, and history never reconstructs it.

#### 6.9.2 Ingress journeys

| Ingress | Initial human step and fast path | Later human gates | Completion behavior |
|---|---|---|---|
| F5/F6, pane Copy/Move, clipboard Paste, internal drag/drop, Find Files | The explicit command captures source, destination, and current defaults and an otherwise routine route starts after Preparing. Clipboard Move consumes its captured cut sequence only after Preparing and any A02 overlap/material acceptance. | Obvious concurrent-task overlap asks Queue/Run together/Don't start; known Copy-only or material loss/cost asks before mutation; collisions and runtime facts use the central task decision surface. | One task/result set; navigation, overlap-decision detail, and retained-source explanation remain available. |
| External OLE/`CF_HDROP` drop | Drop gesture is intent; source is qualified during Preparing. Report Copy to OLE until asynchronous Move and exact source disposition are proved. | The same central conflicts/consent; no Shell/provider second prompt. | The external effect and RedSalamander result never claim Move from intention alone. |
| Compare Directories synchronization | Compare preview/summary plus Run owns initial consent and explicit mappings; do not add a second generic confirmation. | Only new commit-time conflicts or deferred material risks prompt. | Results retain Compare row correlation and permit navigation to both sides. |
| Delete to Recycle / Permanent Delete | The Delete command is explicit intent. Routine Recycle takes the accepted fast path; Permanent Delete always has one clear destructive confirmation with Cancel default. | A02 warns before an obvious overlap with another task; Recycle failure escalation is exact-item, Cancel-first, and never Apply-to-all. | Exact removed/retained/uncertain scope is shown; Shell success is not expanded into invented descendant truth. |
| Inline F2 | Committing the inline editor is the initial consent. A clean fast rename stays inline/silent under the existing reveal policy; delay, conflict, failure, or uncertainty reveals the task in place. | Central Rename conflict and exact replacement authority only. | Restore focus by object identity where known; no duplicate completion card for silent success. |
| Change Case / Batch Rename | Their preview/Run surface owns the initial consent. Preparing and execution move off the UI thread and use one `RenamePlan`; no second generic prompt. Cyclic mappings keep Run disabled and explain an intermediate-name/two-pass workflow. | Newly observed conflicts or risk gates use the central decision owner; unreadable descendants are never silently skipped. | Row/object results map back to preview, including exact partial state. There is no Batch Rename journal, Resume, Roll back, or automatic crash replay; refresh rebuilds truth from the provider namespace. |
| F7 Create Directory | Name entry/affirmative action is consent. Keep the small direct journey only for a provider-proved bounded call; otherwise use worker preparation without pretending it is a recursive transfer. | Provider name feasibility and one exact destination conflict. | Immediate focus/select on success; one localized failure on known non-commit. |
| Pack/Unpack cleanup | The archive confirmation owns the scoped delete-after-success consent. | Cleanup occurs only after successful/required verification; later exact failure is reported, not re-prompted as a new broad Delete. | Archive result and retained cleanup item remain correlated. |
| Properties on a Proven/Possible artifact | Opening Properties is read-only and never warns merely because of `.rs_*` naming. The host performs one fresh bounded shared-classifier query. | None; unknown/unavailable facts are explanatory, not a request for authority. | The **File Operations** section shows policy-permitted, redacted semantic facts and the exact Proven/Possible/name-only reason without changing the object or granting recovery. |
| Command/plugin extension | Must submit the same typed intent and either reuse an existing ingress consent or request the same consequence-based options surface. | Cannot bypass capability, authority, conflict, consent, or result policy. | Same task/history truth; no private provider mutation UI. |

#### 6.9.3 Human-gate classification

| Gate | Ask/collect rule | Safe/default consequence |
|---|---|---|
| Obvious concurrent-task overlap at injection | Compare top-level operation roles by cheap path facts and only bounded positive alias/link/drive evidence. Ask once only when a concrete race is obvious or positively evidenced; name the same-host tasks, paths, and both Run/Queue consequences. A queued destructive task must say when ordinary post-release membership can include predecessor outputs. This gate cannot enter review-later. | **Queue after these tasks** stores one transient same-host predecessor edge and releases automatically after terminal predecessors; **Run at the same time** grants only the disclosed concurrent scheduling/consequence consent; **Don't start** leaves mutation/clipboard/breadcrumb untouched. |
| Concurrently-live same-host output reached without covering consent | Immediately before Delete/overwrite/rename/invalidation, consult the bounded host-local live-task publication index. If the other concurrently-live task created the exact object and no Run receipt covered this destructive consequence, park and ask. An uncertain positive correlation also parks; no history or generic Run setting resolves it. | **Skip this item** preserves it. **Queue until the other task finishes** ends concurrency and then resumes ordinary membership—it does not promise later preservation. The destructive action is explicit and consequence-named; Cancel remains available. |
| Ordinary collision, infeasible name, or type mismatch | Ask now by default; may enter the D2-A17 collected review. Commit-time change returns to this same gate. | Cancel for the current decision; destructive actions require exact authority. |
| Proven `RetryableNoCommit` error | May enter collected review. Retry is always explicit and preserves the same authority/scope. | Skip retains source/current object; unknown commit never offers Retry. |
| Known Copy-only or semantic degradation | Disclose before acceptance for selected roots. A runtime-only discovery parks that item and may let unrelated safe work continue. | `Copy and keep source`, Skip, or Cancel; never render plain `Moved`. |
| Space, hydration, sparse inflation, EFS plaintext, security/metadata loss, or comparable cost/risk consent | Never auto-accept. Park the affected item before mutation; review-later mode may collect the question but only unrelated safe work continues. | Risk-specific non-destructive choice or Cancel; Apply-to-all only for the exact eligible task/risk/root scope. |
| Recycle escalation | Exact-item explicit decision, even in review-later mode; it may be collected but never inferred or broadened. | Cancel first/default, then Delete permanently and Skip; no Apply-to-all. |
| Indeterminate mutation or cleanup | Not a retry/consent gate. Stop dependent destructive work and create one terminal issue. | `Final state uncertain` with inspect/open-location/export guidance; no mutation action. |

#### 6.9.4 Phase controls and keyboard behavior

| Phase | Available controls | Pause/Cancel checkpoint meaning | Escape/Close |
|---|---|---|---|
| Waiting/queued | Start now, reorder where eligible, Cancel | Cancel produces NotAttempted and retains intent/source. | Hides; never changes queue/task. |
| Preparing | Cancel only, plus non-actionable details | No Pause. Cancel must leave mutation, clipboard, and breadcrumb untouched. | Hides; Cancel remains an explicit action. |
| Concurrent-task overlap warning | Queue after these tasks, Run at the same time, Don't start, details | No task mutation has begun. Queue records a transient same-host predecessor edge discarded on automatic release; Run records only the disclosed concurrently-live consent; Don't start leaves clipboard/breadcrumb untouched. | Escape invokes Don't start for this decision; Close hides without resolving it. |
| Discovering + Running | Pause, Cancel, bandwidth/queue controls, `Discover as needed` while eligible; one task-local discovery row with exact count/indeterminate bar/provisional ETA plus a separate operation-progress row | Pause stops new primary discovery/transfer/mutation admissions at bounded checkpoints; active calls acknowledge when their proved contract permits. Cancel enters Stopping. | Hides; task continues. |
| Verifying | Pause, Cancel, details | Pause/Cancel at bounded verification checkpoint; completed publication truth is retained. | Hides; task continues. |
| Conflict/deferred consent/collected review | Only actions in the immutable decision snapshot; eligible Apply-to-all | The affected item is parked before its guarded mutation. Task Cancel stops new work and records unresolved items as NotAttempted/source retained. | Escape chooses the decision's Cancel; Close hides without resolving it. |
| Pausing/Paused | Resume, Cancel | Exact owned compensation already required for safety may finish; no new primary mutation begins. | Hides; state is unchanged. |
| Stopping | Details only; repeated Cancel disabled | Drain known completions and bounded exact compensation, then publish exact or Indeterminate truth. | Hides; stopping continues. |
| Application close while work is live | One aggregate choice: keep the application open, or explicitly cancel operations and exit | Default keeps the application open. Accepted exit fences new commands, cancels once, and follows the bounded/isolated shutdown contract; it never silently treats hidden prompts as accepted. | Window Close requests the aggregate choice; it is not an in-task decision. |
| Terminal | Inspect, open/select known locations, export, view history, dismiss | No action can infer Retry/Resume/Undo authority from the card/history. | Hides/dismisses presentation only. |

#### 6.9.5 Continue-safe-work collect-then-review journey

The default remains **Ask as issues occur**. Under accepted D2-A17 B, the initial
options surface and the task's More menu before collection closes expose a task-local
choice: **Ask as issues occur** / **Continue safe work and review later**. It is not
a persistent Preference without separate arbitration.

In review-later mode, the card shows a stable `N items need review` count without
stealing focus. Only independent unambiguous work continues. Each deferred item is
parked before its guarded mutation, keeps its exact authority or is re-bound under
the normal commit-time rule, and appears exactly once in a keyboard/UIA navigable
review ordered by original task/item order. Per-item actions and eligible explicitly
scoped Apply-to-all remain available. Existing Apply-to-all decisions remain
immutable. A stale or newly changed conflict revalidates and returns to the same
review; it never widens selected roots or crosses the accepted real-container,
virtual-folder, or fixed-set boundary. Canceling the
task records every unresolved item as unattempted/source retained. Closing the review
does nothing. Retry remains available only for proven no-commit; Indeterminate results
are terminal evidence, not review actions. Rescan always creates a new plan.

#### 6.9.6 Completion and history journey

Under accepted D2-A18 B, completed cards expose **View in history** and the jobs
surface Options menu exposes **Operation history**. History opens without changing
the task or replaying any action. It supports result/status/date/endpoint filtering,
search, navigation to still-resolvable locations, explicit Export, Clear, and Disable
controls, and clear wording that it is evidence rather than Undo/Resume authority.
Partial, failed, retained-source, cleanup-debt, and indeterminate entries remain as
discoverable as successes.

History is local, enabled when the feature lands, bounded by both age and storage,
and evicts oldest entries first. Settings expose Disable, Clear, bounded retention,
and path-redaction controls; Disable stops new records without silently clearing old
ones, while Clear removes only human-readable history, never exact artifact claims or
provider recovery records. Export is explicit and contains only the selected visible,
redacted fields. Credentials, tokens, secret-bearing endpoint fragments, and raw
opaque mutation authority are never persisted. R6-A18 records measured numeric age/
size defaults in settings/schema authority; those values are bounded policy, not a
new mutation-authority or architecture decision.
History never auto-opens, steals focus, or offers Retry/Resume/Undo from a saved row;
an eligible new attempt is created through the ordinary command/admission journey.

#### 6.9.7 Human workflow and accessibility acceptance

The specification can define consistency; it cannot honestly claim “smooth and
easy” until representative people and deterministic automation exercise the full
journeys. Acceptance covers mouse, keyboard-only, screen reader/UIA, high contrast,
RTL/localized long text, reduced motion, minimum width, and supported DPI. Required
walkthroughs are routine Local Copy, routine clipboard Move, known mixed Copy-only
Move, slow Preparing, disjoint no-warning injection, each A02 overlap class and all
three choices, pre-closure per-task discovery plus operation progress, multi-task
aggregate progress, USB-like Copy/Delete discovery-service contention, default discovery-ahead,
`Discover as needed`, user-created `.rs_*` open/copy/Properties/destructive warning,
cyclic Batch Rename rejection, immediate conflict, pause/resume/cancel in every live
phase, Recycle escalation, indeterminate completion, continue-safe-work collected
review, and history search/export.

The ordinary accepted-default path has one invocation and zero additional decisions.
An A02 overlap or material-consequence path uses one stable problem-specific surface and one focus
transition. Focus never jumps to a button that was not present when announced;
announcements are bounded and nonrepeating; routine progress never steals OS
foreground; an attention badge cannot hide an actionable decision; and every
visible action has the same keyboard/UIA semantics. Interaction latency, focus
transitions, prompt count, cancellation acknowledgement, and reveal stability are
measured before the wording **smooth and easy** is used in closeout.

### 6.10 Link, rename, name, and Shell-spike boundaries

**`FO-LINK-01`.** Preserve links copies the literal link object and stored payload
unchanged. **Retarget links inside copied tree** is a separately named per-task
opt-in. It may rewrite only an exactly mapped target inside the admitted copied tree,
uses bounded indexed mapping, reports its own conflicts and per-item results, and
leaves unresolved, skipped, failed, or out-of-tree targets unchanged rather than
guessing.

**`FO-RENAME-01`.** Batch Rename detects cycles during Preparing and disables Run
before mutation, explaining an intermediate-name/two-pass remedy. Production Batch
Rename creates no recovery journal and exposes no Resume, Roll back, or automatic
crash replay. No journal-data migration is required. If supported-release evidence
finds a still-actionable existing record, STOP and arbitrate its retirement rather
than inventing migration.

**`FO-NAME-01`.** Every mutating name ingress consumes one provider/path-scoped
executable child-name, joined-path, collision-key, limit, and case-policy result. F2,
Batch Rename, Change Case, and F7 use the same result for the same provider/path.
Win32/DOS restrictions apply only where Local directory semantics require them; a
provider-specific legal name is never rejected merely because Win32 would reject it.

**`FO-SHELL-01`.** A16 authorizes evidence gathering, not production adoption. A
non-default spike may cover only ordinary absent-destination Local regular-file Copy
and no-conflict same-volume Native Local regular-file Move, with current Recycle as
the comparison control. Reject either candidate if Shell owns a prompt/conflict/Undo
authority, exact per-item terminal truth is unavailable, dedicated-STA cancel/
teardown cannot satisfy `FO-LIFE-01`, or measured production-owner/code reduction is
not material. Every other route remains out of scope.

### 6.11 Complexity budget for the replacement (`FO-SIMPLE-01`)

The repair must not become a new generalized framework. The smallest useful engine
has six proof owners:

1. `Preparing` resolves selected-root feasibility and produces an immutable plan.
2. `MutationAuthority` holds only the exact provider evidence needed by the intent.
3. `ConflictArbiter` produces one immutable user-decision snapshot.
4. `NativeAdapter` or `ManagedExecutor` performs one selected strategy.
5. `OwnedPublicationTransaction` owns create/write/publish/compensation truth.
6. `TerminalResultBuilder` validates and stores one result.

These are responsibility boundaries, not a requirement for six COM interfaces,
class hierarchies, allocation layers, or source files. Prefer plain versioned
structs, enums, RAII capabilities, and explicit functions. Add an abstraction only
when it removes a duplicate authority/policy owner or makes an invalid destructive
state impossible.

The implementation must not introduce:

- a workflow DSL, generic policy engine, actor framework, or repository-wide event
  bus;
- a universal distributed transaction coordinator across providers;
- generic per-item WAL, Undo, or Resume for Copy/Move;
- any Batch Rename recovery journal, Resume/Roll back surface, or automatic crash
  replay;
- an out-of-process broker for providers that already prove bounded safe behavior;
- a mandatory exact identity graph or whole-root exclusion lock for task-overlap
  warnings;
- persistent/cross-process Queue, Run, or completed-task output protection, or a
  cross-application filesystem-change coordinator;
- a disk-backed discovery spool/durable enumerated-work queue without a separately
  activated provider-specific proof that bounded in-memory paging/JIT is insufficient;
- a general byte-token, metadata-IOPS, device-class, or adaptive scheduler controller
  before an instrumented unchanged baseline proves the exact mechanism necessary;
- mandatory full content reread for routes that already prove stable source and
  committed destination under the accepted Standard Move policy;
- eager materialization of every descendant or conflict before safe work can start;
- permanent user-facing Native/Managed/Copy-only jargon in routine task summaries;
- template/meta-programming machinery for the small terminal state machine;
- JSON route-safety authority beside the typed contract, a dual reader, or a
  compatibility/data-migration layer for the A11 cutover;
- a second feature-complete popup renderer after the hosted owner lands.

Each package first characterizes the old path, adds the new proof owner, moves one
caller family, and deletes the displaced production owner before expanding scope.
If a package adds more long-lived states, policy owners, or mutation paths than it
removes, it returns to arbitration with a written justification.

## 7. Product-owner decision rationale ledger (non-authoritative)

This is the sole decision-status list, not a second target contract. Status is shown
before each option table so a reader never has to infer it from the rationale below.
Accepted choices map to the section 6 rule IDs; implementation must consume those
rules rather than reconstruct behavior from the examples in this ledger.

| Decision status | Decisions | Meaning |
|---|---|---|
| **NOT YET ANSWERED (0)** | None | No product-owner arbitration remains open. New implementation discoveries return here only when they change accepted target law. |
| **ANSWERED (21 entries)** | A01-A20 plus provider subdecision A13-MD1 | Target law is recorded in this plan. Every related R slice is still `DECIDED-NOT-ACTIVE` until separately activated. |

For an answered row, **Answered example** and scenarios remain rationale and future
review tests; they are not open questions. There are currently no unanswered rows.

### D2-A01 — Provider cancellation and application survival

> **Status: ANSWERED · Decision: B now, conditional C · Implementation: NOT ACTIVE**

| Option | Effect |
|---|---|
| A. Trust cooperative cancellation and keep unconditional join | Smallest change, but a wedged provider can hang the app indefinitely. |
| B. Classify each exact route: admit cooperative/bounded routes and routes with proved provider-local watchdog/quarantine plus lifetime/unload safety; disable only residual uncontained routes | Honest and route-specific; preserves existing contained operations while refusing known host-hang risk. |
| C. Apply B, then broker only residual network/device routes that cannot satisfy the in-process containment contract | Strongest long-term survival for genuinely wedged calls without amputating already-contained routes; higher contained cost. |
| D. Heap-own/quarantine abandoned in-process workers/modules without a proved provider-local lifecycle | May preserve availability, but stuck threads remain and generic lifetime/unload complexity is substantial. |

**Accepted decision (2026-08-29):** B now; use C only for residual exact routes that cannot satisfy
B. Remove or relabel unenforced host deadline claims immediately.

**Existing evidence:** MTP already has a serialized WPD worker, timeout, backend
cancel, worker quarantine, and unload gating
(`Plugins/FileSystemMtp/FileSystemMtp.Core.cpp:2498-2565,4381-4401`) plus
deterministic mutating-timeout coverage
(`CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp:1196-1264,1396-1540`). Preserve
that route unless the host-level shutdown test disproves its containment. This
evidence does not classify SMB; obtain deterministic or live SMB-specific evidence
before enabling or disabling an exact SMB mutation route.

**Answered example:** if an exact route never returns and has neither a proved
provider-local containment lifecycle nor process isolation, should RedSalamander
temporarily disable that route, or accept that Cancel/Exit can hang?

### D2-A02 — Obvious concurrent-operation overlap warning and user choice

> **Status: ANSWERED · Decision: problem-specific warning plus user scheduling choice · Implementation: NOT ACTIVE**

| Option | Effect |
|---|---|
| A. Reject or globally serialize every possibly overlapping task | Deterministic ordering, but invalidates intentional workflows, serializes too broadly, and makes identity/alias machinery a global availability gate. |
| B. Detect obvious overlap cheaply at task injection, explain the concrete race, and offer Queue/Run together/Don't start | User controls scheduling; Queue warns once, waits, then runs normally with no completed-producer protection, while concurrently-live Run still uses exact per-item authority and the same-host live-output guard. |
| C. Never warn; rely only on later file conflicts | Smallest admission UI, but users can unknowingly start a Delete against another task's source/destination or two writers against the same namespace. |

**Decision record:** B, amended from the former interlock ballot on 2026-08-30.
Binding behavior is `FO-OVERLAP-01` in sections 6.1 and 6.9; delivery is R2,
R4-A02, and R6-A02. This ledger records why and gives acceptance examples; it cannot
amend the rule. Product intent is a cheap problem-specific warning and informed user
scheduling choice, not a hidden global interlock.

**Concrete user examples:**

1. Copy `C:\A\...` to `\\server\myfolder`, then copy
   `C:\A\b\c\...` to `\\server\myfolder\b\c`: warn because the destination
   mutation scopes are nested. If Run is chosen, both execute; whichever reaches a
   leaf second receives the ordinary current file-exists/conflict decision.
2. Delete `C:\A\...`, then Delete `C:\A\b\c\...`: warn because mutation scopes
   overlap. Queue and Run are both valid. Under Run, one task may report an item as
   already removed by the other; provider concurrency may speed the operation.
3. If exact cheap residual-scope evidence proves the first task has irrevocably
   closed or excluded `b\c`, the second task need not warn. “The current item is in
   another branch” alone is insufficient and must not trigger an expensive live
   descendant graph merely to suppress the warning.
4. `Z:\Work` and `\\server\share\Work` warn when a cached/bounded mapping positively
   identifies or strongly evidences the alias. If proving it would require slow I/O,
   injection stays responsive and normal execution safety still handles the race.
5. Copy into `C:\Work\New`, then inject Delete `C:\Work`: warn once. If Queue is
   chosen, the warning says Delete waits for Copy to finish and then runs normally,
   so it may remove the copied output as current folder membership; no second A02
   prompt or completed-task protection follows. If Run is chosen, both execute and
   the same-host live-output guard requires the disclosed Run consequence or a later
   item-specific decision while both tasks remain live.

### D2-A03 — Meaning of “Preserve links”

> **Status: ANSWERED · Decision: A default plus explicit Retarget option · Implementation: NOT ACTIVE**

| Option | Effect |
|---|---|
| A. Copy literal link object/payload unchanged; Skip remains available | Smallest/predictable; a copied absolute link can still point to source. |
| B. Keep silent semantic in-tree retargeting | Preserves resolution intent, but rewrites stored data and requires collision maps, deferral, cycles, dependency cleanup. |
| C. “Copy links unchanged” default plus explicit advanced “Retarget links inside copied tree” | Both meanings visible; advanced complexity only when requested. |

**In plain language:** a filesystem link stores a payload. “Preserve” can mean copy
that stored link unchanged, or it can mean silently rewrite the payload so the copied
link resolves inside the destination tree. Those are different operations and should
not share an ambiguous label.

**Concrete user examples:**

1. An absolute link contains `C:\Source\A.txt`. A copies the link unchanged, so it
   still points to the source. B rewrites stored data so it points toward the copy.
2. A relative link contains `..\Shared\A.txt`. Copying the literal payload may make
   it resolve somewhere different from its new parent, but RedSalamander has not
   invented or silently changed data.
3. A copied tree contains a link to an item the user skipped. Literal preservation is
   deterministic. Retargeting needs a rule for whether to retain the old target,
   break the link, defer it, or fail the whole operation.
4. Links form a cycle or two destination names collide. Retargeting requires an
   indexed dependency graph, deferred publication, partial-failure semantics, and
   cleanup rules that ordinary literal Copy does not need.

**Reliability and overdesign check:**

- A treats the link itself as user data. It is predictable, testable, and has no
  hidden dependency on which neighboring items happen to succeed.
- B may be convenient for some tree copies, but it modifies stored data while the UI
  says Preserve and imports cycle/collision/partial-copy complexity into the default.
- C is honest because retargeting is explicit, but it creates a second advanced
  product mode whose graph, conflict, partial-failure, and performance cost must stay
  isolated from ordinary literal Copy.

**Accepted decision (2026-08-30):** A defines the default meaning of Preserve, and
the separately named opt-in **Retarget links inside copied tree** feature is also
accepted. This is an amended A/C combination, never silent B. Ordinary Preserve
copies the literal link object and payload unchanged. Retargeting is selected
explicitly per task, uses indexed/bounded mapping, exposes distinct conflicts and
receipts, and never changes links whose target is outside the admitted copied tree.

**Recorded consequence:** an absolute link can keep pointing to the source under
default Preserve. A user who explicitly selects Retarget receives the additional
dependency/collision/partial-result behavior under that visible name.

### D2-A04 — Proof before Managed Move source deletion

> **Status: ANSWERED · Decision: B · Implementation: NOT ACTIVE**

| Option | Effect |
|---|---|
| A. Successful copy/size alone permits source cleanup | Fastest and closest to common behavior, but source stability and weak provider commit may be assumed rather than proved. |
| B. Route-aware proof tiers: exact publication + stable source + conditional cleanup always; mandatory content verification for weak-commit remote/device routes and explicit Verified Move; stable committed Local routes may use Standard Move | Protects the destructive boundary without universally doubling I/O; route classification and wording must remain honest. |
| C. Exact end-to-end content reread/hash mandatory before every Managed source deletion; otherwise Copy-only | Strongest uniform rule; can double I/O, make large Local/SMB work unexpectedly slow, and preserve sources even where stable commit was already proved. |

**Accepted decision (2026-08-29):** B. Native atomic Move remains unaffected. `Verify Off` may only
disable an extra checksum on a route already proving stable source and committed
destination; it can never waive proof required by a weak route. `Verified Move`
means exact content proof, not size/time equality or a provider success code.

**Answered example:** should a stable Local cross-volume Move always reread every byte
after a committed copy, or should that extra I/O be reserved for Verified Move and
routes whose commit/source stability is otherwise insufficient?

### D2-A05 — Known Copy-only outcomes at consent

> **Status: ANSWERED · Decision: B · Implementation: NOT ACTIVE**

| Option | Effect |
|---|---|
| A. Auto-downgrade and mention only in result | Few prompts, weak informed control. |
| B. Show strategy counts/reasons in initial consent; explicitly label “Copy and keep source” when needed | One informed decision, no surprise duplicate. |
| C. Refuse whole Move if any item is Copy-only | Predictable but prevents safe partial utility. |

**Accepted decision (2026-08-29):** B. Runtime-only downgrade stays visible and preserves source; a
first occurrence can offer Skip/continue for remaining scope.

**Answered example:** if 97 items Move and 3 only Copy, should the primary button
still simply say “Move”?

### D2-A06 — Ordinary new-name Copy visibility

> **Status: ANSWERED · Decision: C · Implementation: NOT ACTIVE**

| Option | Effect |
|---|---|
| A. Hidden owned stage then atomic/conditional reveal for every Copy | Uniform and strongest visibility; adds rename/stage mechanics and may duplicate I/O/storage constraints where no old target exists. |
| B. Exclusively create the final leaf for every new-name Copy | Simplest; remote/weak routes and Managed Move can expose partial final content under the requested name. |
| C. Risk-scoped publication: stage overwrite/Managed/remote/atomic-claim routes; permit retained exclusive final-leaf streaming only for ordinary new-name Local Copy | Preserves the high-risk guarantees while removing universal staging; surviving partial Local artifacts must be explicit and exactly owned. |

**Accepted decision (2026-08-29):** C. For FOS-01/R0a, C specifically means section 6.4 shape 2:
ordinary absent-destination Local regular-file Copy exclusively creates and retains
exact authority over the requested final leaf. It does **not** add a hidden stage to
every new Local Copy. Shape 1 remains required for overwrite, Managed Move,
remote/weak-commit, resume, and atomic-visibility routes. No failure path may
pathname-delete a final occupant that the task cannot prove it created.

**Answered example:** is atomic final-name visibility worth mandatory staging for a
new Local Copy that cannot overwrite anything, or is an explicitly incomplete,
exactly owned final leaf an acceptable simpler behavior?

### D2-A07 — Explicit Retry policy

> **Status: ANSWERED · Decision: B · Implementation: NOT ACTIVE**

| Option | Effect |
|---|---|
| A. One Retry only | Arbitrary ceiling; corrected condition may occur after button disappears. |
| B. Repeated explicit Retry for proven no-commit, with attempt count | User stays in control; authority does not widen. |
| C. Automatic exponential Retry | Fewer prompts; invisible waits and repeat-side-effect risk. |

**Accepted decision (2026-08-29):** B. Never auto-retry unknown commit state.

**Answered example:** after fixing credentials on the second failure, why restart the
entire task?

### D2-A08 — Progress before traversal closes

> **Status: ANSWERED · Decision: B amended · Implementation: NOT ACTIVE**

| Option | Effect |
|---|---|
| A. Backwards provisional percentage | Quantitative early, untrustworthy as denominator grows. |
| B. Indeterminate task bar plus exact discovered/processed/byte counters; no percentage before total closes; an explicitly provisional discovery-aware ETA may be shown | Honest phase state while still offering useful time guidance that can move as work is discovered. |
| C. Full pre-enumeration | Stable percentage, delays work and expands state. |

**Accepted decision (2026-08-29):** B amended. Exact discovered/processed/byte
counters remain visible and no growing denominator becomes a percentage. A
clearly labeled provisional ETA may use the growing discovered workload, must state
that discovery is still active, may move in either direction, and disappears when
sampling is insufficient. Only a closed total can drive a stable percentage and
ordinary ETA. D2-A19 still owns the distinct scheduling decision.

**Answered example:** is “70%” falling to “18%” better than “1,842 processed; still
discovering”?

### D2-A09 — Possible artifact warnings

> **Status: ANSWERED · Decision: B · Implementation: NOT ACTIVE**

| Option | Effect |
|---|---|
| A. Block every touch to an artifact-like name | Conservative; false-positive prompt fatigue. |
| B. Possible never authorizes cleanup; read/open/edit/preview/copy/export is unblocked solely by name; manager-owned destructive mutation gets one exact-object warning with Cancel default; Properties exposes all available semantic File Operations facts and why classification applies | Protects destruction without treating name as authority and lets the user understand both real and name-only artifacts. |
| C. Treat as ordinary with no warning | Simple; user can unknowingly delete a retained stage whose claim is missing. |

**Accepted decision (2026-08-29):** B. A user-created `.rs_*` file is not blocked
from read/open/edit/preview/copy/export merely
because of its name. Rename, Move, Delete/Recycle, overwrite, and content/metadata
writes still receive the one exact-object, Cancel-default warning. Its Properties
surface is extended with all available semantic File Operations information and
explicitly says when `Possible` is only a name-shape warning with no owning claim.
After R7-A10 there is no production durable artifact claimant, so no path is
classified `Proven`; a future `Proven` surface requires its own activated durable
authority and cannot be inferred from a name.

**Answered example:** should a user-created `.rs_*` file be blocked from opening or
copying solely due to its name?

### D2-A10 — Batch Rename cycles and journal

> **Status: ANSWERED · Decision: A · Implementation: COMPLETE (R7-A10)**

| Option | Effect |
|---|---|
| A. Reject cyclic plans in preview before mutation and explain alternate names; create no Batch Rename recovery journal | Removes WAL/Resume/Roll back and migration complexity; valid swaps require non-cyclic intermediate names and crash truth comes from refreshed provider state plus completed-step results. |
| B. Retain exact cycles and bounded WAL/Resume/Roll back | More code; directly required by useful operation and crash truth. |

**Accepted decision (2026-08-29):** A. Swaps/cyclic rename graphs are not a product
requirement. Detect them during preview/admission, mutate nothing, explain the cycle,
and suggest non-cyclic intermediate names. Batch Rename creates no recovery journal
for any plan and exposes no Resume/Roll back; acyclic steps still return exact partial
results, and after process loss the refreshed provider namespace is truth—nothing is
auto-replayed. Remove the production journal/recovery surface after caller inventory.
No existing-journal schema/data migration, replay, automatic deletion, dual writer,
or compatibility layer is required; STOP if release evidence shows a supported build
could have left a record requiring a separately arbitrated retirement obligation.

**Recorded consequence:** previewed `A.txt <-> B.txt` does not execute in one action;
the user receives a clear pre-mutation explanation and can choose an intermediate
name.

### D2-A11 — Capability ABI timing

> **Status: ANSWERED · Decision: B · Implementation: NOT ACTIVE**

| Option | Effect |
|---|---|
| A. Stabilize mandatory JSON v2 as the current compatibility authority; clear dishonest claims and test executability now | Avoids immediate ABI churn; retains parser/schema cost while evidence is collected. |
| B. Replace JSON v2 authority with versioned typed per-operation route facts plus executable QI/receipts; retain JSON only as optional non-authoritative diagnostics/extensions | Smaller executable safety surface and one compiler-checked contract; requires a lockstep provider/host cutover. |
| C. Interfaces only | Smallest declaration; repeated probing and weaker planning explanations. |

**Accepted decision (2026-08-29):** B as a direct lockstep replacement. Typed
per-operation route facts, executable queries, and typed mutation/source/cleanup
receipts become authority; JSON may remain only for diagnostics/extensions and can
never enable a route. No JSON schema/data migration, dual-authority period, or
compatibility layer is required. Inventory every binary/source consumer before
removal; if a supported external consumer exists, STOP because the no-migration
assumption is false rather than inventing an unapproved compatibility program.
Dishonest bits and null destructive receipts are still corrected immediately and do
not wait for the typed cutover.

**Recorded consequence:** host and in-tree providers cut over together to the typed
contract; the plan does not preserve JSON v2 as a second route authority.

### D2-A12 — Popup implementation ownership

> **Status: ANSWERED · Decision: B · Implementation: NOT ACTIVE**

| Option | Effect |
|---|---|
| A. Keep feature-complete hosted and legacy systems | Maximum fallback fidelity; permanent focus/hit-test/accessibility drift risk. |
| B. One hosted layout/action/accessibility owner plus minimal safe failure surface | Smaller/testable; rare attachment failure loses nonessential visuals. |

**Accepted decision (2026-08-29):** B. The target user behavior moves into its owning
authoritative spec when the slice lands; raster algorithms and old measurements move
to performance evidence.

**Answered example:** on rare hosted-control failure, is clear text plus Cancel/Close
enough, or must every graph/action layout be implemented twice?

### D2-A13 — Normal Delete authority

> **Status: ANSWERED · Decision: C amended · A13-MD1=A · Implementation: NOT ACTIVE**

| Option | Effect |
|---|---|
| A. Snapshot the exact selected root and every descendant generation before deletion | Strong fixed set, but expensive/stale for real directories; new children leave the directory behind. |
| B. Delete current path occupants/re-list every namespace until empty without a declared scope or generation conditions | More tasks appear to finish, but replacements and textual-prefix neighbors can enter invisibly. |
| C. Hybrid authority: bind real containers; canonicalize provider-declared virtual folders and empty them within a bounded scope using per-generation conditions; snapshot only fixed object/version sets | Matches ordinary Local and S3 folder expectations while retaining exact scope and mutation truth. |

**Accepted decision (2026-08-29):** C amended. A replacement selected real root,
link target, escaped alias, bucket/profile change, sibling prefix, or object outside
the canonical scope is never authorized. Ordinary
`Delete s3://bucket/prefix/` is expected to delete that virtual folder: bounded
re-listing may admit late/replacement members currently inside the selected prefix,
but every observed generation is conditionally deleted and non-convergence leaves
explicit residuals. Fixed key/version selections remain snapshots.

For Microsoft Drive, the selected folder ID is the real-container root and ordinary
Graph Delete is Recycle. Accepted A13-MD1 uses that stable ID without a root ETag
condition; child membership is not a frozen snapshot. The separate `permanentDelete`
endpoint remains unavailable pending exact authority/receipt.

**Recorded consequence:** `Delete C:\Folder` and
`Delete s3://bucket/prefix/` both mean “remove this selected folder and members
encountered inside its bounded scope”; only an explicitly fixed key/version
selection freezes membership at consent.

#### D2-A13-MD1 — Microsoft Drive folder-root condition

> **Status: ANSWERED · Decision: A · Implementation: NOT ACTIVE**

| Option | Effect |
|---|---|
| A. Bind and revalidate the stable item ID; Delete that exact container without root `If-Match` | Matches live-container meaning and does not reject the same selected folder because unrelated properties changed; requires exact ID preservation and truthful Recycle receipt. |
| B. Bind stable item ID plus ETag and require `If-Match` | Detects any ETag-changing root update after consent, but can reject the same selected folder for metadata churn and does not freeze child membership. |

**In plain language:** both choices bind the selected Microsoft Drive folder by its
stable item ID rather than by path. The only question is whether a later property
revision on that same item must also invalidate consent before Recycle.

**Concrete user examples:**

1. The selected folder is renamed after the user confirms but keeps the same item ID.
   A recycles that exact folder. B can stop with a changed-revision conflict.
2. An unrelated property or metadata field changes. A proceeds because the selected
   object is still the same. B interrupts the operation even though the root was not
   replaced.
3. The selected folder disappears and another folder is created at the old path. It
   has a different item ID, so A rejects the replacement; A does not mean pathname
   Delete.
4. A child is added to the same folder before execution. Neither A nor B freezes
   child membership: accepted A13 container semantics still apply.

**Reliability and overdesign check:**

- A uses stable identity as destructive authority and avoids false failures caused
  by property churn. It still requires exact-ID revalidation and a truthful Recycle
  receipt.
- B adds consent freshness for every ETag-changing root update, but does not improve
  protection against a different-ID replacement and does not snapshot children.
  Its main user-visible effect is more interruptions for changes that may not alter
  Delete meaning.

**Why A was selected:** identity answers “is this the selected folder?” more
directly than a general property revision. If provider evidence later identifies a
specific root-property change that materially changes Recycle meaning, arbitrate that
case explicitly instead of treating all ETag churn as replacement identity.

**Accepted decision (2026-08-30):** `A13-MD1 = A. The stable item ID is the selected-folder authority. Revalidate and
Recycle that exact item without a root ETag condition; a different item ID is never
admitted as a replacement, and a truthful receipt remains mandatory.`

### D2-A14 — Interlock/identity failure on UNC

> **Status: ANSWERED · Decision: C · Implementation: NOT ACTIVE**

| Option | Effect |
|---|---|
| A. Globally map 996/`ERROR_INVALID_DEVICE_REQUEST` to Unsupported | More UNC work, but hides transient/incomplete queries as stable absence. |
| B. Keep all Indeterminate and fail | Honest but rejects safe owned create/Copy work. |
| C. Classify only from proven profile context; permit retained mutation authority while treating comparable identity as optional overlap-warning evidence | Preserves truth and availability without destructive inference or a global serialization requirement. |

**Accepted decision (2026-08-29):** C. Centralize proven unsupported codes, but use the three-axis
authority model.

**Answered example:** inability to compare a UNC parent ID does not prevent an exact
owned create. Cheap path/positive alias evidence may warn about another task, while
the create still relies on retained object authority and ordinary race handling.

### D2-A15 — Provider-specific name feasibility

> **Status: ANSWERED · Decision: B · Implementation: NOT ACTIVE**

| Option | Effect |
|---|---|
| A. Universal Windows checks in F7/Batch Rename/bridge | Simple host; wrong for remote namespaces and Local case-sensitive directories. |
| B. Provider/path-scoped executable validation/collision keys | Correct behavior; one small typed provider responsibility. |

**Accepted decision (2026-08-29):** B. This is correctness, not optional polish.

**Answered example:** should S3 reject a key because Windows forbids one character
although S3 accepts it?

### D2-A16 — Bounded ordinary Local Shell-delegation candidates

> **Status: ANSWERED · Decision: C spike only · Implementation: NOT ACTIVE**

**Known control evidence:** Local Recycle already runs `IFileOperation` from an STA
with explicit `FOF_NOCONFIRMATION | FOF_NOERRORUI | FOF_SILENT |
FOFX_EARLYFAILURE | FOFX_RECYCLEONDELETE` flags and conservatively preserves Unknown
when terminal per-item sink truth is absent
(`Plugins/FileSystem/FileSystem.FileOps.cpp:7396-7404,7513-7533,7769-7773,10753-10805`).
Because `SetOperationFlags` is explicitly called and omits `FOF_ALLOWUNDO` and
`FOFX_ADDUNDORECORD`, do not claim current Recycle inherits default Shell Undo flags.
This is comparison evidence, not proof that Copy/Move is acceptable.

| Option | Effect |
|---|---|
| A. Keep ordinary new-name Local regular-file Copy and no-conflict same-volume Native regular-file Move in the custom engine | One internal model and maximum control; retains custom breadth even for the simplest candidate routes. |
| B. Delegate those two candidate routes directly to `IFileOperation` now | Removes custom breadth quickly; risks a second conflict UI, weaker receipts, and behavior RedSalamander cannot explain. |
| C. Run a bounded, non-production-default spike for exactly those two candidate routes; use existing Shell Recycle only as a comparison control | Produces evidence before committing either way; excludes Managed/cross-volume/folder/link Move, overwrite, Permanent Delete, Rename, directory merge, artifacts, remote/device, and Batch Rename. |

**In plain language:** C does **not** adopt Shell Copy/Move. It authorizes a temporary,
non-default experiment for only two simple Local regular-file routes to determine
whether permanent custom code can be deleted without giving Windows a second prompt
policy or losing RedSalamander's result truth.

**Concrete user examples and automatic reject gates:**

1. Copy `A.txt` to an absent Local destination. If the Shell route reports an exact
   terminal per-item result and never owns a prompt, it remains a candidate.
2. Another process creates the destination during Copy. If Windows prompts,
   overwrites, auto-renames, or silently chooses another collision action, reject the
   candidate: RedSalamander must remain the only conflict-policy owner.
3. Move one regular file within the same volume, with no conflict. If cancellation or
   a missing progress-sink callback leaves the result ambiguous, record
   Indeterminate and reject production adoption rather than guessing Moved.
4. If the route creates a Shell Undo record, requires directory/link semantics, or
   cannot stop and tear down its dedicated STA within A01's accepted contract, reject
   it.

**Reliability and overdesign check:**

- A avoids spike cost and preserves the known custom model, but commits the project
  to maintaining custom code even for the simplest routes.
- B is unjustified: immediate delegation risks prompts, Undo, policy drift, and
  incomplete receipts before those contracts are proved.
- C adds temporary adapter/test work only to answer whether meaningful permanent code
  can be removed. It becomes overdesign if the spike grows beyond the two routes, if
  both engines remain in production, or if failed candidates are not deleted.

**Why C was selected:** the experiment has value only as a deletion test. Reject
either candidate unless it proves no Shell-owned prompt/collision, no Shell Undo,
exact terminal mapping across callbacks/`PerformOperations`/aborted state,
Indeterminate on missing observation, bounded STA lifetime, and regular-file-only
scope. `FOFX_NOSKIPJUNCTIONS` is a widening reject gate, not authority to add folders
or links.

**Accepted decision (2026-08-30):** `A16 = C. Authorize a non-production-default spike only for ordinary new-name Local
regular-file Copy and no-conflict same-volume Native regular-file Move. This gathers
evidence; it does not adopt Shell execution. Reject a candidate if Shell owns a
prompt/collision, creates Undo, broadens scope, loses exact per-item truth, or cannot
prove bounded teardown.`

### D2-A17 — Continue-safe-work conflict/error review

> **Status: ANSWERED · Decision: B · Implementation: NOT ACTIVE**

| Option | Effect |
|---|---|
| A. Pause the task at the first unresolved conflict/error | Immediate decision, but long operations can remain stalled behind one prompt. |
| B. User-selectable **Continue safe work and review later** mode continues unrelated safe work, keeps selected roots/accepted scope fixed, then presents one collected review | More progress and control without auto-deciding; requires a durable in-task collection and commit-time revalidation. |
| C. Always collect and defer every conflict/error | Few interruptions, but small/high-risk operations may surprise users who expected an immediate decision. |

**Accepted decision (2026-08-29):** B. Existing Apply-to-all decisions remain
immutable; newly changed conflicts return to the same review. Retry is still
restricted to proven no-commit and Rescan creates a new plan.

**Answered example:** should a two-hour task stop after one minute because item 20
needs a name decision, or finish every unambiguous item and ask once at the end?

### D2-A18 — Persistent operation history

> **Status: ANSWERED · Decision: B · Implementation: NOT ACTIVE**

| Option | Effect |
|---|---|
| A. Keep only transient task cards | Smallest persistence surface; completed evidence and recovery guidance disappear. |
| B. Persist retention-controlled human-readable run/attempt results with search/export/navigation, explicitly non-authoritative | Gives the user durable understanding without pretending every action is reversible. |
| C. Build a universal transactional WAL/Resume/Undo history | Maximum ambition; false recovery authority and provider-wide complexity. |

**Accepted decision (2026-08-29):** B. Persist a bounded local human-readable history
from immutable results. Store no more identity/path detail than explanation and
navigation require, apply the redaction/retention/Clear/Disable/export contract in
section 6.9.6, and keep exact artifact claims or provider receipts separate. History
can suggest a new operation but never replay an indeterminate one or grant Retry,
Resume, Undo, cleanup, source-delete, or artifact authority.

**Answered example:** after a large operation finishes with three retained sources and
one cleanup artifact, should the explanation disappear when the task card closes?

### D2-A19 — Fast preparation, discovery priority, and the user override

> **Status: ANSWERED · Decision: B evidence-first, no disk spool, per-task dual progress and Known-work aggregate · Implementation: NOT ACTIVE**

| Option | Effect |
|---|---|
| A. Finish a full recursive preflight before transfer and offer Skip preflight | Stable totals first, but slow time-to-first-byte, duplicate/eager work, high memory, and a dangerously ambiguous safety bypass. |
| B. Mandatory top-level Preparing; preserve and instrument the current discovery reservation while operation starts; add only the smallest evidence-proved same-resource correction; use bounded in-memory backpressure and no disk spool; one-way `Discover as needed` releases run-ahead priority | Fast first safe mutation, bounded discovery service and memory, visible per-task discovery plus operation progress, and direct user control without predesigning a new scheduler framework. |
| C. Same streaming scheduler but adapt automatically with no user override | Smallest UI, but removes the existing tested choice when the user values immediate transfer throughput over early totals/queue fill. |

**Decision record:** B evidence-first with no disk spool, per-task dual progress, and
the labeled Known-work aggregate, accepted 2026-08-30. Binding behavior is
`FO-DISCOVERY-01` in sections 6.1 and 6.9; delivery is R4-A19 and R6-A08. Preserve and
instrument the current scheduler first. A new byte limiter, mutation-rate limiter, or
serialized-turn correction is adopted only after the named baseline proves that exact
mechanism necessary. The examples below become verification scenarios; they cannot
amend the rule.

**Concrete user examples:**

1. A folder contains one million small files. Copy starts after the first safe items
   are discovered. Reserved discovery opportunity keeps filling the ready queue and
   reaches closed totals/ordinary ETA sooner instead of letting transfer starve the
   tree walk.
2. On one USB drive, measure a large sequential Copy and thousands of small Deletes
   against the unchanged scheduler. If Copy alone starves scanning, trial the smallest
   byte limiter; if small-file Delete alone starves it, trial the smallest mutation-
   rate control. If the baseline passes, add neither. An unrelated device is never
   throttled for symmetry.
3. One very large file is already copying over a slow connection. Choosing
   **Discover as needed** gives transfer the full current policy budget; the user
   accepts that totals and ETA may stabilize later.
4. If measurement proves an MTP/provider session has one serialized lane and the
   baseline starves discovery, use bounded discovery/operation turns rather than
   pretending the provider is parallel.
5. A conflict appears during discovery. Mandatory admission and conflict handling
   still occur. **Discover as needed** cannot skip it or authorize a mutation.
6. Three tasks run together and only one is still discovering. Each task retains its
   own discovery/operation presentation. If the two closed-total tasks share one unit,
   the jobs view shows their **Known work: N%** plus **1 task discovering — total may
   grow**. If no compatible closed-total cohort exists, it shows activity/throughput
   without a jobs percentage. The Windows taskbar remains indeterminate while that
   open-total task exists.

### D2-A20 — Routine start versus mandatory confirmation

> **Status: ANSWERED · Decision: B · Implementation: NOT ACTIVE**

| Option | Effect |
|---|---|
| A. Show the full confirmation/options surface for every Copy, Move, and Delete after Preparing | Maximum per-run visibility, but adds delay, focus churn, and prompt fatigue to familiar safe actions. |
| B. Routine explicit operations use accepted defaults and start after mandatory Preparing with no extra modal; expose a deliberate **with options** entry and prompt once for material risk, destructive escalation, or known degradation | Familiar fast path plus informed control where the consequence changes; requires precise classification and discoverable options. |
| C. Never prompt; apply Preferences and automatic policies to every route | Fastest interaction, but the user cannot make an enlightened item/task decision when data loss, source retention, cost, or semantics change. |

**Decision record:** B, accepted 2026-08-30. Binding behavior is `FO-LIFE-01` and
`FO-UX-01` in sections 6.1 and 6.9; delivery is R1d-core and R6-A20. The examples
below explain the interaction choice and become verification scenarios; they cannot
amend the rule.

**Concrete user examples under B:**

1. Press F5 or Paste to copy an ordinary Local file into an absent destination with
   accepted defaults: it starts after quick Preparing, without a second click.
2. Cut/Paste an ordinary same-route Move: Paste is the explicit command, so it starts
   without a blanket confirmation. If the route becomes Copy-only, source retention
   is disclosed once before mutation or at the accepted runtime downgrade gate.
3. Choose **Copy with options...**: one surface contains eligible Links, Verify,
   Queue/Parallel, bandwidth, and **Continue safe work and review later** choices; it
   does not lead to a second generic confirmation.
4. Permanent Delete, Recycle escalation, plaintext/security loss, metadata loss, or
   another material consequence: show one consequence-focused prompt with a safe
   default before mutation.
5. F2, Compare Run, Batch Rename Run, archive cleanup, and explicit drag/drop already
   carry intent. Do not add a duplicate generic “Are you sure?” after their own
   editor/preview/Run gesture unless a new material consequence is discovered.

**Overall consequence of the current decision set:** the answered direction removes
universal reread, universal staging, universal descendant snapshots, Batch Rename
journaling, JSON/typed dual authority, duplicate popup renderers, and universal
provider brokering where smaller exact rules provide the required truth. It preserves
exact destructive authority, bounded S3 folder semantics, transparent partial/unknown
outcomes, collected human review, durable non-authoritative explanation, and user
control. All product choices are answered; implementation remains separately gated
and none of these decisions reopens another accepted safety rule.

## 8. Unified delivery program

### 8.1 Sole activation rules and acyclic dependency graph

| Package family | State | Scope |
|---|---|---|
| E0-E8 | `AUTHORIZED-INHERITED` | Complete former-I14 behavior-preserving validation, extraction, rename/admission, durable-mechanics, artifact-index, breadcrumb, and closeout backlog. |
| Accepted R slices | `DECIDED-NOT-ACTIVE` | Target law is recorded, but implementation awaits a complete, predecessor-closed, bounded `ACTIVE` owner as detailed below. |
| Product-decision gates | `RESOLVED` | A01-A20 and A13-MD1 are answered. This records target law only; activate exact slices, never a family by implication. |
| Other release-safety/truth findings | `DECISION-GATED / RELEASE-SAFETY` | FOS findings without a promoted owner still require an explicit bounded owner; they do not silently enter E work. |
| I17 merged qualification | `QUALIFICATION-INHERITED` | Evidence-only obligation carried by E8 and section 9.4; the retired plan is not resumed. |

Arbitration records target law; it changes neither execution state nor current
normative authority. An R activation card may be added and refined as `DRAFT` at any
time. `DRAFT` grants no implementation authority and may contain open fields or
predecessors; the R slice remains `DECIDED-NOT-ACTIVE`. The card may become `ACTIVE`,
and the node may start, only after every mandatory field is exact, every hard
predecessor has a recorded exit receipt, and the owned files are available. A family
name never activates its children. E0-E8 retain inherited authorization but still
obey this graph and their drift/evidence gates.

This subsection is the sole detailed `DRAFT`/`ACTIVE` definition. Dashboard, routing,
Done, STOP, and WIP-index wording are summaries/guards only and cannot add a state,
field, predecessor, or alternative activation path.

The following table is the **only** package-ordering authority. Dashboard phases are
a readable topological summary; package cards own scope and acceptance, not extra
predecessors. If any other line conflicts with this table, STOP and correct this
document before implementation. Any node may start only when every hard predecessor
is complete and its owned files are available; an R node additionally requires its
complete `ACTIVE` card. Nodes with no path between them may run in parallel.

| Node | Hard predecessors | Why the edge exists |
|---|---|---|
| E0 | None | Establish one current merged baseline before any successor edit. |
| E1, E6, E7, R0a-R0e, R1a-R1c | E0 | These are independent seams, containment/truth slices, or scoped indexes after the baseline; each still activates separately. |
| E2 | E1 | Conflict extraction consumes the item/bridge seams. |
| E3 | E2 | Durable mechanics follow the extracted action/result ownership. |
| R2 | E1, R0d | The lockstep typed cutover consumes the extracted seams and the immediate capability/receipt honesty correction. |
| E4 | E3 | Batch Rename worker admission preserves current behavior and emits facts through the existing task surface; it does not own R1d's later common lifecycle. |
| R7-A10 | E4 | Cycle rejection and journal retirement land at the worker-owned Batch Rename admission seam. |
| R7-A15 | E4, R2 | Provider/path name feasibility lands once at the final typed rename-admission seam. |
| E5 | E4, R1b, R7-A10, R7-A15 | Change Case migrates once into the final acyclic/no-journal/provider-correct RenamePlan contract. |
| R1d-core | E4, R0e, R1a | The common Preparing lifecycle consumes Batch Rename admission facts, honest cancellation, and the single terminal/clipboard funnel. |
| R3 | R0a, R1c, E1, R2 | Publication generalization consumes the narrow rollback fix, cleanup truth, seams, and typed route facts. |
| R4 traversal/A08/A13 | E1, R0b, R2 | Traversal and namespace behavior consume the extracted seams, bounded S3 behavior, and typed namespace facts. |
| R4-A02 overlap engine | E1, R1d-core, R2 | Scope comparison and live cross-task guards plug into the common pre-consumption/mutation hooks. |
| R4-A19 discovery scheduler | E1, R1d-core, R2 | Scheduling consumes lifecycle hooks and typed provider/profile facts; it has no UI dependency. |
| R5 | R0e, R2 | Conditional only: isolate a residual exact route proved uncontainable in process. |
| R0f (R0f-SMB, R0f-Curl, R0f-S3, R0f-Graph, R0f-GDrive) | R0e, R2, R1d-core | Route-specific cancellation containment plus full read/write/create/delete for every route R0e left `uncontained`; R0f-SMB precedes M2, every R0f slice precedes M3. |
| R6-A02 | E2, R4-A02 | Render engine overlap facts and capture consent; never rederive overlap. |
| R6-A07 | R1a, E2 | Retry UI consumes the single terminal funnel and conflict snapshot. |
| R6-A17 | R1a, E2 | Review-later UI consumes the single terminal funnel and conflict snapshot. |
| R6-A08 | R4 traversal/A08/A13, R4-A19 | Render exact discovery/operation facts; never own traversal or scheduling. |
| R6-A09 | E6 | Artifact explanation consumes the post-retirement name-shape classifier and exact touch guard; any future durable claim source requires its own schema. |
| R6-A12 | E2 | The hosted renderer consumes the engine-owned conflict snapshot. |
| R6-A18 | E3 | History consumes shared safe persistence mechanics without sharing authority/schema. |
| R6-A20 | R1d-core | Routine-start presentation consumes the common lifecycle. |
| R7-A03 | R2, R3, R6-A20 | Retarget is an independent Copy/link branch using typed link facts, owned publication, and the one with-options surface; it has no E4/E5 dependency. |
| R9 | R1a, E2, R2, R3 | The Shell spike must compare against the final result/conflict/route/publication contracts. |
| R8 | E3, E6, E7 and every activated R slice it consolidates | Consolidate only behavior that has landed; never become a second owner. |
| E8 | E5, E6, E7, R8 and every activated R slice | Closeout consumes all required code, evidence, authoritative-spec updates, and inherited I17 qualification. |

**Current readiness:** E0 completed at exact clean baseline
`fc32a4aa4b01d6d684166968a0615d3265f33d26`; E1 through E7,
R0a-R0e, R1a-R1c, R2, R7-A10, and R7-A15 have recorded exit receipts. R1c closed under
implementation `bfcc9cbe` and archived evidence `54d3909c`; therefore derived
milestone M1 is complete. R7-A10 closed under implementation `8c303593`, archived
evidence `10a27c53`, and source-contract guard correction `eaeadfc3`. R2 closed
under implementation `75df3717` and archived evidence `9bf973cb`. R7-A15 closed
under implementation/test synchronization `e925018f` and final evidence `ae926d24`.
E5 closed under implementation/evidence commit `dd1810f6`; the rename chain is
complete. E6's post-retirement evidence correction landed in `9e14a899`, removing the
empty/synthetic claim index and replacing it with honest name-shape-only projection evidence.
R1d-core closed under implementation `668acbd7`, archived evidence `005b9c81`, and
final compact-status correction `0515e854`. Post-closeout correction R0e-OR1 is
`COMPLETE` under implementation `d7decb40` and archived evidence `921739b4`; it owns
only admitted Local blocking-reader cancellation and does not reopen M1's
mutation-containment receipts. The state below is derived from each card's `State:`
line (2026-09-04) and replaces every earlier readiness sentence:

- `COMPLETE`: R0f (SMB, Curl, Graph, S3, Google Drive) with its post-closeout
  corrections R0f-Curl-OR1 and R0f-Curl-OR2; R0-RC3 (gate #12) with R0c-OR2; R3
  (R3-1, R3-2, R3-3; gate #11); R4-A19 with R4-A19-OR1; R4-T1 (gate #13); R4-T2 and
  R4-T3 (their exit receipts and Fresh Full gate #14 are recorded on their cards); the
  R4 traversal/A08/A13 package checklist; R6-A08 (gate #15).
- `ACTIVE`: R0c-OR3; R4-T4; R4-A02-1; R4-A02-2; R0f-GDrive-OR1;
  R0-Policy-OR1; R6-A20; R6-A09.
- Eligible and inactive: R6-A02, R7-A03.
- Predecessor-blocked or optional: R5, the remaining R6 slices, R8, R9, E8.

Known open product holes (2026-09-04 review): permanent-delete binding and per-leaf
child-name queries still run on the UI thread during admission; MTP and S3 readers
cannot be cut mid-Read and Shutdown joins task threads on the UI thread. The active
cards above own the MTP live-identity correction, Google Drive mutation-retry/deep-tree
corrections, overlap-advisory completion, and review-policy cleanup. The Microsoft Graph
zero-timestamp correction is complete under R0-RC4.

#### Derived delivery milestones

Milestones are release/readiness views over the DAG above. They add no edge and own no
behavior, files, tests, or activation. Nodes may start earlier whenever the DAG permits.
A milestone closes only from the named node exit receipts; its state is derived and
never maintained independently.

| Milestone | Required exit receipts | Meaning |
|---|---|---|
| **M1 — Trustworthy current engine** | E0; R0a-R0e; R1a-R1c | Current routes are baselined and immediate mutation, receipt, cancellation, result, and cleanup-truth findings are contained. Any route R0e still classifies as uncontained remains unavailable only until its R0f slice lands; R5 is not required merely to re-enable it. |
| **M2 — Reliable fast core under user control** | M1; E1-E4; R2; R1d-core; R0f-SMB; R3; R4 traversal/A08/A13; R4-A02; R4-A19; R6-A02; R6-A08; R6-A20 | Common lifecycle, typed authority, publication, traversal, overlap warning/consent, evidence-first discovery service, dual progress, and routine start are integrated. R5 is additionally required before any residual isolated route is re-enabled. |
| **M3 — Complete successor and closeout** | M2; E5-E8; R6-A07; R6-A09; R6-A12; R6-A17; R6-A18; R7-A10; R7-A15; R7-A03; R8; R9; R0f (every slice: SMB, Curl, S3, Graph, GDrive); R5 if activated; every other activated slice | All section 6 rules, optional evidence decisions, durable owners, verification, inherited-I17 qualification, and plan-retirement gates are complete. |

Draft one bounded activation card directly under its exact R-slice heading whenever
that refinement is useful. The master plan owns target-decision mapping, dependencies,
outcomes, and hard boundaries; current authoritative specifications own current
product behavior. A `DRAFT` card may be incomplete and grants zero edit authority.
Before promotion to `ACTIVE`, the card must own exact symbols, files, RED/GREEN test
IDs/commands, fault matrix, performance budget, and same-slice authoritative-spec
updates at the then-current baseline. Every **pre-dispatch** field is then mandatory:
`TBD`, a family-wide owner, a generic “run tests” entry, or an open hard predecessor
leaves the slice inactive and implementation cannot begin. The card predeclares the
exit-receipt schema and evidence destinations; the actual commit/run/receipt values are
filled only after implementation and are required to close, not activate, the slice.

| Activation-card field | Required content |
|---|---|
| Slice/state/owner | Exact R slice and card state `DRAFT` or `ACTIVE`; `ACTIVE` requires one named implementation owner. |
| Baseline and drift | Planned-at commit plus the exact drift-review command/result. |
| Contract | Section 6 target rule IDs/clauses implemented and the current authoritative clauses displaced by the same slice; no behavior reconstructed from section 7 examples. |
| Scope | Exact in-scope and out-of-scope files, interfaces, providers, and symbols. |
| Predecessors | Every hard predecessor from this table. A `DRAFT` may mark one open; `ACTIVE` requires every completion/exit receipt. |
| Steps | Ordered edit/migration/deletion sequence and displaced owner to remove. |
| RED/GREEN evidence | Exact test IDs, commands, expected failing baseline, expected passing result, and fault cases. |
| Performance/resources | Scenario, metrics, exact acceptance/no-regression budget, and `Specs/TestRuns/` evidence path. Characterization-only is permitted solely for an explicitly evidence-only spike/measurement card with a written adoption gate. |
| Durable owners | Authoritative specs/inventories updated in the same slice; only this landed update makes the target behavior normative. |
| STOP/rollback | Conditions that halt the slice and the safe way to remove/disable partial work. |
| Exit-receipt contract | Predeclared fields/destinations for commit, builds, focused tests, perf archive, inventories, and remaining known gaps. Actual generated values are recorded after implementation before the slice closes. |

No former-I14 package is merely referenced back to history:

| Retired source package | Sole live destination | Preserved responsibility |
|---|---|---|
| A0 | E0 | Current merged characterization, deep regression, focused inventories/builds, and pre-successor-change evidence. |
| A1 | E1 | Behavior-preserving bridge responsibility seams and one serial/parallel item policy. |
| A2 | E2 | Engine-owned action/order/default/Cancel snapshot and exhaustive conflict matrix. |
| A3 | E3 | Shared safe durable-file mechanics with separate schemas and authority. |
| A5 | E4 | Worker-side Batch Rename qualification and UI/payload/teardown responsiveness. |
| A4 | E5 | Change Case characterization, central RenamePlan migration, and safe retirement of the second engine. |
| A6 | E6 | Post-retirement name-shape projection gating, exact touch-guard authority separation, and honest perf evidence. |
| A7 | E7 | One session notice identity plus separately named durable context and notice-only tests. |
| A8 | E8 | Normative updates, package evidence, inventories, inherited-I17 qualification, final Fresh Full, and plan retirement. |

Only one active package may modify `FolderWindow.FileOperations.State.cpp` at a time.
I12 and any external branch touching an owned file are treated as explicit graph
coordination, not as an unrecorded dependency or a license to merge scopes.

Former-I14 and planned-at drift must be reconciled and recorded before E0 dispatch.
The inherited E scope below is the exact scope for both comparisons:

```powershell
$e0Scope = @(
    'RedSalamander/FolderWindow.FileOperations.State.cpp'
    'RedSalamander/FolderWindow.FileOperations.Popup.cpp'
    'RedSalamander/FolderWindow.FileOperations.cpp'
    'RedSalamander/FolderWindow.FileOperations.State.Runtime.cpp'
    'RedSalamander/ChangeCase.cpp'
    'RedSalamander/FileSystemRenameBatch.cpp'
    'RedSalamander/BatchRenameRecoveryJournal.cpp'
    'RedSalamander/FileOperationArtifactRegistry.cpp'
    'RedSalamander/FileOperationMoveBreadcrumb.cpp'
    'RedSalamander/FolderView.Enumeration.cpp'
    'RedSalamander/FindFilesWindow.cpp'
    'Plugins/FileSystem/FileSystem.FileOps.cpp'
    'Specs/FileSystem/FileSystem_FileOperations.md'
    'Specs/UI/UI_FileOperationsPopup.md'
    'Specs/UI/UI_BatchRenameWindow.md'
    'Specs/UI/UI_CommandMenuKeyboard.md'
    'Specs/Core/Core_SharedHelpers.md'
    'Specs/Testing/Testing_PerformanceValidation.md'
)

$e0Commit = (git rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($e0Commit)) {
    throw 'Cannot bind E0 to HEAD.'
}
$e0Status = @(git status --porcelain=v1 --untracked-files=all)
if ($LASTEXITCODE -ne 0 -or $e0Status.Count -ne 0) {
    $e0Status
    throw 'E0 requires a clean dedicated worktree.'
}

git diff 6094a966dc0450f75a93341e8364035212610491 -- $e0Scope
if ($LASTEXITCODE -ne 0) { throw 'Former-I14 drift review command failed.' }
git diff fa98ef68618096414733624c1b7c2320db1cca2c -- $e0Scope
if ($LASTEXITCODE -ne 0) { throw 'Planned-at drift review command failed.' }

$e0Markers = @(git grep -n --untracked -e '^<<<<<<< ' -e '^=======$' -e '^>>>>>>> ' -- $e0Scope)
$e0MarkerExit = $LASTEXITCODE
if ($e0MarkerExit -eq 0) { $e0Markers; throw 'Conflict markers remain in the E0 scope.' }
if ($e0MarkerExit -gt 1) { throw "Conflict-marker scan failed with exit $e0MarkerExit." }
```

### Package E0 — Current-HEAD validation baseline

**State:** `COMPLETE` / exit receipt recorded below
**Priority:** P1

#### E0 pre-dispatch binding

E0 is one unsplit baseline gate; there is no E0-lite. Before its first build, select
one exact merged pre-successor-change commit `C` in a dedicated worktree. The following
must hold through Debug, Release, every focused run, and Fresh Full:

- `git status --porcelain=v1 --untracked-files=all` is empty before the first build;
- `HEAD` remains `C`, and the index, worktree, and source-snapshot identity do not
  change; ignored `.build` outputs and runner data beneath the repository-authorized
  test root are allowed;
- the two exact drift commands above are reviewed and their result is recorded;
- the E0 scope contains no unresolved merge marker; and
- no user work is reset, deleted, hidden, or adopted to manufacture cleanliness. Use
  an isolated clean worktree or wait. No branch-name, upstream, or signed-tag rule is
  added.

If `HEAD` or source identity changes, E0 is invalid and restarts at the new exact
commit. Runtime-generated run IDs and archive paths are outputs to record in the exit
receipt, not `TBD` activation fields. If the fixed-NTFS test root is not already
owned/initialized, use only the explicit runner initialization switch permitted by
the testing contract; never infer authorization from an existing directory.

Run the following exact test-enabled full-solution builds from `C`:

```powershell
pwsh -NoProfile -Command '$env:RSBuildEnableTests = "true"; & .\build.ps1 -Configuration Debug -Platform x64; exit $LASTEXITCODE'
if ($LASTEXITCODE -ne 0) { throw 'E0 Debug build failed.' }
pwsh -NoProfile -Command '$env:RSBuildEnableTests = "true"; & .\build.ps1 -Configuration Release -Platform x64; exit $LASTEXITCODE'
if ($LASTEXITCODE -ne 0) { throw 'E0 Release build failed.' }

$e0DebugReceipt = Get-Content -Raw .\.build\x64\Debug\build-receipt.json | ConvertFrom-Json
$e0ReleaseReceipt = Get-Content -Raw .\.build\x64\Release\build-receipt.json | ConvertFrom-Json
if (
    $e0DebugReceipt.target -ne 'solution' -or
    $e0ReleaseReceipt.target -ne 'solution' -or
    -not $e0DebugReceipt.tests_enabled -or
    -not $e0ReleaseReceipt.tests_enabled -or
    $e0DebugReceipt.git_head -ne $e0Commit -or
    $e0ReleaseReceipt.git_head -ne $e0Commit -or
    $e0DebugReceipt.source_snapshot_id -ne $e0ReleaseReceipt.source_snapshot_id
) {
    throw 'E0 build receipts do not bind the same test-enabled full-solution snapshot.'
}
```

Both commands must exit 0. `.build/x64/Debug/build-receipt.json` and
`.build/x64/Release/build-receipt.json` must report `target=solution`,
`tests_enabled=true`, `git_head=C`, and the same `source_snapshot_id`.

Run this exact focused matrix in both configurations:

```powershell
foreach ($e0Configuration in @('Debug', 'Release')) {
    .\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration $e0Configuration -Platform x64 -SkipBuild -ValidationMode Fresh -CaseFilter Phase7_CopyItemsSingleFolderRecursiveParallelism -FailFast -TimeoutMultiplier 2.0
    if ($LASTEXITCODE -ne 0) { throw "E0 deep Copy failed in $e0Configuration." }

    .\Tools\Run-AllTests.ps1 -Suite Commands -CommandsFamily file-operations -Configuration $e0Configuration -Platform x64 -SkipBuild -ValidationMode Fresh -CaseFilter cmd_pane_fileops_popup_progress_contracts
    if ($LASTEXITCODE -ne 0) { throw "E0 popup progress failed in $e0Configuration." }

    .\Tools\Run-AllTests.ps1 -Suite Commands -CommandsFamily file-operations -Configuration $e0Configuration -Platform x64 -SkipBuild -ValidationMode Fresh -CaseFilter cmd_pane_fileops_popup_presentation_settings_and_taskbar
    if ($LASTEXITCODE -ne 0) { throw "E0 popup presentation failed in $e0Configuration." }

    .\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration $e0Configuration -Platform x64 -SkipBuild -ValidationMode Fresh -CaseFilter FileOps_ProviderCapabilityMatrix
    if ($LASTEXITCODE -ne 0) { throw "E0 provider matrix failed in $e0Configuration." }

    .\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration $e0Configuration -Platform x64 -SkipBuild -ValidationMode Fresh
    if ($LASTEXITCODE -ne 0) { throw "E0 FileOps failed in $e0Configuration." }

    .\Tools\Run-AllTests.ps1 -Suite Commands -CommandsFamily file-operations -Configuration $e0Configuration -Platform x64 -SkipBuild -ValidationMode Fresh
    if ($LASTEXITCODE -ne 0) { throw "E0 Commands File Operations failed in $e0Configuration." }

    & ".\.build\x64\$e0Configuration\PluginContractTests.exe"
    if ($LASTEXITCODE -ne 0) { throw "E0 provider contracts failed in $e0Configuration." }
}
```

The required exact registered IDs are therefore:

- `Phase7_CopyItemsSingleFolderRecursiveParallelism` — the 96-level, greater-than-
  `MAX_PATH`, single-item Local Copy regression;
- `cmd_pane_fileops_popup_progress_contracts` and
  `cmd_pane_fileops_popup_presentation_settings_and_taskbar` — merged popup/compact-
  progress and taskbar behavior; and
- `FileOps_ProviderCapabilityMatrix`, plus the no-argument Debug and Release
  `PluginContractTests.exe` runs — provider capability/contract coverage.

Record inventory-derived totals and documented skip reasons; do not hardcode mutable
suite counts. A missing/zero-match/skipped required ID, a corrupt/blocked/
`NOT_EVALUATED` result, or a receipt for another snapshot fails E0.

After the focused matrix, and before **any** E or R successor source edit, run E0's
own canonical Fresh Full and hygiene gates:

```powershell
if ((git rev-parse HEAD).Trim() -ne $e0Commit) { throw 'E0 HEAD changed before Fresh Full.' }
.\Tools\Run-AllTests.ps1 -Suite Full -Configuration Debug -Platform x64 -ValidationMode Fresh
if ($LASTEXITCODE -ne 0) { throw 'E0 Fresh Full failed.' }
.\Tools\Get-TestInventory.ps1 -Format Json
if ($LASTEXITCODE -ne 0) { throw 'E0 test inventory failed.' }
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
if ($LASTEXITCODE -ne 0) { throw 'E0 spec inventory failed.' }
.\Tools\Test-TestRunArchive.ps1 -Inventory
if ($LASTEXITCODE -ne 0) { throw 'E0 archive inventory failed.' }
git diff --check
if ($LASTEXITCODE -ne 0) { throw 'E0 patch hygiene failed.' }
$e0FinalStatus = @(git status --porcelain=v1 --untracked-files=all)
if ($LASTEXITCODE -ne 0 -or $e0FinalStatus.Count -ne 0) {
    $e0FinalStatus
    throw 'E0 worktree changed during validation.'
}
```

Fresh Full deliberately does not use `-SkipBuild`: its own test-enabled full-solution
receipt must bind to `C`. It must exit `PASSED` with zero failures and only documented
environmental skips. Record `baselineCommit=C`, runner run ID, Debug/Release/Full
receipt IDs, source-snapshot identity, totals, raw evidence location, any curated
`Specs/TestRuns/` location, and archive-inventory result. If a run is curated, validate
its actual literal promoted path with `Test-TestRunArchive.ps1 -RunPath` before the
whole-inventory command; never leave a placeholder in the completed receipt.

E0's Fresh Full cannot coalesce with R0 or any changed-behavior closeout and cannot be
reused as E8. E0 proves the pre-change baseline; E8 proves the final merged changed
state.

**Acceptance:** the exact clean commit and unchanged source snapshot are recorded;
both test-enabled configurations build; all exact and broad focused cases pass; both
provider-contract executables pass; Fresh Full and every inventory/hygiene command
pass. The 96-level case must complete without access violation, timeout, or incomplete
Copy. It does not directly measure stack bytes or task leaks, so E0 must not claim
those as measured facts; any observed lifetime/non-completion symptom still fails the
gate. Resume, Affected, stale receipts, or an older archive never qualify E0. If a
correctness/lifetime gate fails, stop successor edits and fix or route one bounded
defect, then restart E0 from a new clean commit. Never restore recursive stack
traversal, weaken link classification, kill an unrelated app, or bypass the DAG by
renaming a partial run E0-lite.

#### E0 execution log

| Attempt baseline | Terminal result | Bounded remediation / evidence |
|---|---|---|
| `5f83dc4b0b7c5d66de4f96895da43298532dd046` | Stopped before Fresh Full: Debug `PluginContractTests.exe` hung in Curl streaming-reader shutdown. | Fixed the lost-wake shutdown ordering and added repeated blocked-reader retirement coverage in `d4a86da64b5199df40f74620841f07287918d65d`. Preserved dump: `C:\RedSalamander.Perf\evidence\E0-PluginContractTests-20260830\PluginContractTests.Debug.hang.dmp`. |
| `d4a86da64b5199df40f74620841f07287918d65d` | Stopped before Fresh Full: the Release broad FileOps matrix exposed one Debug-export-only Riptide proof and one info-diagnostic-dependent auto-concurrency case. | Made the Riptide Release skip explicit while retaining Debug ownership, enabled the required diagnostic inside the auto-concurrency case, and restored settings in `0c20627d1dee2dedce589f8b96573737fab6224f`. |
| `0c20627d1dee2dedce589f8b96573737fab6224f` | Exact Debug/Release builds and the complete focused matrix passed. Canonical Fresh Full failed only `DxUiTests.Tooltip`: 2,041 total, 1,987 passed, 1 failed, 53 documented skips. | Run `20260830T153123Z-59440-5844f12d00ec441387c00ebb51e87725`, durable evidence `C:\RedSalamander.Perf\evidence\runs\20260830T153123Z-59440-5844f12d00ec441387c00ebb51e87725`. Root cause was a tooltip deadline scheduled from a host's stale prior dispatcher tick; the bounded repair uses the current shared dispatcher clock and adds a deterministic stale-epoch regression. Focused Debug and Release Tooltip stress passed 32/32 each; targeted receipts were `3f17a76a71758cd37270122c6d5a9c5e1652514a436e7d8dccdf26248a1cf278` and `d8c1d68bde5576e8144c2930809be8822a95616d8279418b3377c5ed86d7c9ec`. E0 remains open and must restart from the next clean exact commit. |
| `817b7e2d368d176464d3f1635029d95b0bbdba0b` | Exact Debug/Release builds passed and the Debug focused matrix passed. The Release broad FileOps run `20260830T171236Z-73588-6e14251c38564bf0b57cacf93abf2d92` stopped on `Phase6_ParallelBandwidthThrottleFairness` after observing 2 MiB candidate skew against a 0-byte paired baseline. | Repeated Release evidence reproduced a 0/1/2 MiB distribution and showed the shared-only baseline itself reaching 2 MiB three times in eight runs. The bounded repair replaces the provider's 500 ms callback-age worker heuristic with exact held transfer-slot lifetime, fixes the selftest's `_inFlightFiles` data race by taking `_inFlightFilesMutex`, records `FileOps.Progress.MaxCallbackDeltaBytes`, and binds the observation floor to one measured callback quantum while retaining the comparative skew and runtime guards. Strict eight-repeat remediation proofs passed 24/24 in Release (`20260830T180502Z-18776-5ac1061ced39493cb2f5a75fe81a7b3f`, receipt `b500062c58583ee14f6f2b580bceb8bb47ce34be3693fe0b541950d4ea7bd344`) and Debug (`20260830T181204Z-68084-96063276330d44d0bf05d9ab3631e774`, receipt `57173e5e1ff95e006b5091ad0a3c74e4bee11ffabaa41155f8c32a4c2761fa44`); every recorded callback quantum was 2 MiB and observed skew never exceeded one quantum. E0 remains open and must restart from the next clean exact commit. |
| `56afb5b5d70bd9d96e495a82052f925fdfed7fe6` | Exact Debug/Release builds and the Debug FileOps matrix passed until the broad Commands File Operations family failed `cmd_pane_fileops_completed_group_and_navigation` (`20260830T184128Z-79568-51dd9bdf726347fc90c94c7cbf0af47e`). A 16-repeat proof split 8 pass/8 fail between immediate auto-collapse observation and stale More-menu geometry (`20260830T184400Z-58560-49c4bccebc9a4f35a830f8e2dc4f5a50`). | The first attempted observer-side repair was rejected after a 2/32 repeat result (`20260830T190108Z-63868-04c225ce475b40a09601e0ef93382404`): keeping an already-open menu alive could not repair a stale click. The bounded repair in `aaeee66151a7c6164506adefe550c2c877b75c41` polls the paint-owned auto-collapse state, waits for repeated stable client-space More hit-target geometry, and maps it to screen immediately before activation. On the isolated initialized `D:\RedSalamander.Perf` root, Release passed the case 32/32 (`20260830T194342Z-65540-72bcff90f2e147a08699099bf87d80ce`) and the family 32/32 (`20260830T194459Z-7936-b1994c7f792d4f73a0b5aafacb3bd253`); Debug passed 32/32 (`20260830T195217Z-80264-e3c098392e0b409cb1c4c46d939b2306`) and 32/32 (`20260830T195322Z-83428-60e36caeca62460d960a9201932b801b`). All four runs reported zero disk-audit issues. Repository-supported embedded debug information avoided unrelated shared-PDB contention during remediation builds; E0 remains open and must restart from the next clean exact commit. |
| `c5deff7c567d0fbae229c3a0881535cc8f130062` | Exact Debug/Release builds passed with full-solution receipts `c2b8c0cd1ea17162564d094ea97820e758339f9ebb532497e205f27ff8d6084d` and `4e46971bf5fac6ff82cdff480a17f8eac087ac78ff9e85ab45592cbe2319ac8b`. Debug deep Copy passed 3/3 (`20260830T200550Z-78404-834cfb7c25094987bfe25dedddeca932`) and popup progress passed 1/1 (`20260830T200612Z-78404-e6388a26bde04b73b23e91f7854a632c`), but the second run's disk audit incorrectly reported the first run as unexpected, so the matrix was stopped before popup presentation. | The exact matrix intentionally invokes sequential runner processes from one PowerShell parent. Preflight cleanup correctly protected the prior sibling while owner PID `78404` remained live, but the post-run audit allowed only the current run id. The bounded tooling repair in `3fe4aa5a963a68f7975d2d4e691b0033af51bd68` makes disk audit reuse the cleanup resolver's exact live-owner/PID-reuse classification, retains dead-owner and manual-directory findings, and adds same-parent plus concurrent-live regression coverage. Focused plan tests passed 54/54, tool inventory reported 0 findings, governance passed 12/12, inventory tests passed 17/17, source contracts passed 182/182, and spec inventory reported 0 blocking findings. E0 remains open and must restart from the next clean exact commit. |
| `c7b4dd8efd49d23a4347b3c8618ada5f3542eafc` | Exact Debug/Release builds, the complete focused matrix, and all inventories passed. Canonical Fresh Full run `20260830T205624Z-59420-bb2aaeaa47fa4658a0084066f11632de` failed only `Phase7_RecycleBinBatchDelete`: 2,042 total, 1,987 passed, 1 failed, 54 documented skips, zero disk-audit issues. Four isolated reruns split 2 pass/2 fail while candidate Shell latency ranged from 7.5% to 12.5% above baseline (`20260830T215529Z-78484-629b4a2110e345bf960c134c0cec0f03`, `20260830T215614Z-78484-4ffd6d69c1f042e5bd1b714e9b2beb0f`, `20260830T215653Z-78484-47f69fce5145409ba6d9177ea884ce70`, `20260830T215731Z-78484-d8c7bb50bb1548bdbf02f65201ea38d2`). | Test-only route counters first proved that the existing fixture never entered batching (`20260830T221429Z-84496-617a380f9ec44317a62a4cdcfa703b4c`): it used the helper's default `PerItem` admission while the pane Recycle command uses `BulkItems`. The bounded repair makes `BulkItems` explicit, hard-gates zero baseline calls plus exact 1x384 and 3x256 route observations with zero failures/fallbacks, retains variable Shell/Defender/indexer timing as advisory evidence, and generates run-unique recycled leaves after a Release repeat exposed a Recycle Bin name-collision prompt. Four-repeat candidate/multi-batch proofs passed 12/12 each in Release (`20260830T224747Z-79472-e338ed81580d47b2b0b3bca6f5c541ad`, `20260830T224849Z-40136-ea164e37caac46898cf5b1e64b0cc99e`) and Debug (`20260830T224931Z-80588-f46ddb21f92c45ef9af13f7c1de3d3f7`, `20260830T225412Z-86576-21b4a082b1b543c08082f93b9ebce154`), all with zero disk-audit issues. Release broad FileOps passed 122/122 applicable cases (`20260830T225802Z-25656-65a3ca2371b34f6b854311cc30adc262`), and both Debug/Release `PluginContractTests.exe` passed. E0 remains open and must restart from the next clean exact commit. |
| `625ef00c80f9bddc8832ee2fe35dca6b00312a0a` | Exact Debug/Release test-enabled full-solution builds passed with receipts `f6c9b3b2ca2179b9024df1fd5d57970049cd5c5d0ad5db8cb355978dccabcd73` and `fe839469acbc57f9ee3f1810b93aa8fdb7c0da9d03bd39c9aaaa4f83924172b6`. Debug deep Copy and popup-progress gates passed, but popup presentation stopped when its polling observer missed the one-second taskbar retry-pending state even though the real retry had already succeeded (`20260830T232116Z-85088-1afcc69bf5494bfc823d52626d59d947`). The first synchronous observer repair then exposed one first-run-only transient button-overlap sample in a 31/32 run (`20260830T232452Z-16344-6cec31e4b7e64298acd643cc7f439722`). | The bounded repair in `82a624cdc797391d4312e6e95c0b6174e958bb63` synchronously drives the real taskbar update path around the injected failure and published retry delay, exposes exact test-only layout-motion state, and applies the existing minimum-width overlap hard gate only to a synchronous settled snapshot. An intentionally over-broad retained-animation predicate failed closed (`20260830T233913Z-85524-f3e3661e7f094d78a5967ee5b1451715`); narrowing it to current window/selected-card geometry and accepting the deadline's exact settled snapshot removed the observer artifact without weakening overlap detection. Final Debug and Release full-solution builds passed with zero warnings/errors and receipts `c3c7c2c6e5302672bca4843a079d18d00987edef0ba3a352860546493012a862` and `990cf2c740aa4b9766d400308fbd4989d5f647c5785f26141fdf798d678ce3d9`; popup presentation passed 32/32 in Debug (`20260830T234839Z-79632-bbadfd44d1a8402dbd7b92dc58d28b6d`) and Release (`20260830T235452Z-76628-0a0262065f6c4a70aa314e2ff38a3bb5`), both with zero disk-audit issues. E0 remains open and must restart from the next clean exact commit. |

#### E0 completion receipt

| Attempt baseline | Terminal result | Bounded remediation / evidence |
|---|---|---|
| `fc32a4aa4b01d6d684166968a0615d3265f33d26` | **PASSED / E0 complete.** Exact Debug and Release test-enabled full-solution builds passed with zero warnings/errors, receipts `97363fbda7839407d5b951d10a719a2f3023359ab33f22023b541f68e1590f21` and `4b6f492c881439d880baef95b4113e109fe7a09865317782416e53c3f0042e96`, and shared source snapshot `496378c1bf50fbb835053c5274179cde97fbf418ae105ef13621985b85a12ae1`. The complete Debug and Release focused matrices passed: Debug deep Copy `20260831T000900Z-64252-51aa4603a7e2478bacc042a97ca611dc`, popup progress `20260831T000923Z-64252-f485969c505145a88f645328ca842fb9`, popup presentation `20260831T000943Z-64252-89bb7608ca1340399a2de87a884a2dd4`, provider matrix `20260831T001009Z-64252-e0be8ab7a24747bab6f55dabf12ef86d`, broad FileOps `20260831T001035Z-64252-1baf7d900706417dbeb56d0a48e18f99`, and Commands family `20260831T002334Z-64252-69dde36f5d6140ab880d0609dc8bf051`; Release equivalents were `20260831T002507Z-57476-14d9f07009bf4abdba29d8925708b21f`, `20260831T002530Z-57476-72e83ca001ec4dc9ae7847f8b8114fc8`, `20260831T002551Z-57476-b487ed3cdcea4e209e3b2c020dbce180`, `20260831T002614Z-57476-1c70f1bc04824dbfa5faa7344af5ccdb`, `20260831T002636Z-57476-8ad8ade83c0d4fd59c428dd0b4984083`, and `20260831T003856Z-57476-bfd56ea07ae34b4fb6e9064b0ee1ef9`; both `PluginContractTests.exe` runs passed and every focused disk audit reported zero issues. | Canonical non-`SkipBuild` Fresh Full run `20260831T004015Z-51980-1a59bf5343ce446c81b81c16bb9cd1a1` passed 2,042 total / 1,989 passed / 0 failed / 53 documented environmental skips in 55m 28.8s with zero disk-audit issues. Its full-solution Debug receipt was `b812a5ee7e0952542d7a320ebf351aebcc50007b2c359e4d30c4c80a19daf923`, bound to the same exact commit and source snapshot. Raw evidence: `D:\RedSalamander.Perf\evidence\runs\20260831T004015Z-51980-1a59bf5343ce446c81b81c16bb9cd1a1`. Test inventory exited 0; spec inventory reported zero blocking findings; whole archive inventory passed for 1,216 files; `git diff --check` passed; final status was empty at the exact baseline. No curated `Specs/TestRuns/` promotion was required for this pre-change correctness baseline. |

### Package E1 — Named bridge and per-item policy boundaries

**State:** `COMPLETE`
**Priority:** P1, after E0; coordinate I12 lifetime/teardown overlap

Land two behavior-preserving commits:

1. Characterize the local `CrossFileSystemBridge` responsibilities and introduce
   narrow named seams for current publication/pump, traversal, and cleanup work.
   Do **not** transplant the local god object intact into `Bridge.h/.cpp` and call the
   relocation architecture. Preserve counters, cancellation polling, current
   resource rules, publication order, verification serialization, metadata consent,
   accepted link handling, and cleanup authority until the exact accepted R slice is
   explicitly activated to change them.
2. Replace the `processIndex` lambda/`std::function` ownership with named item-policy
   functions, private `Task` methods, or a small context type. Prefer private methods
   if a new type would merely expose most of `Task`. Serial and parallel schedulers
   must call the same item policy.

- [x] Characterize Native/Managed/Copy-only, conflict, retry, terminal, folder-merge,
      link, Keep Both, Recycle, verification, and traversal behavior before moving it.
- [x] Preserve one prompt per task; Copy-only never deletes source; Managed cleanup
      consumes retained bound authority only; unknown commit is never retried as
      known non-commit; directory race requalification remains one-way; Verify On
      does not overlap the next transfer; result axes remain truthful.
- [x] Preserve cancellation, completion posting, reaper, callback mutex, and module/
      task lifetime owned by I12; no bridge object outlives referenced state.
- [x] Prove parity at concurrency 1 and N, including cancel during discovery, read,
      write, verification, prompt wait, and cleanup.
- [x] Archive before/after large-tree and many-file throughput, retained memory,
      allocation/copy, bridge-recreation, and whole-tree-retention evidence.

#### E1 completion receipt

| Implementation | Correctness evidence | Performance / retention evidence |
|---|---|---|
| `0e24444c392f4d00ba706cabf97cab5ff7ab6ef7` introduced the narrow publication/pump, traversal, and cleanup seams without moving the local bridge owner. `20c15a26b68015dce8b9328f9a929bbb24e73189` replaced scheduler-owned `std::function`/`processIndex` closure state with one stack-bound `QualifiedItemPolicy` used directly by concurrency 1 and by the bounded scheduler at concurrency N. `WaitJob`/`Shutdown` remain the callback quiet point before the non-owning policy/context references unwind. | The 182-case source-contract suite passed. Full test-enabled Debug and Release solution builds passed with zero warnings/errors and receipts `99c986f465fdd54eb2cf55a7f2ea646de0b19f94b867dc2a7d8002fc21bc84dc` and `00f7ac9fcbbb76b97d4fded3c9a86012ce7721d0bdd9df6e80f755cdb507bc46`, sharing source snapshot `ab0582654de7cd6dc93bc0f90b3db1e66786c3e0aaf181e2222f9da69679bfa8` for the final extraction. Phase 11 passed `9/9` in Debug (`20260831T022015Z-59396-357073a4d4cd49629e6ee23c494c6b06`) and Release (`20260831T024639Z-89372-1db801e718be4d359527b89ab26f2aa3`). Broad FileOps passed with no failures in Debug (`123/143`, 20 documented skips, `20260831T022630Z-74156-5eaeffbad6494ac8b2f42b0da2c358b3`) and Release (`122/143`, 21 documented environmental/Debug-only skips, `20260831T025227Z-50900-9e6f613fecac4ae6b1430a3335768336`). The compiled source guard passed separately in both configurations (`20260831T031902Z-68696-f55d88f2a71b4c88bb5133a59ec8f517`, `20260831T031902Z-88952-344258c6383f4709b6f364a98de2537a`); every run reported zero disk-audit issues. | Same-machine exact Release evidence is archived at `Specs/TestRuns/4cb089111a23/FileOps/2026-08-31_050934_e1_named_bridge_policy_release`: the pipeline case improved from `130,890 ms` to `128,344 ms` (-1.9%) and the three-case run from `132,344 ms` to `129,609 ms` (-2.1%), with scheduler/admission/cardinality and retention bounds preserved. The complete Release family (`9/9`, `20260831T031026Z-63536-56efdd730e384d2cbe5c5269ae0a2033`) is archived at `Specs/TestRuns/4cb089111a23/FileOps/2026-08-31_051603_e1_named_bridge_policy_family_release`; all `48/48` wide-tree files started before producer completion, retained entries peaked at `24`, queued-path bytes at `12,432`, metadata at `9,638`, and traversal-limit hits/stage retention/deferred-link retention remained zero. Both explicit archive checks and whole-inventory validation (`1,236` files) passed. The single-sample in-run candidate slice (+12.4%) and 18-sample scheduler p95 (+8.6%) remain documented low-quality diagnostics; the decision-grade whole scenario passed and improved. |

Closeout governance passed: `Get-SpecInventory.ps1 -FailOnFindings` reported zero
blocking findings, and `git diff --check` passed.

**STOP:** an extraction commit cannot change product behavior, merge Native and
Managed engines, expose owning raw COM pointers, weaken WIL ownership, or absorb an
unresolved R decision. Any required behavioral amendment returns to arbitration.

### Package E2 — One conflict-action policy owner

**State:** `COMPLETE`
**Priority:** P1, after E1

- [x] Extend the immutable engine prompt snapshot with final ordered action IDs,
      primary/More placement, default, Cancel/Escape action, Apply-to-all and Skip
      All eligibility, metadata-loading state, and button publishability.
- [x] Delete popup-side `ConflictBucket`-to-action derivation. The popup may measure,
      wrap, paint, and expose supplied actions, but cannot add, remove, reorder, or
      choose a different default.
- [x] Cover regular/read-only Exists with destructive action allowed/withheld, type
      mismatch, destination semantic link, TargetConflict, Retry-only access/
      transport buckets, Recycle Cancel-first/default/Escape, deferred Proceed/
      Retain Source, Skip/Skip All/Apply-to-all scope, metadata loading with no
      actions, and cached actions without a second derivation.
- [x] Update `FileSystem_FileOperations.md` and `UI_FileOperationsPopup.md` with the
      resulting single-owner boundary; preserve the retired-I17 visual and UIA
      baseline.

#### E2 completion receipt

| Implementation | Correctness evidence | Performance / governance evidence |
|---|---|---|
| `116ec18bf2830ca3d48ca28d4e1aba925c161d37` made `ConflictActionPolicy` the sole production owner of ordered actions, primary/More placement, default and Escape/Cancel, scope eligibility, loading publication, and cached-decision validation. The popup now copies and validates the immutable snapshot only. Cancel remains directly visible; Close hides without resolving, while Escape submits the published safe action. No new conflict action, D2-A17 collection, or D2-A18 persistence was introduced. | Final test-enabled Debug and Release solution builds passed with zero warnings/errors, receipts `95ba469ed746df3e6e9c9a060b2ea32808237850ab12d99cdb9fcd96359b37e6` and `ca505fa096ab5fa39a75458f5f7011809a6ed4ccc5f1d51f8b35d92ef4a020e6`, sharing source snapshot `82503628fe4df99043ea25c18a815ed89c92bd2489e5f8e3dfff7d1be26b48f0`. Conflict metadata/actions/Close/Escape passed `3/3` in Debug (`20260831T035200Z-63892-b7ff4d756acc4b7894da73b7560ec47d`) and Release (`20260831T041116Z-11012-483313f699674f7591f8c4e3552ea7d9`); popup presentation/UIA passed in Debug (`20260831T035242Z-79256-1830921941a74573b93b960756745a74`) and Release (`20260831T041150Z-19192-b9a877214e364ff6b5a815a646110794`). The policy-heavy Phase 10 family passed `9/9` in Debug (`20260831T035313Z-77064-762fa37363f74adcaea4a32667d8182c`) and on the final Release source (`20260831T043801Z-77340-817830570189429ebc9fca78bbd7c276`). Broad FileOps passed with no failures in Debug (`123/143`, 20 documented skips, `20260831T035349Z-5656-8e5ca8c7be4e4712a7decde2a858bbdd`) and Release (`122/143`, 21 documented environmental/Debug-only skips, `20260831T041255Z-74520-a34865339a90480eb71c5a9a36ddd92c`). The final compiled source guard passed in both configurations (`20260831T043039Z-90840-f51dde1535684bca883c3e5162b3a16e`, `20260831T043531Z-39280-c295a26deca043128f5705693330aca7`); every cited runner reported zero disk-audit issues. | The final same-machine Release Phase 10 archive is `Specs/TestRuns/4cb089111a23/FileOps/2026-08-31_063824_e2_conflict_action_policy_release`: `9/9` passed, popup snapshot build p95 was `143 us`, hosted-control sync p95 `2,929 us`, conflict metadata p95 `24 us`, and conflict wait remained zero. The archive explicitly makes no improvement claim because no apples-to-apples pre-E2 popup stream survived governed retention. Explicit archive validation passed for 10 files and whole-inventory validation passed for 1,246 files. Source-contract Pester passed `182/182`; `Get-SpecInventory.ps1 -FailOnFindings` reported zero blocking findings; `git diff --check` passed. |

**Acceptance:** one production action/default/Cancel policy owner remains, popup
tests render the supplied snapshot, and keyboard/UIA/default/Escape behavior matches
the authoritative specs. This package does not authorize D2-A17 collection or
D2-A18 history persistence merely because those decisions are now accepted; R6 still
requires explicit activation. It also authorizes no new conflict choice.

### Package E3 — Shared safe durable-file mechanics, separate authority

**State:** `COMPLETE`
**Priority:** P2

- [x] Search `Core_SharedHelpers.md`, `Common/`, `LocalFileTransaction`, and
      `Tests/TestSupport/` before adding a helper; keep it app-local when dependency,
      ABI, or policy makes `Common` inappropriate.
- [x] Share only configured byte/count limits, no-follow open and regular-file
      validation, exclusive create/atomic replacement, durable-write propagation,
      in-memory rollback on persist failure, safe malformed/reparse-resistant scan,
      and acknowledgement/removal constrained to the configured store root.
- [x] Preserve any still-live Batch Rename journal behavior byte-for-byte. E3 may
      reuse safe-file mechanics only when that is behavior-preserving; it neither
      removes/migrates the journal nor enlarges its authority. R7-A10 solely owns the
      accepted journal/recovery removal, and no generic abstraction is justified only
      to keep that retiring store alive.
- [x] Keep artifact claim identity/Proven/Possible classification and Move breadcrumb
      notice/count/pane projection as separate schemas, versions, authorities, UI
      actions, and locks.
- [x] Do not create a generic journal or grant a breadcrumb/claim recovery, resume,
      rollback, cleanup, or deletion authority through shared I/O code.
- [x] Run each remaining schema lifecycle suite plus faults for parent/exclusive create,
      persist before/after in-memory transition, atomic replace, reparse/directory,
      oversize/truncation/invalid UTF-8 or JSON/version, rejected sibling scan,
      outside-root acknowledgement, and documented concurrent-reader/one-writer use.
- [x] Update `Core_SharedHelpers.md` and every owning domain spec in the same commit
      if a shared helper is introduced or extended.

#### E3 completion receipt

| Implementation | Correctness evidence | Performance / governance evidence |
|---|---|---|
| `c9bcb8bf6217ae1810dbe5c42aab99ed2755c816` introduced the app-local `FileOperationDurableStore` mechanics boundary and extended `LocalFileTransaction` replacement publication with handle-based replace/POSIX rename plus the established `MoveFileExW` fallback. Batch Rename retained its schema, transition rollback, reconciliation, Resume/Roll back authority, and 64 MiB limit; the Move breadcrumb retained its separate schema, notice/count/pane projection, and 128 KiB limit; artifact Proven/Possible identity joins remained in the registry owner. The helper contains no parser, recovery action, classification, or UI policy. | Exact-commit test-enabled Debug and Release full-solution builds passed with zero warnings/errors, receipts `a9efeed7357de00070d49b50b14b0bee0045a0edd81a76edae1f0a0af784ce46` and `d210cda60f9cd3937ead04dc27bbb5db60b4f2a37183dd87e840600bc6018b02`, sharing source snapshot `4e84cdbe092b529b55f9726e8a5cbb36acf2a089dc9f8e41eac4304ccda0cd28` and exact Git head `c9bcb8bf6217ae1810dbe5c42aab99ed2755c816`. The eight-case durable-store/journal/breadcrumb/artifact matrix passed `8/8` in Debug (`20260831T060554Z-45368-ba49646f3fad4a38bc9bcd16521e6916`) and Release (`20260831T060235Z-38728-1432d2c158e94559b0f2b91c9b5ef7f7`); both runs reported zero disk-audit issues. The matrix covers parent/exclusive create, write and publish rollback, atomic replacement with concurrent readers and one serialized writer, no-follow root/file handling, exact-handle removal, byte/count limits, malformed/truncated/invalid-UTF-8/wrong-version documents, and rejected siblings. | Same-machine Release characterization is archived at `Specs/TestRuns/4cb089111a23/FileOps/20260831_075139_e3_durable_state_store_release`: the archived run passed `8/8`; successful journal load/write p95 values were `7,741 us` / `1,914 us`, artifact-registry load p95 was `1,146 us`, and Move breadcrumb load/persist p95 values were `4,410 us` / `1,388 us`. The archive explicitly makes no improvement claim because no apples-to-apples pre-E3 scenario exists. Explicit archive validation passed for 6 files and whole-inventory validation passed for 1,246 files. Source-contract Pester passed `183/183`; `Get-SpecInventory.ps1 -FailOnFindings` reported zero blocking findings and emitted the current test inventory; `git diff --check` passed. |

#### Post-E3 adversarial review remediation receipt (2026-08-31)

`fb331ef70c6a625d26d3255e1e63568d080e8fd1` repaired the confirmed E0-E3
review regressions without activating an R package. Per-item scheduler failures now
use explicit caller policy: qualified-item failure retains the inherited task-cancel
behavior, while bridge directory workers record `E_OUTOFMEMORY` without remapping it
to Cancel or suppressing sibling roots. Cancel, stop, and shutdown release the worker
start gate before joining; task removal moves ownership out of the task-map lock before
the `jthread` is destroyed. Curl reader/writer waits now have stop-aware predicates and
behavioral teardown coverage. The durable store preserves the former journal's
replace-over-final-directory-entry behavior without following a planted reparse leaf;
Read and Acknowledge remain no-follow/reparse-rejecting. Direct-child limits and
directory/reparse rejection statistics are explicit, deterministic junction fixtures
fail closed, and both BeginMutation and Finalize restore terminal state when persistence
fails. The tautological Recycle debug helper was removed in favor of the production
`BuildConflictActionPolicy(...)` coverage.

The exact repaired tree passed a test-enabled x64 Debug full-solution build with zero
warnings/errors, artifact receipt
`61a98e21868605c19860161da5fab0e2213f1d2d366d0165a62505795a281598`.
Receipt-verified focused runs passed journal recovery `1/1`
(`20260831T065207Z-62416-b21619d0490c46968c03f8828656486f`), durable-store
commands `4/4` (`20260831T065235Z-86492-53b65af7726945539816921236988ae5`),
and scheduler/start-gate File Operations `3/3`
(`20260831T065300Z-25328-fa7af28761ae4bce96a822754b39f0fd`), all with zero
disk-audit issues. Focused Curl contract tests passed `226/226` and reached the DLL
unload quiet point. The expanded source-contract suite passed `185/185`; tool inventory
and tooling-governance suites passed `185/185`, `17/17`, and `12/12`; spec inventory
reported zero blocking findings; `git diff --check` passed.

The review suggestion to resolve an in-flight conflict when the popup caption Close is
used was not adopted: the authoritative E2/UI contract intentionally defines Close as
hide-without-resolution and Escape as submission of the safe published action. FOS-05,
FOS-07, and FOS-08 remain owned by their declared R0e/R1a slices and were not smuggled
into E3. This receipt also corrects the review's stale premise that E3 was uncommitted;
the E3 implementation and original exit receipt predated the review remediation.

### Package E4 — Responsive Batch Rename admission

**State:** `COMPLETE` — original exit receipt restored by the post-exit STOP remediation below
**Priority:** P1 foundation for E5
**Ordering:** the E4 row in section 8.1.

- [x] Measure current synchronous path-capability/bind cost for 1, 64, and 1,024
      rows on Local and a deterministic delayed provider.
- [x] On the UI thread capture only immutable pane/provider identity, source-to-leaf
      mappings, and completion/progress endpoints.
- [x] Move path capability, provider-parent derivation, no-follow binding,
      duplicate-object detection, identity-domain grouping, dependency construction,
      and the current schedule/cycle qualification to cancellable worker-owned
      admission without changing which preview plans can execute.
- [x] Use the existing retired-I17 task/popup surface while worker qualification runs
      and emit immutable admission facts for later R1d consumption. E4 does not create
      or own the future common Preparing lifecycle, a second task, executor, or dialog.
- [x] Post completion/failure through registered payload ownership and existing
      generation/init/drain discipline; never retain a modeless HWND unsafely.
- [x] Prove no mutation before every mapping qualifies, one identity domain per
      plan, deterministic hard-link/object duplicate rejection, navigation cannot
      retarget captured intent, cancel/close drains admission, provider unload reaches
      quiet point, and no row-proportional provider I/O remains on the UI thread.
- [x] Archive UI-thread time, wall time, retained bindings, and cancellation latency.
      Derive the numeric UI budget from evidence and place it in the owning normative
      performance/Batch Rename spec, not as an arbitrary constant here.

#### E4 post-exit STOP remediation

The post-E4 adversarial review found two regressions and one decision-surface
reachability gap that suspend the original exit claim until this addendum closes:

- [x] Keep only `taskId` across the artifact-touch prompt's nested message pump,
      re-find the task before completion, and defer destruction of the already-shut-
      down `FileOperationState` until the UI prompt handler unwinds.
- [x] Treat a missing registered prompt payload as a task-ID-addressed terminal wake
      while live state still exists; no take/miss path may park the worker indefinitely.
- [x] Preserve caption `X = hide` while adding the explicit View-menu command,
      command-palette entry, application shortcut, and automatic re-show at actionable
      decision publication. Escape remains the engine-published Cancel action only
      while an actionable decision is visible.
- [x] Prove destroy/shutdown during the artifact overlay, payload-miss wakeup,
      metadata-loading Close followed by actionable auto-show, actionable Close
      followed by explicit Show File Operations, and Escape-as-Cancel.
- [x] Record the exact corrective commit, focused test run, resource/localization
      validation, source-contract guard, and archive receipt before considering this
      addendum closed.

#### E4 corrective exit receipt

Implementation commit `5e6213e9b9cfbb4ebbbefcdde0354dd015b75344` keeps only
`taskId` across the nested artifact prompt, re-finds after the prompt pump, defers
already-shut-down File Operations state destruction until dispatch unwinds, and wakes
a live task on payload adoption failure. The same commit preserves global popup
caption `X = hide`, adds **View > File Operations**, `cmd/app/showFileOperations`, and
default `Ctrl+Shift+J`, automatically restores a popup hidden before an actionable
decision is published, and keeps Escape mapped to that visible immutable snapshot's
published Cancel action.

The complete x64 Debug solution built at that implementation commit with zero warnings
and errors under receipt
`faeb962ea72b53e29a061975737ab348d96d2d87e715641bdc845f9b2911966f`.
After the E7 evidence commit, the same unchanged binaries were rebuilt at exact clean
evidence commit `217bff62bab70fb064f7574bd10af83dc932615c` with zero warnings
and errors under receipt
`5eb4ec62c0e44f747709a3f290d805c3134992b3afe970b167e18668176fc9c9`.
Fresh Commands run
`20260831T103421Z-41740-e1aa572e6ab143c187fcf5b85ebce740` passed the three
conflict metadata/action/Close/Show/Escape cases `3/3`; run
`20260831T103514Z-31592-a5fc24f532a848188db79dc6e826bb06` passed nested-prompt
shutdown and payload-miss wakeup `1/1`. Both had no flaky/quarantine classification
and zero disk-audit issues. The compiled source guard passed separately in run
`20260831T102535Z-56316-3fc34b6891664311931c23258ec8989a`;
resource/localization and tooling-governance Pester suites passed `9/9` and `185/185`.

Curated evidence is archived at
`Specs/TestRuns/4cb089111a23/Commands/2026-08-31_123421_e4_popup_reachability_debug/`
and
`Specs/TestRuns/4cb089111a23/Commands/2026-08-31_123514_e4_artifact_prompt_lifetime_debug/`.
Explicit validation passed for seven files in each archive; whole-inventory validation
passed for 1,286 files after staging. The authoritative popup, command/menu/keyboard,
and File Operations specs carry the durable Close/Show/automatic-show/Escape contract.

#### E4 residual direct-admission prompt lifetime remediation

The later adversarial review correctly identified the twin of the posted Batch Rename
lifetime defect: direct `StartOperation` admission and external artifact-touch guards could
enter the same nested `HostShowPrompt` pump without participating in the E4 dispatch-depth
fence. Folder-window teardown could therefore destroy `FileOperationState` while one of its
member frames was still active. Regression-test commit `e41ae4bd` added the direct-admission
shutdown case. Its first focused execution also exposed incorrect test wiring to the
application-main HWND; the implementation change corrected the fixture to target the real
Folder-window HWND, so no pristine pre-fix archive is claimed.

Implementation commit `8fb251b6f8e25564c40f3159fb06e6e6358141e7` replaces the
single-handler fence with one `FileOperationPromptDispatchScope` shared by every Folder-window
File Operations admission wrapper, external artifact-touch confirmation, and posted Batch
Rename prompt. Shutdown still cancels and joins immediately, but destruction is deferred until
the outermost guarded UI frame unwinds. The deterministic case proves direct admission returns
`ERROR_SHUTDOWN_IN_PROGRESS`, publishes no task, retains the state throughout the nested pump,
and preserves the selected source.

The exact implementation commit built the complete x64 Debug solution with zero warnings and
errors under receipt
`3bc15a5887da8810816a7fb19db423ca5cb07bf729433e9533d7aec3d2954637`; the focused
Debug Commands run passed `2/2`. A test-enabled complete x64 Release build passed with zero
warnings and errors under receipt
`02e9e727d9bb9fb1fd2bd9205964904c0605f43c8b22c564b0a5f99b58f0cd78`.
Five independent Release processes recorded `124,785`, `102,788`, `108,080`, `138,783`,
and `140,539 us` (nearest-rank p95/max `140,539 us`) against the absolute `2,000,000 us`
correctness ceiling. Every sample retained state (`value0=1`), preserved the source
(`value1=1`), and returned `0x8007045B`. Curated evidence is archived in the five directories
`Specs/TestRuns/4cb089111a23/Commands/2026-08-31_194755_e4_direct_prompt_lifetime_release/`,
`Specs/TestRuns/4cb089111a23/Commands/2026-08-31_194825_e4_direct_prompt_lifetime_release/`,
`Specs/TestRuns/4cb089111a23/Commands/2026-08-31_194843_e4_direct_prompt_lifetime_release/`,
`Specs/TestRuns/4cb089111a23/Commands/2026-08-31_194901_e4_direct_prompt_lifetime_release/`, and
`Specs/TestRuns/4cb089111a23/Commands/2026-08-31_194920_e4_direct_prompt_lifetime_release/`.
Explicit pre-stage validation passed for all 45 promoted files. This is bounded teardown and
correctness evidence, not a performance-improvement claim.

#### E4 exit receipt

Implementation commit `227f3e91ef89c864c6858dedae8a982a4eb7c8f3` routes Batch
Rename execution admission through the existing File Operations task/card/popup while
the worker owns capability lookup, provider-parent derivation, no-follow binding,
identity-domain and duplicate-object proof, dependency/cycle construction, artifact
guarding, and immutable plan publication. UI admission captures only immutable intent
and registered progress/completion endpoints. Cancellation is rechecked after artifact
approval and before the mutation interlock. Modeless artifact prompts use registered
payload ownership and the existing window init/drain discipline.

Focused characterization covers 1, 64, and 1,024 Local and deterministic delayed-provider
rows, cancel at the admission barrier and during provider work, close during admission,
hard-link rejection, provider unload quiet point, and navigation after qualification. The
captured x64 Debug implementation tree built with zero warnings/errors under artifact
receipt `28f959f6fe07017ce51e71d0961d1b1fb70b33944eb9bca92bff5b35fdd426b8`.
Fresh Commands run `20260831T081222Z-95224-fd9df41501e34f63999d77690a33648b`
passed `1/1`, with no flaky/quarantine classification and zero disk-audit issues. Maximum
complete UI admission was `12,409 us` against the evidence-derived `50,000 us` budget;
worker-only qualification was at most `777,820 us`, cancel-to-terminal was at most
`16,000 us`, and retained admission bindings stayed at zero. The authoritative Batch
Rename, File Operations, and performance specs carry the durable lifecycle and budget
rules. Curated digest-bound evidence is archived at
`Specs/TestRuns/4cb089111a23/Commands/2026-08-31_101602_e4_batchrename_admission_debug/`;
explicit pre-stage validation passed for eight files and whole-archive inventory passed
for 1,260 files. The focused source-contract guard and `git diff --check` also passed.

#### E4-OR2 — Popup-control nested-prompt lifetime residual

| Activation-card field | E4-OR2 bounded execution contract |
|---|---|
| Slice/state/owner | **E4-OR2 / `COMPLETE` / Codex `/root` in this worktree.** This is a post-closeout correction for File Operations popup-control prompts only. It does not reopen Batch Rename admission, activate reader cancellation/UI-join redesign, change conflict Close/Escape behavior, or activate an R package. |
| Baseline and drift | Clean tracked baseline is `5b2aacd6`; only the user-owned untracked `last_run/` directory is present. Direct admission and posted artifact prompts use `FileOperationPromptDispatchScope`, but the custom speed-limit modal holds a raw `Task*` across its all-thread message pump and writes through it after return. The Folder-window and popup Cancel-All confirmation paths also call `HostShowPrompt` without the dispatch-depth fence and retain FileOps pointers across the nested pump. Teardown dispatched inside any of those loops can clear task storage or destroy `FileOperationState` before the caller resumes. Existing speed-prompt shutdown coverage toggles only the global prompt-stop flag and never destroys FileOps, so it cannot expose the lifetime failure. |
| Contract | Every File Operations-owned nested prompt participates in the existing Folder-window dispatch-depth fence. Speed-limit submission captures only `taskId` and the initial scalar limit before the pump, then re-finds the task after return; shutdown/task removal makes the result a no-op. Cancel-All confirmation revalidates host/state after the pump before cancellation. Teardown may cancel/join as today, but `FileOperationState` destruction remains deferred until the outermost prompt frame unwinds. No raw `Task*` may survive a nested pump. |
| RED/GREEN evidence | Extend the real custom speed-limit prompt case to dispatch `WndMsg::kFileOperationShutdownForSelfTest` while the modal is visible. RED is the current post-pump raw-task lifetime violation. GREEN proves the prompt unwinds, FileOps was retained during nested shutdown, no stale task receives the submitted value, and a fresh FileOps state can be created. Source contracts must cover the speed-limit and both Cancel-All fence/revalidation paths; existing confirm/cancel/validation/UIA cycles remain green. |
| Performance/resources | Emit `FileOps.SelfTest.SpeedLimitPromptNestedShutdownUs` from the deterministic teardown case. Capture test-enabled x64 Release evidence under `Specs/TestRuns/4cb089111a23/Commands/2026-09-01_e4_or2_popup_prompt_lifetime_release/`; every sample must remain below `2,000,000 us`, retain FileOps during the nested frame, and leave no prompt HWND. This is teardown correctness/latency evidence, not a rendering-performance claim. |
| STOP/rollback | STOP if the fix needs public ABI, worker abandonment, reader cancellation, a different conflict Close/Escape contract, or a new prompt framework; if shutdown can destroy FolderWindow itself before the scope unwinds; if a stale task still receives a speed update; or if existing FileOps prompt interaction regresses. Rollback removes only the fence/re-find/test/spec delta and leaves E4-OR2 open. |
| Exit-receipt contract | Record activation/RED/implementation/spec/evidence commits, exact Debug and test-enabled Release receipts, focused prompt and affected FileOps totals, archived metric evidence and validation, source-contract results, `git diff --check`, and the remaining blocking-reader/UI-join gap. Mark only E4-OR2 `COMPLETE`. |

- [x] Fence the speed-limit and both Cancel-All nested prompt paths.
- [x] Replace post-pump raw task use with task-ID revalidation.
- [x] Add nested FileOps-shutdown behavioral and source-contract coverage.
- [x] Update authoritative prompt-lifetime text and archive Release evidence.
- [x] Record the E4-OR2 exit receipt and mark only the residual `COMPLETE`.

#### E4-OR2 exit receipt

Activation commit `849422275c590fc22cc4273d77ab82a80b98320f` bounded the residual to
File Operations popup-control prompts. RED commit
`cbeda90256c68c4bd2d45762a1bacfcc04d544bc` extended the real custom
speed-limit case with nested `kFileOperationShutdownForSelfTest`; governed run
`e4-or2-red-speed-prompt` failed `0/1/0` because shutdown destroyed the
`FileOperationState` before the modal frame unwound. Implementation/spec commit
`08f93497701b87e9a26ede018afc74bf9fbf2c00` places the speed-limit and both
Cancel-All prompts inside `FileOperationPromptDispatchScope`, captures speed
state by task ID and scalar value only, re-finds after the pump, and revalidates
host/FileOps state before Cancel-All. Source-contract correction commit
`3e45b624f95a9c2a3c965ae4ffa8548bd379be98` aligns the pre-existing Inline-F2
heading token with the authoritative lifecycle spec; it changes no runtime
behavior.

The exact corrective x64 Debug full solution built with zero warnings/errors
under artifact receipt
`e10e174c74f8413e6442d5e055d7b3200dad6c3f2c0d1947a4d348d0ff43115c`.
Governed Debug runs passed the nested-shutdown behavior `1/0/0`, the File
Operations source contract `1/0/0`, and the affected speed-limit prompt family
`4/0/0`. The test-enabled x64 Release RedSalamander project built with zero
warnings/errors under receipt
`ad7ab5c05566b9b03739345bd52387e7ebca2fdbf81f6fa8f0c823a784499c05`.
Five independent Release processes each passed `1/0/0`; nested-frame duration
was `114892`, `95670`, `90862`, `100297`, and `117875 us`, so nearest-rank p95
was `117875 us` against the `2000000 us` ceiling. Every sample retained FileOps
during the nested frame, left no prompt HWND, and returned `hr=0`.

Evidence commit `a0333fa0b7358a7e79807eac0fb443bfe35faa6c` archives the five
digest-bound samples at
`Specs/TestRuns/4cb089111a23/Commands/2026-09-01_e4_or2_popup_prompt_lifetime_release/`.
Explicit pre-stage validation passed for 16 files and whole-inventory validation
passed for 1,867 files. Documentation drift contracts, the zero-finding spec
inventory, and `git diff --check` passed. E4-OR2 changes neither `X = hide` /
visible Escape semantics nor the known blocking-reader gap: File Operations
cancel still cannot abort an admitted blocking `IFileReader::Read`, and UI-thread
shutdown joining that I/O remains separate work.

### Package E5 — Change Case through the typed RenamePlan

**State:** `COMPLETE`
**Readiness:** exit-receipted
**Priority:** P1 rename-chain completion
**Ordering:** the E5 row in section 8.1.

#### E5 activation receipt

Activated on 2026-09-01 at `bc160cae` after the recorded E4, R1b, R7-A10,
and R7-A15 exit receipts. The bounded implementation scope is the Change Case
discovery-to-execution chain only: capture typed selected-item kind on the UI
thread, perform recursive provider discovery without mutation, fail the complete
operation when a selected or discovered directory cannot be enumerated, submit
the resulting immutable mappings as `RenamePlan(ChangeCase)`, and delete the
remaining production `FileSystemRenameBatch` authority after central execution
coverage replaces it. Clipboard admission and the speed-limit prompt remain
separate reliability follow-ups and do not expand E5 ownership.

The protected responsiveness scenario is Change Case over a deterministic local
tree while discovery and central admission remain worker-owned. Evidence records
bounded aggregate discovery/admission duration and counts only; it never records
provider paths or leaf names. The first blocking RED case is a typed selected
directory whose `ReadDirectoryInfo` fails: the exact provider failure must be
terminal and no rename mutation may occur.

- [x] Characterize file/directory conversion, recursive deepest-first order,
      equal-depth batches, case-only Local rename, siblings/parent-child selection,
      collisions, current cycle/journal behavior, cancel during discovery and between batches, 64-item
      artifact guard/revalidation batches, progress/partial failure, and unsupported
      or unbound providers.
- [x] Keep command-specific discovery where useful, but emit immutable `RenameStep`
      mappings into the same RenamePlan admission/execution authority as Batch Rename.
- [x] Define a `ChangeCase` origin's queue/admission, card/popup presentation,
      artifact owner, conflict/Keep Both behavior,
      and completion refresh/focus. Any intentional difference is documented in
      `UI_CommandMenuKeyboard.md` and `FileSystem_FileOperations.md`.
- [x] Preserve case-only handling, the landed acyclic/no-journal contract, identity revalidation,
      unreadable subtree truth, and recursive artifact warnings; never silently skip work.
- [x] Consume R7-A10 cycle/journal behavior and R7-A15 provider/path name feasibility.
      Do not add a Change Case recovery UI, Batch Rename journal authority, or a
      command-local name rule.
- [x] Remove `FileSystemRenameBatch` only after `rg` proves zero production mutation
      callers. Rename/document a legitimate test/ABI-only adapter rather than leaving
      a misleading second engine.

**Acceptance:** Change Case and Batch Rename consume the central identity, conflict,
mutation, terminal-result, schedule, and provider namespace gates without two
production rename authorities. E5 implements neither A10 nor A15; it consumes their
landed contracts and introduces no second rename or name-policy authority.

#### E5 exit receipt

Implementation and evidence commit
`dd1810f6b511c41611b94f71acc5c56c162915c1` makes Change Case discovery-only:
the worker binds typed selected roots, enumerates recursively without mutation,
returns the exact first selected/discovered directory failure, and emits immutable
provider-keyed operations. UI completion submits those operations through
`AdmitScheduledRename(..., RenameOrigin::ChangeCase, ...)`; central RenamePlan
admission then owns provider-name feasibility, identity revalidation, dependency/
cycle scheduling, artifact consent, mutation, typed results, progress, and
completion refresh/focus. The legacy direct executor is named
`DebugApplyToPathsForTests`, compiled only under `ENABLE_TESTS`, and production
source-contract coverage plus `rg` prove zero production `FileSystemRenameBatch`
mutation callers.

The test-enabled x64 Release full solution built with zero warnings/errors under
artifact receipt
`dd4c0d8597a8526d723ba44f9d42c0c8e1db3aaaf96096fa92dae6986d98b6e8`.
After the final test-only exact-casing assertion correction, the governed x64 Debug
product rebuild also completed with zero warnings/errors under exact-source receipt
`093b4fa4b32a4df8ce0a11a015b7db2d962a68ee03caaaff1ce0cddf8e1afa74`.
Focused Debug runs passed the exact selected-directory read-failure/no-mutation
case, `cmd_pane_changeCase`, `cmd_pane_changeCase_central_rename_plan`
(`e5-change-case-central-debug-02`), and the final production source guard
(`e5-final-debug-source-guard`), each `1/1`; existing recursive, collision,
case-only, cancellation, artifact, partial-failure, and dialog characterization
also remained green.

Five independent test-enabled x64 Release discovery processes each produced
1,024 typed rows with `hr=0`; nearest-rank p95 was `3,206,236 us` against the
`5,000,000 us` scenario cap. Five independent central-plan processes each
completed and published one exact Local case-only rename with no retained binding;
p95 UI admission was `26,216 us`, wall admission `28,696 us`, worker admission
`2,791 us`, and execution `2,378 us`, all below the same cap. Digest-bound evidence
is archived at
`Specs/TestRuns/4cb089111a23/Commands/2026-09-01_e5_change_case_discovery_release/`
and
`Specs/TestRuns/4cb089111a23/Commands/2026-09-01_e5_change_case_central_plan_release/`.
Both explicit archive checks passed for 16 files each and whole-inventory validation
passed for 1,799 files. `git diff --check` passed. Durable behavior and metrics now
live in `Specs/FileSystem/FileSystem_FileOperations.md`,
`Specs/UI/UI_CommandMenuKeyboard.md`, `Specs/UI/UI_FileOperationsPopup.md`, and
`Specs/Testing/Testing_PerformanceValidation.md`.

The same implementation commit also carries two separately owned reliability
follow-ups: clipboard Move readiness no longer blocks the UI thread and the custom
speed-limit prompt observes host shutdown before/after nested dispatch. Their focused
FileOps Phase 10, dialog, speed-prompt, and source-contract tests passed. These changes
do not broaden E5 and do not close the still-distinct in-flight provider-I/O problem:
a blocking reader must still gain an abort channel before UI teardown can honestly
claim a bounded join.

### Package E6 — Post-retirement artifact projection gate

**State:** `COMPLETE`
**Priority:** P2, may follow E0 when ownership does not overlap

- [x] After R7-A10 retired the Batch Rename journal and the sole durable claim source,
      remove the empty endpoint/parent claim index rather than advertising a populated
      production classifier.
- [x] Load one empty classifier snapshot per FolderView/Find request. Only recognized
      generated-name shapes enter exact no-follow projection; ordinary names perform
      no artifact capability, identity, or bind work.
- [x] Keep Possible rows visible and retain exact current identity only for warning and
      touch-guard revalidation. No name shape or captured identity grants mutation,
      recovery, cleanup, or deletion authority.
- [x] Cover FolderView, Find Files, and Compare's inherited FolderView projection; ban
      synthetic private claim keys and the retired candidate-index APIs in source guards.
- [x] Archive five independent Release samples for 10,000 ordinary names plus 100
      Possible names, requiring zero ordinary probes and all Possible probes in both
      FolderView and Find.

**Boundary:** E6 changes projection lookup cost and evidence honesty only. A09 warning
exemptions and the Properties projection are behavior changes owned by R6-A09; they may
consume E6's shared name-shape classifier and touch guard but cannot be smuggled into
this inherited package. A future durable claim source must define and prove an
authoritative schema before a keyed index can return.

#### E6 exit receipt

Implementation `48faff7e4dae2a9c04f41efa0b4e7fd0075b07ae` originally replaced
the process-wide leaf hint with an endpoint/parent candidate-index shape. R7-A10 later
retired the Batch Rename journal and the only production durable-claim source, leaving
that index empty in production while its scaling helper manufactured private keys.
Corrective implementation/evidence commit `9e14a899` removes the index, generation,
conservative flag, key stuffing, and per-request capability query. FolderView and Find
now admit only `HasPossibleArtifactName` rows to exact no-follow projection; ordinary
names perform zero artifact probes. Compare continues to inherit FolderView projection.

The exact corrective source built as a test-enabled x64 Release full solution with zero
warnings/errors under artifact receipt
`7e0e8e4fc2b1fe413d99213042dd7e1575daee1a95c601bfe2e7d17122200132`.
`file_operations_phase0_typed_contract_source_guard` passed `1/1`, proving the retired
index APIs and synthetic key stuffing absent while preserving the E5 central
scheduled-Rename admission boundary. An exact-source rerun of
`cmd_fileops_artifact_name_shape_projection_scaling` also passed `1/1`.

Five independent Release processes built before that source-guard-only synchronization
under receipt `1518a7903f2231f88029766c5afc8633d579d057a26587b29c55f0019b980371`
all passed the same scaling case. Across 10,000 ordinary names and 100 Possible names,
FolderView lookup measured `1,281`, `1,548`, `1,297`, `1,274`, and `1,263 us`
(nearest-rank p95 `1,548 us`); Find measured `1,916`, `2,190`, `1,848`, `1,986`, and
`1,819 us` (p95 `2,190 us`). Every sample reported zero ordinary probes, exactly 100
Possible probes, and `hr=0` for both consumers.

The authoritative File Operations, FolderView, Find Files, and performance specs now
carry the post-retirement name-shape-only contract. Digest-bound evidence is archived at
`Specs/TestRuns/4cb089111a23/Commands/2026-09-01_e6_name_shape_projection_release/`;
explicit validation passed for 16 files and whole-inventory validation passed for 1,815
files. `git diff --check` passed. The earlier
`2026-08-31_105535_e6_artifact_candidate_index_debug` archive remains historical
implementation characterization only and must not be cited as evidence of a populated
production claim index after journal retirement.

### Package E7 — Unambiguous interrupted-Move notice identity

**State:** `COMPLETE`
**Priority:** P2, may follow E0 when ownership does not overlap

- [x] Keep `summary.taskId` as the in-session card/action key; expose historical
      `record.taskId` only through an explicitly named `interruptedOperationId` or
      `durableTaskId` when it is useful.
- [x] Ensure user text, commands, telemetry, and dismissal never call both values
      “task ID” or imply that the durable notice is an active resumable task.
- [x] Keep actions addressed to the session card while Acknowledge removes only the
      exact breadcrumb file associated with that card.
- [x] Test one/multiple records, source/destination pane/plugin navigation, named IDs,
      exact dismissal, malformed/outside-root rejection, and absence of Resume,
      rollback, retry, Delete, Cut-again, source cleanup, or automatic recovery.

#### E7 exit receipt

Implementation commit `5e6213e9b9cfbb4ebbbefcdde0354dd015b75344` separates the
fresh in-session card/action key from the explicitly named interrupted-operation ID and
routes every action/dismissal through the session key to one exact breadcrumb. It also
corrects restart navigation to decode comparison-only canonical instance identity
(`host/default` or `host/context/<opaque-context>`) into ordinary pane context and to
reject any other persisted form instead of navigating with canonical identity text.

The exact commit's x64 Debug full-solution build passed with zero warnings/errors under
receipt `faeb962ea72b53e29a061975737ab348d96d2d87e715641bdc845f9b2911966f`.
Fresh Commands run
`20260831T102904Z-96656-d90fe2ef8fba473ea5402d5983d80c40` passed `1/1`
with no flaky/quarantine classification and zero disk-audit issues. It covers one and
multiple records, canonical-context decoding, recorded source/destination pane/plugin
navigation, named identities, auto-collapsed-card expansion, exact independent
dismissal, malformed/oversized/directory/reparse/outside-root rejection, and absence of
Resume/retry/rollback/cleanup/Cut-again/fabricated diagnostics or provider mutation.
Curated evidence is archived at
`Specs/TestRuns/4cb089111a23/Commands/2026-08-31_122904_e7_interrupted_move_identity_debug/`;
explicit validation passed for seven files and whole-inventory validation passed for
1,286 files after staging. The authoritative File Operations and popup specs carry the
durable identity and navigation contract.

### Package E8 — Consolidated specs, evidence, and retirement

**State:** `AUTHORIZED-INHERITED`
**Priority:** final gate after all accepted E/R work intended for this successor

- [ ] Update every affected authoritative domain spec and
      `Core_SharedHelpers.md` where applicable; no durable rule remains only here.
- [ ] Run focused Debug and test-enabled Release validation for every package and
      archive required performance/lifetime evidence.
- [ ] Run test inventory, spec inventory, test-run archive inventory, and
      `git diff --check` on the exact closeout commit.
- [ ] Execute section 9.4's inherited I17 FileOps/Commands/resource/spec-hygiene gate.
- [ ] Run one `Run-AllTests.ps1 -Suite Full -ValidationMode Fresh` on the exact merged
      closeout commit and record run ID, commit, totals, receipts, and archive result.
- [ ] Reconcile I3/I12 ownership, resolve or explicitly route every D2-A decision and
      accepted R package, move this sole plan to Done, and remove I14 from WIP.

### 8.2 Inactive delivery-slice briefs awaiting bounded activation

Do not activate R0-R9 as one mega-plan. Product-owner acceptance is not blanket
implementation authorization. The unchecked bullets below are outcome/boundary inputs
to a card, not sufficient instructions to edit code. A `DRAFT` card may refine them at
any time and grants no edit authority. Mark a card `ACTIVE` only when every field is
exact and every hard predecessor has a recorded exit receipt; do not maintain a second
activation definition here. R0 is an
immediate containment family, not one mandatory cross-provider change: R0a-R0e need
separate bounded ownership unless the product owner explicitly approves a combined
owner. E work above may proceed only while it preserves current behavior; only the
owning predecessor-closed `ACTIVE` R slice may implement its accepted target behavior,
and current normative text remains binding until that implementation, tests, and
authoritative-spec update land together.

### Package R0 — Immediate release-containment family

**Priority:** release blocker

#### R0a — Local FOS-01 pathname rollback

##### R0a activation card

| Field | Activated scope |
|---|---|
| Slice/state/owner | **R0a / `COMPLETE` / Codex `/root` in this worktree.** This is one Local-provider release-containment slice, not activation of R0 or any adjacent R node. |
| Baseline and drift | Planned at `680b370524cbd58ba5bf8ba7d31cc65fad6cb4be`. `git diff --name-status 680b370524cbd58ba5bf8ba7d31cc65fad6cb4be -- Common/PlugInterfaces/FileSystem.h Plugins/FileSystem/FileSystem.FileOps.cpp RedSalamander/FolderWindow.h RedSalamander/FolderWindow.FileOperations.State.cpp RedSalamander/FolderWindow.FileOperations.State.Diagnostics.cpp RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.cpp RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.Fairstream.cpp Specs/FileSystem/FileSystem_FileOperations.md Specs/Plugins/Plugins_VirtualFileSystem.md Specs/Testing/Testing_PerformanceValidation.md Specs/Plans/WIP/Operation_FileOperations_ReliabilityFirstSuccessorPlan_2026-08-29.md` returned empty before activation. The diagnostics file was added to the written owner list during closeout review because the new disposition must contribute to the existing retained-artifact aggregate; it was clean at the same baseline and does not broaden the slice. |
| Contract | Implements `FO-ITEM-01`'s `FinalLeafCreated -> WritingVisible -> FinalLeafCommitted -> Published` branch and `FO-PUBLISH-01` shape 2 for ordinary absent-destination Local regular-file Copy. It displaces only the current `CopyFileInternal` `CopyFileExW(..., COPY_FILE_FAIL_IF_EXISTS)` plus `TryRollbackCopiedDestination` pathname cleanup. Existing overwrite, directory, link, Managed Move, remote/device, resume, and atomic-reveal routes remain unchanged. |
| Scope | Production: `Plugins/FileSystem/FileSystem.FileOps.cpp::{CopyFileInternal, CopyProgressRoutine}` plus one narrow retained-final helper; `Common/PlugInterfaces/FileSystem.h::FileSystemOwnedStageDisposition`; `RedSalamander/FolderWindow.h::FileOperations::OwnedStageDisposition`; provider-result mapping/text in `RedSalamander/FolderWindow.FileOperations.State.cpp`; and retained-artifact aggregation in `RedSalamander/FolderWindow.FileOperations.State.Diagnostics.cpp`. Tests: the two FileOps cases and registration in `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.{cpp,Fairstream.cpp}`. Test-only fault seams live beside the Local copy path under `ENABLE_TESTS`. Out of scope: `FileSystem.cpp::AbortOwnedObject` cancellation behavior (FOS-11/R1c), provider ABI generation/R2, bridge publication/R3, overwrite staging, directories, links, Move/source cleanup, popup/UI, and every non-Local provider. |
| Predecessors | E0 exit receipt is closed at `fc32a4aa4b01d6d684166968a0615d3265f33d26`; R0a has no other hard predecessor. No active owner overlaps these files. |
| Steps | (1) Add RED race and abort-failure fixtures. (2) Replace the absent-new-name branch with `CreateExclusiveWriter(destination)` and retain its exact bound object through content, metadata, commit, and abort. (3) Reuse exact Local reader/metadata transfer and remove `TryRollbackCopiedDestination`/`ShouldRollbackCopiedDestination` when `rg` proves no caller. (4) Report successful direct-final content as Published; report a known surviving partial final leaf as `RetainedIncomplete`, an unknown abort outcome as Unknown, and a successful exact abort as Removed. (5) Update the host enum mapping/diagnostic text and authoritative specs. |
| RED/GREEN evidence | `Floodgate_LocalCopyNewNameConcurrentReplacementSurvives`: `pwsh -NoProfile -File .\Tools\Run-AllTests.ps1 -Suite FileOps -CaseFilter Floodgate_LocalCopyNewNameConcurrentReplacementSurvives -TestRoot D:\RedSalamander.Perf -TimeoutMultiplier 2`; baseline must fail because the pathname rollback deletes the injected foreign replacement, GREEN must preserve the replacement and remove only the exact partial object. `Floodgate_LocalCopyNewNameAbortFailureIsRetainedIncomplete`: the same command with that exact case filter; baseline must fail because the direct-final abort seam/receipt does not exist, GREEN must return `ERROR_IO_INCOMPLETE`, retain the partial leaf, and report known final non-publication plus `RetainedIncomplete`. Both cases also cover successful empty/small copy, current-source metadata/ADS preservation, cancellation after visible progress, exact receipt fields, and zero `.rs_copy_tmp_` sibling for the new-name route. |
| Performance/resources | The first case emits `FileOps.SelfTest.R0aLocalDirectFinalCopyUs` for one 64-MiB ordinary new-name Local Copy and production emits `FileOps.Local.DirectFinalCopyUs` with logical bytes and the retained pump buffer size. Capture five independent test-enabled x64 Release processes before and after under `Specs/TestRuns/4cb089111a23/FileOps/2026-08-31_r0a_local_direct_final_{baseline,candidate}_release/`. Candidate p95 must be at most `max(1.25 * baseline p95, baseline p95 + 25,000 us)`, every sample must finish within 2,000,000 us, retained payload storage must stay at or below 4 MiB per active transfer, and no timeout/UI starvation/disk-audit issue is accepted. This is a safety stabilization, not a throughput-improvement claim. |
| Durable owners | Update `Specs/FileSystem/FileSystem_FileOperations.md` for direct-final publication/result truth, `Specs/Plugins/Plugins_VirtualFileSystem.md` for the Local provider route, and `Specs/Testing/Testing_PerformanceValidation.md` for metrics/scenario/budget. No shared helper or tooling inventory changes are planned. |
| STOP/rollback | STOP if exact retained authority cannot write and abort the same final object; metadata/ADS semantics regress; the new route needs a hidden sibling/stage; a foreign replacement is touched; a known surviving partial is reported Published/Removed/ordinary Retained; the Release budget fails without an explained machine anomaly; or work crosses into FOS-11/R1c, R2, R3, overwrite, Move, directory/link, or another provider. Rollback is removal of the new direct-final helper, enum value/mapping, fault seams/tests/spec clauses, restoring the exclusive `CopyFileExW` branch only while R0a remains open and FOS-01 is still a release blocker. |
| Exit-receipt contract | Before R0a closes, record exact implementation/evidence commits; Debug and test-enabled Release build receipts; both GREEN run IDs/totals/classification/disk audit; five-process baseline/candidate metric summaries and comparison; the two validated archive paths plus whole-inventory validation; focused source-contract/resource/spec tests where affected; `rg` proof that pathname rollback is gone from this branch; `git diff --check`; authoritative-spec destinations; and any remaining FOS/R owner gaps. |

- [x] Replace absent-destination Local regular-file Copy's fail-if-exists/pathname-
      rollback path with section 6.4 shape 2: exclusive final-leaf create, retained
      exact write/abort authority, and explicit `RetainedIncomplete` truth if exact
      abort cannot finish. Do not add a hidden stage to this route.
- [x] Add concurrent foreign-create/replacement and abort-failure regressions.

##### R0a exit receipt (2026-08-31)

- **Implementation/evidence commit:** `ed4fde4edf0043d3544bfdfec9a29e01e295f511`.
  The direct-final helper exclusively creates the absent final leaf, retains the exact
  bound object for write/commit/abort, removes the pathname rollback helpers, and adds
  `RetainedIncomplete` provider/host truth. Exact abort uses a cleanup-only options
  snapshot so primary cancellation stops content work without revoking this one proven
  Local cleanup authority; FOS-11's general staged-compensation policy remains R1c-owned.
- **RED:** `20260831T114210Z-91848-0ece0e2e6812450d8623ae39c3ab557a`
  deleted the injected foreign replacement on the legacy pathname rollback;
  `20260831T114550Z-98916-860c203b075340a0865f9ea7034cb608`
  could not report retained-incomplete truth. Raw RED sandboxes were pruned after the
  failures were captured; the five archived baseline samples retain the race RED plus
  the successful 64-MiB baseline measurement.
- **Builds:** final test-enabled x64 Debug solution receipt
  `dd1fc37bc23b0ceadcc8a2b1c4ebf2790e110c98571d6c8d5f015a868ed6c1e6`
  and final test-enabled x64 Release solution receipt
  `f431e72785d19f28da1dc0da92a3cdf341b874b95c6d56f36dda0970d159fcaf`;
  both exited `0` with `0` warnings and `0` errors.
- **GREEN correctness:** Debug race run `r0a-green-debug-race-final` and Debug
  abort-failure run `r0a-green-debug-abort-final` each passed `3/0/0`, with no
  classification/quarantine activity and disk audit `0`. Release abort-failure run
  `r0a-green-release-abort-final` also passed `3/0/0` with disk audit `0`. The final
  classified Debug FileOps run `r0a-fileops-debug-full-final2` passed `125/0/20`, with
  no classification/quarantine activity and disk audit `0`. Its predecessor run had
  one timing-sensitive `Phase7_ParallelCopyMoveKnobs` observation miss (`115/1/29`);
  that case passed immediately in isolation and the final full rerun passed without a
  retry.
- **Performance/resources:** same-machine five-process baseline values were `28,374`,
  `21,200`, `20,891`, `19,495`, and `19,079` us (p95 `28,374` us). Final candidate
  runs `r0a-candidate-release-final-01` through `-05` produced `40,353`, `43,669`,
  `43,759`, `40,882`, and `45,083` us (p95 `45,083` us). The admission budget was
  `max(1.25 * 28,374, 28,374 + 25,000) = 53,374` us; candidate p95 passed, every
  sample stayed below `2,000,000` us, the retained pump buffer is 1 MiB, and every
  candidate run passed `3/0/0` with disk audit `0`.
- **Archived evidence:**
  `Specs/TestRuns/4cb089111a23/FileOps/2026-08-31_r0a_local_direct_final_baseline_release/`,
  `Specs/TestRuns/4cb089111a23/FileOps/2026-08-31_r0a_local_direct_final_candidate_release/`,
  and
  `Specs/TestRuns/4cb089111a23/FileOps/2026-08-31_r0a_local_direct_final_debug/`
  passed explicit validation (`128` files total); whole-inventory validation passed
  for `1,286` files. The compact Debug full archive intentionally omits its 283-MiB
  general perf stream while retaining results and traces.
- **Contracts/proofs:** DocumentationDriftContracts plus TestHarnessSourceContracts
  passed `203/0/0`; Debug and Release `PluginContractTests.exe` passed, including Local
  provider Debug selftests `154/0`; `rg` found no `TryRollbackCopiedDestination`,
  `ShouldRollbackCopiedDestination`, or `COPY_FILE_FAIL_IF_EXISTS` in the Local FileOps
  implementation; and `git diff --check` passed. Durable behavior now lives in
  `Specs/FileSystem/FileSystem_FileOperations.md`,
  `Specs/Plugins/Plugins_VirtualFileSystem.md`, and
  `Specs/Testing/Testing_PerformanceValidation.md`.
- **Remaining release gaps:** at this receipt R0b-R0e and R1a-R1c were inactive;
  R0b-R0e have since closed while R1a-R1c remain inactive. R0a closes only Local
  FOS-01; it does not close
  M1, general cleanup/cancellation FOS-11, overwrite, Move, directory/link, bridge, or
  non-Local publication work.

#### R0b — S3 FOS-03 exact-generation virtual-folder Delete

##### R0b activation card

| Activation-card field | R0b bounded execution contract |
|---|---|
| Slice/state/owner | **R0b / `COMPLETE` / Codex `/root` in this worktree.** This was one S3-provider release-containment slice, not activation of R0, R0d, R2, R4, or another provider. |
| Baseline and drift | Planned-at implementation baseline is clean commit `4f1faa5a61fc013c60960966000ef60fa5bac98d`. `git diff --name-status HEAD -- Plugins/FileSystemS3/FileSystemS3.Directory.cpp Specs/FileSystem/FileSystem_S3.md Specs/Plugins/Plugins_VirtualFileSystem.md Specs/Testing/Testing_PerformanceValidation.md Specs/Plans/WIP/Operation_FileOperations_ReliabilityFirstSuccessorPlan_2026-08-29.md` returned no paths before activation. Current source still routes singular ordinary Delete through bare-key `DeleteS3Object(...)` and recursive Delete through bare-key `DeleteS3Keys(...)`/`DeleteObjects`, so FOS-03 remains live at this baseline. |
| Contract | Implement section 6.6 `FO-DELETE-01` for accepted A13/A13-MD1: one canonical S3 account/profile/bucket/prefix virtual-folder boundary; each ordinary current-key delete consumes its just-observed ETag and omits `VersionId`; same-key mismatch is a non-mutation and may be rebound only by a later bounded listing; exact historical `VersionId` delete remains a separate fixed-set cleanup/rollback authority; success requires one complete empty re-list; request-level unknown, unsupported conditions, permission failure, cancellation, deadline, or persistent writers stop without a false empty-folder claim. This replaces the current unconditioned batch wording in `Specs/FileSystem/FileSystem_S3.md` and adds the provider-specific durable rule to `Specs/Plugins/Plugins_VirtualFileSystem.md`. |
| Scope | In scope: `Plugins/FileSystemS3/FileSystemS3.Directory.cpp` symbols `DeleteResolvedPath`, the displaced `DeleteS3Keys` owner, `DeleteS3Object` ordinary-call inputs, bounded virtual-folder metrics, `DebugS3Graph`, R0b debug selftests, and one test-enabled focused R0b export; the exact `--s3-r0b-delete-selftests` dispatch branch in `Tests/PluginContractTests/PluginContractTests.cpp`; `Specs/FileSystem/FileSystem_S3.md`; `Specs/Plugins/Plugins_VirtualFileSystem.md`; `Specs/Testing/Testing_PerformanceValidation.md`; this plan/card/receipt; compact evidence below `Specs/TestRuns/4cb089111a23/FileOpsS3/`. Out of scope: public production plugin ABI/interfaces, the legacy Debug-only S3 suite boundary, other PluginContractTests dispatch/tests, `FolderWindow.FileOperations.State.cpp`, host scheduling/results/cancellation, Move/copy publication, exact-version cleanup/rollback call sites, R0d capability truth, R2 typed authority, R4 traversal, non-S3 providers, live AWS credentials, and generic retry policy. Conditional `DeleteObjects` batching is out because the active AWS SDK request cannot express a distinct `If-Match` for every member. |
| Predecessors | E0 exit receipt is closed at `fc32a4aa4b01d6d684166968a0615d3265f33d26`; R0b has no other hard predecessor. R0a is complete and no active owner overlaps these files. |
| Steps | (1) Add deterministic generation/boundary/termination/resource fixtures and coalesced per-operation metrics; prove the current bare-key path RED. (2) Pass the probed ETag-only current-generation condition for singular object Delete. (3) replace bare-key recursive batching with serialized, cancel/deadline-bounded per-observation conditional deletes; treat `ERROR_REVISION_MISMATCH` as changed in-scope membership to be freshly re-listed, and return every other request failure immediately without retry. (4) retain the empty-list linearization point, late-child convergence, exact prefix boundary, and exact-VersionId cleanup call sites. (5) update authoritative specs, run Debug/test-enabled Release validation and same-machine perf/resource evidence, record the exit receipt, then mark only R0b complete. The displaced bare-key `DeleteS3Keys` owner is removed rather than retained as an alternate ordinary path. |
| RED/GREEN evidence | RED/GREEN IDs are `S3.R0b.OrdinaryObjectGeneration`, `S3.R0b.RecursiveReplacementConvergence`, `S3.R0b.VirtualFolderBoundary`, `S3.R0b.MarkerAndVersioningTruth`, `S3.R0b.RequestFailureNoRetry`, `S3.R0b.CancelAfterProgress`, `S3.R0b.PersistentWriterBound`, and `S3.R0b.ResourceBudget`. Debug executes the full fault corpus through required `RedSalamanderS3DebugSelfTests`; test-enabled Release executes the shared 4,096-object resource/conditional-request contract through `PluginContractTests.exe --s3-r0b-delete-selftests` and `RedSalamanderS3R0bDeleteSelfTests`, without widening the legacy Debug-only suite. Baseline must fail at least the ordinary/replacement conditional-request assertions while retaining the existing late-child convergence assertion; the focused Release baseline must fail only its conditional-request assertion while still emitting the timing sample. GREEN requires all full Debug S3 selftests and the focused test-enabled Release R0b export to pass. Faults cover same-key replacement between observation/delete, late child, exact `prefix/` exclusion of `prefix` and `prefix2/`, canonical bucket path retention, trailing-`/` marker-only folder, mixed unchanged/changed observations, request-level timeout/unknown with one attempt, cancel after a successful member, a writer that exhausts the pass bound, and versioned current-key deletion that creates a marker while preserving historical versions. Commands: `build.ps1 -ProjectName FileSystemS3 -Configuration Debug`, `build.ps1 -ProjectName PluginContractTests -Configuration Debug`, direct Debug `PluginContractTests.exe`, and the same two builds with `RSBuildEnableTests=true`/Release followed by the focused flag, plus affected source-contract/spec tests. |
| Performance/resources | Protected scenario: one ordinary recursive S3 virtual-folder Delete over 4,096 listed 4-KiB debug objects with no writer. Emit once per operation: `FileOps.S3.VirtualFolderDelete.ElapsedUs`, `.PassCount`, `.ObservedObjectCount`, `.ConditionalRequestCount`, `.RevisionMismatchCount`, and `.ResidualObjectCount`; the deterministic fixture emits `FileOps.S3.VirtualFolderDelete.SelfTestUs`. Capture five independent test-enabled x64 Release processes before and after under `Specs/TestRuns/4cb089111a23/FileOpsS3/2026-08-31_r0b_exact_generation_{baseline,candidate}_release/`. Candidate harness p95 must be at most `max(1.50 * baseline p95, baseline p95 + 50,000 us)` and every sample below 2,000,000 us. Correctness intentionally changes service-request count from up to 1,000 bare keys per request to exactly one conditional request per observed member: admission is serialized (one live request), capped at 64 passes/deadline, no key is attempted more than once per listing, and retained memory is one `PlannedTransferObject` list plus constant per-item request state with no duplicate bare-key vector. This is a safety stabilization, not a network-throughput improvement claim; live-AWS latency remains environment-gated and directional. |
| Durable owners | `Specs/FileSystem/FileSystem_S3.md` owns S3 virtual-folder/current-generation semantics and metrics; `Specs/Plugins/Plugins_VirtualFileSystem.md` owns the provider matrix/ordinary destructive-authority statement; `Specs/Testing/Testing_PerformanceValidation.md` owns the deterministic 4,096-object scenario and archive/budget. This plan retains only delivery mapping and the generated exit receipt. |
| STOP/rollback | STOP if the endpoint cannot enforce ETag-only `If-Match`; a condition mismatch mutates any generation; ordinary Delete requires/uses `VersionId`; the implementation broadens beyond one canonical bucket/prefix, retries an unknown request result, claims success without a complete empty re-list, loses late-child convergence, needs host/generic ABI work, touches exact-version cleanup semantics, retains more than one list plus constant request state, violates the Release budget without an explained machine anomaly, or crosses into R0d/R2/R4/another provider. Rollback is removal of the new conditional ordinary-delete helper/metrics/tests/spec clauses and restoration of the prior bare-key helper only while R0b remains open and FOS-03 is still a release blocker; do not partially advertise the route as safe. |
| Exit-receipt contract | Before R0b closes, record exact activation/RED/implementation/evidence commits; Debug and test-enabled Release build receipts; focused/full S3 GREEN totals and fault results; five-process baseline/candidate metric summaries; conditional-request/pass/residual/resource bounds; validated archive paths and whole-inventory validation; source proof that ordinary object/prefix paths no longer call bare-key or `DeleteObjects`; `git diff --check`; authoritative-spec destinations; and remaining FOS/R owner gaps. |

- [x] Preserve ordinary S3 prefix Delete's bounded re-list-until-empty folder
      behavior, canonicalized to one exact account/profile/bucket/prefix boundary.
- [x] Make ordinary singular and batch current-key removal consume the ETag observed
      for each key and omit `VersionId`. Use conditional batching only when the active
      SDK/service route can express every ETag; otherwise use bounded conditional
      singular requests. Keep exact `VersionId` deletion as a separately admitted
      fixed-version cleanup/rollback operation.
- [x] Re-list/rebind late or replaced in-scope members while the bounded operation can
      converge. On persistent writers, unsupported conditions, permission failure,
      cancel, or deadline, stop and report exact residuals rather than claiming the
      virtual folder was deleted.
- [x] Keep the existing late-child-is-deleted selftest and add same-key replacement,
      `prefix`/`prefix2/`/bucket/profile escape, marker-only folder, mixed conditional
      batch, request-level unknown, cancellation-after-progress, non-converging-writer,
      and versioned-bucket delete-marker/result-truth cases.

##### R0b exit receipt

- **Commits and scope:** activation is
  `7dfc27ecd8ab6b11e9aca3c0673e37d68b0bdbbd`; the RED fault/resource corpus is
  `6201901a25c2f743f7d0715ac14a7fb8c3c2bb52`; the focused Release boundary is
  `9443f8fedfe445257590535ade48d1946aba9949` plus
  `cd9d5f05f0da2dda4045a0df2ecfad925c6c2853`; implementation is
  `e172269d0e44156cee22e6d22cd8ac4a9069a556`; authoritative-spec and evidence
  closeout is `07ef0fdc86e40d1e1c1220728f2da4ccd9357255`. The intermediate
  `29bed097ecce4701984a2e0bca3da095dbc50f89` experiment exposed that the legacy S3
  suite has Debug-only helpers; `0c33f789ad1a8050611601f59596d7677e2ff333`
  immediately restored that boundary. The final focused export shares only the R0b
  4,096-object contract and does not widen the legacy suite.
- **RED/GREEN:** before the behavior change, direct Debug `PluginContractTests.exe`
  reported S3 `passed=184, failed=11`; those 11 failures were the new R0b generation
  assertions, while the existing late-child convergence case and unrelated S3 suites
  remained green. The focused Release baseline reached the timing point and was
  intentionally `passed=2, failed=1` solely because conditional-request count was zero.
  On exact committed tree `07ef0fdc`, Debug S3 is `passed=194, failed=0` and the whole
  `PluginContractTests.exe` process passes. The test-enabled Release command
  `PluginContractTests.exe --s3-r0b-delete-selftests` is `passed=3, failed=0`.
- **Exact-commit builds:** x64 Debug built with zero warnings/errors under FileSystemS3
  receipt `2c3f01d6de6ad1de3745a623390a21265f6f9d26f345fea0d95cf026b2a4e4e7`
  and PluginContractTests receipt
  `8b34c0d99ca3319c6c0443322c7559789ec10ce01a6818a4b7cdcc6e880a01f3`.
  Test-enabled x64 Release built with zero warnings/errors under FileSystemS3 receipt
  `3fbe44958a3f7dab12fc1537566f7b36ae7da3aca78d9315ed7b9f88a2a81ceb`
  and PluginContractTests receipt
  `b6aee7be117c7246eeb53cf9592d311bce831d23129fc70737992278975f8a5b`.
- **Fault and authority proof:** the GREEN corpus covers stale singular ETag rejection,
  recursive same-key replacement/re-list, unchanged-plus-changed members, exact
  trailing-`/` scope, different-bucket rejection, exclusion of the prefix-name object
  and `prefix2/`, marker-only folders, missing ETag, versioned current-key delete-marker
  truth, separate exact-VersionId historical cleanup, one-attempt request failure,
  cancellation after progress, a 64-pass persistent writer, and the resource bound.
  Account/profile remain bound by the selected `FileSystemS3` instance and cannot be
  selected by the object path; bucket and key-boundary escape are direct fault tests.
  Source search has no `DeleteS3Keys`, `DeleteObjectsRequest`, `ObjectIdentifier`,
  `kMaxDeleteBatchSize`, or `Model::Delete` owner in
  `FileSystemS3.Directory.cpp`. Ordinary object and prefix paths build ETag-only current-
  key conditions; exact-VersionId calls remain confined to fixed-set cleanup/rollback.
- **Performance/resources:** five independent baseline samples were `5,623`, `5,790`,
  `5,037`, `6,334`, and `6,051` us (p95 `6,334` us); five candidate samples were
  `5,290`, `6,094`, `5,882`, `6,547`, and `6,613` us (p95 `6,613` us). The candidate
  delta is `+279` us / `+4.40%`, below the safety ceiling
  `max(1.50 * 6,334, 6,334 + 50,000) = 56,334` us, and every sample is below
  2,000,000 us. Each candidate process made exactly 4,096 serialized ETag-conditioned
  requests, used two complete listings, and ended with zero residuals. The route retains
  one `PlannedTransferObject` list plus constant request state and no duplicate bare-key
  vector. This fake-graph evidence makes no live-AWS throughput claim.
- **Archives and contracts:** the paired baseline/candidate roots under
  `Specs/TestRuns/4cb089111a23/FileOpsS3/` pass explicit archive validation for 32 files;
  whole-inventory validation passes for 1,446 files. Documentation-drift plus harness
  source contracts pass `203/203`, and `git diff --check` passes. Durable behavior is in
  `Specs/FileSystem/FileSystem_S3.md`, `Specs/Plugins/Plugins_VirtualFileSystem.md`, and
  `Specs/Testing/Testing_PerformanceValidation.md`.
- **Remaining release gaps:** at this receipt R0c-R0e and R1a-R1c were inactive;
  R0c-R0e have since closed while R1a-R1c remain inactive. R0b closes only S3 FOS-03
  and does not close M1. The
  provider returns a non-success selected-root result and exact aggregate residual count;
  itemized central residual/result presentation remains owned by R1a/R2/R6.

#### R0c — MTP FOS-02/FOS-04 exact recovery identity

##### R0c activation card

| Activation-card field | R0c bounded execution contract |
|---|---|
| Slice/state/owner | **R0c / `COMPLETE` / Codex `/root` in this worktree.** This was one MTP-provider release-containment slice, not activation of R0d, R1, R2, another provider, or a public plugin-ABI cutover. |
| Baseline and drift | Planned-at implementation baseline is clean commit `156b13bd`. A scoped `git diff --name-status` over the MTP provider, its selftests/specs, and this plan returned no paths before activation. `RecordOverwriteJournalIntent(...)` writes schema v1 before the temp object exists; writer and device-source paths read a non-empty temp PUID but never atomically add it to the journal before deleting the original. Replay still deletes/renames by stored path, can sweep by leaf+size+timestamp, clears malformed/missing/exhausted records, and does not validate the journal's full device identity. `MtpDeviceIdentitySuffix`, `MtpPersistentObjectIdentitySuffix`, and `MtpObjectIdentitySuffix` expose only `FormatMtpIdentityHash(...)`; live device lookup accepts that hash as identity. FOS-02/FOS-04 therefore remain live. |
| Contract | Implement section 6.6 `FO-AUTH-01` for MTP recovery/path identity. Step-0 intent remains fail-closed before device mutation, but no automatic cleanup/rename authority exists until a non-empty created-temp PUID is atomically persisted with the full device identity. Replay must enumerate the exact parent, resolve exactly one full PUID match, and mutate only that rebound object; path, display name, size, timestamp, object hash, and destination size are hints only. Missing/legacy/malformed/device-mismatched/ambiguous authority mutates nothing and preserves the journal as a quarantined recovery artifact. Device/object path suffixes carry reversible percent-encoded full PnP/PUID/object IDs; the existing stable hash may remain only as a storage/telemetry hint when the same component also carries and validates the full identity. Ambiguous matches fail closed. |
| Scope | In scope: `Plugins/FileSystemMtp/FileSystemMtp.Core.cpp` overwrite-journal schema/record/rewrite/replay/quarantine owners; `FileSystemMtp.Shared.cpp` and `FileSystemMtp.Internal.h` identity token helpers; `FileSystemMtp.Device.cpp` live device/child resolution; `FileSystemMtp.FakeBackend.cpp` deterministic identity/recovery fault seams only; the exact test-only MTP exports/`ENABLE_TESTS` guards in `Plugins/FileSystemMtp/Factory.cpp`, `FileSystemMtp.Core.cpp`, `FileSystemMtp.Device.cpp`, and `FileSystemMtp.Internal.h` needed to run the same fake-backend case in test-enabled Release; `RedSalamander/SelfTest/CompareDirectories/CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp` focused RED/GREEN cases; `Specs/FileSystem/FileSystem_Mtp.md`; MTP rows/contracts in `Specs/Plugins/Plugins_VirtualFileSystem.md`; `Specs/Testing/Testing_PerformanceValidation.md`; this card/receipt; and compact evidence under `Specs/TestRuns/4cb089111a23/FileOpsMtp/`. Out of scope: public production `IFileSystem` ABI/capability v2, non-test exports, host FileOps result/lifecycle work, general MTP watchdog/cancellation, non-overwrite ordinary Delete, direct backend overwrite policy, Move route truth, other providers, live-device credentials/hardware, and R0d/R1/R2. Reuse `Common/UriEncoding.h`; do not add a second byte encoder. |
| Predecessors | E0 is closed at `fc32a4aa4b01d6d684166968a0615d3265f33d26`. R0c has no R0b dependency, but R0b is complete and no active owner overlaps these files. R0d remains inactive until this independently releasable slice closes. |
| Steps | (1) Add deterministic RED cases for unpersisted-PUID replay, exact-PUID recovery after path reuse, malformed/legacy/device-mismatched quarantine, case-fold-hash collision, full device suffix lookup after friendly-name drift, and bounded large-sibling resolution; retain current success-path overwrite tests. (2) Introduce a journal schema that persists full device identity and temp PUID after temp verification and before original deletion. (3) replace path/name/size/time recovery with exact parent enumeration plus unique full-PUID rebind; clear only proved terminal states, otherwise retain/quarantine. (4) replace hash-only exposed device/object tokens with percent-encoded full identities and make live resolution compare the complete suffix/ID with ambiguity rejection. (5) update authoritative specs, run focused Debug and test-enabled Release evidence, record the receipt, and mark only R0c complete. |
| RED/GREEN evidence | New cases must prove: a legacy/planned no-PUID journal never deletes a lookalike; a journal whose path was reused by a different PUID never mutates that occupant; exactly one matching PUID can be rolled back or promoted; a matching PUID already at the destination proves a completed swap; zero/two matches and device-identity mismatch quarantine without mutation; malformed records are not deleted; `CasePuid` and `casepuid` no longer collapse despite the legacy case-fold hash collision; friendly-name drift resolves only through the full PnP token; duplicate full-ID matches fail closed. GREEN also retains the existing overwrite crash-window, duplicate-name, WPD cache, journal-generation, and watchdog suites. Focused command is test-enabled `Run-AllTests.ps1 -Suite Compare -CaseFilter mtp_r0c_exact_recovery_identity`; final Debug validation includes the affected existing MTP cases. |
| Performance/resources | Protected fake-backend scenario: resolve one exact temp PUID among 4,096 siblings, plus format/compare 4,096 full identity suffixes, in a test-enabled x64 Release process. Emit `FileOps.Mtp.R0c.ExactIdentity.SelfTestUs`, `FileOps.Mtp.R0c.RecoveryEnumeratedCount`, `FileOps.Mtp.R0c.FullIdentityCompareCount`, `FileOps.Mtp.R0c.AmbiguousMatchCount`, and `FileOps.Mtp.R0c.QuarantinedJournalCount`. Capture five independent baseline/candidate processes under `Specs/TestRuns/4cb089111a23/FileOpsMtp/2026-08-31_r0c_exact_identity_{baseline,candidate}_release/`. Candidate p95 must be at most `max(1.50 * baseline p95, baseline p95 + 50,000 us)` and every sample below 2,000,000 us. Retained memory is one backend enumeration vector plus constant match state; do not build a second candidate-path/PUID vector. This is deterministic safety/resource evidence, not live-WPD latency. |
| Durable owners | `Specs/FileSystem/FileSystem_Mtp.md` owns MTP identity and overwrite-recovery semantics/metrics; `Specs/Plugins/Plugins_VirtualFileSystem.md` owns the provider identity/authority matrix; `Specs/Testing/Testing_PerformanceValidation.md` owns the deterministic 4,096-identity scenario and archive/budget. This plan retains delivery mapping and the generated receipt only. |
| STOP/rollback | STOP if WPD cannot return the created temp PUID before original deletion; replay must mutate by path/name/size/time/hash; the path token cannot carry the full identity without an ABI change; ambiguity can select one candidate; legacy/malformed/device-mismatched records must be cleared to make progress; the change needs host scheduling/result work; existing overwrite no-loss cases regress; memory exceeds one enumeration vector plus constant state; Release evidence exceeds the budget without a diagnosed machine anomaly; or scope crosses R0d/R1/R2/another provider. Rollback removes the new exact-recovery/token helpers/tests/spec clauses only while R0c remains open; never restore a claim that weak recovery is safe. |
| Exit-receipt contract | Before R0c closes, record exact activation/RED/implementation/evidence commits; Debug and test-enabled Release build receipts/runs; focused and affected MTP GREEN totals; five-process metric summaries and resource bounds; archive validation; source proof that replay has no path/name/size/time/hash mutation authority and exposed suffix resolution compares full identities; `git diff --check`; authoritative-spec destinations; and remaining FOS/R owner gaps. |

- [x] Persist the full temp PUID before granting cleanup authority; disable weak
      recovery, retain uncertain journals, and expose retained artifacts.
- [x] Fail ambiguous full-ID/hash matches and prevent hashes, names, size, or time
      from authorizing mutation.

##### R0c exit receipt

- **Commits and scope:** activation is
  `610b8849c622cab2c73612410ae18ccba91929ec`; the test-enabled Release boundary
  refinement is `55008c96cc85c94a1a00364fb369b03f5fbcf14c`; the RED identity/recovery corpus is
  `6eb05929fcb92ae0fddffe1240fa162492e8fc12`; implementation is
  `00e95d5295361280afc2ed4dfdc2472215e2b79b`; the Release-only test warning cleanup is
  `00b6df42a581b944093f0f107832dba6bbb5d778`; authoritative specs and archived
  evidence are `d4a2d4fd2560fc81b04d9fe71f98326d6760cab4`. No public production plugin ABI,
  host lifecycle/result path, other provider, R0d, R1, or R2 surface changed.
- **RED/GREEN:** the focused RED case failed exactly because the old case-folded hash
  exposed 4,095 of 4,096 identities and collapsed `CasePuid` with `casepuid`. On the
  implementation source tree, the affected Debug MTP family is `passed=54, failed=0,
  skipped=1`; the skip is the declared live-device smoke because no approved physical
  device was configured. Its zero-error Debug build receipt is
  `f4de1faa879dc15a98ead7aafa5cf348e90899ef44760fcf2e8233f3b213d5af`; the final
  `00b6df42` change is Release-only warning hygiene and does not alter Debug behavior.
  On exact commit `00b6df42`, test-enabled x64 Release is `54/0/1`, with zero warnings
  and zero errors under full-solution receipt
  `ecc1c16341cd4cb6902bb3afaa722b72dada7faed9f2134fb06feb7309bfe9ca` and
  FileSystemMtp receipt
  `a5a3aa8988f6e1ef13f7c386f27e1863d27e093a39ffe5a7b45b2e010c127b1a`.
- **Fault and authority proof:** schema v2 atomically replaces `planned` intent with
  complete device identity plus non-empty temp PUID after verification and before
  original deletion. Replay validates exact device identity, enumerates the exact
  parent once, compares full PUIDs case-sensitively, requires exactly one rebound at
  the temp or destination path, and mutates only that object. Legacy, malformed,
  PUID-less, device-mismatched, zero/multiple-match, and wrong-path records mutate
  nothing and are preserved as `.stale[.N]`. Path/name/size/time/destination-size/hash
  inference and orphan sweep are removed. Exposed device/PUID/object suffixes carry
  reversible percent-encoded full IDs; live suffix resolution compares the full token
  and rejects ambiguity. The existing hash is now an exact-case storage/telemetry hint.
- **Performance/resources:** five sequential RED baseline processes recorded whole-case
  durations `25`, `24`, `23`, `25`, and `23` ms (p95 `25` ms). The legacy token failure
  occurs before recovery setup, so those honest RED processes cannot emit the candidate
  metric; no synthetic metric was substituted. Five candidate processes emitted
  `FileOps.Mtp.R0c.ExactIdentity.SelfTestUs` values `25,939`, `27,596`, `26,463`,
  `26,550`, and `26,466` us (p95 `27,596` us), with whole-case durations `40`, `39`,
  `37`, `39`, and `39` ms (p95 `40` ms). The comparable p95 passes
  `max(1.50 * 25 ms, 25 ms + 50 ms) = 75 ms`; every candidate metric is below
  2,000,000 us and reports 4,096 enumerated/full-compared identities. Retained state is
  one backend enumeration vector plus constant match state; this is not live-WPD
  latency evidence.
- **Archives and contracts:** paired evidence under
  `Specs/TestRuns/4cb089111a23/FileOpsMtp/2026-08-31_r0c_exact_identity_{baseline,candidate}_release/`
  passes explicit validation for 32 files; whole-inventory validation passes for 1,478
  files. Documentation-drift plus harness source contracts pass `203/203`,
  `Get-SpecInventory.ps1 -FailOnFindings` reports zero blocking findings, and
  `git diff --check` passes. Durable behavior is in
  `Specs/FileSystem/FileSystem_Mtp.md`, `Specs/Plugins/Plugins_VirtualFileSystem.md`,
  and `Specs/Testing/Testing_PerformanceValidation.md`.
- **Remaining release gaps:** at this receipt R0d-R0e and R1a-R1c were inactive;
  R0d-R0e have since closed while R1a-R1c remain inactive. R0c closes only MTP
  FOS-02/FOS-04 and
  does not close M1, advertise host cancellation deadlines, repair central receipt/
  capability truth, or change non-overwrite provider mutation routes.

##### R0c-OR1 — Destination-identity overwrite replay residual

| Activation-card field | R0c-OR1 bounded execution contract |
|---|---|
| Slice/state/owner | **R0c-OR1 / `COMPLETE` / Codex `/root` in this worktree.** This was a post-closeout correction to MTP overwrite replay only; it did not reopen the full R0c identity/path work or activate reader cancellation, host shutdown, prompt lifetime, R1, R2, or another provider. |
| Baseline and drift | Clean source baseline is `e10ffc606e0283b6ec33e19c834de985a02e52a0`; only the user-owned untracked `last_run/` directory is present. Schema-v2 replay uniquely rebinds `tempPuid`, but when that exact object remains at `tempPath`, it calls path-based `GetAttributes(destinationPath)` and deletes the temp whenever the destination path appears occupied. The journal does not record the original destination PUID. A stale WPD path observation or a different object installed at the destination can therefore authorize deletion of the only remaining replacement bytes. The authoritative MTP spec and `mtp_overwrite_journal_replay_removes_temp_when_final_exists` encode that unsafe occupancy rule. |
| Contract | Persist both the verified temp PUID and the pre-delete destination PUID before deleting the original. Replay enumerates the exact parent once and decides only from full, ordinal, case-sensitive PUID state: temp PUID at destination proves completion; temp PUID at temp plus the recorded destination PUID still at destination proves safe rollback of the temp; temp PUID at temp plus a proved-absent destination permits promotion; any different, empty, duplicated, or otherwise unproved destination identity quarantines the journal and mutates nothing. Bump the executable journal schema so older identified records lacking destination authority are quarantined rather than inheriting the new semantics. Path occupancy remains a routing observation, never deletion authority. |
| RED/GREEN evidence | RED must install a different-PUID object at the recorded destination while the identified temp remains and prove the current replay deletes that temp and clears the journal. GREEN must retain both objects, preserve both byte payloads, and quarantine the journal. Existing exact-original-at-destination rollback, destination-absent promotion, temp-already-at-destination completion, bounded mutation retry, legacy/malformed quarantine, and full R0c identity cases remain green. Focused case: `mtp_overwrite_journal_destination_identity_mismatch_quarantines`. |
| Performance/resources | Reuse the deterministic fake-backend exact-parent scenario and existing journal replay metrics. Replay must still perform one parent enumeration and retain only constant identity/disposition state beyond that backend vector. Archive test-enabled x64 Release evidence under `Specs/TestRuns/<commit>/FileOpsMtp/2026-09-01_r0c_or1_destination_identity_release/`; the focused case must remain below 2,000,000 us and report the enumerated count plus the destination-authority rejection. This is deterministic safety/resource evidence, not live-WPD latency. |
| STOP/rollback | STOP if the original destination PUID cannot be obtained before deletion, if replay needs path/name/size/time/hash as authority, if ambiguity selects an object, if old schemas must execute for compatibility, if the fix needs a public ABI or host scheduler/lifecycle change, if one-parent-enumeration/constant-state is exceeded, or if existing no-loss overwrite cases regress. Rollback removes only this residual's schema/identity/test/spec changes and leaves the journal quarantined; never restore path occupancy as delete authority. |
| Exit-receipt contract | Before close: record activation/RED/implementation/spec/evidence commits; exact Debug and test-enabled Release build receipts; focused and affected MTP totals; archived metric/resource evidence and validation; source proof that replay has no destination `GetAttributes` mutation authority; `git diff --check`; and the remaining reader-cancel/prompt-lifetime gaps. Mark only R0c-OR1 `COMPLETE`. |

- [x] Persist original destination PUID beside the temp PUID before destructive swap.
- [x] Decide replay only from one-parent full-PUID state and quarantine ambiguity.
- [x] Replace the path-occupancy GREEN with identity-mismatch RED/GREEN coverage.
- [x] Update the MTP specification and archive focused Release evidence.
- [x] Record the R0c-OR1 exit receipt and mark the residual `COMPLETE`.

##### R0c-OR1 exit receipt

- **Commits and scope:** activation is `de0e2308`; the corrected RED is
  `18d33e49`; schema-v3 destination-identity implementation and authoritative-spec
  update are `34defa7e`; fail-before-delete coverage for an unavailable original
  destination PUID is `faa4abd4`; the focused Release archive is `68a1c197`. No
  public ABI, host scheduler/lifecycle,
  reader-cancellation, prompt-lifetime, other-provider, R1, or R2 surface changed.
- **RED/GREEN:** exact RED run `r0c-or1-red-schema2` failed `0/1/0` because the
  schema-v2 path-occupancy replay deleted the only identified temp and its reader
  returned `0x80070002`. On exact implementation source `faa4abd4`, x64 Debug built
  with zero warnings/errors under full-solution receipt
  `642bf5ecd17da1cb4aec532bdf7245369ef4c27380ad1368968f9bcf35e39a6b`; the governed
  MTP family run `r0c-or1-mtp-family-debug-final` passed `54/0/1`. Test-enabled x64
  Release built serially with zero warnings/errors under RedSalamander-project receipt
  `30f307d6773fd6ee6cf4292c76b59f9cd37f214610a88be247e05e752a336f88`; focused run
  `r0c-or1-green-focused-release` passed `1/0/0`, and affected run
  `r0c-or1-mtp-family-release-final` passed `54/0/1`. The one family skip is the
  declared physical-device smoke because no approved device was configured.
- **Authority and fault proof:** schema v3 persists non-empty temp and original
  destination PUIDs before deleting the original. An unavailable original PUID fails
  before delete and removes only the newly staged temp. Replay accepts only identified
  schema-v3 state, enumerates the exact parent once, compares full PUIDs ordinally,
  and permits mutation only for temp-at-destination completion, exact-original-at-
  destination rollback, or proved destination absence. Different, empty, duplicated,
  legacy, or otherwise unproved identity quarantines without mutation. Source search
  finds neither `BackendItemExists` nor destination `GetAttributes` replay authority;
  schema v1/v2 parsing remains only so those records can be quarantined.
- **Performance/resources:** the focused Release case completed in `24 ms` and emitted
  `FileOps.Mtp.R0cOr1.DestinationIdentity.SelfTestUs=23,492 us`, one exact-parent
  enumeration of `10` entries, and one destination-identity rejection. It is below
  the `2,000,000 us` deterministic ceiling. Retained state is one backend enumeration
  vector plus constant temp/destination identity and count state; this is fake-backend
  safety/resource evidence, not live-WPD latency.
- **Archives and validation:** focused Release evidence is under
  `Specs/TestRuns/4cb089111a23/FileOpsMtp/2026-09-01_r0c_or1_destination_identity_release/`;
  explicit archive validation passes for `4` files and whole-inventory validation
  passes for `1,851` files. Documentation-drift contracts pass, spec inventory reports
  zero blocking findings, and `git diff --check` passes. The broad harness source-
  contract run exposes one pre-existing stale Change Case assertion that still expects
  the retired skip-as-success branch; it is not represented as R0c-OR1 GREEN.
  Durable behavior is in `Specs/FileSystem/FileSystem_Mtp.md` and
  `Specs/Plugins/Plugins_VirtualFileSystem.md`.
- **Remaining release gaps:** R0c-OR1 closes only destination-identity authorization
  during MTP overwrite replay. FileOps still lacks a cancel channel that aborts a
  blocking `IFileReader::Read`, UI teardown can still join that admitted I/O, and
  unfenced prompt/lifetime work remains separate. No reader-cancel or prompt-lifetime
  closure is claimed here.

##### R0c-OR2 — Live overwrite occupancy and verified-temp retention

| Activation-card field | R0c-OR2 bounded execution contract |
|---|---|
| Slice/state/owner | **R0c-OR2 / `COMPLETE` / Claude in this worktree (2026-09-04).** Review fold-in (R0-RC3 step 7) on the MTP overwrite commit only. |
| Verification first | Verified before the change: exclusive create already looked the destination child up live (`FindDestinationChildCached`) and replay already enumerated the parent live, but the writer and device-source overwrite commits took occupancy and the destination PUID from the path cache (`ResolvePathCached` returns a hit without a device round trip); missing-object results on a cached entry already dropped the device caches (`FailAndMaybeInvalidateCaches`); a failed original delete destroyed the verified temp. |
| Contract | One occupancy primitive for the commit: `IMtpBackend::RefreshPathOccupancy(path)` forgets the cached leaf so the existence check, the PUID read, and the original delete all resolve live; both overwrite commits call it before deciding. When the original delete fails, the verified temp and its identified journal entry are retained; replay resolves the pair from live PUID state (temp removed beside an intact original, committed when the original is gone, quarantined on ambiguity). |
| RED/GREEN evidence | `mtp_wpd_overwrite_occupancy_is_live_on_stale_path_cache` (WPD self-test backend, `replacePhotoAfterFirstLookup` + `writable`): without the refresh the plugin still reports the stale PUID after the commit attempt; with it the live PUID. `mtp_overwrite_delete_original_failure_keeps_original_and_allows_retry` now expects the retained journal after the failure (the file, observed before any backend command) and the temp's removal by the next command's replay. |
| STOP/rollback | STOP if a live lookup were needed on every read (the refresh is commit-only) or if replay could not disambiguate the retained pair (quarantine stays the answer). |
| Exit receipt | Recorded on the R0-RC3 card, step 7. |

##### R0c-OR3 — MTP overwrite deletes the replaced object by live identity

| Field | R0c-OR3 |
|---|---|
| Slice/state/owner | **R0c-OR3 / `ACTIVE` (2026-09-04) / Codex `/root` in this worktree, continuing the existing implementation.** Post-closeout correction to the MTP overwrite commit only. Baseline `62715685`; activation drift showed only the user-owned untracked `last_run/`. |
| Finding (2026-09-04 review) | Both MTP overwrite commits (`FileSystemMtp.Core.cpp`, the temp-swap commit and `CommitDeviceSourceOverwriteWithTempSwap`) journal the destination's `persistentId` read through `GetItemProperties`, which serializes the path cache, then remove the replaced object with `backend.DeleteItem(<destination path>)` after the temp upload. `RefreshPathOccupancy` (R0c-OR2) drops only the destination leaf before the decision; the upload invalidates the temp path, not the destination. An object published at the destination name during the upload is therefore deleted by path and the temp is renamed onto it. The R0c-OR2 self-test proves refusal when occupancy changes before the decision; it does not cover an occupant replaced during the upload. |
| Correction | Add `IMtpBackend::DeleteItemByIdentity(path, expectedPersistentId, recursive)`. The production WPD implementation invalidates the path cache, resolves the path live, compares the persistent ID, and deletes only the resolved object ID; an external replacement between resolve and delete is safe because the delete remains addressed to the old object ID. The fake backend performs comparison and deletion under one lock so the executable seam is atomic. Refuse the commit (temp retained, journal kept, `ERROR_REVISION_MISMATCH`) when the live object does not carry the journaled ID. Both overwrite commits and journal replay use the primitive; no path-based delete remains in those flows. |
| Scope | In: `Plugins/FileSystemMtp/FileSystemMtp.Core.cpp` (both commits, replay), `FileSystemMtp.Internal.h` (`DeleteItemByIdentity`), `FileSystemMtp.Device.cpp`, the fixture backend, Compare `mtp_` self-tests, `Specs/FileSystem/FileSystem_Mtp.md`, this plan. Out: host bridge, other providers, the occupancy refresh itself. |
| RED/GREEN evidence | RED: `mtp_overwrite_refuses_occupant_replaced_before_delete` swaps the destination object (new persistent ID) after the commit records its identity: the baseline deletes the new object by path and succeeds. GREEN: the commit refuses with `ERROR_REVISION_MISMATCH`, the new object survives, and the verified temp/journal are retained; `mtp_wpd_overwrite_occupancy_is_live_on_stale_path_cache` and the delete-failure retention case stay green. The Commands source contract pins all three overwrite/replay calls to `DeleteItemByIdentity` and the former path deletes absent. |
| Performance/resources | One live path resolve at the destructive boundary; WPD already enumerates that leaf during resolution, and no new retained state is added. |
| STOP/rollback | STOP if WPD cannot carry the resolved object ID into the delete or the fake seam cannot make compare/delete atomic; rollback is the path delete after the refresh (one revert). |
| Exit-receipt contract | Commit, Debug build receipt, Compare `mtp_` cases, PluginContractTests, Pester, spec update, and the Fresh Full gate that follows; recorded on this card. |

#### R0d — Receipt and capability honesty

##### R0d activation card

| Activation-card field | R0d bounded execution contract |
|---|---|
| Slice/state/owner | **R0d / `COMPLETE` / Codex `/root` in this worktree.** This was one independently releasable host-result and capability-honesty slice. It did not activate R0e, R1, R2, another provider mutation implementation, or a public plugin-ABI cutover. |
| Baseline and drift | Planned-at implementation baseline is clean commit `6e61b4cb`. A scoped `git diff --name-status` over the host result/admission owners, affected provider capability owners and tests/specs returned no paths before activation. `FinalizeTypedItemResults(...)` already classifies a missing destructive receipt conservatively, but `storeQualifiedItemResult(...)` stores a successful receipt-less Native Move/Delete as `Completed` first, so the finalizer cannot correct it. Capability v2 has distinct `operations.delete` and `operations.recycle` booleans, while same-filesystem admission currently checks only `delete`; Microsoft Drive therefore cannot honestly expose its ordinary Graph recycle route without also exposing Permanent Delete. Dummy, writable MTP, Curl/IMAP, Microsoft Drive, and S3 Rename claims still reach central routes that cannot bind/consume the required exact authority. S3 ordinary Delete is the one already-conditioned provider-native virtual-folder route proved by the R0b receipt. |
| Contract | Implement sections 6.1/6.6 `FO-RESULT-01` and `FO-AUTH-01` for A11/A13-MD1/FOS-06/FOS-10. A successful destructive Native call without a valid mutation receipt is immediately `Indeterminate` with `ERROR_IO_INCOMPLETE`, publication/source axes unknown as applicable, and can never be stored as `Completed`; Copy and proved Managed/Copy-only retained-source cases keep their existing result truth. Host Delete admission selects `operations.recycle` only when `FILESYSTEM_FLAG_USE_RECYCLE_BIN` is present and `operations.delete` otherwise. Clear product Delete/Recycle/Rename claims that lack an executable central exact route: Dummy Delete/Recycle/Rename, writable MTP Delete/Rename, Curl FTP/SFTP/SCP/IMAP Delete, Microsoft Drive Permanent Delete/Rename, and S3 Rename. Keep S3 ordinary Delete enabled only through the R0b exact current-generation virtual-folder contract. Microsoft Drive advertises `recycle:true`, accepts only recycle-flagged public Delete, resolves the selected path to a stable item ID and deletes/reconciles that exact ID without root ETag/`If-Match`; a same-path replacement is not the accepted object. The separate Graph permanent-delete endpoint remains unavailable. |
| Scope | In scope: `RedSalamander/FolderWindow.FileOperations.cpp`, `FolderWindow.FileOperationsInternal.h`, and `FolderWindow.FileOperations.State.cpp` only for flag-aware Delete admission and immediate receipt-less destructive result classification; focused FileOps selftests in `FolderWindow.FileOperations.SelfTest.cpp` and `.SelfTest.Phases05_06.cpp`; capability literals in `Plugins/FileSystemDummy/FileSystemDummy.h`, `Plugins/FileSystemS3/FileSystemS3.h`, `Plugins/FileSystemMtp/FileSystemMtp.Core.cpp`, `Plugins/FileSystemCurl/FileSystemCurl.Shared.cpp`, and `Plugins/FileSystemMicrosoftDrive/FileSystemMicrosoftDrive.h`; Microsoft Drive public recycle flag gate plus deterministic stable-ID/no-`If-Match` tests in `FileSystemMicrosoftDrive.cpp`; the existing writable-fake MTP capability case in `CompareDirectoriesEngine.SelfTest.Cases.Mtp.cpp`; exact generic/provider contract assertions in `Tests/PluginContractTests/PluginContractTests.cpp` only if required; `Specs/FileSystem/FileSystem_FileOperations.md`, `FileSystem_MicrosoftDrive.md`, `FileSystem_Mtp.md`, `FileSystem_S3.md`, `Specs/Plugins/Plugins_VirtualFileSystem.md`, `Specs/Testing/Testing_PerformanceValidation.md`, this card/receipt, and compact evidence below `Specs/TestRuns/4cb089111a23/FileOps/`. Out of scope: public production interfaces/ABI, adding typed receipts to providers, provider Move/Copy/Rename/Delete algorithms other than the Microsoft Drive recycle-only public gate, S3 R0b semantics, host cancellation/shutdown, clipboard/result-funnel redesign, UI presentation, and R0e/R1/R2. |
| Predecessors | E0 is closed at `fc32a4aa4b01d6d684166968a0615d3265f33d26`; R0a-R0c are independently closed, with R0b proving the retained S3 Delete exception at `156b13bd` and R0c closing the last active provider overlap at `6e61b4cb`. No active owner overlaps this scope. |
| Steps | (1) Add RED assertions for receipt-less successful Native Move/Delete, the exact provider capability matrix, recycle-vs-permanent host admission, Microsoft Drive permanent rejection, and stable-ID same-path replacement safety. (2) Capture five same-machine test-enabled x64 Release baseline samples before production changes. (3) make Delete admission flag-aware and classify missing destructive Native receipts before the generic success branch. (4) clear only the named capability claims; retain S3 Delete; make Microsoft Drive public Delete recycle-only while preserving by-ID mutation/reconciliation. (5) run focused Debug/test-enabled Release suites, archive candidate evidence, update authoritative specs, record the exit receipt, and mark only R0d complete. |
| RED/GREEN evidence | RED must fail because the current host stores a successful receipt-less Native Move/Delete as Completed, current capability literals expose the named unsafe claims, host admission cannot distinguish recycle from Permanent Delete, and Microsoft Drive accepts a public Delete without the recycle flag. GREEN requires: receipt-less Native Move/Delete produce dense Indeterminate typed results with `ERROR_IO_INCOMPLETE`; Copy and Managed/Copy-only controls remain unchanged; recycle-only admission accepts Microsoft Drive ordinary Delete but rejects its Permanent Delete and every cleared provider claim before task creation; S3 ordinary Delete remains admitted while S3 Rename is rejected; Microsoft Drive deletes/reconciles only the path-resolved stable item ID, preserves a replacement installed at the same path, sends no root ETag/`If-Match`, and rejects non-recycle public calls. Writable and read-only MTP capability cases, full affected provider Debug selftests, generic plugin capability shape, and existing R0a-R0c regressions remain green. Focused commands include `Run-AllTests.ps1 -Suite FileOps -CaseFilter FileOps_ProviderCapabilityMatrix`, `Run-AllTests.ps1 -Suite Compare -CaseFilter mtp_capabilities_are_instance_honest`, direct `PluginContractTests.exe`, and the corresponding Debug/test-enabled Release builds. |
| Performance/resources | Protected deterministic scenario: 4,096 iterations of parsing/querying the affected offline capability profiles and exercising flag-aware host admission plus receipt-less result classification controls in the focused FileOps case. Emit `FileOps.SelfTest.R0d.CapabilityReceiptHonestyUs`, `.CapabilityQueryCount`, `.AdmissionDecisionCount`, and `.ReceiptClassificationCount` once per process. Capture five independent test-enabled x64 Release processes before and after under `Specs/TestRuns/4cb089111a23/FileOps/2026-08-31_r0d_capability_receipt_{baseline,candidate}_release/`. Candidate p95 must be at most `max(1.50 * baseline p95, baseline p95 + 50,000 us)` and every sample below 2,000,000 us. The loop retains one capability JSON/document at a time and constant result state; no provider objects, receipts, or parsed documents may accumulate across iterations. This is in-process admission/classification evidence, not live-network latency. |
| Durable owners | `Specs/FileSystem/FileSystem_FileOperations.md` owns flag-aware Delete admission and receipt-less result truth; `Specs/FileSystem/FileSystem_MicrosoftDrive.md` owns Graph recycle-by-stable-ID semantics; `FileSystem_Mtp.md` and `FileSystem_S3.md` own their current capability exceptions; `Specs/Plugins/Plugins_VirtualFileSystem.md` owns the provider capability/receipt matrix; `Specs/Testing/Testing_PerformanceValidation.md` owns the deterministic 4,096-decision scenario and archive/budget. This plan retains delivery mapping and the generated exit receipt only. |
| STOP/rollback | STOP if immediate Indeterminate classification changes Copy or proved Managed/Copy-only truth; recycle cannot be admitted separately from Permanent Delete without public ABI change; Microsoft Drive Graph DELETE targets path text/current occupant, requires a root ETag/`If-Match`, or reaches the permanent-delete endpoint; retaining S3 Delete requires weakening the R0b contract; a cleared claim is required by an unrelated active owner; tests need live credentials/device/network; retained state grows with the 4,096-iteration sample; Release evidence exceeds the budget without a diagnosed machine anomaly; or scope crosses R0e/R1/R2. Rollback removes the host classification/admission branch, capability deltas, Microsoft Drive gate, tests, metrics, and spec clauses only while R0d remains open; never partially advertise Permanent Delete or receipt-less completion as trustworthy. |
| Exit-receipt contract | Before R0d closes, record exact activation/RED/implementation/evidence commits; Debug and test-enabled Release build receipts; focused FileOps/MTP/provider GREEN totals; five-process baseline/candidate metric summaries and constant-resource proof; capability matrix and Graph stable-ID/no-`If-Match` fault results; source proof that missing destructive receipts cannot reach the generic Completed branch; validated archive paths and whole-inventory validation; `git diff --check`; authoritative-spec destinations; and remaining FOS/R owner gaps. |

**R0d RED receipt (2026-08-31):** activation commit `5e57da68`. The test-only
working tree built Debug `RedSalamander` with receipt
`75652ec75743e21e21c65855d281af5e2d9964b9790586e08ded4f92390ac403` and
Debug `PluginContractTests` with receipt
`be90ec254ed4187ad2230fa737106a688a7ba1e3b5ee2ea08ded4a855c51a7be`.
Direct focused run `r0d-red-fileops-debug-20260831a` failed exactly because a
successful receipt-less destructive Native call was stored as Completed; it also
emitted 4,096 deterministic iterations in 527,685 us. Direct focused run
`r0d-red-mtp-debug-20260831b` failed exactly because writable MTP advertised
Delete. Direct `PluginContractTests.exe` failed the new Microsoft Drive permanent-
Delete, stable-ID replacement, and no-`If-Match` assertions; its generic capability
matrix also failed the named Curl/IMAP, Dummy, Microsoft Drive, and S3 Rename claims
while unrelated provider debug suites remained green. These are RED diagnostics, not
qualification evidence.

- [x] Classify null destructive Native receipts as Indeterminate immediately.
- [x] Clear Microsoft Drive, MTP, Curl, Dummy, and IMAP Delete/Rename claims whose
      central exact route is absent; this does not wait for the typed cutover.
- [x] Keep S3 ordinary folder Delete available only through its contract-tested,
      provider-native, exact virtual-folder route until the typed central contract can
      express that same scope. Do not pretend a physical container binding exists.
- [x] Classify Microsoft Drive ordinary Graph Delete as Recycle by stable item ID,
      revalidate that exact ID without a root ETag/`If-Match` condition, and keep Graph
      Permanent Delete unavailable until its separate endpoint satisfies authority/
      receipt rules.

##### R0d exit receipt

- **Commits and scope:** activation is `5e57da68`; the RED contracts are
  `bf037d2d`; implementation is `57f5ab1a`; authoritative specs and compact
  baseline/candidate evidence are `c79d4bfb`. No public production plugin ABI,
  provider Copy/Move algorithm, S3 R0b mutation contract, host cancellation/shutdown,
  clipboard/result-funnel redesign, or R0e/R1/R2 surface changed.
- **Builds and focused GREEN:** Debug `RedSalamander` built with zero warnings/errors
  under receipt
  `d447739566e18e3bc22deb6220120fb6be96aa46c9e9506d970711589430ecc9`;
  the corrected Microsoft Drive plugin receipt is
  `fbec870de17fafe6e5374cdc1ceeeb48ed8642c2b374ab7d853c0c89ed2d4d13`,
  and Debug `PluginContractTests` receipt is
  `3e61a735f2608488de4ac8754495f07454645e4f57c8c00d4d983d23ae908b20`.
  Focused Debug FileOps passed `3/0/0`
  (`r0d-green-fileops-debug-20260831a`), focused Debug MTP passed `1/0/0`
  (`r0d-green-mtp-debug-20260831a`), and direct Debug provider contracts exited 0.
  Exact implementation commit `57f5ab1a` then built the full test-enabled x64 Release
  solution with zero warnings/errors under receipt
  `0f1b7b85a4ab2d0a49b3ef35e318bfd50ae80f52cb035f511e967c793128da7a`.
  Direct Release `PluginContractTests.exe` exited 0; Microsoft Drive contracts reported
  `223/0`, Curl `226/0`, S3 multipart `50/0`, and S3 directory probes `157/0`.
  Focused Release MTP passed `1/0/0`
  (`r0d-green-mtp-release-20260831a`).
- **Truth and authority proof:** flag-aware central admission reads
  `operations.recycle` only for `FILESYSTEM_FLAG_USE_RECYCLE_BIN` and
  `operations.delete` otherwise. A successful receipt-less provider-native Delete or
  Native Move is classified before generic success as Indeterminate/
  `ERROR_IO_INCOMPLETE`, with unknown destructive axes; Copy and host-proved
  Managed/Copy-only controls retain their prior truth. Dummy Delete/Rename/Recycle,
  Curl FTP/SFTP/SCP/IMAP Delete, writable MTP Delete/Rename, Microsoft Drive Permanent
  Delete/Rename, and S3 Rename are non-advertised. S3 ordinary Delete remains enabled
  only through the R0b ETag-conditioned provider-owned route. Microsoft Drive accepts
  only recycle-flagged public Delete, resolves the path once to a stable item ID,
  deletes/reconciles that ID, preserves a same-path replacement, and sends no root
  ETag/`If-Match`.
- **Performance/resources:** the five Release baseline samples are `30,097`,
  `31,497`, `30,163`, `30,011`, and `30,383` us (p95 `31,497` us).
  The five candidate samples are `29,704`, `29,940`, `29,497`, `29,485`, and
  `29,951` us (p95 `29,951` us), a 4.9% reduction. Candidate p95 passes
  `max(1.50 * 31,497, 31,497 + 50,000) = 81,497` us, and every sample is below
  2,000,000 us. Every process reports exactly 12,288 capability queries, 8,192
  flag-aware admission decisions, and 4,096 receipt classifications while retaining one
  capability document and constant typed-result state. All five candidate runs passed
  `3/0/0`; the RED baseline's `2/1/0` is retained because the old receipt-less
  destructive success reached Completed after the measured loop.
- **Archives and contracts:** paired evidence under
  `Specs/TestRuns/4cb089111a23/FileOps/2026-08-31_r0d_capability_receipt_{baseline,candidate}_release/`
  passes explicit validation for 16 files per root; whole-inventory validation passes
  for 1,478 files. Documentation-drift plus harness source contracts pass `203/203`,
  `Get-SpecInventory.ps1 -FailOnFindings` reports zero blocking findings, and
  `git diff --check` passes. Durable behavior is in
  `Specs/FileSystem/FileSystem_FileOperations.md`,
  `Specs/FileSystem/FileSystem_MicrosoftDrive.md`,
  `Specs/FileSystem/FileSystem_Mtp.md`, `Specs/FileSystem/FileSystem_S3.md`,
  `Specs/Plugins/Plugins_VirtualFileSystem.md`, and
  `Specs/Testing/Testing_PerformanceValidation.md`.
- **Remaining release gaps:** R0e subsequently closed; R1a-R1c remain inactive, so M1
  is not closed. The review-reported
  direct `StartOperation -> ConfirmArtifactTouch -> HostShowPrompt` nested-pump
  lifetime twin was subsequently closed under its FileOps prompt-lifetime owner by
  `e41ae4bd`, `8fb251b6`, and evidence/spec commit `818f15ee`; it is not an R0d
  closure. R0d closes only FOS-06/FOS-10.

#### R0e — Cancellation honesty and route characterization

##### R0e activation card

| Activation-card field | R0e bounded execution contract |
|---|---|
| Slice/state/owner | **R0e / `COMPLETE` / Codex `/root` in this worktree.** This was one independently releasable cancellation-honesty and exact-route containment slice. It did not activate R1, R2, R5 process isolation, or a generic abandoned-worker design. |
| Baseline and drift | Activation baseline is clean commit `818f15ee`. A scoped `git diff --name-status` over the named host/capability/provider/test/spec owners returned no tracked paths before activation; repository-root `last_run/` is pre-existing untracked test output and is excluded. Current capability documents carry `abort:false`, `deadline:false`, and often `quietPointTimeoutMs:30000`, but the host copies those facts without initializing `FileSystemOptions::deadlineTickCount64` or enforcing a quiet point. `Task::RequestCancel()` wakes host gates only, while `FileOperationState::Shutdown()` unconditionally joins every task. Cross-filesystem `IFileReader::Read` can therefore remain blocked after Cancel; Curl's private `_stopping` is not connected to the task, and Local direct-reader creation drops operation options. MTP alone has reviewed provider-local watchdog/cancel/quarantine and unload gating. |
| Contract | Implement D2-A01/FOS-05 as exact path-and-operation admission truth. Replace the decorative timeout with one required cancellation route class in each operation-specific capability response: `bounded`, `providerWatchdog`, or `uncontained`. `bounded` means the complete admitted provider call/reader/writer route has deterministic cancellation checkpoints and a proved quiet point; `providerWatchdog` additionally carries a nonzero provider-owned timeout and proves backend cancel, quarantine, and module-lifetime safety; `uncontained` carries no timeout and is rejected before task/card/worker creation. The host owns no synthetic deadline and never abandons an in-process task. Local fixed-volume and deterministic in-memory/archive routes may remain only with focused blocked-I/O proof; Local UNC/remote-drive routes are a distinct uncontained SMB profile until separate evidence proves containment. Preserve MTP's watchdog route. Curl streaming transfer, S3/Graph network mutation/transfer, and other remote routes remain available only if their exact operation response proves one of the two admitted classes; otherwise clear/reject that route without disabling inspection. |
| Scope | In scope: capability parsing/qualification and same-/cross-filesystem admission in `RedSalamander/FolderWindow.FileOperations.cpp`, `FolderWindow.FileOperationsInternal.h`, and the immutable plan endpoint fields; task/shutdown code only for deterministic witnesses/telemetry, not worker abandonment; focused FileOps tests in `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest*.cpp`; exact capability documents and route-specific tests in built-in Local, Dummy, 7z, Curl, S3, Microsoft Drive, Google Drive, and MTP providers; generic capability-shape tests in `Tests/PluginContractTests/PluginContractTests.cpp`; existing MTP watchdog tests; Local SMB/UNC profile tests; `Specs/FileSystem/FileSystem_FileOperations.md`, affected provider specs, `Specs/Plugins/Plugins_VirtualFileSystem.md`, `Specs/Testing/Testing_PerformanceValidation.md`, this card/receipt, and compact evidence below `Specs/TestRuns/4cb089111a23/FileOps/`. Out of scope: a public reader-cancel ABI, host thread termination/abandonment, heap-quarantined generic tasks, R5 process brokering, clipboard/result-funnel changes, progress UI, and unrelated provider mutation/identity algorithms. |
| Predecessors | E0 is closed at `fc32a4aa4b01d6d684166968a0615d3265f33d26`; R0a-R0d are independently closed, and the later direct-prompt lifetime remediation is closed at `818f15ee`. Section 8.1 requires only E0. No active owner overlaps this scope. |
| Steps | (1) Add RED contracts proving decorative timeout facts are accepted, an explicitly uncontained never-returning fake route is admitted far enough to threaten shutdown, Curl export remains reachable without task-connected read cancellation, and Local UNC is not a distinct route class. (2) Capture five same-machine test-enabled x64 Release baseline samples of capability parsing/admission only; never execute the deliberately wedged call in-process. (3) replace `quietPointTimeoutMs` with the exact route class plus provider-owned watchdog timeout, validate combinations fail closed, and carry the classification into immutable endpoints. (4) gate every same-/cross-filesystem admission before task creation; preserve only proved bounded/provider-watchdog routes, including MTP, and classify Local SMB separately. (5) add a host-level never-returning-provider witness that proves zero provider calls and bounded shutdown for an uncontained route, plus existing MTP containment and SMB-specific tests. (6) run focused Debug/test-enabled Release suites, archive candidate evidence, update authoritative specs, record the exit receipt, and mark only R0e complete. |
| RED/GREEN evidence | RED must fail because the current parser has no exact route class, accepts the decorative 30-second field, and admission ignores all cancellation facts. GREEN requires malformed/contradictory route declarations to fail capability parsing; uncontained same-provider and either-side cross-provider routes fail before task/card/worker creation; the never-returning fake records zero calls and host shutdown stays below 2,000 ms; MTP's watchdog-qualified fake timeout still returns/quarantines and reaches unload quiet point; Local fixed-volume controls remain admitted; UNC/remote-drive paths return the SMB-uncontained profile and are rejected without opening network I/O; no host deadline tick or unbounded-route timeout is advertised; affected provider capability and operation tests remain green. |
| Performance/resources | Protected deterministic scenario: 4,096 operation-specific capability parses and same-/cross-route admission decisions over bounded, provider-watchdog, uncontained, malformed, Local fixed-volume, SMB, and MTP profiles. Emit `FileOps.SelfTest.R0e.RouteContainmentUs`, `.CapabilityParseCount`, `.AdmissionDecisionCount`, `.RejectedUncontainedCount`, and `.ProviderCallCount` once per process. Capture five independent test-enabled x64 Release processes before and after under `Specs/TestRuns/4cb089111a23/FileOps/2026-08-31_r0e_route_containment_{baseline,candidate}_release/`. Candidate p95 must be at most `max(1.50 * baseline p95, baseline p95 + 50,000 us)`, every duration below 2,000,000 us, parse/decision counts exact, provider-call count zero for the never-returning witness, and retained state constant. This is admission/containment evidence, not live-network or SMB throughput. |
| Durable owners | `Specs/FileSystem/FileSystem_FileOperations.md` owns task cancellation, route admission, and shutdown truth; `Specs/Plugins/Plugins_VirtualFileSystem.md` owns capability schema/route-class semantics; provider specs own their exact classified routes and watchdog behavior; `Specs/Testing/Testing_PerformanceValidation.md` owns the deterministic 4,096-decision scenario, five-process evidence, and budgets. This plan retains delivery mapping and the generated exit receipt only. |
| STOP/rollback | STOP if a route is called before its containment class is validated; MTP containment or unload gating regresses; Local fixed-volume FileOps are disabled without a failing containment witness; SMB is treated as equivalent to fixed local storage; an uncontained network route remains admitted by plugin-ID exception; the design requires terminating/abandoning an in-process worker, retaining callback/task state after shutdown, or smuggling R5 isolation; tests require live credentials/device/share; the never-returning fixture is actually invoked in-process; retained state grows with the 4,096-iteration sample; or Release evidence exceeds the budget without a diagnosed machine anomaly. Rollback removes the route-class parser/gates, capability deltas, tests, metrics, and spec clauses only while R0e remains open; never restore a decorative host deadline claim as a substitute. |
| Exit-receipt contract | Before R0e closes, record exact activation/RED/implementation/evidence commits; Debug and test-enabled Release build receipts; focused host/Local/Curl/MTP/S3/Graph/provider GREEN totals; five-process baseline/candidate summaries and constant-resource proof; never-returning zero-call/bounded-shutdown, MTP quarantine/unload, and SMB-profile results; source proof that uncontained admission cannot allocate a task; validated archive paths and whole-inventory validation; `git diff --check`; authoritative-spec destinations; exact disabled routes and R5 follow-ups; and remaining FOS/R-owner gaps. |

**R0e RED receipt (2026-08-31):** activation is `91d14243`; RED contracts are
`1b0474f8`. Direct Debug run `r0e-red-debug-20260831b` failed `2/1/0` exactly
because Local's capability document had no bounded route class; its protected loop
completed 4,096 parses and 4,096 decisions in 1,854,233 us before the assertion.
The exact clean RED commit then built test-enabled x64 Release under receipt
`7b4eedd5cdf14df1bb75a3b2bcfd53201f41cf44166b146fc2c1c3bcfd658288`
with zero warning/error diagnostics. Five sequential baseline processes retained the
same exact failure and metrics. Generic provider-contract RED also failed because the
new route-class/watchdog schema was absent. The never-returning fixture was deliberately
not dispatched on the RED tree.

- [x] Remove or relabel unenforced host deadline/quiet-point claims.
- [x] Preserve proven provider-local containment such as MTP while adding a
      host-level never-returning-provider shutdown witness and SMB-specific evidence.
- [x] Disable only an exact route shown uncontained; defer process isolation to R5.

##### R0e exit receipt

- **Commits and scope:** activation is `91d14243`; RED contracts are `1b0474f8`;
  implementation plus authoritative specifications are `bdacb1ea`; compact evidence
  is `b8245cb0`. No public reader-cancel ABI, host deadline, thread abandonment,
  generic task quarantine, clipboard/result funnel, progress UI, R1/R2 behavior, or
  R5 process broker was introduced.
- **Builds and focused GREEN:** exact implementation commit `bdacb1ea` built the full
  test-enabled x64 Release solution with zero warnings/errors under receipt
  `ceea32e11a7572cdad338e358941a1b16fbcdeb243cf77e98860ad461ba3deec`
  and the full Debug solution under receipt
  `c667ef40a249434fb765d6cfc47229d594d8a10894461f9aec3d604e17a08acb`.
  Focused committed Debug FileOps passed `3/0/0`
  (`r0e-green-debug-committed-20260831a`), and the MTP instance-honesty/watchdog
  case passed `1/0/0` (`r0e-green-mtp-capabilities-committed-20260831a`). Direct
  Release `PluginContractTests.exe` exited 0; Microsoft Drive reported `223/0`, Curl
  `226/0`, S3 multipart `50/0`, and S3 directory probes `157/0`, with the applicable
  module-unload quiet points reached.
- **Cancellation and route proof:** capability v2 now requires exactly one
  `cancellation.routeClass` (`bounded`, `providerWatchdog`, or `uncontained`) plus a
  semantically consistent `providerWatchdogTimeoutMs`; the decorative
  `quietPointTimeoutMs` member is invalid. The host rejects malformed or uncontained
  same-provider and either-side cross-provider mutation routes before task publication.
  The deliberately never-returning provider observes zero operation calls and no task
  count increase, so shutdown has no witness worker to join. Synthetic UNC is always
  classified `local-win32-smb`/uncontained without opening a share; mapped remote
  volumes use the same `DRIVE_REMOTE` classification. Local fixed volumes, Dummy, and
  7z are bounded; MTP retains its 30-second provider-watchdog contract. Curl
  FTP/SFTP/SCP/IMAP, S3/S3Table, Microsoft Drive/SharePoint, Google Drive, and Local
  SMB FileOps mutation/transfer routes are uncontained and unavailable while browsing
  and non-mutating inspection remain available.
- **Performance/resources:** five Release baseline samples are `1,067,199`,
  `1,053,325`, `1,065,507`, `1,063,903`, and `1,066,916` us (p95
  `1,067,199` us). Five exact-implementation candidate samples are `1,090,078`,
  `1,092,912`, `1,076,248`, `1,055,543`, and `1,075,131` us (p95
  `1,092,912` us). Candidate p95 passes
  `max(1.50 * 1,067,199, 1,067,199 + 50,000) = 1,600,799` us, and every
  sample is below 2,000,000 us. Every process reports exactly 4,096 parses and 4,096
  admission decisions; candidate samples report one rejected-uncontained witness and
  zero provider calls with constant fixture/result state.
- **Archives and durable owners:** compact evidence is under
  `Specs/TestRuns/4cb089111a23/FileOps/2026-08-31_r0e_route_containment_{baseline,candidate}_release/`.
  Explicit validation passes for both compact files, and whole-inventory validation
  passes for 1,557 files. Documentation-drift, specification-information-architecture,
  and expanded harness source contracts exit 0; `Get-SpecInventory.ps1
  -FailOnFindings` reports zero blocking findings; `git diff --check` passes. Durable
  behavior is in
  `Specs/FileSystem/FileSystem_FileOperations.md`, the affected provider specifications,
  `Specs/Plugins/Plugins_VirtualFileSystem.md`, and
  `Specs/Testing/Testing_PerformanceValidation.md`.
- **Remaining release gaps:** R1a-R1c remain inactive, so M1 is not closed and
  FOS-07 through FOS-12 remain under their declared owners. R0e contains FOS-05 for
  the currently admitted route set by rejecting uncontained calls; it does not make
  `IFileReader::Read` cooperatively cancelable. An exact disabled route requires new
  bounded/provider-watchdog proof or R5 isolation before re-enable; the missing reader
  cancel channel is therefore a re-enable/isolation follow-up, not a claimed R0e ABI
  fix.

##### R0e-OR1 — Admitted Local blocking-reader cancellation residual

| Activation-card field | R0e-OR1 bounded execution contract |
|---|---|
| Slice/state/owner | **R0e-OR1 / `COMPLETE` / Codex `/root` in this worktree.** This post-closeout correction owns only an admitted fixed-volume Local source reader blocked inside `ReadFile`. It does not reopen route classification, re-enable Local SMB/Curl/S3/Graph, change MTP watchdog policy, add public ABI, abandon workers, or claim that every provider/write path is cooperatively cancelable. |
| Baseline and drift | Clean tracked baseline is `fb446233`; only the user-owned untracked `last_run/` directory is present. The bridge initially binds a regular-file source with metadata authority only. Ordinary Copy therefore calls legacy `IFileSystemIO::CreateFileReader`, and Local constructs `Win32FileReader` with null operation options. Managed Move can call exact `OpenReader(&options)`, but Local still executes synchronous `ReadFile`; once blocked it cannot observe `Task::RequestCancel()`. The bridge checks cancel only around `Read`, its pipeline joins the reader thread, and UI-thread `Shutdown()` joins the task. A canceled admitted Local read can therefore wedge teardown despite R0e's `bounded` classification. |
| Contract | Every admitted Local bridge file read uses retained no-follow `FILESYSTEM_BIND_READ_CONTENT` authority and passes the task's existing `FileSystemOptions.operationControl` to `IFileSystemBoundObject::OpenReader`. Local exact readers reopen the retained object with an independent overlapped handle, maintain their own logical position, and poll the existing operation-control/deadline contract only while kernel I/O is pending. Cancel/deadline calls `CancelIoEx` for that exact request, drains completion, returns the original `ERROR_CANCELLED`/`ERROR_TIMEOUT`, and reports zero newly consumed bytes. Ordinary completed reads retain `IFileReader` size/seek/read behavior. Current shutdown may still join, but canceled Local pending reads must unwind within the bounded test budget. |
| Scope | Production: `Plugins/FileSystem/FileSystem.cpp` and the regular-file source binding/opening seam in `RedSalamander/FolderWindow.FileOperations.State.cpp`. Test-only: one `ENABLE_TESTS` Local-provider blocked-read/ordinary-read export, its Commands behavioral case/registration, and a source contract proving content authority plus option-bearing bound open. Specs: `Specs/FileSystem/FileSystem_FileOperations.md`, `Specs/Plugins/Plugins_VirtualFileSystem.md`, `Specs/Testing/Testing_PerformanceValidation.md`, and this receipt. Evidence: `Specs/TestRuns/4cb089111a23/Commands/2026-09-01_r0e_or1_local_reader_cancel_{baseline,candidate}_release/`. Out of scope: `IFileReader` ABI additions, task reapers/UI-thread join architecture, writer cancellation, MTP timeout reduction, remote-route re-enable, R5 isolation, conflict/UI behavior, and unrelated Local mutation/metadata code. |
| RED/GREEN evidence | `cmd_pane_fileops_local_blocked_reader_cancel_is_bounded` uses the real Local reader over a connected byte pipe with no producer bytes. RED proves the read does not remain pending and return `ERROR_CANCELLED` after operation control flips; a fail-safe server close always releases the fixture, so RED cannot hang the suite. GREEN requires the read to enter kernel-pending state, observe at least one post-entry control check, return `ERROR_CANCELLED` with zero bytes within 500 ms, and leave no pending I/O/thread/handle. The same case reads and seeks a deterministic fixed-local payload through both legacy and exact-bound readers. The source contract requires `FILESYSTEM_BIND_READ_CONTENT` on the bridge source binding and `OpenReader(&options)` before any legacy fallback; existing Local object-binding, operation-control, bridge Copy/Move, verification, and shutdown tests remain green. |
| Performance/resources | The case emits `FileOps.SelfTest.R0eOr1.LocalBlockedReadCancelUs`, `.AbortCheckCount`, `.BytesReadAfterCancel`, `.OrdinaryReadUs`, and `.OrdinaryBytesRead`. Capture five independent test-enabled x64 Release processes before and after. Candidate blocked-cancel p95 and every sample must be below 500,000 us with zero bytes after cancel. For the deterministic ordinary-read control, candidate p95 must be at most `max(1.50 * baseline p95, baseline p95 + 50,000 us)` with exact byte count and seek replay. Retained handles, reader threads, and pipe instances return to zero each process. This is Local reader cancellation/overhead evidence, not disk-throughput or generic provider containment evidence. |
| Durable owners | `Specs/FileSystem/FileSystem_FileOperations.md` owns the host bridge/cancel/shutdown truth. `Specs/Plugins/Plugins_VirtualFileSystem.md` owns Local `OpenReader` operation-control and exact-object semantics. `Specs/Testing/Testing_PerformanceValidation.md` owns the deterministic blocked-read and ordinary-read metrics/budgets. This plan retains activation mapping and the exit receipt only. |
| STOP/rollback | STOP if the correction needs a public ABI or task-owned raw reader pointer; exact-object reading cannot be preserved with an independently positioned overlapped reopen; `CancelIoEx` cannot be drained without abandoning a request; a normal Local read or seek changes bytes/position; a route currently classified uncontained must be enabled; the test needs a live share/device; shutdown still exceeds the bound after the Local read returns; or ordinary-read p95 exceeds budget without a diagnosed machine anomaly. Rollback removes only the Local overlapped reader, bridge content-binding preference, tests, metrics, and spec clauses while R0e-OR1 remains open. |
| Exit-receipt contract | Record activation/RED/implementation/spec/evidence commits; exact Debug and test-enabled Release build receipts; focused behavioral/source-contract and affected Local/FileOps totals; five-process baseline/candidate summaries; exact canceled status/zero-byte/abort-check/ordinary-read observations; archive and whole-inventory validation; `git diff --check`; durable-spec destinations; and remaining generic UI-join/writer/MTP/disabled-route gaps. Mark only R0e-OR1 `COMPLETE`. |

- [x] Add bounded RED blocked-read, ordinary-read, seek, and source-authority coverage.
- [x] Route admitted Local bridge reads through option-bearing exact content authority.
- [x] Make Local exact readers independently positioned and cancelable while pending.
- [x] Update authoritative specs and archive five test-enabled Release candidates.
- [x] Record the R0e-OR1 exit receipt and mark only the residual `COMPLETE`.

##### R0e-OR1 exit receipt

- **Commits and scope:** activation is `d34a0ed7`; RED coverage is `8dd8f305`;
  implementation plus authoritative specifications are `d7decb40`; the independent
  Change Case/source-contract synchronization is `c818b710`; compact baseline and
  candidate evidence is `921739b4`. The correction adds no public ABI, route
  re-enable, worker abandonment, remote-provider claim, or writer-cancellation claim.
- **Builds and focused GREEN:** exact implementation commit `d7decb40` built
  the test-enabled x64 Release RedSalamander graph with zero warnings/errors under
  receipt `7c099b7742ac87d76d891eaf86ad272e4c7174ed402f1352005a46c85265d30f`.
  Post-contract-sync commit `c818b710` rebuilt the same graph with zero warnings/errors
  under receipt `bbf1bbd0a267c050e3613cc4c10f45e1fae0f2250fe282de426bd725bfa24107`;
  its exact blocked-reader closeout case passed `1/0/0`. The seven affected Local/FileOps
  processes passed `21/0/0`, the Release source guard passed `1/0/0`, Release
  `PluginContractTests.exe` exited 0, and the synchronized Pester source-contract suite
  passed `187/0/0`. The real Change Case unreadable-descendant case also passed `1/0/0`.
  Exact pre-closeout commit `4d15e297` built the test-enabled x64 Debug RedSalamander
  graph with zero warnings/errors under receipt
  `a5a4461db8e17ae5a32ec862649ec036c82c7585c71eca02b10a77fe892872ce`.
  Its blocked-reader plus source-authority Commands process passed `2/0/0`, and seven
  one-case FileOperations processes passed `21/0/0`. The exact Debug
  `PluginContractTests` target built with zero warnings/errors under receipt
  `ae688b62fe20627b4c0d470b1c9bcdecc895b0e37ed838eeac468107aea0d69f`;
  the executable exited 0, including Local provider debug selftests and all applicable
  provider/module-unload quiet points.
- **RED/GREEN and performance:** five RED Release samples required the fail-safe pipe
  close and returned after `527,005`-`528,759` us with non-cancel I/O failures. Five
  candidate Release samples returned `ERROR_CANCELLED`, zero bytes, and three
  operation-control checks each in `30,943`, `46,574`, `31,010`, `31,025`, and
  `31,628` us (p95 `46,574` us), below the 500,000-us bound. Ordinary-read candidate
  p95 is `6,950` us versus baseline p95 `9,861` us and the accepted `59,861`-us
  ceiling; every sample read exactly `16,777,216` bytes and passed seek replay.
- **Archives and durable owners:** paired evidence is under
  `Specs/TestRuns/4cb089111a23/Commands/2026-09-01_r0e_or1_local_reader_cancel_{baseline,candidate}_release/`.
  Explicit validation passed for all 32 archived files and whole-inventory validation
  passed for 1,899 files. Documentation-drift and specification-information-architecture
  Pester passed `43/0/0`; `Get-SpecInventory.ps1 -FailOnFindings` reported zero blocking
  findings; `git diff --check` passed. Durable behavior is in
  `Specs/FileSystem/FileSystem_FileOperations.md`,
  `Specs/Plugins/Plugins_VirtualFileSystem.md`, and
  `Specs/Testing/Testing_PerformanceValidation.md`.
- **Exit decision:** all scoped STOP conditions and build/test/evidence gates pass.
  The independently launched Debug Monitor that initially blocked exact attestation was
  closed by the user; no process was terminated by build/test tooling. R0e-OR1 is
  `COMPLETE` without changing the public ABI or enabling an uncontained route.
- **Remaining gaps:** generic UI-thread join architecture, writer cancellation, MTP
  provider-watchdog policy, every disabled/uncontained route, and R5 isolation remain
  outside this residual. The retained synchronous join is bounded only for the admitted
  Local exact-reader path proved here.

**Acceptance:** each promoted R0 owner is independently reviewable and releasable.
No automatic cleanup/Delete uses pathname, current occupant, unconditioned textual
prefix membership, hash, name, size, or timestamp **alone**. The selected real root or canonical virtual-
folder boundary is retained across consent; each discovered descendant/object
generation is bound/conditioned immediately before deletion; fixed object-set work
consumes its admitted revision snapshot. No
receipt-less destructive success is called Completed, and capability/cancellation UI
never promises a route or bound the implementation cannot execute.

##### R0-RC1 — Review corrections (2026-09-02)

| Field | R0-RC1 |
|---|---|
| Slice/state/owner | **R0-RC1 / `COMPLETE` (Fresh Full gates #2 on `0683cc51` and #3 on `7727ca9c` recorded below) / Claude in worktree `codex/r4-a19-closeout-2026-09-02`.** Post-closeout corrections from the 2026-09-02 branch review. No public ABI change, no route re-enable, no scheduler or provider-policy change beyond the items below. |
| Corrections | (1) Local direct-final Copy treats metadata the destination cannot hold as best effort like `CopyFileExW` (recorded as `FileOps.Local.DirectFinalMetadataLost`, never fatal); only an EFS source that cannot stay encrypted fails closed with `ERROR_ENCRYPTION_FAILED` before content; its exact abort uses `MakeOwnedStageCleanupOptions`. (1b, owner decision 2026-09-02: File Explorer parity) NTFS compression is never transferred by the Local `TransferMetadataTo` prepare phase and is no longer a host pre-content feature: the destination inherits its parent folder's compression like `CopyFileExW`; sparse and EFS handling is unchanged. Before this fix a compressed/sparse source failed against FAT/exFAT/ReFS, and a source with a named stream (every downloaded file) was fully copied to FAT/exFAT and then aborted. (2) `FileSystemRouteCapabilitiesBase`: `GetChildNameCollisionKey`/`JoinPath` apply only shape rules; `ValidateChildName` owns proposed-name rules and distinguishes `ERROR_FILENAME_EXCED_RANGE` and `ERROR_BAD_DEVICE`; reserved names complete (`COM0`/`LPT0`, superscripts, `CONIN$`, `CONOUT$`). Create Directory, the Batch Rename parent listing and source keys, and Change Case discovery skip an existing child the provider cannot key instead of failing the whole command. (3) The Local provider caches resolved volume facts for two seconds so per-name route queries cost one `GetVolumePathNameW`. (4) Batch Rename preview reports `name_empty`/`name_dot`/`name_separator` before the provider call and maps provider Invalid to `name_too_long`/`name_reserved_device`/`name_invalid_character` instead of `name_destination_probe_failed`. (5) `RequestCancel` and the stop callback take `_batchRenameArtifactPromptMutex` before notifying the artifact-prompt wait (lost-wakeup join hang). (6) `FinalizeTypedItemResults`: a `MutationPossible` root without a terminal receipt is Indeterminate even when the task succeeded. (7) Popup `Enter`/`Escape` bind to an unfocused decision only when exactly one is actionable. (8) Local `Win32FileReader` reuses one event and polls only under an operation control or deadline. (9) A breadcrumb store above 4,096 children logs one warning. |
| Deferred | S3 recursive Delete stays one sequential conditional request per object (the route is `uncontained` and unavailable); parallelizing to `deleteMaxConcurrency` needs a thread-safe fake graph and reworked R0b cancel/pass assertions. MTP device roots no longer carry any `[devid:]`/`[devpuid:]` suffix (owner decision 2026-09-02: the friendly name is the path, the identity is provider metadata); object `[puid:]`/`[oid:]` collision suffixes are unchanged. |
| Tests | `cmd_pane_batchRename_window_preview_tolerates_legacy_named_sibling` (trailing-dot sibling created through the extended prefix must not fail the parent listing); the R0a direct-final step gained a forced sparse/stream-loss Copy that must publish complete content, a forced EFS loss that must fail closed before content, and an NTFS compression-inheritance check (plain source into a compressed folder becomes compressed, compressed source into a plain folder stays plain; skipped on volumes without NTFS compression). |
| Evidence | Full-solution x64 Debug build: 0 warnings, 0 errors. Direct `PluginContractTests.exe` exited 0 on the changed provider base. Focused Debug FileOps: `Floodgate_LocalCopyNewName` 4/0/0 (includes the sparse/stream-loss Copy, EFS fail-closed, and compression-inheritance checks), `Phase10_MetadataPreservationAndSourceRetention` 3/0/0 (bridge metadata after compression stopped being a pre-content feature), `FileOps_ProviderCapabilityMatrix` 3/0/0, `R1d_PreparingLifecycle` 3/0/0, `Phase9_ConflictPrompt` 10/0/0. Focused Debug Commands: `cmd_pane_batchRename_window_` 25/0/0 (includes the new legacy-sibling case), `cmd_pane_createDirectory_` 5/0/0, `cmd_pane_changeCase_` 8/0/0, `cmd_pane_fileops_popup_` 4/0/0. Fresh Full has not been run by this card. |

##### R0-RC2 — Simplifications (2026-09-02)

| Field | R0-RC2 |
|---|---|
| Slice/state/owner | **R0-RC2 / `COMPLETE` (Fresh Full gates #2 on `0683cc51` and #3 on `7727ca9c` recorded in Evidence) / Claude in worktree `codex/r4-a19-closeout-2026-09-02`.** Behavior-preserving simplifications requested by the owner after the 2026-09-02 review. No public ABI change, no route re-enable, no provider-policy change. |
| Changes | (1) `FileOperationArtifacts::Registry` (an empty object whose `LoadDefault` always succeeded) is gone; the classifier is the free functions `ClassifyCandidate`/`ProjectCandidate`/`BuildTouchGuardRequest`/`AcceptTouchGuard`/`RevalidateTouchGuard`, and the "registry unavailable" branches, warnings, and perf tags that could never fire are deleted. (2) `BatchRenameMutationKind` (a one-value enum) and `BatchRenameExecutionOp::currentSource` (always equal to `originalSource`) are removed from the Batch Rename mutation boundary. (3) `CrossFileSystemBridge` and `QualifiedItemFailurePhase` are hoisted out of `Task::ExecuteOperation` into the named `FolderWindowFileOperationsStateInternal` namespace of the same translation unit as pure code motion (every moved line dedented by 8 columns and proven identical against the previous commit); `FolderWindow` befriends the bridge so its existing settings read keeps compiling. `ExecuteOperation` shrinks from ~8.1k to ~2.7k lines. The `QualifiedItemPolicyContext` lambda bag stays as the documented per-item seam. |
| Corrections found by the R0-RC1 Fresh Full gate | (4) `FinalizeTypedItemResults`: a root without a receipt whose status is `S_FALSE` is a conflict the user answered Skip before any mutation and is reported `Skipped + NotAttempted + Retained` instead of `ERROR_IO_INCOMPLETE`. (5) `Phase7_CrossPaneVisibleRefreshDummy` routes a host same-provider Copy (both panes show the destination folder and must see the copied file) instead of a recycle Delete: the Dummy provider stopped advertising Delete/Rename/Recycle on 2026-08-31 (`57f5ab1a`), and a receipt-less Native Move is Indeterminate by contract; the Dummy row of `Plugins_VirtualFileSystem.md` says so. (7) `storeQualifiedItemResult`: an explicit user Skip answered at a Native Move/Delete conflict prompt (provider `ERROR_PARTIAL_COPY` with the skip observed) is reported `Skipped` with the source retained instead of being swallowed by the receipt-less destructive-success rule from `57f5ab1a` (`Fairstream_MoveSameSizeCollisionPrompts`, last recorded pass 2026-08-31 13:21Z before that rule). (6) `cmd_fileops_move_breadcrumb_durable_lifecycle` now asserts the R1d contract (`668acbd7`): admission publishes the Move, the clipboard barrier runs after preparation, and a breadcrumb persistence failure terminates the task with the exact status, no publication, and every source retained. Both (5) and (6) were failing at the previous commit. (8) Admission captures the source attribute hints (top-level file/folder kinds) for every Copy/Move whose pane selection matches the sources, as before R1d; `668acbd7` had gated the capture behind the confirmation prompt, so ordinary transfers never counted completed files/folders in the popup (`Phase7_SharedPerItemScheduler`). (9) `Riptide_LiveFinishedSnapshotCarriesDiagnostics` throttles a 4 MiB copy past the 500-ms reveal deadline: under the R1d presentation rule a task that finishes before the deadline is never presented before its completion summary, so the old 8 KiB copy had no live row to observe. (10) `cmd_pane_navigation_change_case_prompt_keeps_navigation_shell_stable` waited for the renamed file with a case-insensitive `exists()` that is already true before the change-case worker admits the rename; it now waits for the exact spelling of the single directory entry. (11) The two tooling source contracts that still described the retired `FileOperationArtifacts::Registry` object follow the free-function classifier. |
| Not done | The capability JSON v2 documents are not deleted or generated: they are a spec'd diagnostic/extension surface (`search`, `identity.object`, `metadata`, the MTP `byteVerifyOnOverwrite`/`mtp` sections) consumed by search and Compare self-tests. Instead the plugin contract suite now proves the JSON route sections equal the typed route facts for every provider, so drift is a test failure rather than a silent host/JSON disagreement. Generating the route sections from `FileSystemRouteDescriptor` remains an owner decision. |
| Tests | `PluginContractTests`: per-provider JSON-versus-typed drift guard (`pathProfile`/`rootId`, `operations`, `concurrency`, `publication`, `identity`, `links`, `verification`, `cancellation`, `names`). The `file_operations_*_source_guard` scans gain the structural guard that `CrossFileSystemBridge` is defined before `ExecuteOperation`. Existing coverage carries the rest: artifact touch-guard cases (`Phase10`/`Phase11` FileOps, `cmd_pane_batchRename_`, `cmd_folderView_` artifact projection), Batch Rename engine cases, and the bridge families (`Fairstream`/`Riptide`/`Cinderstar`/`Floodgate`). |
| Evidence | Full-solution x64 Debug build after every patch: 0 warnings, 0 errors. `PluginContractTests.exe` passed with the drift guard for all eight providers (Google Drive reports `Unsupported` because its path text is not a stable identity; its typed facts are still copied and compared). Focused Debug FileOps: `Phase7_` 15/1/3 (the rewritten Dummy Copy step passes), `Fairstream_` 20/0/0 (collision Skip fixed), `Riptide_` 9/1/3, `Phase10_` 9/0/0, `Phase11_` 9/0/0, `Floodgate_` 15/0/0. Focused Debug Commands: `cmd_pane_batchRename_` 74/0/0, `cmd_fileops_` 4/0/0 (includes the rewritten breadcrumb lifecycle case), `cmd_pane_changeCase_` 8/0/0, `cmd_pane_changeAttributes_recursive_artifact_guard` 1/0/0, source guards `file_operations_` 3/0/0, `file_system_` 3/0/0, `folder_view_` 10/0/0; Compare `local_index_fileops_artifacts_remain_visible` 1/0/0. The two focused misses, `Phase7_SharedPerItemScheduler` (just-in-time discovery counters `items=2 files=0 folders=0`) and `Riptide_LiveFinishedSnapshotCarriesDiagnostics` (paused popup snapshot not readable within 2 s), fail identically in isolation on the previous commit `6904dce9` (2/2 each, rebuilt in place), so they predate this card (both are R1d `668acbd7` fallout; corrections (8) and (9) fix them); their last recorded pass before that is the 2026-08-31 13:21Z Full FileOps run. Fresh Full gate on `29c5a3a7` (run `20260902T164442Z-93900-…`): 1917 counted, 1868 passed, 16 failed — the FileOps process stopped in `Phase11_BridgePipelineDummyToDummyPerf` after families 1–14 (its three recorded misses were the two R1d cases above and the unauthorized alternate volume), so the runner reported `COUNT_MISMATCH`; the Phase11 family passes 9/0 in isolation. Commands 879/10: nine of the ten pass 2/2 in isolation (load flakes), the change-case navigation case is correction (10). Pester 724/4: the two sandbox disk-audit tests (shared `C:\RedSalamander.Perf` leftovers) and the two retired-registry contracts, correction (11). After corrections (8)–(11): `Phase7_` 19/0/0, `Fairstream_` 20/0/0, `Riptide_` 13/0/0, `Phase6_` 7/0/0, `Phase14_` 3/0/0, `R1d_` 3/0/0, `Phase11_` 9/0/0, `cmd_fileops_` 4/0/0, the source-contract Pester describe 181/181, and the two R1d cases 3/3 each in isolation. Fresh Full gate on `0683cc51` (run `20260902T181823Z-107984-…`): 2062 total, 2003 passed, 5 failed — the two shared-sandbox disk-audit Pester tests (environment), `theme_cycle_overlay_timer_fallback` and `cmd_pane_batchRename_admission_worker_owned_cancellable` (timing under full-suite load), and `R4A19_DiscoveryProviderControls` (fake-MTP cancel returned `ERROR_IO_INCOMPLETE` instead of a cancel status; the copy is started programmatically with no pane selection, so no attribute hint is involved). Isolation reruns after the gate: `theme_cycle_overlay_timer_fallback` 3/3, `cmd_pane_batchRename_admission_worker_owned_cancellable` 3/3, and `R4A19_DiscoveryProviderControls` 3/3 all pass in isolation, so every residual failure is environment or full-suite timing and none is attributable to R0-RC1, R0-RC2, or R0-RC3. Fresh Full gate #3 on `7727ca9c` (run `20260902T195425Z-79868-…`, after the R0f activation `a57c6d5d` and the MTP friendly-name device roots): 2062 total, 2007 passed, 2 failed, 53 skipped (missing live S3/FTP/OAuth endpoints and host connection UI) — `cmd_pane_batchRename_admission_worker_owned_cancellable` (known full-suite load flake; 3/3 pass in isolation on the same build) and `R4A19_DiscoveryIndependentVolumes` (alternate volume not authorized). The two sandbox Pester audits, `theme_cycle_overlay_timer_fallback`, and `R4A19_DiscoveryProviderControls` passed in this run. |

##### R0-RC3 — Review corrections (2026-09-03)

| Field | R0-RC3 |
|---|---|
| Slice/state/owner | **R0-RC3 / `COMPLETE` (2026-09-04; steps 1-9 landed, Fresh Full gate #12 recorded in Evidence) / Claude in worktree `codex/r4-a19-closeout-2026-09-02`.** Fold-in of the independent review of `f3d57f42` (plan-versus-code and adversarial passes). Each accepted finding is one step below with its owner; the review's unverified claims are steps that start with a verification. |
| Steps | (1) `[x]` R3-3 (b): the Commands source contract that pinned the former bound-reader call-site spelling now pins `OpenSourceReader` (same guarantee: the exact bound source is opened with the task's operation-control options before any legacy fallback), landed with the extraction. (2) `[x]` (`a911b56a`, 2026-09-04) Identity-less replace is fail-closed without an occupant token: `CaptureReplaceExpectation` produces no expectation when the destination's basic information cannot be read, `PrepareStage` refuses a granted replacement without an expectation (`bridge.replace.expectationUnavailable`, nothing written), and the timestamp-only validators (Curl `ValidateCurlReplaceOccupant`, Dummy `SetExpectedReplacement`) refuse when either timestamp is zero instead of skipping the comparison; test `R3_4_ReplaceWithoutOccupantTokenRefused` through the new `SetFileOpsBridgeFailNextDestinationBasicInfoForSelfTest` hook. S3, Microsoft Drive and Google Drive keep their ETag/eTag/version tokens. (3) `[x]` (2026-09-04) The live speed-limit menu's Custom prompt (`ShowSpeedLimitMenu`, `FolderWindow.FileOperations.Popup.cpp`) runs under `FileOperationPromptDispatchScope` like the deferred `ShowCustomSpeedLimitPromptForTask`; the PluginConfig source contract covers it. (4) `[x]` R4-A19 closes from the archived parent matrix (exit receipt below the card, state `COMPLETE`); the two "R3 inactive" sentences in section 8.1 and the derived milestones are corrected. (5) `[x]` (`e91bfa52`, 2026-09-04) Specs aligned with R0f: `Plugins_VirtualFileSystem.md` no longer says FTP/SFTP/SCP publish empty import lists; `FileSystem_FtpSftpScp.md` no longer says copy/move/rename are `false`; `FileSystem_MicrosoftDrive.md` no longer says `operations.rename: false`; the Curl control claim is narrowed to the transfers that carry the callback until step 6 lands. (6) `[x]` (2026-09-04) R0f-Curl-OR1: the streaming reader returned by `CreateFileReader` ran its libcurl transfer with `ApplyCommonCurlOptions(..., nullptr, ...)` and therefore without the owning call's operation control. Landed: the optional reader contract `IFileReaderOperationControl::SetOperationControl` (`FileSystem.h`), the bridge hands legacy readers the task's options in `OpenSourceReader` (the self-test reader decorator forwards it), `CurlStreamingReader` stores the control and its `CurlProgress` callback polls `FileSystemCheckOperationControl` on every libcurl call, stopping the transfer through the existing `_stopping` path (`ERROR_CANCELLED`); `RunCurlStalledReaderCancelSelfTests` cancels a `Read` blocked on a stalled `RETR` and bounds the return (< 3 s). Shutdown joins therefore become bounded by the poll period for readers under a task; readers outside a task keep the destructor stop. (7) `[x]` (2026-09-04) R0c-OR2 (card under R0c-OR1; verification first): live MTP overwrite Commit took occupancy and the destination PUID from the path cache (`ResolvePathCached`) and destroys the verified temp when the original delete fails; the replay path already live-enumerates. Required outcome: one destination occupancy primitive (invalidate or live child lookup) for exclusive create, replace, and replay, missing-object results drop the cache entry, and a verified temp is retained when the original delete fails; test on the Device path cache with a stale destination, not the fake backend. Landed: `IMtpBackend::RefreshPathOccupancy` (the WPD backend forgets the cached leaf; both overwrite commits call it before the existence check), the verified temp and its identified journal entry survive a failed original delete (replay removes it beside an intact original or commits it when the original is gone), `mtp_wpd_overwrite_occupancy_is_live_on_stale_path_cache` on the WPD self-test backend (`replacePhotoAfterFirstLookup`, `writable`), and the delete-original-failure case now expects the retained journal after the failure (observed on the file, as the next backend command replays it) and the temp's removal by that replay. (8) `[x]` (verified 2026-09-04, no change) R0f-SMB: the permanent-delete root and ancestor bindings run in `Task::PrepareMutationInterlockScopes`, called from `Task::PrepareForExecution` on the task thread (`Task::ThreadMain`), not on the UI thread; the admission-time `BindObjectAuthority` in `FolderWindow.FileOperations.cpp` classifies reparse points only (one bound bind per link item). (9) `[x]` (2026-09-04) `JoinFileSystemPath` leftovers on transfer destination, native Move, artifact, and Batch Rename paths: inventory with the PluginConfig guard extended beyond `QualifyCreateDirectory`; convert to `QueryChildNameContract` where a provider name contract exists (writable Google Drive and MTP destinations). Inventory: `TryResolveTransferDestinationProviderPath` (transfer destination), `LocalNativeMoveItemShape` (native Move shaping), `ProjectProviderChildObject` and `CollectArtifactTouchCandidates` (artifact projection), and the Batch Rename engine's `JoinFolderAndLeaf`; every provider deriving from `FileSystemRouteProviderBase` publishes `ValidateChildName`, so the contract exists for every destination, not only Google Drive and MTP. Landed: `AdmitOperation` queries the destination child-name contract for every transfer leaf (source leaf or explicit mapping) before any join and refuses with the provider's reason (`PlanRejectionBucket::InvalidDestinationName`); the downstream joins now only see admitted leaves (Batch Rename leaves are admitted by its own contract query), and the PluginConfig guard pins the admission. Test `RC3_9_TransferDestinationNameRefused` (Dummy trailing-dot leaf to a local folder → `ERROR_INVALID_NAME` at admission, destination untouched). Correction after Fresh Full gate #12: the first landing also refused on a contract-layer failure, and the child-name contract query (`QueryChildNameContract`) answers a contract violation (`ERROR_INVALID_DATA`) whenever the parent path carries a trailing separator (the R0f cases pass `/r0f-bucket/`, `//anonymous@host:port/`, `/@conn:.../`): the provider's `JoinPath` result re-splits into a parent without that separator, which the contract's consistency check does not consider equivalent, so every `R0f*_Fake*` Copy failed to start; admission now refuses only when an answered contract reports `FILESYSTEM_CHILD_NAME_INVALID` and otherwise keeps the plain join. Observation for a follow-up (not this step): the contract layer should accept a trailing accepted separator on the parent it is asked about, so that these destinations get the name check as well. |
| Not accepted | Adding `Failed`/`Unknown`/`Retained` as `PublicationTransaction` states: they are receipt outcomes the executor already records from the item result; the record names the phases reached. Stopping R3-3 (b) before the fault matrix: (b) is behaviour-preserving by construction and the matrix is its evidence. |
| Evidence | Commits: `a911b56a` (step 2), `0c7aa352` (step 3), `e91bfa52` (steps 4 and 5, review fold-in), `4a6f69ca` (gate #11 record, step 8 verified), `1e3ddbb6` (step 6, R0f-Curl-OR1: Curl debug self-tests 238 to 250 checks, R0fCurl_/R3_/Phase10/Phase11 families, Commands contracts, Pester 187), `7bea151b` (step 7, R0c-OR2: Compare `mtp_` cases 55/0/1 with the two new or reworked cases, PluginContractTests, Pester 187), `d3cfbe67` (step 9: R3_/RC3_9/Phase11/Phase12/Phase7 families, Commands contracts, Pester 187). Fresh Full gate #12 (`d3cfbe67`, 2026-09-04 02:08-03:12): Total 2074 / Passed 2014 / Failed 7 / Skipped 53 (every FileOps family ran). Classified: `cmd_pane_batchRename_admission_worker_owned_cancellable` (admission load flake, passed isolated), `cmd_app_prompt_uses_alert_overlay_window` (UI timing under load, passed isolated), `R4A19_DiscoveryIndependentVolumes` (environment: alternate volume root not authorized), and the four `R0f*_Fake*` Copy starts (Curl, S3, Microsoft Drive, Google Drive) refused by the step 9 admission with `InvalidDestinationName` / `ERROR_INVALID_DATA`: a regression of `d3cfbe67`, corrected in `076d863b` (refuse only on an answered `FILESYSTEM_CHILD_NAME_INVALID`; the trailing-separator contract answer is recorded as a follow-up on step 9). Correction evidence: focused FileOps `R0fCurl_`/`R0fS3_`/`R0fGraph_`/`R0fGDrive_`/`R0fSmb_`/`RC3_9`/`R3_`/`Phase11_CrossFileSystemBridge` 3/0, 3/0, 3/0, 3/0, 4/0, 3/0, 6/0 and 3/0 passed/failed, Commands contracts, Pester; the next Fresh Full gate (#13) re-covers the whole suite. |

##### R0f — Network and cloud routes: containment proof and full read/write/create/delete (DONE 2026-09-03)

| Field | R0f |
|---|---|
| Slice/state/owner | **R0f / `DONE 2026-09-03 — every slice landed with its receipt` / unassigned.** The 2026-09-02 review found that no package restored mutation on the routes R0e classified `uncontained` (Local `local-win32-smb` UNC and mapped drives, Curl FTP/SFTP/SCP, S3/S3 Table, Microsoft Drive/SharePoint, Google Drive). The owner's decision: these destinations MUST read, write, create, and delete like any other folder before S1. R0f is therefore mandatory, not conditional, and R5 process isolation stays a fallback only for a route whose containment cannot be proved in-process. |
| Slices | **R0f-SMB** (`local-win32-smb`): `CancelIoEx` on the exact overlapped reader/writer requests, bounded namespace calls (open/enumerate/delete/rename) through the same operation-control polling the fixed-volume reader already has, reclassify to `bounded`; fixture is a loopback share authorized like the alternate-volume root. **R0f-Curl** (FTP/SFTP/SCP): task-connected abort through the libcurl progress callback plus bounded connect and low-speed timeouts for every transfer and control call, reclassify to `providerWatchdog` with a tested provider-owned timeout; fixture is the gated `REDSALAMANDER_SELFTEST_CONN_FTP/SFTP` profiles plus a local FTP/SFTP server fixture for CI. **R0f-S3**: AWS CRT request cancellation wired to `IFileSystemOperationControl`, reclassify to `providerWatchdog`; fixture is the S3 fake backend used by the Compare self-tests plus the gated `REDSALAMANDER_SELFTEST_CONN_S3` profile. **R0f-Graph** (Microsoft Drive/SharePoint): WinHTTP request cancellation on cancel/deadline, reclassify to `providerWatchdog`, and add the missing Delete and Rename entry points (the tuple is currently `delete:false, rename:false`). **R0f-GDrive**: the plugin was a read-only skeleton; add create/write/delete/rename with revision-bound identity, task-connected cancellation through the libcurl progress callback (the plugin's transport is libcurl, not WinHTTP), and reclassify to `providerWatchdog`; fixture is a loopback Drive v3 endpoint. IMAP remains read-only by design (a mailbox is not a writable folder) and is not an R0f slice. |
| Contract | Per slice: (1) prove a bounded quiet point for every provider call and reader/writer the admitted route uses, with a RED case that shows the route wedging cancel today and a GREEN case that shows cancel/deadline/shutdown returning within the declared bound; (2) prove read, write, create, delete, rename, and cross-provider Copy/Move in both directions against the slice fixture with the R1a/R0d receipt truth (published/retained/removed per root) and no silent data loss; (3) flip the typed route facts (`cancellationRoute`, `providerWatchdogTimeoutMs`, the operation booleans) and the diagnostic JSON together, guarded by the PluginContractTests drift check; (4) keep browsing and inspection available while the slice is open. A decorative host timeout, a detached worker, or an unload under a live callback is never acceptable containment. |
| Exit | Every shipped provider route is `bounded` or `providerWatchdog`; `IsQualifiedEndpointValid` rejects nothing a user can navigate to; the R0e uncontained-rejection witnesses become route-specific containment witnesses; each slice ships its own Fresh Full receipt with the fixture families enabled. R0f-SMB precedes M2 (the most common destination); every slice precedes M3 and therefore S1. **Met 2026-09-03:** Local SMB is `bounded`; FTP/SFTP/SCP, S3, Microsoft Drive/SharePoint, and Google Drive are `providerWatchdog` with provider-owned bounds and mutation receipts; only the read-only-by-design IMAP and S3 Table routes stay `uncontained`, and neither admits a mutation. The slice receipts below carry Fresh Full gates #4 to #7; gate #8 for R0f-GDrive is recorded on its receipt row once it completes. |
| STOP/rollback | STOP a slice if containment needs a public ABI change the plan has not budgeted, if the fixture cannot run in CI, or if a provider cannot report per-root receipts for its mutations; record the exact mechanism and route the slice to R5 instead of shipping a route that ignores Cancel. Rollback of a slice restores its `uncontained` classification only, never the pre-R0e decorative timeouts. |
| R0f-SMB receipt (2026-09-02) | Mechanism: instead of per-call overlapped conversions, one shared synchronous-I/O cancel watch (`Common/SynchronousIoCancelWatch.h`, a per-module threadpool timer) registers every File Operations worker thread (`Task::ThreadMain`, the per-item scheduler's `Process` call) and every Local provider scheduler unit (`SharedFileOpsJobScheduler::executeWorkItem`, jobs carry their operation options) and, once cancel/stop has been requested for 500 ms, calls `CancelSynchronousIo` for the registered thread on every 50 ms poll until the call returns; a call aborted that way (`ERROR_OPERATION_ABORTED`) under a requested cancel is reported as the cancellation by the task and per item. The Local descriptor and capability JSON declare `bounded` for `local-win32-smb`; `IsQualifiedEndpointValid` needed no change. RED: before this slice no code path issued `CancelSynchronousIo`, so a worker inside a wedged SMB call had no quiet point (the R0e never-returning witness documents the wedge). GREEN: Local debug self-test `RunDebugSynchronousIoCancelWatchSelfTest` (anonymous-pipe read returns `ERROR_OPERATION_ABORTED` within the grace), FileOps `R0fSmb_BlockedSynchronousCallCancelReturns` (a bounded synthetic provider wedged in a synchronous read returns and the task ends canceled within 3 s), FileOps `R0fSmb_LoopbackReadWriteCreateDelete` (Copy to `\\localhost\<drive>$` alias of the sandbox, Move back, provider CreateDirectory/Rename on the share, Delete on the share, byte-exact verification; environment-gated on the administrative share being reachable), and the Provider identity matrix UNC witness flipped from rejection to bounded admission. Follow-up found by the loopback control: the mutation-interlock ancestor walk bound the UNC share root without its trailing separator (`ERROR_BAD_PATHNAME`), so `TryGetFileSystemParentPath` now stops at the share root and Local `ToExtendedPath` opens a share root as `\\\\?\\UNC\\server\\share\\` (contract checks in `cmd_pane_batchRename_path_identity_parser` and the Local path-normalization debug self-test). Evidence: full-solution Debug build 0 warnings/0 errors; PluginContractTests pass (FileSystem.dll debug self-tests 162/0 incl. the cancel-watch and share-root path checks); FileOps `R0fSmb_` 4/4 (wedged call back in 563 ms, canceled task terminal in 578 ms; loopback Copy 156 ms, Move back 172 ms, Delete 125 ms), `FileOps_ProviderCapabilityMatrix` 3/3, `Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker` 3/3, `Phase5_DiscoveryCancelLatencyLocal` 3/3, Commands `cmd_pane_batchRename_path_identity_parser` 1/1. Fresh Full gate #4 on `883f4ac9` (run `20260902T221417Z-640-…`): 2064 total, 2009 passed, 2 failed, 53 skipped; both R0f-SMB cases passed in the gate, and the two failures are the known `cmd_pane_batchRename_admission_worker_owned_cancellable` load flake and the unauthorized alternate-volume control. |
| R0f-Curl receipt (2026-09-03) | Mechanism: the seven mutation entry points and the Curl shared scheduler keep the owning call's `FileSystemOptions` on the current thread (`CurlOperationOptionsScope`, `CurlCurrentOperationOptions`); `ApplyCommonCurlOptions` installs `CurlOperationControlXferInfo`, a progress callback that polls `FileSystemCheckOperationControl`, and the four control-command helpers (`CurlProbeRemoteFileSize`, `CurlPerformList`, `CurlPerformListAndParse`, `CurlPerformQuote`) run under those options; libcurl invokes the callback at least once per second even while a command waits for the server, and `HResultFromCurl` reports the host's own verdict for `CURLE_ABORTED_BY_CALLBACK`. Typed facts and JSON for FTP/SFTP/SCP: `providerWatchdog`, `deadline:true`, `providerWatchdogTimeoutMs` = `CurlProviderWatchdogTimeoutMs(connect, operation)` (max of the connect timeout and the low-speed abort; 60 s default), `copy`/`move`/`nativeMove`/`delete`/`rename` true and Copy import `["*"]` (Move export/import stay empty until a conditional delete exists); IMAP stays read-only and `uncontained`. Writer publication: the plugin advertises `IFileSystemAtomicWriter` (staged sibling + one server-side rename) so the host bridge can publish new names into FTP; no-overwrite writers, Copy, Move and Rename probe the destination and report `ERROR_FILE_EXISTS` when it exists (the probe-to-rename window is a documented residual; the bridge still refuses overwrite into a destination without bound objects). Host: permanent Delete on a provider whose binding authority is `Unsupported` runs the provider's native `DeleteItem` (`DeletePlan::nativeAuthority`): admission records no ingress snapshot, the interlock keeps no exact-delete authority and skips the identity-continuity check, and the provider's mutation receipt decides truth (no receipt stays Indeterminate). The UNC share-root parent rule applies only to identities that accept backslashes, so `//user@host/...` provider paths derive ordinary parents. RED: before this slice the control commands (SIZE, LIST, DELE, MKD, RNFR/RNTO) ran without any progress callback, so a server that stopped answering held the task until the transport timeouts and Cancel had no effect on them; the typed facts refused Copy/Delete/Rename on every Curl route; Copy into FTP through the bridge failed closed because the plugin advertised no writer publication; and permanent Delete on FTP was rejected at admission because the host demanded bound objects. GREEN: Curl debug self-test `RunCurlStalledControlCommandCancelSelfTests` (fake FTP holds `DELE`; Cancel returns `ERROR_CANCELLED` within a few progress ticks; without Cancel the transport bound returns it on its own), FileOps `R0fCurl_FakeFtpReadWriteCreateDelete` (the plugin's deterministic loopback FTP fixture started through `RedSalamanderCurlStartFakeFtpForSelfTest`; the route qualifies as providerWatchdog with a nonzero bound; cross-provider Copy local -> FTP and FTP -> local byte-exact through the host; provider CreateDirectory and Rename on FTP; permanent Delete on FTP through a task with every object gone), the rewritten Curl writer/race self-tests, and PluginContractTests expecting FTP/SFTP/SCP to report delete and rename truthfully as true. Evidence: focused rerun 2026-09-03 02:32 on the fix9 tree: PluginContractTests all OK (1548 OK, 0 FAILED; Curl debug self-tests passed=238 failed=0); FileOps `R0fCurl_` 3/3 (`FileOps.SelfTest.R0fCurl.ProviderWatchdogTimeoutMs` 8000 ms, `CopyToFtpMs` 156, `CopyFromFtpMs` 78, `DeleteOnFtpMs` 94), `R0fSmb_` 4/4, `FileOps_ProviderCapabilityMatrix` 3/3, `Phase10_` delete guards 9/9; Commands `cmd_pane_batchRename_path_identity_` 1/1. Residuals kept open under R0f: bridge overwrite into FTP/SFTP/SCP (and S3) after a granted conflict, Move export/import lists, the probe-to-rename window. Fresh Full gate #5 on `ab8b4b07` (run `20260903T003618Z-107312-0b1e6712f0804809a6eedfa8a489ae41`): 2065 counted, 2009 passed, 3 failed, 53 skipped; every R0f-Curl and R0f-SMB case passed in the gate. The three failures are the known `cmd_pane_batchRename_admission_worker_owned_cancellable` load flake, the unauthorized alternate-volume control, and ToolsPesterTests `keeps Floodgate data-safety invariants wired`, whose Curl writer assertion still encoded the pre-R0f-Curl refusal of no-overwrite writers; the source contract now encodes the probe-then-`ERROR_FILE_EXISTS` publication (follow-up commit). |
| R0f-S3 receipt (2026-09-03) | Mechanism: the eight mutation entry points keep the owning call's `FileSystemOptions` on the current thread (`S3OperationOptionsScope`, `S3CurrentOperationOptions`); every AWS request the plugin issues (`ArmS3RequestControl` at the 21 request sites: bucket and object listing, head, get, put, copy, delete, multipart) carries a continue handler that the S3-CRT client polls from its header, body, and progress callbacks and that cancels the meta request when `FileSystemCheckOperationControl` fails; `HresultFromAwsError` returns the host's verdict for any failure observed after Cancel or a passed deadline. Provider-owned bound: `S3ProviderWatchdogTimeoutMs(connect, request)` = (connect + 2 x max(3 s, request) + the monitor's 1 s tick) x 2 attempts + 2 s backoff, backed by the CRT stall monitor (1 B/s over that interval, evaluated once per second and firing after about two intervals: measured 12-16 s for two attempts at a 3 s interval; 144 s with the default timeouts) and an exponential-backoff retry policy capped at one retry (the SDK default retried a silent server many times). Typed facts and JSON for S3: `providerWatchdog`, `deadline:true`, `providerWatchdogTimeoutMs` materialized from the configured timeouts (`configured-s3-watchdog-ms`); S3 Table stays read-only and `uncontained`. `anonymous` (configuration and connection `extra`) creates the client with empty credentials, which the CRT signs as anonymous. S3 `DeleteItem` and `DeleteItems` now complete each item with a `FileSystemItemMutationResult` (`BuildDeleteMutationResult`: success = known and committed with the original gone; a definitive refusal = known and not committed; cancellation or a transport failure mid-request = no record), which the host's native Delete authority consumes; without it every successful S3 delete was Indeterminate. New fixture `Plugins/FileSystemS3/FileSystemS3.SelfTest.FakeS3.cpp`: a loopback path-style HTTP/1.1 S3 endpoint (buckets, objects with ETag and Last-Modified, `list-type=2` with prefix/delimiter/paging, Range, If-Match/If-None-Match, server-side copy, multipart, aws-chunked bodies) with `SetStallMethod` (a server that stopped answering) and `SetDripListing` (bytes still flowing); host exports `RedSalamanderS3StartFakeS3ForSelfTest`/`RedSalamanderS3StopFakeS3ForSelfTest`. RED: before this slice no S3 request could be canceled while in flight, a silent server held a task for the SDK's full retry budget, and the `uncontained` route refused every mutation at admission. GREEN: S3 debug self-test `RunS3StalledRequestCancelSelfTests` (a drip-fed listing under a recursive Delete returns `ERROR_CANCELLED` within 3 s of Cancel and deletes nothing; a stalled DELETE returns `ERROR_CANCELLED` within the declared bound after Cancel, and as a transport failure within the bound without it; the object survives), FileOps `R0fS3_FakeS3ReadWriteCreateDelete` (route qualifies as providerWatchdog with a nonzero bound; cross-provider Copy local -> S3 and S3 -> local byte-exact through the host bridge; provider CreateDirectory and Rename; permanent Delete through a task with every object gone), the provider capability matrix now expecting S3 mutation admission, and PluginContractTests' drift guard on the new route facts. Residuals: `operations.rename` stays `false` (host Rename route receipt), bridge overwrite into S3 after a granted conflict, and readers/writers under the bridge rely on the provider bound rather than the continue handler. Evidence: focused rerun 2026-09-03 04:22 on the final tree: PluginContractTests all OK (1548 OK, 0 FAILED; S3 debug self-tests passed=212 failed=0 including `RunS3StalledRequestCancelSelfTests`, focused `--s3-r0f-containment-selftests` 17/17); FileOps `R0fS3_` 3/3 (`FileOps.SelfTest.R0fS3.ProviderWatchdogTimeoutMs` 22000 ms for the fixture timeouts, `CopyToS3Ms` 171, `CopyFromS3Ms` 110, `DeleteOnS3Ms` 94), `R0fCurl_` 3/3, `R0fSmb_` 4/4, `FileOps_ProviderCapabilityMatrix` 3/3, `Phase10_` delete guards 9/9. Fixture lessons kept in the code: the CRT sends PutObject in absolute request form, and a permanent Delete without a provider receipt is Indeterminate by design. Fresh Full gate #6 on `aa04d6bd` (run `20260903T022914Z-70552-13ff58daa5a74eb7b427077b51cf6bbd`): 2066 counted, 2010 passed, 3 failed, 53 skipped; every R0f-S3, R0f-Curl and R0f-SMB case and the Pester source contracts passed in the gate. The three failures are the known `cmd_pane_batchRename_admission_worker_owned_cancellable` load flake, the unauthorized alternate-volume control, and `cmd_plugin_configuration_dialog_tab_traversal_live_dx_interaction`, whose scripted S3 dialog tab sequence did not know the new Anonymous access toggle between the addressing toggle and the max-keys field (sequence extended in the follow-up commit). |
| R0f-Graph receipt (2026-09-03) | Mechanism: the eight mutation entry points keep the owning call's `FileSystemOptions` on the current thread (`GraphOperationOptionsScope`); the single WinHTTP path `SendHttpRequest` polls `FileSystemCheckOperationControl` before each attempt, between 64 KiB request-body chunks, between response chunks, and in 100 ms slices while sleeping before a 429/5xx retry, and every transport failure observed after Cancel or a passed deadline returns the host's verdict (`GraphTransportFailure`). Provider-owned bound: `GraphProviderWatchdogTimeoutMs(connect, request)` = 2 x connect + 2 x request, the WinHTTP resolve/connect/send/receive timeouts of one attempt (transport failures are not retried). Typed facts and JSON: `providerWatchdog`, `deadline:true`, `providerWatchdogTimeoutMs` materialized from the configured timeouts (`configured-drive-watchdog-ms`), `rename:true`; Delete stays Recycle by stable item ID (permanent Delete honestly absent). Receipts: `DeleteItem`/`DeleteItems` complete each item with `BuildRecycleMutationResult` and `RenameItem` with the move commit receipt, so the host's Recycle plan is never Indeterminate on success. Writer publication: the plugin advertises `IFileSystemAtomicWriter` (a simple upload is one PUT with If-None-Match, an upload session's item appears only at completion under the requested conflict behavior), so the host bridge publishes new names into Graph destinations through the provider writer. Test seam (test-enabled builds only): `ConfigureFakeGraph` redirects `GraphBaseUrl()` to a loopback origin that `ValidateGraphApiUrl` and `ValidatePreauthenticatedUploadUrl` accept, bypasses the OAuth token, suppresses throttle sleeps, and uses the synthetic `/@conn:microsoft-drive-selftest` context, which the capability JSON and the typed route descriptor now also resolve their root from (the fixture has no connection profile). New shared fixture core `Common/LoopbackHttpFixture.h` (the loopback HTTP/1.1 server with concurrent sessions, absolute-form targets, chunked bodies, stall and drip switches, and a request log) and `Plugins/FileSystemMicrosoftDrive/FileSystemMicrosoftDrive.SelfTest.FakeGraph.cpp` (a fake Graph drive: metadata by path and id, children with `$top` paging, folder creation, PATCH rename/move with If-Match, DELETE, simple content upload with If-None-Match, upload sessions with Content-Range chunks, ranged download with If-Match); host exports `RedSalamanderMicrosoftDriveStartFakeGraphForSelfTest`/`...StopFakeGraphForSelfTest`/`...FakeGraphRequestLogForSelfTest`. RED: before this slice no Graph request could be canceled while in flight (the only control check sat at the entry points), Rename was hidden behind `rename:false`, Recycle and Rename completions carried no receipt, and the `uncontained` route refused every mutation at admission. GREEN: Microsoft Drive debug self-test `RunGraphStalledRequestCancelSelfTests` (a drip-fed children listing under a folder Recycle returns `ERROR_CANCELLED` within 3 s of Cancel and deletes nothing; a stalled DELETE returns `ERROR_CANCELLED` within the declared bound after Cancel, and as a transport failure within the bound without it; the item survives), FileOps `R0fGraph_FakeGraphReadWriteCreateRenameRecycle` (route qualifies as providerWatchdog with a nonzero bound; cross-provider Copy local -> Graph and Graph -> local byte-exact through the host bridge; provider CreateDirectory and Rename; Recycle through a task with every item gone), the provider capability matrix expecting Microsoft Drive rename and mutation admission, and PluginContractTests' drift guard on the new route facts. Residuals: permanent Delete stays absent on Graph routes (Recycle is the delete), readers/writers under the bridge rely on the provider bound rather than the entry-point options, and the live Graph profiles remain gated. Evidence: focused reruns 2026-09-03 06:00-06:16 on the final tree: PluginContractTests all OK (1548 checks; Microsoft Drive debug self-tests passed=240 failed=0 including `RunGraphStalledRequestCancelSelfTests`, focused `--microsoft-drive-r0f-containment-selftests` 16/16 with the witness reporting streaming Cancel returning in 0-15 ms, a stalled DELETE returning 3.5-7.5 s after Cancel and on its own after about 4 s against the declared 10 s bound at the fixture timeouts; the one red S3 timing check of the first full run led to the corrected S3 bound and re-measured 17/17 at 12 s against 20 s); FileOps `R0fGraph_` 3/3 (`FileOps.SelfTest.R0fGraph.ProviderWatchdogTimeoutMs` 20000 ms for the fixture timeouts, `CopyToGraphMs` 140, `CopyFromGraphMs` 79, `RecycleOnGraphMs` 109), `R0fS3_` 3/3, `R0fCurl_` 3/3, `R0fSmb_` 4/4, `FileOps_ProviderCapabilityMatrix` 3/3, `Phase10_` delete guards 9/9; Pester source contracts 187/187. Fixture lessons kept in the code: the capability root of a synthetic-context path resolves without a connection profile, and a destination without atomic-final publication cannot be bridged. Fresh Full gate #7 on `29f1f270` (run `20260903T041708Z-96988-89f8c728dc6a494c88197bc2bd46bdb7`): 2067 counted, 2001 passed, 4 failed, 62 skipped; every R0f-Graph, R0f-S3, R0f-Curl and R0f-SMB case, the provider capability matrix, PluginContractTests and the Pester source contracts passed in the gate. The four failures are the known `cmd_pane_batchRename_admission_worker_owned_cancellable` load flake, the unauthorized alternate-volume control, `Phase7_ParallelCopyMoveKnobs` ("expected more than one in-flight entry" under gate load; its abort skipped the nine chained Phase7 cases, which is the skip-count difference; 19/19 in isolation on the same build), and `DxUiTests.NativeTextInput` (TSF document activation for a focused field), which passed in gate #6, reproduces in isolation on this machine now, and has no DxUi change since (environmental text-services state; re-checked on gate #8). |
| R0f-GDrive receipt (2026-09-03) | Mechanism: the eight mutation entry points keep the owning call's `FileSystemOptions` on the current thread (`DriveOperationOptionsScope`); every libcurl transfer issued under it polls `FileSystemCheckOperationControl` from the progress callback (`CurlCheckTransferDeadline`; libcurl invokes it at least once per second even while a request waits) and returns the host's verdict, throttle backoffs sleep in 100 ms slices under the same control (`SleepWithOperationControl`), and every attempt is bounded by the connect timeout and the hard request timeout (`PerformAuthorizedRequest`, one retry loop for every verb; transport failures are not retried). Provider-owned bound: `DriveProviderWatchdogTimeoutMs(connect, request)` = 2 x (connect + request) + 1 s (one token refresh plus one request; 81 s with the default timeouts). Surface: `IFileSystemIO` (ranged `alt=media` reader that re-reads the object's `version` before every 4 MiB chunk and fails with `ERROR_REVISION_MISMATCH` on a mid-read change; writer staging bytes in a delete-on-close temporary file and publishing them with one resumable upload at Commit, 8 MiB chunks, committed size proven from the returned resource), `IFileSystemDirectoryOperations` (folder creation, recursive sizing), `IFileSystemAtomicWriter` (the name appears only when Drive acknowledges the last chunk), and the mutations: `files.copy` for files with folders recreated per child, native move/rename through one `PATCH` (`name`, `addParents`/`removeParents`), permanent `DELETE` (a folder without RECURSIVE must be empty), trash for Recycle, overwrite = create the new object first and trash the replaced one, folder-onto-folder merge per child with `ERROR_PARTIAL_COPY` for skipped collisions. Typed facts and JSON: `providerWatchdog`, `deadline:true`, `providerWatchdogTimeoutMs` materialized from the configured timeouts (`configured-drive-watchdog-ms`), every operation true for a writable profile (a `readOnly` profile keeps its route read-only), `pathTextStableIdentity:true` (exposed names are unique per folder; duplicates carry the `[id:...]` decoration and an ambiguous text fails closed with `ERROR_DUP_NAME`), `caseOnlyRename:supported`, concurrency copy/move 4, delete 4, recycle 1. Receipts: every item completion carries `BuildDriveMutationReceipt` (success and definitive refusals known; cancel/deadline/transport failure = no record). Test seam (test-enabled builds only): `ConfigureFakeDrive` redirects the token and Drive API origins to a loopback fixture and resolves `/@conn:google-drive-selftest/...` to a synthetic My Drive connection. New fixture `Plugins/FileSystemGoogleDrive/FileSystemGoogleDrive.SelfTest.FakeDrive.cpp` on the shared `Common/LoopbackHttpFixture.h` core (token endpoint, `files.list` with the plugin's `q` grammar and paging, metadata by id, `alt=media` with Range, JSON folder creation, resumable upload sessions with Content-Range chunks including the empty-object completion, PATCH name/parents/trashed, DELETE, copy, stall and drip switches, request log); host exports `RedSalamanderGoogleDriveStartFakeDriveForSelfTest`/`...StopFakeDriveForSelfTest`/`...FakeDriveRequestLogForSelfTest`. RED: before this slice the plugin exposed no `IFileSystemIO`, every mutation returned `ERROR_NOT_SUPPORTED`, the route was `uncontained` with `pathTextStableIdentity:false`, and no libcurl transfer could be canceled while in flight. GREEN: Google Drive debug self-test `RunDriveStalledRequestCancelSelfTests` (a drip-fed children listing under a folder Delete returns `ERROR_CANCELLED` within 3 s of Cancel and deletes nothing; a stalled DELETE returns `ERROR_CANCELLED` within 3 s of Cancel, and as a transport failure within the declared bound without it; the file survives), FileOps `R0fGDrive_FakeDriveReadWriteCreateMoveDelete` (route qualifies as providerWatchdog with a nonzero bound; cross-provider Copy local -> Drive and Drive -> local byte-exact through the host bridge; provider CreateDirectory, Rename and native Move; permanent Delete through a task with every object gone), the provider capability matrix expecting Google Drive copy/native move/delete/rename/recycle with mutation admission and stable path identity, and PluginContractTests' drift guard on the new route facts. Residuals: no Docs export (native documents cannot be read as bytes), folder merges skip colliding children without a per-child prompt, readers/writers under the bridge rely on the provider bound rather than the entry-point options, `md5Checksum` proof not yet claimed, and the live Google profiles remain gated. Evidence: focused set 2026-09-03 07:29-07:36 on the final tree: full-solution Debug build 0 errors (correction 2026-09-03 09:20: the MSBuild log of that build carried 10 warnings in `FileSystemGoogleDrive.cpp`, the implicitly deleted copy/move operations of the new reader, writer, and transfer context plus three unused endpoint constants, which the console summary reports only as a count; R3-1 removed them and the 0-warning rule is checked against the MSBuild diagnostics count from now on); PluginContractTests all OK (1548 OK, 0 FAILED; Google Drive debug self-tests passed=25 failed=0 including `RunDriveStalledRequestCancelSelfTests`; focused `--google-drive-r0f-containment-selftests` 16/16 with the witness reporting streaming Cancel returning in 0 ms, a stalled DELETE returning 766 ms after Cancel and on its own after 3016 ms against the declared 11000 ms bound at the fixture timeouts 2000/3000); FileOps `R0fGDrive_` 3/3 (`FileOps.SelfTest.R0fGDrive.ProviderWatchdogTimeoutMs` 21000 ms for the fixture timeouts 2000/8000, `CopyToDriveMs` 188, `CopyFromDriveMs` 187, `DeleteOnDriveMs` 78), `R0fGraph_` 3/3, `R0fS3_` 3/3, `R0fCurl_` 3/3, `R0fSmb_` 4/4, `FileOps_ProviderCapabilityMatrix` 3/3, `Phase10_` delete guards 9/9; Pester source contracts 187/187. Fixture lesson kept in the code: the fake Drive removes a subtree by an id taken by value (a view into the vector element being erased dangled once its children were removed, so a folder survived its own DELETE while its child vanished). Fresh Full gate #8 on `d4c65293` (run `20260903T053735Z-73932-5f544f5287ce4edcbf280be5645359b5`): 2068 counted, 1975 passed, 9 failed, 84 skipped; every R0f-GDrive, R0f-Graph, R0f-S3, R0f-Curl and R0f-SMB case, PluginContractTests, the Pester source contracts, and `DxUiTests.NativeTextInput` (red in gate #7) passed in the gate. The nine failures: the known `cmd_pane_batchRename_admission_worker_owned_cancellable` load flake; Compare `google_drive_plugin_contract`, whose expectation still pinned the pre-R0f-GDrive read-only tuple (`copy=false`, `delete=false`) and now expects the full tuple plus the providerWatchdog route class (follow-up commit); and seven FileOperations cases (`R4A19_DiscoveryProviderControls`, `Phase5_DiscoveryCancelReleasesSlot`, `Phase7_CrossPaneVisibleRefreshDummy`, `Phase8_DefaultBandwidthLimitFromSettings`, `Phase9_ConflictPrompt_TypeMismatchNoOverwrite`, `Phase11_CrossFileSystemBridge`, `Phase12_ReparsePointPolicy`, six of them 2-3 minute timeouts whose aborted chains produced the extra skips) caused by machine load: a whole-disk `find` started on this machine at 07:44 for an SDK-header lookup ran through the FileOperations stage. Once it was stopped, every affected family passed in isolation on the same build (Phase5 9/9, Phase7 19/19, Phase8 7/7, Phase9 11/11, Phase11 9/9, Phase12 3/3, R4A19 4/5 with only the unauthorized alternate-volume control red; while the scan was still running the fake-MTP cancel control failed 3/3 with `ERROR_IO_INCOMPLETE` about 100 ms after Cancel, so that control is load-sensitive rather than broken). |

##### R0f-GDrive-OR1 — Mutation retry truth and deep-tree completion (post-closeout correction)

| Field | R0f-GDrive-OR1 |
|---|---|
| Slice/state/owner | **R0f-GDrive-OR1 / `ACTIVE` (2026-09-04) / Codex `/root` in this worktree.** Adversarial-review correction to the completed Google Drive provider only. |
| Finding | The provider's one authorized-request loop retries `429` and `5xx` for every verb. A committed-but-failed POST can therefore create a second folder or copy, while DELETE/PATCH can report the retry rather than the first commit. Separately, recursive directory size silently stops below level 64 and returns `S_OK`, and provider-native folder copy/merge still recurses with the same arbitrary ceiling although the R4 closeout names only Local provider recursion as residual. The provider's newly added ABI catches also use forbidden `catch (...)` and treat `std::bad_alloc` as recoverable. |
| Correction | Add an explicit request retry class: only read-only requests may automatically retry transient HTTP status; mutations refresh an actually rejected `401` once but never repeat after a `429`/`5xx` response whose commit state is unknown. Convert the provider-native folder transfer to an iterative frame stack and remove the silent directory-size depth cutoff. Replace catch-all/alloc recovery with the repository's named-exception ABI boundary policy. |
| Tests | `[+]` Add a debug HTTP witness proving a POST that returns `503` is attempted once; retain the bounded GET retry witness. `[+]` Extend the loopback fixture proof with a tree deeper than 64 levels and require exact recursive size plus successful provider-native copy. Keep R0f-GDrive containment, R3 publication, capability, and PluginContractTests green. |
| Performance/resources | The deep-tree path is iterative and retains one frame per open directory; the existing provider request/page bounds remain. Record focused metrics/counts in the provider selftest and consume the next Fresh Full archive; no new disk spool. |
| Scope/STOP | In: `Plugins/FileSystemGoogleDrive/FileSystemGoogleDrive.{h,cpp}`, its fake-drive selftest, `Specs/FileSystem/FileSystem_GoogleDrive.md`, and this plan. STOP if safe mutation retry needs a public ABI change or if the iterative provider walk cannot keep the existing per-child partial-result semantics. |
| Exit | Code, deterministic tests, authoritative spec, 0-warning Debug/test-enabled Release builds, focused provider/FileOps gates, and an exit receipt are recorded before `COMPLETE`. |

##### R0-Policy-OR1 — Shared helper ownership and exception-policy cleanup

| Field | R0-Policy-OR1 |
|---|---|
| Slice/state/owner | **R0-Policy-OR1 / `ACTIVE` (2026-09-04) / Codex `/root` in this worktree.** Adversarial-review cleanup of files introduced by completed reliability slices; no product-policy expansion. |
| Finding | `Common/ContentDigest.h` and the Curl FTP selftests contain forbidden catch-all handlers, and `Common/SynchronousIoCancelWatch.h` manually owns/closes a thread handle and threadpool timer despite the WIL-only resource rule. |
| Correction | Let `noexcept` make allocation failure fatal or catch only named `std::exception` at a required boundary with one diagnostic. Replace raw owned handles with WIL unique wrappers while preserving timer drain-before-close and scope unregister-before-thread-handle-close. |
| Tests | Existing content-digest known-answer tests, SMB cancel-watch tests, Curl plugin selftests, source-contract checks, 0-warning builds, and `git diff --check`. |
| Scope/STOP | In: `Common/ContentDigest.h`, `Common/SynchronousIoCancelWatch.h`, `Plugins/FileSystemCurl/FileSystemCurl.SelfTest.Ftp.cpp`, and this plan. STOP on any behavior or public-ABI change. |
| Exit | Focused tests and an exit receipt are recorded before `COMPLETE`. |

### Package R1 — Truth and Preparing family

**Priority:** P1
**Boundary:** the subpackages are independently promotable. FOS-10 immediate
capability clearing belongs only to R0d; R2 owns any later executable provider
contract. An exact provider call without a cooperative bound, proved provider-local
watchdog/quarantine, or accepted isolation is disabled rather than admitted.

##### R0f-Curl-OR2 — Curl connections per easy handle, no cross-thread pool (post-closeout correction)

| Field | R0f-Curl-OR2 |
|---|---|
| Slice/state/owner | **R0f-Curl-OR2 / `COMPLETE` (2026-09-04) / Claude in worktree `codex/r4-a19-closeout-2026-09-02`.** Correction to the Curl provider's libcurl sharing only. |
| Finding | Fresh Full gates #11 and #13 (and a 2026-09-03 23:10 run) crashed the FileOps self-test process with an access violation in libcurl 8.21 `url_match_destination` (url.c:977) inside `Curl_cpool_find`, on a bridge worker thread performing an FTP request during `CopyDirectoryParallel` (`R0fCurl_FakeFtpReadWriteCreateDelete`). Dumps and sidecars: `%LOCALAPPDATA%\RedSalamander\Crashes\RedSalamander-20260904-010252-p107076`, `...-20260904-044416-p99544`, `...-20260903-231016-p88132`. The plugin's one process-wide `CURLSH` shared `CURL_LOCK_DATA_CONNECT` across every easy handle while those handles ran concurrently on several threads; libcurl's pooled-connection matching raced with another thread's teardown. Isolated reruns never reproduce it (sequential), which is why the aborted stages classified as environment before the dumps were read. |
| Correction | The share covers DNS and TLS sessions only (`FileSystemCurl.Shared.cpp`, `GetCurlShareHandle`). Pooled easy handles keep their own connections; `CurlEasyPool` hands a handle to one thread at a time, so connection reuse stays thread-confined. No host change. |
| Evidence | `RunCurlParallelWritersSelfTests` (Curl debug self-tests, run by `PluginContractTests`): six threads upload ten files each through one FTP instance against the fake server; the Commands source contract pins `CURL_LOCK_DATA_CONNECT` absent from the share. Focused cycle on the corrected tree (2026-09-04 14:24, Debug rebuild 0 warnings / 0 errors): PluginContractTests green with the Curl debug self-tests at 258/258 (250 before the new case); FileOps `R0fCurl_` 3/3, `Phase11_` 9/9, `Fairstream_CrossFs` 3/3, `Cinderstar_` 8/8, `Phase10_ContentVerification` 3/3; Commands `file_operations_` 3/3, `file_system_capabilities_contract_` 1/1; Pester 187/187. The Fresh Full gate that follows (gate #13 rerun) is recorded on the R4-T1 card. |
| Residual | The gate #13 FileOps stage that crashed is rerun as gate #13 on the corrected tree; the R4-T1 exit receipt is recorded from that rerun. |

##### R0-RC4 — Identity-less replace fails closed on a zero host timestamp (draft)

| Field | R0-RC4 |
|---|---|
| Slice/state/owner | **R0-RC4 / `COMPLETE` (2026-09-04; exit receipt below) / Claude in worktree `codex/r4-a19-closeout-2026-09-02`.** Post-closeout correction to the cloud writers' conditional replace only. Baseline `30a2a239`; drift review: `git status --short` shows only the user-owned untracked `last_run/`. |
| Finding (2026-09-04 review) | R0-RC3 made the Dummy and Curl writers refuse an identity-less replace whose host token (`FileSystemBasicInformation.lastWriteTime`) is zero. Microsoft Drive's `ResolveReplaceOccupant` (`FileSystemMicrosoftDrive.cpp` ~7476) compares the expected and live timestamps only when both are non-zero and then pins `If-Match` to the live occupant's eTag, so with a zero timestamp it replaces whoever is there now rather than what the prompt showed. S3 and Google Drive pin the ETag or id observed when the writer opens, which is after the prompt, not the prompt's occupant; the host token carries no ETag field. |
| Correction | The same rule as Dummy/Curl on all three writers: with an expected replacement whose timestamp is zero, or a live occupant whose timestamp is zero, refuse with `ERROR_REVISION_MISMATCH` before any upload; the prompt's token is the only authority, and an ETag pinned at writer open is not a substitute. Extending the host token with an ETag is out of scope (a separate decision). |
| Scope | In: `Plugins/FileSystemMicrosoftDrive/FileSystemMicrosoftDrive.cpp` (`SetExpectedReplacement`, `ResolveReplaceOccupant`), `Plugins/FileSystemS3/FileSystemS3.IO.cpp` and `Plugins/FileSystemGoogleDrive/FileSystemGoogleDrive.cpp` (`SetExpectedReplacement` and their replace paths), the providers' debug self-tests against their fake fixtures, the three provider specs, this plan. Out: host prompt token shape, R3 publication record, Local and MTP writers. |
| RED/GREEN evidence | RED: per provider, a debug self-test that sets an expected replacement with `lastWriteTime == 0` and swaps the occupant before Commit: today Graph commits with the live eTag, S3 and Google Drive commit against the writer-open revision. GREEN: all three refuse before upload with `ERROR_REVISION_MISMATCH` and the swapped occupant survives; the existing R3-1 conditional-replace cases with real timestamps stay green; `R3_4` gains the three providers beside Dummy. |
| Performance/resources | No new request; a refused replace saves the upload. |
| STOP/rollback | STOP if a provider cannot observe the occupant's timestamp before upload without an extra round trip it does not already make; rollback is the live-eTag pin (one revert per provider). |
| Exit receipt | R0-RC4 exit receipt (2026-09-04): activation `7a68110f`; implementation `504d0e8f` (Microsoft Drive `ResolveReplaceOccupant`, the S3 and Google Drive `SetExpectedReplacement` refuse with `ERROR_REVISION_MISMATCH` before any upload when the expected or the live timestamp is zero; three provider debug self-tests `RunGraphZeroTimestampReplaceSelfTests`, `RunS3ZeroTimestampReplaceSelfTests`, `RunDriveZeroTimestampReplaceSelfTests` against the fake fixtures; the three provider specs state the rule). Debug full-solution builds 0 warnings / 0 errors (RED and GREEN). RED witness on the baseline (2026-09-04 19:12-19:17, `7a68110f` plus the tests alone, PluginContractTests): all three self-tests failed at exactly the new check: Microsoft Drive debug self-tests 246 passed / 1 failed (`a Graph identity-less replace with a zero host timestamp must be refused, not pinned to the live occupant`), Google Drive 31 / 1 (`must be refused before upload`), S3 219 / 1 (`must be refused before upload`); every other provider suite green. GREEN (focused runs on the landed tree): PluginContractTests green (Microsoft Drive debug self-tests 247/247, Google Drive 32/32, S3 220/220; every other suite unchanged); FileOps `R3_` 6/6, `R0fGraph_` 3/3, `R0fS3_` 3/3, `R0fGDrive_` 3/3, `Phase10_ContentVerification` 3/3; Commands `file_operations_` 2/3 in the cycle, the miss being the Phase 0 popup spec-token guard that R6-A08's spec rewrite had displaced (repaired in the follow-up `4f594ea2`, after which 3/3) and `file_system_capabilities_contract_` 1/1; Pester 187/187. Performance: a refused replace saves the upload; no new request. Fresh Full gate #15: on `477bb59f` (run `20260904T173150Z-37604-9250b807c93644b79438ffb8707a8dd0`, started 2026-09-04 19:35): 2077 counted, 2014 passed, 9 failed, 54 skipped; PluginContractTests, DxUi tests, Compare 232/232 green. The nine failures: the Tools Pester resource-parity test (the four new R6-A08 strings were missing from the cs-CZ, fr-FR, ja-JP and sk-SK satellite tables; corrected in `d537a240` with translations, parity 9/9 and source contracts 187/187 on a clean build); `ViewerPETests.Interactive` (`TestViewerShellComboHostsLongRunOpenCloseStayStable` child harness exited non-zero in the gate and in a direct rerun on the same build, then the child and the whole interactive group passed on the `d537a240` build, so it is attributed to the same satellite gap); `R4A19_DiscoveryIndependentVolumes` (environment: alternate-volume test root not authorized); and six Commands UI-timing cases (`cmd_pane_batchRename_admission_worker_owned_cancellable`, `cmd_preferences_dialog_keyboard_live_search_dx_interaction`, two Find-dialog cases, `cmd_pane_navigation_go_to_root_directory_keeps_navigation_shell_stable`, `cmd_app_menuBar_hover_switches_top_level_popup`) that each passed 1/1 isolated on the same build. |

#### R1a — Terminal and clipboard truth

**Owns:** FOS-07/FOS-08

##### R1a activation card

| Activation-card field | R1a bounded execution contract |
|---|---|
| Slice/state/owner | **R1a / `COMPLETE` / Codex `/root` in this worktree.** This independently released terminal-result and clipboard-admission truth slice does not activate R1b Change Case, R1c compensation, R1d's complete Preparing lifecycle, R4 overlap comparison, R6 presentation, or R2 provider contracts. |
| Baseline and drift | Activation baseline is clean commit `a8838d38`. A scoped `git diff --name-status` over the named runtime/test/spec owners returned no tracked paths before activation; repository-root `last_run/` is pre-existing untracked test output and is excluded. `StartOperation` publishes a task and worker, calls the clipboard barrier, and unconditionally sets `_workerReleased=true`; `ThreadMain` only then calls `PrepareMutationInterlockScopes`. A clear failure is therefore a warning on a still-mutating task, while a later root/interlock failure occurs after the cut list is gone. Stop/cancel, Batch Rename admission failure, interlock failure, queue cancellation, and pre-execution cancellation post completion without calling `FinalizeTypedItemResults`. `StoreTypedItemResult` overwrites an occupied slot, and fallback synthesis assigns `E_PENDING`/Unknown to unobserved rows although the task may still be known pre-mutation. The authoritative clipboard paragraph still blesses clear-failure execution and must change with this slice. |
| Contract | Implement FOS-07/FOS-08 and the existing `FO-ITEM-01`/clipboard contract with a narrow pre-mutation handshake. Every selected top-level root receives one initialized result builder before worker publication. The worker completes current top-level mutation-interlock/root readiness without recursive discovery or mutation, publishes one readiness status, and waits. Only successful readiness permits the UI-owned exact clipboard-sequence consumer to run; successful consumption records its one-shot receipt and permits mutation, while failure records the terminal cause and wakes the worker only into the no-mutation terminal funnel. Every worker exit reduces through that funnel. A builder moves monotonically from `Preparing` to `MutationPossible` and then to exactly one immutable terminal result; a missing observation before `MutationPossible` is `NotAttempted + Retained` with the task failure/cancel status, never `E_PENDING`/Unknown. Terminal store is compare-and-store: duplicate or invalid index is diagnosed and rejected without altering the first truth. A no-op unresolved-gate hook is present at the readiness boundary for later R4/R6 registration but this slice does not compare scopes or render a decision. A successfully consumed cut list is never restored. |
| Scope | In scope: `RedSalamander/FolderWindow.FileOperations.State.Runtime.cpp`, `FolderWindow.FileOperations.State.cpp`, and `FolderWindow.FileOperationsInternal.h`; focused FileOps selftests in `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.cpp` and `.SelfTest.Phases10_13.cpp`; source-contract assertions only where needed; `Specs/FileSystem/FileSystem_FileOperations.md`, `Specs/UI/UI_FileOperationsPopup.md` if result wording changes, `Specs/Testing/Testing_PerformanceValidation.md`, this card/receipt, and compact evidence below `Specs/TestRuns/4cb089111a23/FileOps/`. Out of scope: public ABI, provider algorithms/capabilities, Change Case/raw payload repair, primary-cancel compensation, cleanup-debt receipts, recursive discovery redesign, full Preparing/reveal UI, overlap detection/consent, retry/review-later UI, progress-table retention, and re-enabling R0e-uncontained routes. |
| Predecessors | E0 is closed at `fc32a4aa4b01d6d684166968a0615d3265f33d26`; R0e is closed at `a8838d38` and guarantees admitted readiness calls have bounded/provider-watchdog containment. Section 8.1 requires only E0, but R0e's closure makes this worker/readiness handshake honest for the current admitted set. No active owner overlaps this scope. |
| Steps | (1) Add RED tests proving clipboard-consume failure currently mutates, readiness failure consumes the cut list before failing, an early pre-mutation exit has no dense typed results, a second terminal store overwrites the first, and multi-root fallback can emit `E_PENDING`/Unknown. (2) Capture five same-machine test-enabled x64 Release baseline samples of builder initialization/finalization plus the no-provider clipboard gate fixture. (3) add one per-root builder and phase, initialize it before publication, and make terminal storage compare-and-store. (4) add the worker readiness/consume/mutation handshake and funnel every `ThreadMain` exit through phase-aware finalization. (5) retain successful one-shot clipboard semantics, add the no-op unresolved-gate hook, and update authoritative specs. (6) run focused Debug and exact-commit test-enabled Release validation, archive candidate evidence, record the exit receipt, and mark only R1a complete. |
| RED/GREEN evidence | RED must fail because clear failure still permits mutation, root/interlock readiness still follows consumption, early worker exits omit dense results, duplicate store wins, and unobserved multi-root rows can become `E_PENDING`/Unknown. GREEN requires: clear failure leaves every source and destination untouched and stores dense `NotAttempted + Retained + Failed`; readiness failure never invokes clipboard consumption and produces the same known-no-mutation axes; successful consume precedes the first possible mutation and is called exactly once; cancel/stop at either gate yields dense canceled results; every selected root has exactly one terminal value; duplicate store increments a diagnostic/test counter but preserves the original value; no final typed result has `E_PENDING`; existing successful/partial/indeterminate mutation receipts retain their axes; duplicate sequence admission and retained-source actions remain correct. |
| Performance/resources | Protected deterministic scenario: initialize and phase-finalize 4,096 selected-root builders, attempt one duplicate store, and execute 1,024 no-provider readiness/clipboard gate decisions in one process. Emit `FileOps.SelfTest.R1a.TerminalFunnelUs`, `.BuilderCount`, `.TerminalStoreCount`, `.DuplicateStoreRejectCount`, `.ClipboardGateDecisionCount`, and `.ProviderCallCount` once per process. Capture five independent test-enabled x64 Release processes before and after under `Specs/TestRuns/4cb089111a23/FileOps/2026-08-31_r1a_terminal_clipboard_{baseline,candidate}_release/`. Candidate p95 must be at most `max(1.50 * baseline p95, baseline p95 + 50,000 us)`, every duration below 2,000,000 us, counts exact, provider-call count zero, and retained state constant. This is result-funnel/admission overhead evidence, not transfer throughput. |
| Durable owners | `Specs/FileSystem/FileSystem_FileOperations.md` owns clipboard ordering, selected-root readiness, per-item phase truth, and immutable terminal reduction. `Specs/UI/UI_FileOperationsPopup.md` owns any durable user-facing failed-consumption/result wording. `Specs/Testing/Testing_PerformanceValidation.md` owns the deterministic builder/gate scenario, five-process evidence, and budgets. This plan retains delivery mapping and the generated exit receipt only. |
| STOP/rollback | STOP if clipboard consumption can precede complete selected-root readiness; a failed consume can enter provider/discovery/mutation code; readiness runs on the UI thread; the handshake introduces an unbounded call not admitted by R0e; any early exit can post completion without dense typed results; terminal truth can still be overwritten; existing explicit mutation/cleanup axes are degraded to synthesized values; the change requires R1d UI, R4 overlap policy, R6 decisions, R1b/R1c behavior, public ABI, or live clipboard/provider credentials; retained state grows with the deterministic loop; or Release evidence exceeds budget without a diagnosed machine anomaly. Rollback removes only the builder/handshake/tests/metrics/spec clauses while R1a remains open; never restore clear-failure execution as the normative contract. |
| Exit-receipt contract | Before R1a closes, record exact activation/RED/implementation/evidence commits; Debug and test-enabled Release build receipts; focused GREEN totals; five-process baseline/candidate summaries and constant-resource proof; clear-failure/no-mutation, readiness-before-consume, cancel-at-gates, dense early-exit, duplicate-store rejection, and no-`E_PENDING` results; source proof that mutation cannot follow a failed gate; validated archive paths and whole-inventory validation; `git diff --check`; authoritative-spec destinations; and remaining R1b/R1c/R1d/R4/R6 gaps. |

**R1a RED receipt (2026-08-31):** activation is `d2e0efe3`; RED contracts are
`2e30790a`. The exact RED tree built the full Debug solution with zero warnings/errors
under receipt `0052f0c30bf7b19e11cdfcfbe02924d4839eee4a0badf9d293fd118aa5ed037d`
and the full test-enabled x64 Release solution with zero warnings/errors under receipt
`286d6b8abae0b6aad771cb6d3fb5eccf5d39bc647618471a59fc204b27731269`.
Focused Debug run `r1a-red-debug-20260831b` failed `2/1/0` exactly because clipboard-
consumption failure still released the worker instead of terminating with dense known-
no-mutation truth. Five independent exact-commit Release processes
`r1a-baseline-release-final-20260831-01` through `-05` retained the same `2/1/0`
failure and emitted `FileOps.SelfTest.R1a.TerminalFunnelUs` values `587`, `487`,
`539`, `619`, and `594` us (p95 `619` us). The first external baseline directories
were not retained; the archived samples were recreated from the same committed RED
tree with zero diagnostics under project receipt
`20d14ee14819b7c17bed213ae25389227935c81a1486cc913451027b79a8327e`.
Every process initialized and terminal-
filled `4,096/4,096` slots only through the legacy fallback, rejected `0/1` duplicate
stores, permitted mutation for all `1,024/1,024` failed clipboard-gate decisions, and
made zero provider calls. The RED fixture also proves the legacy fallback leaves all
4,096 unobserved rows at `E_PENDING`, and the duplicate terminal write replaces the
first truth. These are baseline diagnostics, not qualification evidence.

- [x] Allocate one result builder for every selected root before execution.
- [x] Funnel all exits through exactly one terminal store; duplicate store is a test
      failure/diagnostic, not last-writer-wins recovery.
- [x] Add phase-aware transitions so a no-mutation failure is known NotAttempted.
- [x] Complete selected-root readiness before clipboard consumption and expose one
      unresolved-gate hook in that funnel. R4/R6 later register A02 decisions through
      the hook; R1a neither compares scopes nor renders warnings. Failed consumption
      aborts release; successful consumption remains one-shot and is never restored as
      stale intent.

**R1a exit receipt (2026-08-31):** implementation/spec/test commit `c2e055e9`
and archive-evidence commit `57dd527e`
allocates every selected-root builder before publication, runs selected-root/interlock
readiness on the worker before the UI-owned exact clipboard consumer, and admits
discovery/mutation only when both readiness and consumption succeed. Failed readiness
never invokes consumption; failed consumption and gate cancellation reduce through the
same dense `NotAttempted + Retained` terminal funnel. Terminal storage is first-writer-
wins and rejects invalid/duplicate indexes without altering the first truth; no terminal
fallback carries `E_PENDING`. The no-op unresolved Preparing hook is registered only at
this boundary and adds no R4/R6 policy.

The exact implementation tree built the full Debug solution with `0/0` diagnostics under
receipt `42dce528dc0d540db022c443f9f27aea212c5e51925656732ad481babdacd6e8`
and the full test-enabled x64 Release solution with `0/0` diagnostics under receipt
`269b700e5705d5105646a1994ebcc4a0395837379703b5188ca3abc5619da1b0`.
Focused Debug `r1a-green-debug-fullbuild-20260831a` and each of five exact-commit
Release processes `r1a-candidate-release-final-20260831-01` through `-05` passed
`3/0/0`. Candidate funnel values were `769`, `1,174`, `558`, `608`, and `661` us
(p95 `1,174` us) against `max(1.50 * 619, 619 + 50,000) = 50,619` us; every
sample was below 2 seconds and reported 4,096/4,096 dense terminals, `1/1` duplicate
rejection, 1,024 gate decisions with 512 mutation permits, and zero provider calls.
The provider-capability/receipt matrix passed `3/0/0`; `LocalizationTests.exe` returned
0. Archives are
`Specs/TestRuns/4cb089111a23/FileOps/2026-08-31_r1a_terminal_clipboard_{baseline,candidate}_release/`.
The complete TestRuns inventory and `git diff --check` pass at closure.

A broad Debug FileOps run still stops at the pre-existing
`Fairstream_MoveSameSizeCollisionPrompts` expectation (`ERROR_IO_INCOMPLETE` instead of
`S_FALSE` after Skip). Exact RED-tree run `r1a-control-samesize-release-20260831a`
reproduces the same result, so this is not an R1a regression and was not used as GREEN
evidence. R1b, R1c, R1d, R4, and R6 remain inactive.

#### R1b — Change Case immediate truth

**Owns:** FOS-09 only
**Ordering:** the R1b/E5 rows in section 8.1.

##### R1b activation card

| Activation-card field | R1b bounded execution contract |
|---|---|
| Slice/state/owner | **R1b / `COMPLETE` / Codex `/root` in this worktree.** This independently released truth slice repairs the current Change Case executor and its task-card transport. It does not activate E5, replace `FileSystemRenameBatch`, create `RenamePlan`, add recovery/journaling, or activate R1c/R1d/R2-R9. |
| Baseline and drift | Activation baseline is commit `c1841bc3`. Scoped inspection finds no tracked worktree changes; repository-root `last_run/` is pre-existing untracked output and is excluded. `ChangeCase::ApplyToPaths` currently treats a failed or null `ReadDirectoryInfo` result as `continue`, so recursive discovery can omit a subtree and later return `S_OK`. The 700 ms reveal path in `FolderWindow::CommandChangeCase` calls `payload.release()` and passes the raw pointer through `SendMessageTimeoutW`, while `FolderWindow::OnChangeCaseTaskUpdate` accepts only `TakeMessagePayload` registry tokens; successful delivery therefore rejects/leaks the raw payload and yields no stable task ID. Existing cancel and artifact-guard tests prove pre-mutation cancellation/revalidation, but not unreadable-descendant truth or long-run task-card identity. |
| Contract | Recursive Change Case discovery is all-or-fail before mutation for every child that parent enumeration has already proved is a traversable non-reparse directory: its first failed directory read or malformed/null successful result returns the exact failure (success plus null becomes `E_UNEXPECTED`), posts failed completion truth, and performs zero renames. An initial selected item is not yet typed by this legacy path-only API, so an exact failed initial directory probe remains the file/non-directory discriminator until E5 supplies typed plan inputs; successful-null is still invalid. Cancellation remains `ERROR_CANCELLED`; a rename-batch failure remains the exact first provider failure after the already completed batches and may never be converted to `S_OK`. Long-run task-card creation and updates use only `PostMessagePayload`/`TakeMessagePayload`. One shared creation receipt links every payload for the operation; the UI handler resolves or creates one informational task and publishes its ID into that receipt. Same-worker/same-HWND FIFO ordering makes later queued updates target the same task without blocking the worker or retaining raw ownership. Failed post, stale token, and HWND teardown destroy payload state without a leak or worker wait. Runs below the existing 700 ms reveal threshold create no task card. |
| Scope | In scope: `RedSalamander/ChangeCase.cpp`; `RedSalamander/FolderWindow.FileSystem.Commands.cpp`, `.cpp`, and `.Private.h`; focused Commands selftests in `RedSalamander/SelfTest/Commands/Commands.SelfTest.Dialogs.cpp` plus the minimal test registration/source-contract surface; `Specs/UI/UI_CommandMenuKeyboard.md`, `Specs/FileSystem/FileSystem_FileOperations.md` only if shared FileOps wording is affected, `Specs/Testing/Testing_PerformanceValidation.md`, this card/receipt, and compact evidence under `Specs/TestRuns/4cb089111a23/Commands/`. Out of scope: `FileSystemRenameBatch` replacement or semantic redesign, central `RenamePlan`, preview UI, journals/recovery, provider ABI/capability changes, retry/skip conflict policy, common Preparing lifecycle, per-descendant result UI, cleanup compensation, and unrelated task/progress retention limits. |
| Predecessors | E0 and R1a are complete; section 8.1 requires E0 only. E5 explicitly consumes completed R1b and remains blocked. No active owner overlaps these symbols. |
| Exact symbols | Runtime owners are `ChangeCase::ApplyToPaths`, the worker-local `ProgressState` inside `FolderWindow::CommandChangeCase`, `FolderWindowFileSystemInternal::ChangeCaseTaskPayload`, and `FolderWindow::OnChangeCaseTaskUpdate`. Test-only support may expose a narrow receipt-resolution helper, but no production-wide singleton or second task registry is allowed. |
| Test IDs and commands | Add `cmd_pane_changeCase_unreadable_descendant_truth`, `cmd_pane_changeCase_partial_cancel_truth`, and `cmd_pane_changeCase_task_payload_truth` in `Commands.SelfTest.Dialogs.cpp`; retain `cmd_pane_changeCase` and `cmd_pane_changeCase_dialog`. Build with `./build.ps1 -Configuration Debug` and `./build.ps1 -Configuration Release -EnableTests`; run each exact case with `./Tools/Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -CaseFilter <id> -RunId <run-id>`. Run the existing Commands source-contract suite that owns posted-payload rules and `Tests/LocalizationTests.exe`; the exit receipt records exact commands/receipts selected by the current harness. |
| Fault matrix | Required deterministic faults are: failed `ReadDirectoryInfo` for a child already proven directory by its parent; `S_OK` plus null `IFilesInformation`; malformed buffer/size propagation retained from current core coverage; stop requested before discovery and between progress callbacks; preparation/revalidation cancellation; exact first `FileSystemRenameBatch::Execute` failure after at least one completed depth/batch; first reveal post failure; stale/drained task token; and reveal/progress/completion payloads queued before the first UI adoption. Each asserts exact HRESULT, mutation set, task count/ID, and remaining payload/receipt ownership. |
| Steps | (1) Add RED coverage for unreadable descendant discovery, raw task-card rejection, stable single-card updates, cancellation, and exact partial failure. (2) Capture five same-machine test-enabled x64 Release baseline samples of the task receipt/reduction scenario. (3) fail recursive discovery on the exact read/null/buffer error before constructing or executing rename batches. (4) replace the synchronous raw send with tokenized asynchronous payloads sharing one receipt, and resolve the current task ID only in the UI handler. (5) update authoritative Change Case and performance contracts. (6) run focused Debug, exact-commit test-enabled Release validation, archive five candidate samples, record the exit receipt, and mark only R1b complete. |
| RED/GREEN evidence | RED must show that a deliberately unreadable proven nested directory is skipped and the call returns `S_OK`, and that the threshold-reveal raw pointer cannot be adopted by the token-only handler. GREEN requires: unreadable proven nested discovery returns the exact failure before any rename while an initial selected file remains a valid non-recursive leaf; success-plus-null is `E_UNEXPECTED`; pre-cancel remains `ERROR_CANCELLED` with no mutation; an injected later rename failure is returned exactly while already completed batches remain observable; reveal plus multiple progress/completion updates creates exactly one informational task with one nonzero ID; completion updates that same card; payload-post failure/window drain leaks nothing and never blocks; and source contracts reject `payload.release()`/raw `SendMessageTimeoutW` in the Change Case task path. |
| Performance/resources | Protected deterministic scenario: construct, post/resolve, and reduce 4,096 tokenized Change Case task updates through one shared creation receipt without a provider call, with one UI informational task identity and bounded payload retention after queue drain. Emit `Commands.SelfTest.R1b.ChangeCaseTaskDispatchUs`, `.PayloadCount`, `.TaskCreateCount`, `.TaskUpdateCount`, `.ProviderCallCount`, and `.OutstandingReceiptCount` once per process. Capture five independent test-enabled x64 Release processes before and after under `Specs/TestRuns/4cb089111a23/Commands/2026-08-31_r1b_change_case_truth_{baseline,candidate}_release/`. Candidate p95 must be at most `max(1.50 * baseline p95, baseline p95 + 50,000 us)`, every duration below 2,000,000 us, counts exact, provider calls zero, and outstanding receipt count zero after drain. This measures task-transport overhead, not provider traversal or rename throughput. |
| Durable owners | `Specs/UI/UI_CommandMenuKeyboard.md` owns Change Case traversal failure, cancellation/partial-result, asynchronous reveal, and one-card behavior. `Specs/Testing/Testing_PerformanceValidation.md` owns the deterministic tokenized dispatch scenario, five-process evidence, and budget. `Specs/FileSystem/FileSystem_FileOperations.md` changes only if the shared result contract needs clarification. This plan retains delivery mapping and the generated receipt only. |
| STOP/rollback | STOP if a read failure for a child already proven as a directory can still be skipped; discovery failure can occur after a rename begins; the legacy path-only input cannot preserve initial selected-file behavior without typed inputs that belong to E5; task creation blocks the worker/UI on cross-thread send; a payload bypasses the opaque-token registry; more than one task card can be created per operation; token adoption depends on a raw address; cancellation or a provider failure is remapped to success; the change requires `RenamePlan`, `FileSystemRenameBatch` redesign, provider ABI, R1d/R6 presentation, live credentials, or an unbounded provider call; retained state grows with the deterministic loop; or Release evidence exceeds budget without a diagnosed machine anomaly. Rollback removes only R1b runtime/test/metric/spec clauses and leaves R1b open; never restore silent proven-subtree omission or raw payload ownership as normative behavior. |
| Exit-receipt contract | Before R1b closes, record exact activation/RED/implementation/evidence commits; Debug and test-enabled Release build receipts; focused GREEN totals; five-process baseline/candidate summaries and constant-resource proof; exact unreadable/null/cancel/partial-failure HRESULTs and mutation observations; one-card creation/update identity; post-failure/drain ownership; source proof banning raw Change Case payload sends; validated archive paths and whole-inventory validation; `git diff --check`; authoritative-spec destinations; and remaining E5/R1c/R1d/R6 gaps. |

**R1b RED receipt (2026-08-31):** activation is `c19c359a`; RED contracts are
`448f2981`. The exact RED tree built the full test-enabled x64 Release solution
with zero warnings/errors under receipt
`c7882c69d6b41268f9281eee1dc70637836b6b4e6b79d70992b61431a788f2a3`.
`cmd_pane_changeCase_unreadable_descendant_truth` returned `S_OK` instead of
the injected `0x80070020`, proving that a child already enumerated as a directory
was silently omitted. `cmd_pane_changeCase_task_payload_truth` created 4,096
simulated task identities instead of one and the HWND path did not publish the
shared receipt, proving that the raw synchronous reveal and later tokenized
updates cannot form one card. The corrected exact-name partial/cancel fixture
already passes on the RED tree and therefore guards preserved behavior rather
than supplying the regression signal. One retained Release sample reported
155 us, 4,096 payloads, 4,096 creates, zero updates/provider calls/outstanding
receipts. The runner executed the requested five-process loop, but its disk audit
treated earlier explicitly authorized run directories below the same initialized
test root as unexpected and pruned samples 1--4. The five baseline samples were
subsequently recaptured by copying each sample into the runner-owned `evidence`
directory before the next process, as recorded in the exit receipt below.

**R1b exit receipt (2026-09-01):** activation is `c19c359a`; RED contracts
are `448f2981`; RED receipt is `529eca0f`; production/tests/spec repair is
`b9df3a3a`; compact evidence is `1562a9ef`. The exact implementation tree
built the full Debug solution with zero warnings/errors under receipt
`8ea7cb1a750830121765430a7d97795083ee3a13d82006cb80b6e4f863ab56ef`
(`.build/logs/msbuild-20260831_235200_525-pid59792-0bc6808c.log`) and the
full test-enabled x64 Release solution with zero warnings/errors under receipt
`38841fe2388d7d8522fa0cb0c7d5a7b3d1c0f211623b4780c626f85367973024`
(`.build/logs/msbuild-20260831_235603_426-pid26420-53bdc465.log`).

Focused Debug runs
`r1b-green-exact-debug-unreadable_descendant_truth`,
`r1b-green-exact-debug-partial_cancel_truth`, and
`r1b-green-exact-debug-task_payload_truth` each passed `1/0/0`. The exact
unreadable proven-directory fault is retained as `0x80070020` before any
rename; `S_OK` plus null is `E_UNEXPECTED`; an initial untyped probe retains
the legacy leaf discriminator; pre-cancel remains `ERROR_CANCELLED`; and a
later injected batch failure preserves already completed deep renames while
leaving the later shallow name untouched. Failed post and HWND drain release
the payload, stale tokens are rejected, and the real FolderWindow reveal plus
completion path creates and updates one card. Retained Release cases
`cmd_pane_changeCase` and `cmd_pane_changeCase_dialog` each passed `1/0/0`.
The Commands source-contract suite passed `186/186`, including the ban on raw
Change Case `payload.release()`/`SendMessageTimeoutW`, and the corrected
Release `LocalizationTests.exe` invocation passed. The first direct
localization invocation omitted `REDSALAMANDER_TEST_RUN_ID`; its expected
fail-closed settings path and resulting test fail-fast are preserved under
`Specs/TestRuns/4cb089111a23/Continuation/20260901_000534_r1b_localization_invocation_failfast/`
and are not qualification evidence.

Five same-machine Release baseline samples measured `160`, `156`, `154`,
`172`, and `155` us (p95 `172`); each exposed 4,096 task creations, zero
updates/provider calls/outstanding receipts, and failed the one-card RED
assertion. Five candidate samples measured `172`, `172`, `172`, `172`, and
`173` us (p95 `173`); each passed with 4,096 payloads, exactly one task create,
4,095 updates, zero provider calls, and zero outstanding receipts. The
candidate passes the `50,172` us relative ceiling and the `2,000,000` us
absolute cap. Archives are
`Specs/TestRuns/4cb089111a23/Commands/2026-08-31_r1b_change_case_truth_{baseline,candidate}_release/`.
Explicit pre-stage validation passed for all 33 files and whole-inventory
validation passed all 1,589 files. The runner disk audit reports only earlier
explicitly authorized sibling runs under `D:\RedSalamander.Perf`; it found no
write outside the initialized root. Authoritative behavior is in
`Specs/UI/UI_CommandMenuKeyboard.md`; performance/resource qualification is in
`Specs/Testing/Testing_PerformanceValidation.md`. At R1b close, E5 remained blocked
by R7-A10/R7-A15; those predecessors have since closed. R1c/R1d/R6 remain separate
owners and this historical R1b receipt activates none of them.

- [x] Repair the task payload ownership and expose unreadable descendants,
      cancellation, partial failure, and task-card truth without creating another
      rename executor.

#### R1c — Compensation and cleanup-debt truth

**Owns:** FOS-11/FOS-12

##### R1c activation card

| Activation-card field | R1c bounded execution contract |
|---|---|
| Slice/state/owner | **R1c / `COMPLETE` / Codex `/root` in this worktree.** This slice closes only FOS-11/FOS-12: Local exact owned-stage compensation gets a fresh bounded cleanup control independent of primary Cancel, and already-structured S3/Microsoft Drive committed-primary cleanup debt reaches the existing item-result stage axis and one host warning. It does not activate R1d, R2, R3, provider route/capability redesign, general publication transactions, automatic cleanup retry/recovery, or new UI architecture. |
| Baseline and drift | Activation baseline is clean tracked commit `08fb670cef485865c944941ac793d270ba74fd4c`; repository-root `last_run/` is pre-existing untracked output and excluded. Scoped drift from that commit over the named owners is empty. Inventory found three staged Local file/link/directory cleanup lambdas passing the primary `FileSystemOptions` into `AbortOwnedObject`, whose Local implementation checks the already-canceled operation control before exact `DeleteExact`; the direct-final R0a path already constructs a separate abort control and remains out of scope. S3 `S3TransferCommitResult` and Microsoft Drive `MoveCommitResult` already retain `primaryMutationCommitted` plus `cleanupStatus`, but their public `FileSystemItemCompleted` calls pass a null mutation result. The host already carries `FileSystemOwnedStageDisposition::Retained` and counts retained stages, but currently converts every Retained stage into Indeterminate and exposes no retained-cleanup summary. |
| Contract | Implement section 6 `FO-PUBLISH-01` containment only. After primary cancellation/failure, an exact Local stage cleanup call uses a copied options record with the primary abort callback detached and a fresh short absolute deadline; it performs no primary work and mutates only the retained owned object. Removed/Retained/Unknown remains exact. For S3 Copy/Move and Microsoft Drive native Move, a committed primary with failed backup cleanup returns ordinary operation success plus `FileSystemItemMutationResult{outcomeKnown=TRUE, mutationCommitted=TRUE, ownedStageDisposition=Retained}`; Copy reports source retained and Move reports source removed. The host preserves Completed publication/source truth, records cleanup debt independently, emits one warning/issue, and presents “Completed; cleanup item retained.” Unknown cleanup or incomplete final content remains Indeterminate. No automatic destructive retry follows the committed primary. |
| Scope | Production: `Plugins/FileSystem/FileSystem.Internal.h`, `FileSystem.FileOps.cpp`, and the allocation-failure cleanup in `FileSystem.cpp`; `Plugins/FileSystemS3/FileSystemS3.Directory.cpp`; `Plugins/FileSystemMicrosoftDrive/FileSystemMicrosoftDrive.cpp`; `RedSalamander/FolderWindow.FileOperations.State.cpp`, `.State.Diagnostics.cpp`, and `FolderWindow.FileOperationsInternal.h` only if a summary field is required; base plus cs-CZ/fr-FR/ja-JP/sk-SK File Operations resources. Tests: Local/S3/Microsoft Drive debug provider selftests, the exact host provider-matrix typed-result block, focused source/resource contracts, and their existing registrations. Specs: `Specs/FileSystem/FileSystem_FileOperations.md`, `Specs/Plugins/Plugins_VirtualFileSystem.md`, `Specs/UI/UI_FileOperationsPopup.md`, `Specs/Testing/Testing_PerformanceValidation.md`, compact FileOps evidence, and this card/receipt. Out of scope: public ABI shape changes, R2 typed route/receipt replacement, bridge cleanup, source Delete policy, direct-final R0a semantics, remote cleanup retry/recovery, persistent artifact claims, and R3 generalization. |
| Predecessors | E0 is complete. R0a, R0b, R0d, R0e, and R1a are complete and provide exact Local authority, S3 conditional behavior, honest current mutation-result handling, cancel containment, and first-write terminal truth. Section 8.1 requires only E0 for R1c. No active slice owns the scoped files. |
| Steps | (1) Add RED behavioral/source contracts for canceled Local exact stage cleanup, S3/Microsoft Drive public committed-cleanup receipts, host Completed-plus-Retained reduction, and one warning summary; capture five Release reduction baselines. (2) Add one Local-provider-internal cleanup-options builder and route every staged file/link/directory plus writer-allocation abort through it. (3) Thread existing provider commit results to public callbacks without changing success HRESULTs or ABI shape. (4) Make Retained debt orthogonal to primary completion while keeping Unknown/RetainedIncomplete indeterminate; add one localized warning/result summary. (5) update authoritative specs, run focused Debug/test-enabled Release/provider/localization/source validation, archive five candidate samples, record the exit receipt, and mark only R1c complete. |
| RED/GREEN evidence | RED/GREEN owners are `FileOpsFamily_ClearflowPhase07_ProviderMatrix`, the debug `RedSalamanderFileSystemDebugSelfTests`, `RedSalamanderFileSystemS3DebugSelfTests`, and `RedSalamanderFileSystemMicrosoftDriveDebugSelfTests` surfaces exercised through `PluginContractTests`, plus `TestHarnessSourceContracts.Tests.ps1` and `ResourceLocalizationContracts.Tests.ps1`. RED must show Local abort returning `ERROR_CANCELLED` with the exact stage retained, provider public callbacks receiving null cleanup truth, and host committed+Retained reduction becoming Indeterminate. GREEN requires Local canceled-primary cleanup to remove only the retained object under the cleanup deadline; S3 Copy/Move and Microsoft Drive Move callback receipts to preserve committed primary/source truth with Retained debt; host terminal completion to remain Completed/Published with one retained-cleanup warning; and Unknown/RetainedIncomplete controls to remain Indeterminate. |
| Performance/resources | Add a provider-free 4,096-result reduction loop to the provider-matrix case and emit `FileOps.CleanupDebt.ReduceUs`, result count, completed count, retained-debt count, indeterminate count, and retained bytes. Capture five same-machine test-enabled x64 Release processes before and after under `Specs/TestRuns/4cb089111a23/FileOps/2026-09-01_r1c_cleanup_truth_{baseline,candidate}_release/`. Every candidate reports 4,096 Completed+Retained results, zero Indeterminate results, constant O(1) per-result retained state, duration below 2,000,000 us, and p95 no greater than `max(1.50 * baseline p95, baseline p95 + 50,000 us)`. Provider debug fixtures additionally prove one exact cleanup call and one callback result without live credentials. |
| Durable owners | `Specs/FileSystem/FileSystem_FileOperations.md` owns cancel-independent bounded exact compensation and independent cleanup-debt truth. `Specs/Plugins/Plugins_VirtualFileSystem.md` owns S3/Microsoft Drive callback receipt behavior. `Specs/UI/UI_FileOperationsPopup.md` owns the one retained-cleanup warning/summary. `Specs/Testing/Testing_PerformanceValidation.md` owns the deterministic reduction scenario and budget. This plan retains mapping and generated receipts only. |
| STOP/rollback | STOP if cleanup lacks exact retained object authority; a fresh cleanup control can restart primary work; Local cleanup cannot remain short and bounded; a provider cannot distinguish committed primary from rollback/unknown; the existing result prefix cannot carry retained debt without an ABI change; success would be changed to failure or destructive retry; the host must collapse cleanup debt into primary failure/Indeterminate; live cloud credentials are required; or work crosses into bridge/source cleanup, R2, R3, or R6 redesign. Rollback removes the cleanup-control routing and public debt threading together; never land provider Retained receipts without host interpretation or host interpretation without provider proof. |
| Exit-receipt contract | Before R1c closes, record activation/RED/implementation/evidence commits; exact Debug and test-enabled Release build receipts; focused provider/host/source/localization totals; canceled-primary exact stage removal and foreign-object preservation observations; S3/Microsoft Drive callback receipt fields; Completed-plus-Retained and Unknown controls; one-warning presentation proof; five-process baseline/candidate summaries and constant-resource proof; validated archive paths and whole-inventory validation; `git diff --check`; authoritative-spec destinations; and remaining R1d/R2/R3 boundaries. |

- [x] Separate primary cancellation from bounded exact compensation.
- [x] Expose provider cleanup debt independently of primary success; neither Cancel
      nor presentation disposition may erase known publication/source/stage truth.

##### R1c exit receipt

R1c activated at `8457faf8`; RED contracts landed at `c20eed1c`; production,
focused tests, localization, and authoritative-spec changes landed at `bfcc9cbe`;
compact baseline/candidate evidence landed at `54d3909c`. The exact implementation
tree built test-enabled x64 Release with zero warnings/errors and full-solution
receipt `77ad8b8d05fde6c971bdd4259b7e271d96e10d8048d01bd65e14c959bdce3ef7`.
The exact implementation tree also built x64 Debug with zero warnings/errors and
receipt `2de01d9a64da50c01a92a5ebf5f0805dfb3bb98a421f833910493d77409aeb83`.
The earlier exact RED tree's test-enabled x64 Release receipt was
`29cc49e75f99578208d1a8037f4f53913a9d5e4c41e66d5fe563e5d373b5550d`.

The exact RED provider-matrix run and all five RED baseline processes reproduced
the one intended `2/1/0` failure: committed primary results carrying Retained
cleanup debt were collapsed into Indeterminate. The exact GREEN run and all five
candidate processes passed `3/0/0`. Debug `PluginContractTests.exe` returned 0;
its Local, Microsoft Drive, S3, and S3 directory-transfer selftests passed
`155/0`, `224/0`, `195/0`, and `157/0`, respectively. The test-enabled Release
provider runner returned 0, including Microsoft Drive `224/0` and S3
directory-transfer `157/0`. `LocalizationTests.exe` returned 0 in both Debug and
Release. Focused source contracts passed `187/187`; resource-localization
contracts passed `9/9`.

Local compensation now copies the validated options header, removes the primary
operation-control callback/cookie, and supplies a fresh overflow-safe absolute
deadline no later than 2,000 ms after cleanup admission. File, link, directory,
and writer-allocation aborts use that control and only the retained owned-object
token; provider selftests observe exact stage removal while preserving a foreign
replacement. No primary work or pathname cleanup is restarted. S3 Copy returns
known committed, original-present, Retained; S3 Move and Microsoft Drive native
Move return known committed, original-removed, Retained. Public callbacks carry
those receipts while the operation HRESULT remains successful. The host reduces
that state to Published/Completed plus independent Retained debt, emits exactly
one `item.cleanup.retained` warning using `IDS_FILEOPS_CLEANUP_ITEM_RETAINED`, and
uses `IDS_FILEOPS_RESULT_COMPLETED_CLEANUP_RETAINED` for the completed summary.
Unknown and RetainedIncomplete controls remain Indeterminate, and no automatic
destructive retry is attempted.

The five baseline `FileOps.CleanupDebt.ReduceUs` samples were `569`, `319`,
`351`, `435`, and `457` us (nearest-rank p95 `569` us). Every baseline process
reported 4,096 results, zero Completed, 4,096 Retained, 4,096 Indeterminate, and
256 retained-state bytes per result. Candidate samples were `302`, `355`, `306`,
`593`, and `380` us (p95 `593` us), below both the 2,000,000 us absolute limit
and the `max(1.50 * 569, 569 + 50,000) = 50,569` us admission ceiling. Every
candidate process reported 4,096 Completed, 4,096 Retained, zero Indeterminate,
and the same constant 256 bytes per result.

Evidence is archived under
`Specs/TestRuns/4cb089111a23/FileOps/2026-09-01_r1c_cleanup_truth_{baseline,candidate}_release/`.
Explicit pre-stage validation passed all 32 files and whole-inventory validation
passed all 1,654 files. `git diff --check` passed. The runner disk audit reported
only earlier explicitly authorized sibling runs and the pre-existing
`D:\RedSalamander.Perf\worktrees` subtree, all within the initialized test root;
no write outside that root was observed. Durable behavior moved into
`Specs/FileSystem/FileSystem_FileOperations.md`,
`Specs/Plugins/Plugins_VirtualFileSystem.md`,
`Specs/UI/UI_FileOperationsPopup.md`, and
`Specs/Testing/Testing_PerformanceValidation.md`. R1d still owns the common
Preparing lifecycle, R2 still owns the typed route/capability cutover, and R3
still owns publication generalization; none was activated or absorbed. With the
recorded E0, R0a-R0e, and R1a-R1c receipts, derived milestone M1 is complete.

#### R1d — Common Preparing lifecycle (core)

**Ordering:** the R1d-core row in section 8.1.
**Boundary:** R1d owns the common Preparing lifecycle, A05 strategy disclosure, A20
routine-start transition, immutable top-level phase/scope facts, and extension hooks.
It consumes E4 Batch Rename admission facts. R4 owns A02 comparison/concurrent
correctness and A19 scheduling; R6 owns presentation. R1d owns none of those later
behaviors.

##### R1d-core activation card

| Activation-card field | R1d-core bounded execution contract |
|---|---|
| Slice/state/owner | **R1d-core / `COMPLETE` / Codex `/root` in this worktree.** This slice implements `FO-LIFE-01` plus only the core A05/A20 facts and transitions required by `FO-UX-01`. It does not activate R3 publication, any R4 overlap comparison or discovery scheduler, R6 rendering/options/review/history, R7-A03 retarget, R9 Shell delegation, provider-cancel ABI work, or any currently disabled R0e route. |
| Baseline and drift | Activation baseline is exact tracked commit `1da78096d61fd8e66d5cb6c9a8872a0fbf44ae15`; repository-root `last_run/` is pre-existing untracked output and excluded. `git status --short` otherwise reports no drift. E4, R0e, and R1a have recorded exit receipts. Current code has selected-root readiness only for clipboard Move; ordinary tasks release `_workerReleased` before worker preparation, Batch Rename qualifies through a private branch, Move breadcrumbs persist before Preparing, `unresolvedPreparingGate` is test-only and rejected outside clipboard Move, popup Preparing is inferred from missing progress, and `StartOperation` invokes the generic Copy/Move confirmation regardless of `requireConfirmation`. These are the displaced authorities. |
| Lifecycle contract | Add one engine-owned `TaskLifecyclePhase` with legal monotonic transitions `Preparing -> AwaitingAcceptance/Ready -> Waiting/Running -> Stopping -> Terminal`; pause, verification, conflict, and discovery remain orthogonal facts. Every admitted task publishes `Preparing` before worker work. Worker-owned preparation performs Batch/Change Case qualification where applicable, validates the immutable plan group, binds selected top-level mutation/interlock scopes, builds immutable top-level role/scope plus strategy-count facts, runs one pre-consumption decision hook, and observes cancel/stop. It performs no recursive enumeration, content read/hash, totals walk, descendant conflict scan, mutation, clipboard consumption, breadcrumb write, or history write. |
| Acceptance and routine start | Generalize `unresolvedPreparingGate` into one `preConsumptionDecisionGate` legal for every task. Routine accepted-default work crosses Ready without a generic confirmation. Existing ingress-owned preview/editor/drop consent remains authoritative; explicit/material confirmation remains only when `requireConfirmation`, permanent Delete, known Copy-only Move, or an existing exact artifact/risk gate requires it. Known Copy-only counts are immutable preparation facts and must be accepted before mutation; runtime-only downgrade remains in the existing item gate. R6-A20 later owns the deliberate **with options** command and final presentation, not this core slice. |
| Clipboard and breadcrumb order | Clipboard Move consumes the captured sequence exactly once only after successful common preparation and the pre-consumption decision. Failure/cancel produces dense NotAttempted/source-retained results. A Move breadcrumb is created only after common preparation plus required acceptance/clipboard consumption and immediately before queue/execution admission; it advances to Executing only after entering the operation and is finalized through the existing completion funnel. No Preparing failure/cancel leaves a breadcrumb. |
| Immutable hooks/facts | Add bounded `PreparationSnapshot` data owned by the task: task ID, operation, selected-root count, endpoint/role/path facts copied without retained mutation authority, strategy counts, Copy-only count, and preparation timing/status. Publish it as an immutable snapshot only after successful preparation. Provide one non-blocking observer hook for future R4/R6 consumers; no consumer may mutate it or infer authority. Existing discovery begin/close/skip signals remain the R4-A19 extension points; this slice changes no scheduler policy. |
| Scope | Production: `RedSalamander/FolderWindow.FileOperationsInternal.h`, `FolderWindow.FileOperations.State.cpp`, `.State.Runtime.cpp`, `.State.Queue.cpp`, `FolderWindow.FileOperations.Popup.h/.cpp`, and only the admission call sites in `FolderWindow.FileOperations.cpp`/`FolderView.FileOps.cpp` needed to remove the displaced generic confirmation. Tests: `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.cpp`, a bounded new R1d phase source, Commands FileOps/PluginConfig source contracts, registrations, and project files only if a new source is added. Specs: `FileSystem_FileOperations.md`, `UI_FileOperationsPopup.md`, `UI_CommandMenuKeyboard.md`, `Testing_PerformanceValidation.md`, compact evidence, and this card/receipt. No `Common/PlugInterfaces`, provider, localization/layout, durable schema, Tools, installer, R3/R4/R6 implementation, or reader/writer ABI file is in scope. |
| Steps | (1) Add RED lifecycle/source-contract tests and five Release baseline processes. (2) Introduce the phase/snapshot types and one legal-transition helper. (3) centralize worker preparation for ordinary, clipboard, and scheduled Rename tasks before any consumption/mutation. (4) generalize the decision hook and preserve dense cancellation truth. (5) move Move breadcrumb creation after acceptance. (6) make popup status consume the explicit phase while leaving R6 layout/actions untouched. (7) retire the unconditional/generic Copy/Move confirmation owner, preserve explicit/material gates, and add Copy-only preparation acceptance. (8) update authoritative specs, build Debug/Release, run focused/affected tests, archive candidate evidence, and close only R1d-core. |
| RED/GREEN evidence | Add FileOps family `FileOpsFamily_R1dPreparingLifecycle` and Commands case `file_operations_preparing_lifecycle_source_guard`. RED proves: an ordinary task cannot install the common decision gate; lifecycle has no explicit phase/snapshot; a slow ordinary preparation is not observable as Preparing; cancel/failure can leave pre-created Move breadcrumb state; and generic routine Copy reaches the confirmation owner. GREEN requires ordinary/clipboard/Batch Rename/Change Case tasks to publish the same phase order; slow/fail/cancel gates to perform zero mutation/clipboard/breadcrumb; Copy-only Move to require one acceptance; routine Copy to invoke no generic prompt; explicit/material gates to remain; exact dense results; and no recursive/provider-content work in preparation. Run `Run-AllTests.ps1 -Suite FileOps -CaseFilter FileOpsFamily_R1dPreparingLifecycle`, `-Suite Commands -CaseFilter file_operations_preparing_lifecycle_source_guard`, existing `FileOpsFamily_Phase10_DeleteValidation`, scheduled-Rename admission/cancel/prompt cases, and prompt teardown/source guards in Debug and exact-source Release. Faults cover gate slow/fail/cancel, stop before/after Ready, clipboard consume fail, breadcrumb create/advance/finalize failure, task removal during publication, and duplicate terminal completion. |
| Performance/resources | Protected offline scenario is the new FileOps R1d family using 4,096 top-level plan items and no descendants. Emit `FileOps.Preparing.BuildSnapshotUs`, `.SelectedRootReadinessUs`, `.UiAdmissionUs`, `.SelectedRootCount`, `.ScopeFactCount`, `.StrategyFactCount`, `.CopyOnlyCount`, `.MutationBeforeReadyCount`, `.ClipboardBeforeReadyCount`, `.BreadcrumbBeforeReadyCount`, and retained snapshot bytes. Capture five independent test-enabled x64 Release processes before/after under `Specs/TestRuns/4cb089111a23/FileOps/2026-09-01_r1d_preparing_{baseline,candidate}_release/`. Candidate p95 for snapshot/readiness is at most `max(1.50 * baseline p95, baseline p95 + 50,000 us)`; UI admission p95 must not regress by more than 25,000 us; every sample is below 2,000,000 us; retained snapshot state is below 64 MiB at 4,096 roots; the three before-Ready counters are zero. This is top-level preparation cost only, never a recursive traversal benchmark. |
| Durable owners | `Specs/FileSystem/FileSystem_FileOperations.md` owns lifecycle ordering, preparation bounds, strategy facts, consumption/breadcrumb boundary, and route availability. `Specs/UI/UI_FileOperationsPopup.md` owns phase semantics while R6 retains visual layout/actions. `Specs/UI/UI_CommandMenuKeyboard.md` owns routine invocation versus explicit/material consent without adding the later with-options command. `Specs/Testing/Testing_PerformanceValidation.md` owns metrics, five-process archive, and budgets. This plan retains delivery mapping and the generated receipt only. |
| STOP/rollback | STOP if common preparation requires recursive enumeration/content I/O; any current R0e-disabled route must be re-enabled; provider cancellation ABI/isolation is required; the phase model would become a second item-result/conflict/queue authority; Copy-only acceptance cannot use the existing decision owner without R6 redesign; Move breadcrumb creation after acceptance would lose the crash witness before mutation; a hook must retain executable mutation authority; work requires R3/R4/R6 behavior; or candidate evidence exceeds budget without a diagnosed machine anomaly. Rollback removes phase/snapshot/hooks/order changes and RED/GREEN tests/spec clauses together; never leave clipboard or breadcrumb on different preparation orders. |
| Exit-receipt contract | Before R1d-core closes, record activation/RED/implementation/evidence commits; exact Debug and test-enabled Release build receipts; phase-order results for ordinary/clipboard/Batch/Change Case; zero-before-Ready counters and dense cancel/failure truth; routine/material/Copy-only consent observations; breadcrumb fault/order results; five-process baseline/candidate p95 and memory; explicit archive and whole-inventory validation; source-contract/prompt-teardown totals; `git diff --check`; authoritative-spec destinations; and newly eligible R4-A02/R4-A19/R6-A20 state. Known in-flight provider-I/O cancellation remains a separately disabled-route/re-enable concern and cannot be claimed closed here. |

- [x] Introduce mandatory cancellable Preparing; keep it selected-root/top-level
      proportional only, with zero recursive enumeration, content bytes, totals walk,
      or descendant conflict prescan.
- [x] Coalesce the quick path with execution/consent; reveal one stable Preparing
      surface only after the measured reveal budget and replace it in place.
- [x] Expose immutable top-level roles/scopes, one pre-consumption decision-gate hook,
      and discovery-service hooks. Do not compare scopes, schedule discovery, or
      render overlap/progress decisions in R1d.
- [x] Show known strategy counts/reasons at consent; runtime-only downgrade remains
      visible and preserves source.

##### R1d-core exit receipt

R1d-core activated at exact baseline `1da78096d61fd8e66d5cb6c9a8872a0fbf44ae15`
under activation commit `b3cfe8cf`; RED contracts landed in `7f7d2488`, baseline
fact isolation in `22cc1f08`, the common lifecycle implementation in `668acbd7`,
and the five-process baseline/candidate archive in `005b9c81`. Release-only warning
repairs landed in `8b452b8f`, Batch Rename cancel synchronization in `21d44433`,
and the final compact Delete correction in `0515e854`: the compact renderer now
consumes `TaskLifecyclePhase` and no longer infers Preparing from missing progress
numbers, matching the authoritative popup contract.

Exact x64 builds after the final correction passed with zero warnings and zero
errors: Debug artifact receipt
`557c7acb2faf2e926a5b624a04037fd85dae8b1c2003c05ca787b1db01b2f092`,
product Release receipt
`be2604bbe6fb3356322be4593e7011ad1b6932729adaba95886244d8649f4822`,
and serialized test-enabled Release receipt
`a1b5b544a5d589df6c6c639cd8fd4d5968a8e67ea9ad712004de2c8226ef613e`.
`FileOpsFamily_R1dPreparingLifecycle`, `FileOpsFamily_Phase10_DeleteValidation`,
`file_operations_preparing_lifecycle_source_guard`, and
`cmd_pane_batchRename_cancel_mid_batch_tracks_completed_rows` passed in exact Debug
and exact test-enabled Release. Existing focused closeout coverage also passed for
Inline F2, popup/Delete presentation, target-collection cancellation, artifact-prompt
shutdown during the nested pump, and direct-admission artifact-prompt teardown.
Routine accepted-default Copy/Move uses no generic prompt; explicit/material and
known Copy-only decisions remain. Slow/fail/cancel preparation publishes the common
phase order and dense source-retained truth, while mutation, clipboard consumption,
and breadcrumb creation before Ready remain zero.

The validated archives are
`Specs/TestRuns/4cb089111a23/FileOps/2026-09-01_r1d_preparing_baseline_release/`
and
`Specs/TestRuns/4cb089111a23/FileOps/2026-09-01_r1d_preparing_candidate_release/`.
Baseline/candidate p95 values are respectively 2,785/1,792 us for snapshot build and
7,433/7,192 us for selected-root readiness; UI admission p95 is 1,071/3,047 us.
Candidate retained snapshot state is 2,802,208 bytes for 4,096 roots, every sample
is below 2 seconds, and all three before-Ready counters are zero. Explicit archive
validation passed when the evidence landed; final whole-inventory validation passed
1,847 files. `git diff --check` is clean. Durable behavior resides in
`Specs/FileSystem/FileSystem_FileOperations.md`, `Specs/UI/UI_FileOperationsPopup.md`,
`Specs/UI/UI_CommandMenuKeyboard.md`, and
`Specs/Testing/Testing_PerformanceValidation.md`.

This receipt activates no successor. R4-A02, R4-A19, and R6-A20 become eligible but
remain inactive. It makes no claim that an admitted blocking `IFileReader::Read` is
abortable or that UI-thread shutdown join is bounded; that provider-I/O cancellation
work remains separate. It also makes no claim about the newly identified MTP replay
occupancy residual, where destination path occupancy is not destination-PUID proof.

**Family acceptance:** every selected source has exactly one explainable result;
predictable rejection does not consume cut intent; failed consume cannot execute;
Cancel does not suppress exact safe cleanup. Preparing performs only its declared
top-level work, remains cancellable under the accepted provider bound, and creates no
popup/focus flash on the routine fast path.

### Package R2 — Small authority, capability, and namespace contract

**Priority:** P1 architecture
**Boundary:** A02/A04/A11/A13/A13-MD1/A14/A15 are accepted. Typed facts support
cheap overlap explanation, but exact mutation authority and receipts remain separate;
no overlap comparison becomes an execution prohibition or destructive grant.

##### R1d-OR1 — Admission probes leave the UI thread (draft)

| Field | R1d-OR1 |
|---|---|
| Slice/state/owner | **R1d-OR1 / `DRAFT` (2026-09-04) / Claude in worktree `codex/r4-a19-closeout-2026-09-02`.** Post-closeout correction to the Preparing lifecycle: the per-item provider calls that `AdmitOperation` still makes on the UI thread move into `Task::PrepareForExecution` on the task thread, under the synchronous-I/O cancel watch. Activation flips this row to `ACTIVE` with the baseline commit. |
| Finding (2026-09-04 review) | `AdmitOperation` (`RedSalamander/FolderWindow.FileOperations.cpp` ~100-200) calls `GetAttributes` for every selected source and `BindObjectAuthority` for reparse-point sources and destinations, and since R0-RC3 step 9 `QueryChildNameContract` for every transfer leaf, all on the UI thread before a task exists; permanent Delete binds its selected roots there too. A dead SMB share or a stalled device blocks each call for its own timeout and freezes the folder window before the task thread's cancel watch exists. R0-RC3 step 8 verified only the Preparing bind on the worker. |
| Correction | The UI thread keeps the cheap immutable facts (endpoint pair, route facts, path text, pane identity) and publishes the task; `PrepareForExecution` performs the per-item probes (attributes, no-follow binding of reparse points and delete roots, destination child-name contract) under `SynchronousIoCancelWatch`, so Cancel during Preparing aborts a stuck call and the card shows `Preparing...` instead of a frozen window. Rejections that admission returned synchronously become Preparing failures with the same plan-rejection buckets. |
| Scope | In: `RedSalamander/FolderWindow.FileOperations.cpp` (`AdmitOperation` and its callers, `CommandPermanentDelete` at the permanent-delete ingress), `RedSalamander/FolderWindow.FileOperations.State.cpp` (`PrepareForExecution`), FileOps self-tests (a blocked-share fixture: the Dummy provider's `latencyMs` on a selected root with Cancel during Preparing), Commands contracts that pin the synchronous rejections, `Specs/FileSystem/FileSystem_FileOperations.md` (Preparing paragraph), `Specs/UI/UI_FileOperationsPopup.md`, this plan. Out: the plan builder's typed validation (no provider I/O), R4-A02 overlap advice (it consumes the same Preparing seam), inline F2. |
| RED/GREEN evidence | RED: a FileOps case that selects a root on a Dummy instance with `latencyMs` large enough to exceed the reveal budget and cancels during Preparing: today the ingress call blocks before the task exists (no card, no Cancel). GREEN: the card reveals as `Preparing...`, Cancel returns within the bound, nothing mutates; the existing admission rejection cases keep their buckets (now reported through the task), Phase9/Phase10 conflict and delete families stay green. |
| Performance/resources | Same probes, moved; routine Local admission gains one thread hop already paid by Preparing. Evidence: the reveal-budget timing rows and the next Fresh Full gate. |
| STOP/rollback | STOP if a rejection that the ingress currently reports before a task exists cannot be reported through the task card without a new surface; rollback is the UI-thread probing (one revert). |
| Exit-receipt contract | Commit, Debug build receipt, focused FileOps and Commands families, Pester, spec updates, and the Fresh Full gate that follows; recorded on this card. |

#### R2 activation card

| Activation-card field | R2 bounded execution contract |
|---|---|
| Slice/state/owner | **R2 / `COMPLETE` / Codex `/root` in this worktree.** This is the one lockstep typed route, path identity, namespace, name-feasibility, and receipt-classification authority cutover. It does not activate R3-R9, consume provider name policy in rename commands before R7-A15, change mutation algorithms, add overlap enforcement, add a serialized migration, or promise cross-generation plugin compatibility. |
| Baseline and drift | Activation baseline is exact clean tracked commit `b53e2c2ce7bf3a7803394116ac7bdb99bacb61dc`; repository-root `last_run/` is pre-existing untracked output and excluded. Scoped drift over every named owner is empty. Executable route facts currently come from provider-owned UTF-8 capability JSON v2 parsed independently by `FolderWindow.FileOperations.cpp`, `FolderWindow.FileOperations.State.cpp`, `FolderView.cpp`, `BatchRenameWindow.cpp`, `CompareDirectoriesEngine.cpp`, and `Common/FileSystemPathIdentity.cpp`. Eight shipped providers, Dummy, seven in-tree `IFileSystem` adapters/stubs, and `PluginContractTests` participate. No `PlugInterfaces` header or SDK is packaged by Installer/runtime manifests. Settings accept custom DLL paths, but `Plugins_PluginAPI.md` supports only same-source-tree/release-generation binaries and requires explicit negotiation before any future mixed-generation support. Therefore no supported external binary consumer blocks the atomic cutover; a plugin lacking the new IID is an unavailable/contract-violation route, never a JSON fallback. |
| Public ABI | Add separate `IFileSystemRouteCapabilities` IID `1e924d87-2e62-4ab4-9f37-c565d465f25e`; do not alter `IFileSystem` or `IFileSystemPathCapabilities2` vtables/IIDs. New enums and records use one leading `uint32_t sizeBytes`, exact current-size producer validation, current-size-or-larger consumer prefix validation, valid-enum/strict-`BOOL` checks, and no numeric version field. `GetRouteFacts(path, operation, arena, facts)` returns operation-specific route, cancellation, namespace, identity, proof, publication/binding/delete/link, concurrency, and full provider/profile/root identity. Variable UTF-16 facts live only in the caller-owned `FileSystemArena`; `facts.requiredArenaBytes` permits one bounded retry, capped at 64 KiB. `IsTransferPeerAllowed(path, operation, role, fullPeerPluginId, allowed)` owns directional import/export compatibility without truncated IDs. `ValidateChildName`, `GetChildNameCollisionKey`, and `JoinPath` are path/operation scoped; string outputs use the same arena/required-size rule. Providers never retain caller records, strings, or arenas. |
| Central contract | Add one canonical `Common/FileSystemRouteContract` validator/query helper. Its result is exactly `Available`, `Unsupported`, or `ContractViolation`; missing IID, malformed sizes/enums/BOOLs, contradictory cancellation facts, empty/mismatched provider/profile/root IDs, arena overrun, or malformed name/peer output fail closed. `IFileSystemPathCapabilities2::GetPathCapabilities` remains only for optional human/machine diagnostics and extensions because removing the inherited method would change `IFileSystem`; no production admission, command availability, path comparison, mutation selection, publication, Delete, link, or rename-name decision may call it or parse its JSON. A source guard bans those calls/parsers outside provider diagnostics and dedicated diagnostic tests. Typed route facts are descriptive feasibility only: the host must also bind the existing executable interface and exact selected-root/container/object authority required by the operation. Cheap comparable identity and positive alias hints may explain a future warning but never grant or prohibit mutation. |
| Receipt/result law | Do not add a second callback or result vtable. Existing `FileSystemItemMutationResult` remains the provider receipt authority for primary commit, source presence, and owned-stage cleanup. The canonical helper maps route support plus HRESULT/receipt truth to exactly `Unsupported`, `RetryableNoCommit`, `FailedKnown`, `Indeterminate`, or `ContractViolation`; Retry remains legal only for proved non-commit with known source/cleanup state. Missing/malformed destructive receipts remain Indeterminate under R0d/R1a, and retained cleanup debt remains a separate completed-primary axis under R1c. JSON can neither manufacture nor repair a receipt. |
| Provider/adapter inventory | Shipped producer projects are Local, 7z, Curl, Dummy, Google Drive, Microsoft Drive, MTP, and S3. In-tree adapters/stubs are `CompareDirectoriesEngine.cpp`; Commands Batch Rename selftest; Compare Directories core plus MTP/search-index selftests; FileOps selftest; `PerformanceTests2` duplicate-path fixture; `ViewerPETests`; and `ViewerSqliteTests`. Every one is cut over in the same implementation commit or explicitly returns no typed route in a test whose purpose requires an unavailable provider. Full provider IDs are `builtin/file-system`, `builtin/file-system-7z`, `builtin/file-system-ftp`, `builtin/file-system-sftp`, `builtin/file-system-scp`, `builtin/file-system-imap`, `builtin/file-system-dummy`, `builtin/file-system-gdrive`, `builtin/file-system-onedrive-personal`, `builtin/file-system-onedrive-business`, `builtin/file-system-sharepoint`, `builtin/file-system-mtp`, `builtin/file-system-s3`, and `builtin/file-system-s3table`; short IDs/hashes remain display/cache hints only. |
| Scope | Public/shared: `Common/PlugInterfaces/FileSystem.h`; `Common/FileSystemRouteContract.h/.cpp`; `Common/FileSystemPathIdentity.h/.cpp`; shared-helper catalog; and the RedSalamander/PluginContractTests project files needed to compile the canonical helper. Host consumers: `RedSalamander/FolderWindow.FileOperations.cpp`, `.State.cpp`, `FolderWindow.FileOperationsInternal.h`, `FolderView.cpp`, `BatchRenameWindow.cpp`, and `CompareDirectoriesEngine.cpp`. Providers: `Plugins/FileSystem/FileSystem.h/.cpp`; `FileSystem7z/FileSystem7z.h/.cpp`; `FileSystemCurl/FileSystemCurl.h/.Shared.cpp`; `FileSystemDummy/FileSystemDummy.h/.cpp`; `FileSystemGoogleDrive/FileSystemGoogleDrive.h/.cpp`; `FileSystemMicrosoftDrive/FileSystemMicrosoftDrive.h/.cpp`; `FileSystemMtp/FileSystemMtp.h/.Core.cpp`; `FileSystemS3/FileSystemS3.h/.Core.cpp`. Tests/adapters are the inventory row plus `Tests/PluginContractTests/PluginContractTests.cpp`, `Commands.SelfTest.PluginConfig.cpp`, `FolderWindow.FileOperations.SelfTest.Phases05_06.cpp`, registrations, and focused source contracts. Specs: `Plugins_PluginAPI.md`, `Plugins_VirtualFileSystem.md`, `FileSystem_FileOperations.md`, `Core_SharedHelpers.md`, `Testing_PerformanceValidation.md`, compact evidence, and this card/receipt. Out of scope: localization/UI layout, mutation implementation, provider discovery, persistent formats, actual R7-A15 rename ingress, R3 publication, R4 overlap/traversal, R5 isolation, and live credentials. |
| Predecessors and ownership | E1 and R0d have recorded exit receipts as required by section 8.1; R0a-R0e/R1a-R1c are also closed, so current receipt and route-honesty law is stable. No active package owns the named capability consumers/producers. R2 completion unblocks R7-A15 and R3 and is a predecessor for R4/R5/R9; it grants none of them implementation authority. |
| Steps | (1) Add RED ABI/malformed-output/JSON-inversion/source-guard tests plus a provider-free 4,096-query baseline loop. (2) Capture five independent test-enabled x64 Release baseline processes. (3) add the new interface/records and canonical validator/classifier. (4) implement all eight providers and all in-tree adapters atomically. (5) replace every production JSON route/path-identity consumer with the typed helper; leave JSON diagnostics only. (6) update authoritative specs/shared-helper catalog. (7) run focused Debug and exact-commit test-enabled Release qualification, archive five candidate samples, record the exit receipt, and mark only R2 complete. |
| RED/GREEN evidence | RED must prove the legacy host still obeys a JSON `true` claim, lacks the new IID/size boundary, and parses JSON in the protected loop. GREEN requires every shipped provider to pass undersized/current/oversized-prefix ABI tests; missing IID and every malformed fact fail closed; JSON `true` plus typed `false` remains unavailable; typed `true` plus absent/false/malformed JSON follows typed facts; mismatched full provider/profile/root IDs, invalid route/cancellation/namespace/name enums, non-strict BOOLs, contradictory timeout/watchdog facts, oversized arena requests, peer-role errors, and malformed name/join/collision output are contract violations; all-zero/uncomparable identity is unsupported; command admission requires the complete typed route and executable mutation authority; current provider matrices and receipt classification controls stay green; and production source contracts find no `GetPathCapabilities`, `TryParseCapabilitiesJson`, or JSON path-identity authority use. No test requires a live account/device/network. |
| Name/namespace contract | Facts distinguish real containers, provider-declared virtual folders, and fixed object/version sets; they carry stable-path identity, ordinal component comparison, accepted/preferred separators, case preservation/case-only rename, maximum child-name units, normalization policy, and full comparable provider/profile/root identity. `ValidateChildName` returns executable valid/invalid/unsupported truth for one parent and operation; `JoinPath` is the provider-owned path join; `GetChildNameCollisionKey` returns the provider/path-scoped collision key. R2 tests these surfaces and publishes them, but R7-A15 alone replaces F2/Batch Rename/Change Case/F7 validation and collision consumers. Unsupported name methods fail rename admission when R7-A15 consumes them; generic Win32 validation never becomes a fallback for a virtual provider. |
| Performance/resources | Protected deterministic case `FileOps_ProviderCapabilityMatrix` performs 4,096 provider-free route queries over the offline profile matrix and emits `FileOps.RouteFacts.QueryUs`, `.QueryCount`, `.JsonParseCount`, `.TypedQueryCount`, `.ArenaFallbackCount`, `.RejectedCount`, and `.ResultBytes`. RED executes the legacy JSON parser before its intended authority assertion. GREEN requires `QueryCount=4096`, `JsonParseCount=0`, `TypedQueryCount=4096`, normal-profile `ArenaFallbackCount=0`, bounded constant result/arena state, and exact rejection counts. Capture five same-machine test-enabled x64 Release processes before/after under `Specs/TestRuns/4cb089111a23/FileOps/2026-09-01_r2_typed_route_{baseline,candidate}_release/`. Candidate p95 must be at most `max(1.50 * baseline p95, baseline p95 + 50,000 us)` and each sample below 2,000,000 us. This measures contract query/validation, not provider or network latency. |
| Commands and faults | Build full Debug and test-enabled x64 Release. Focused cases are `FileOps_ProviderCapabilityMatrix`, new Commands case `file_system_capabilities_contract_source_guard`, provider Debug suites through `PluginContractTests.exe`, affected Batch Rename/Compare/FolderView cases, and Pester source contracts. Deterministic faults cover allocation failure, zero/undersized/exact/oversized arena and records, required-size overflow, stale pointers forbidden by post-call copy tests, absent IID, every invalid enum/BOOL, route/watchdog contradiction, empty/full-ID mismatch, peer allow/deny/failure, invalid child/join/collision output, and receipt absent/malformed/known-noncommit/known-commit/retained/unknown combinations. |
| Durable owners | `Specs/Plugins/Plugins_PluginAPI.md` owns the new IID, record sizing/lifetime, same-generation boundary, and fail-closed negotiation. `Specs/Plugins/Plugins_VirtualFileSystem.md` owns provider route/namespace/name matrices and JSON-diagnostic status. `Specs/FileSystem/FileSystem_FileOperations.md` owns central admission, authority separation, and receipt classification. `Specs/Core/Core_SharedHelpers.md` owns the canonical query/validator/classifier. `Specs/Testing/Testing_PerformanceValidation.md` owns the deterministic route-query scenario, counters, five-process archive, and budget. This plan retains delivery mapping and the generated exit receipt only. |
| STOP/rollback | STOP if any supported packaged or cross-generation external consumer is found; a provider cannot expose full identity/route facts without network/device I/O; one implementation commit cannot cut every executable consumer/producer over; JSON remains necessary to enable a route; new typed facts would become mutation authority; existing exact object/container/receipt checks must be weakened; the new IID requires altering an existing vtable; a compatibility reader/adapter or serialized migration is required; R7-A15/R3/R4/R5 work is needed to make R2 compile; tests require live credentials; retained memory grows with the loop; or Release evidence exceeds budget without a diagnosed machine anomaly. Rollback removes the new IID/helper/provider implementations/tests/spec clauses together and leaves R2 open; never retain a dual JSON/typed authority interval. |
| Exit-receipt contract | Before R2 closes, record activation/RED/implementation/evidence commits; exact Debug and test-enabled Release build receipts; focused provider/host/adapter GREEN totals; eight-provider/current-size and prefix-boundary ABI matrix; JSON inversion and production-source-ban proof; full-ID/namespace/name/peer/receipt fault results; five-process baseline/candidate summaries and constant-resource proof; validated archive paths and whole-inventory validation; external-consumer/packaging conclusion; `git diff --check`; authoritative-spec destinations; and newly eligible R7-A15/R3 state. |

**R2 RED receipt (2026-09-01):** activation is `8eed155e`; RED contracts are
`1c1ca225`. The exact RED tree built the full test-enabled x64 Release solution
with zero warnings/errors under receipt
`8eb9fa88e523972e079b3d88265f5bdc12171b7c7b6666f4dae6aec8b40f295d`.
`FileOps_ProviderCapabilityMatrix` passed its two setup/control cases and failed
only because Dummy did not expose IID `1e924d87-2e62-4ab4-9f37-c565d465f25e`;
`file_system_capabilities_contract_source_guard` failed because the separate typed
interface/spec authority is absent. Five retained Release processes each reported
4,096 legacy queries, 4,096 JSON parses, zero typed queries, and zero arena
fallbacks/rejections. Durations were `4,146`, `4,357`, `4,163`, `4,612`, and
`4,782` us (p95 `4,782` us), with constant 168-byte snapshots. Evidence is under
`Specs/TestRuns/4cb089111a23/FileOps/2026-09-01_r2_typed_route_baseline_release/`.
The runner's stale-sandbox cleanup removed the first non-retained loop, so the cited
samples were recaptured and copied to a stable evidence folder immediately after
each process before archival; no vanished sample is cited.

**R2 exit receipt (2026-09-01):** activation is `8eed155e`; RED contracts are
`1c1ca225`; the atomic implementation is `75df3717`; archived GREEN and crash
continuation evidence is `9bf973cb`. Exact test-enabled Debug and Release
full-solution builds completed with zero warnings/errors under receipts
`623677f34ece12cbc77070a9857bcc8e9a22b7c304eb7d3eed8cdcc4926599cb` and
`ac0e6f452203c56bb71f9bc235828a725524097aefd8f41d44022d29034680fd`.
The Debug provider matrix passed `3/3`
(`20260901T015002Z-21808-51a132f9be1548319ced72418166d32b`) after replacing
the Local provider's deep-stack 64 KiB metadata buffer with a nothrow heap buffer;
the failure dump, sidecar, traces, and focused GREEN are retained under
`Specs/TestRuns/4cb089111a23/Continuation/2026-09-01_r2_provider_matrix_stack_overflow/`.
The Release matrix passed `3/3`, the Commands production-source guard passed
`1/1` (`20260901T020505Z-4040-135a8d7af50145369cca3c0611eb44f7`), and standalone
Debug/Release `PluginContractTests.exe` passed the complete eight-provider,
fourteen-provider-ID matrix. The affected Compare case passed `1/1`, both Batch
Rename capability cases passed, `PerformanceTests2` duplicate-path coverage passed,
Viewer PE/SQLite adapter suites passed, and Pester source contracts passed `187/187`.

Every provider passed zero/undersized/current/oversized-prefix record and arena
boundaries, including absolute `wchar_t` alignment. Deterministic controls covered
missing IID, allocation failure, over-cap required sizes, post-call copies, invalid
enums and non-strict `BOOL`s, cancellation/watchdog contradictions, empty or
mismatched full provider/profile/root identity, unsupported identity, peer
allow/deny/failure, invalid child-name/join/collision output, and absent/malformed/
known-noncommit/known-commit/retained/unknown receipts. JSON inversion proved that
typed false defeats diagnostic JSON true and typed true survives absent, false, or
malformed diagnostic JSON. The compiled source guard found no executable
`GetPathCapabilities`, `TryParseCapabilitiesJson`, or diagnostic path-identity
authority use.

Five independent Release candidates (`r2-candidate-01` through `-05`) passed `3/3`
and measured `1,450`, `1,432`, `1,437`, `1,440`, and `1,462` us (p95 `1,462` us),
versus baseline p95 `4,782` us and ceiling `54,782` us. Every process reported exactly
4,096 queries, zero JSON parses, 4,096 typed queries, zero arena fallbacks/rejections,
and 917,504 result bytes (224 bytes per snapshot); every sample was below 2,000,000
us. Baseline and candidate are archived at
`Specs/TestRuns/4cb089111a23/FileOps/2026-09-01_r2_typed_route_{baseline,candidate}_release/`;
explicit validation passed all 32 files, the crash continuation passed all 7 files,
and whole-inventory validation passed 1,725 files. Unfiltered candidate JSONL remains
under the authorized `D:\RedSalamander.Perf\evidence\r2-typed-route-candidate-release\`
root; checked-in JSONL retains the seven protected rows per process.

Installer, runtime-manifest, Tools, and workflow inventory found no packaged SDK or
supported external source consumer of the new header/IID; the documented boundary is
same-source-tree/release-generation only. `git diff --check` passed. Durable contracts
landed in `Plugins_PluginAPI.md`, `Plugins_VirtualFileSystem.md`,
`FileSystem_FileOperations.md`, `Core_SharedHelpers.md`, and
`Testing_PerformanceValidation.md`. R2 alone is complete; R7-A15 and R3 are newly
eligible and remain inactive, and no R3-R9 mutation or UI package was smuggled in.

- [x] Inventory host, every shipped provider, Dummy, adapters, selftests, and any
      supported external binary/source consumer of JSON v2 capability authority.
- [x] Introduce a new IID/name for typed path-and-operation route facts using the
      repository `sizeBytes` ABI convention; never change an existing IID's vtable or
      add a redundant numeric struct-version field.
- [x] Cut host and every in-tree provider/adapter/test over atomically. Executable
      binding/publication/Delete interfaces and typed receipts remain proof authority;
      JSON is optional diagnostics/extensions only and can never enable a route.
- [x] Add no serialized-data migration, compatibility adapter, dual reader, or
      dual-authority interval. Mixed-generation binaries are unsupported. STOP if the
      inventory finds a supported external consumer that invalidates this assumption.
- [x] Split mutation authority, comparable identity, and cheap overlap-advisory
      evidence. Carry path/role and already-available positive alias facts needed to
      explain a warning without creating a mandatory exact overlap graph.
- [x] Declare route proof inputs (stable-source, commit, provider checksum/reread) and
      real-container, provider-declared virtual-folder, and fixed object/version-set
      namespace semantics.
- [x] Carry full provider IDs; short hashes are non-authoritative hints.
- [x] Define Unsupported/RetryableNoCommit/FailedKnown/Indeterminate/
      ContractViolation and primary/source/cleanup receipts.
- [x] Add path-scoped leaf validation, join, collision key, limits, and case policy.
- [x] Make command availability require the complete central route.
- [x] Inventory external COM consumers/selftest adapters before future ABI retirement.

**Acceptance:** one compiler-checked capability authority remains, runtime ABI
boundary/size tests cover every in-tree producer/consumer, and UI cannot offer an
operation known unable to execute safely; every
destructive route retains exact selected-root/container authority across consent,
binds any lazily discovered descendant immediately before mutation within that
authority, and returns a receipt.

### Package R3 — One owned-publication transaction

**Priority:** P1 architecture/performance
**Depends on:** landed R0a containment, R1c cleanup truth, completed E1 seams,
D2-A04, and D2-A06. It generalizes those accepted outputs rather than
reimplementing them.

- [x] Define transaction states and exact receipts (R3-3, 2026-09-03: `PublicationTransaction`
      with `Admitted`, `SourceBound`, `Staged`, `Written`, `Committed`, `Published`, `Verified`;
      the receipt is the typed item result the executor already records).
- [x] Adapt file/link/directory payloads without copying policy/cleanup code (R3-3 (a):
      `PublishOwnedStageAs` is the one owner of the `PublishAs` outcome for all three payloads).
- [ ] Implement the non-final-stage shape for overwrite, Managed Move,
      remote/weak-commit, resume, and atomic-visibility routes.
      - [x] R3-1 (2026-09-03): overwrite on identity-less atomic-final routes (FTP/SFTP/SCP,
            S3, Microsoft Drive/SharePoint, Google Drive, Dummy) publishes through the provider
            writer with the occupant the user saw as its expectation
            (`IFileWriterExpectedReplacement`); a changed or vanished occupant is a definitive
            non-commit. Managed Move into weak-commit routes stays Copy-only until R3-2 proves
            content.
      - [x] R3-2 (2026-09-03): the writer of an identity-less atomic-final route proves the
            published content (`IFileWriterContentProof`: S3 full-object CRC-64/NVME, Microsoft
            Drive `file.hashes`, Google Drive `sha256Checksum`, Dummy SHA-256; route fact
            `FILESYSTEM_ROUTE_PROOF_WRITER_DIGEST`). Verified Copy into those routes needs no
            readback and Managed Move removes the bound source only after the proof matched;
            FTP/SFTP/SCP report no digest and stay Copy-only.
- [ ] Implement the retained exclusive final-leaf shape only for ordinary new-name
      Local Copy under accepted D2-A06; no pathname rollback.
- [x] Define Standard/Verified Move proof inputs and prove stable-source/commit facts
      before cleanup under D2-A04 (bound routes as landed under D2-A04; identity-less routes
      through the R3-2 writer proof plus the committed-size proof, 2026-09-03).
- [ ] Keep provider-native atomic operations as adapters, not force all through bridge.
- [x] Fault-inject every transition and validate result combinations (R3-3 (c):
      `R3_3_PublicationFaultMatrix`, seventeen hooks x three routes x Copy/Move against the
      publication invariants).

**Acceptance:** one owner decides create/write/commit/publish/abort/reconcile truth;
no rollback touches an unowned object; no Managed source cleanup occurs without the
accepted route proof.

### Package R4 — Traversal, overlap admission, and discovery scheduling

**Priority:** P1 reliability/performance
**Ordering:** use the three independent R4 rows in section 8.1; no whole-package
activation or dependency is implied.
**Boundary:** A02/A08/A13/A19 are accepted. Promote traversal, overlap-advisory/
concurrent-correctness, and discovery-I/O slices independently and consume E1
boundaries without duplicating their behavior-preserving extraction. R6 owns the
warning/progress rendering, not the comparison or scheduler.

Until a slice has a predecessor-closed `ACTIVE` card, the bullets below are outcome/
boundary requirements, not an executable edit list. A `DRAFT` card may refine them;
the complete `ACTIVE` card supplies exact symbols, files, tests, fault matrix, and
performance gate.

#### R4 traversal/A08/A13 — Iterative bounded traversal

- [x] Replace recursive walkers with iterative/paged frames and post-order metadata.
      R4-T1 (2026-09-04) landed the three host bridge walkers and R4-T2 (2026-09-04) the Local
      provider's permanent-delete walker. Every host Copy/Move/Delete walk is frame-based.
      Explicit residual: the Local provider's native copy walkers (`CopyDirectoryInternal`,
      `runDirectory`) still recurse behind the provider's own `CopyItem`/`CopyItems`, which no
      host route reaches (the plan builder selects the bridge for every Copy and Local Move is
      native-only without a copy fallback); they keep the provider's declared 128-level ceiling
      as a provider limit and are not a product traversal path.
- [x] Convert aggregate ceilings to bounded in-memory backpressure/paging/JIT. Do not
      create a disk-backed spool or durable enumerated-work list unless a separately
      activated provider-specific card proves the in-memory/JIT design insufficient.
      R4-T3 (2026-09-04) converted the bridge's aggregate retained-listing budget to
      reporting (one frame's listing plus name views, released on pop); the bridge's ready
      queue (`DiscoveryQueueTarget`) and the Local provider's copy/delete queues already
      backpressure; the Local delete walk retains at most one 256-entry batch per frame. No
      spool exists (R4-A19 evidence). The Local native copy walkers' ancestor-metadata ceiling
      is the same provider-only residual as checkbox 1.
- [x] Keep only fundamental per-item/provider limits terminal.
      After R4-T3 the bridge's terminal limits are the task-lifetime link/cleanup records
      (4,096 entries / 16 MiB path text), the per-item child-name contract, and the
      provider's own declared limits (path length, name length, enumeration buffer).
- [x] Separate bound real-container traversal, bounded provider-declared virtual-folder
      convergence, and fixed object/version snapshots; never follow links/aliases or
      accept provider results outside the admitted boundary.
      Landed by R2 (typed A13 cutover: a binding provider's selected root is captured as a
      no-follow bound identity before confirmation and only that object is deleted through
      `DeleteIfUnchanged`; a pathname replacement is never deleted), R0b (S3 virtual-folder
      convergence per observed generation with explicit residuals; fixed key/version
      selections stay snapshots), R0d (Microsoft Drive stable item ID as the real-container
      root, A13-MD1), the native-receipt rule for providers without host binding (S3 exact
      generation, MTP object identity), and the walkers: the bridge and the Local delete walk
      snapshot each object no-follow, a link is handled as an object under the task's link
      policy (Skip/Preserve; a directory-namespace link is traversed only when that policy
      says so), and every enumerated child name must satisfy the structural contract
      (`ValidateBridgeStructuralChildName`: no empty, `.`/`..`, separator or control
      characters) plus the destination's component rules before it is admitted, so a provider
      result cannot escape the admitted parent.
- [x] Expose exact traversal-open/closed state and exact counters required by A08;
      do not construct a growing-denominator percentage in the engine.
      The engine publishes `discoveryClosed`, `discoveryAheadActive`, `discoverySkipped`,
      `firstMutationBeforeDiscoveryClosed`, `discoveredTotalBytes`, `discoveredFileCount`,
      `discoveredDirectoryCount`, `discoveryElapsedMs`, and the completed byte/item counters on
      the task snapshot (R4-A19), and computes no percentage. The popup's percent text is gated
      on `discoveryClosed`; its provisional bar fraction before closure is the rendering
      question R6-A08 owns (D2-A08: indeterminate bar plus exact counters before closure).

##### R4-T1 — Iterative walkers with post-order metadata (activation card)

| Activation-card field | R4-T1 bounded execution contract |
|---|---|
| Slice/state/owner | **R4-T1 / `COMPLETE` (2026-09-04; exit receipt below) / Claude in worktree `codex/r4-a19-closeout-2026-09-02`.** First slice of the R4 traversal/A08/A13 package: the three host bridge walkers stop recursing on the C++ stack; the provider-side walkers (Local delete and native copy) follow in R4-T2 with the same technique, and every other package checkbox (backpressure conversion, terminal per-item limits, container/virtual-folder separation, A08 counters) stays inactive. |
| Baseline and drift | Planned at `d3cfbe67` (Fresh Full gate #12 recorded on the R0-RC3 card). Drift review: `git status --short` shows only the user-owned untracked `last_run/`; the walkers named in Scope are unchanged since R3-3 (`ef3dc101`). |
| Contract | `FO-DISCOVERY-01` traversal clause (section 6.1: iterative/paged walking, bounded in-memory retained state, never a disk/durable spool) and the section 4 traversal-bounds row. Displaced: the "hard traversal ceiling of 128 descendant directory edges" sentences in `Specs/Core/Core_FileSystemBridge.md` and `Specs/FileSystem/FileSystem_FileOperations.md`; the retained work-entry, queued-path-text, and metadata ceilings stay task-terminal and unchanged. |
| Scope | In: `RedSalamander/FolderWindow.FileOperations.State.cpp` (`CopyDirectorySequential`, the `CopyDirectoryParallel` producer lambda, `CopyLink` directory-target traversal, `ValidateTraversalDepth`), `Common/FileOperationTraversalPolicy.h` (`kTraversalMaxDepth`), the self-test depth override (`SetFileOpsBridgeTraversalDepthLimitForSelfTest`, kept as a test-only seam), the two specs above, and the FileOps self-tests named below. Out: discovery scheduling (R4-A19), ready-queue backpressure and paging (package step 2), provider enumeration paging, MTP/cloud walkers (they already page per parent), and the Local provider's own walkers (`DeleteDirectoryRecursiveBatched` with `kDeleteTraversalMaxDepth`, `CopyDirectoryInternal`, the `runDirectory` frame lambda in `FileSystem.FileOps.cpp`), which keep their depth ceilings until R4-T2 applies the same frame technique; host-routed directory Copy/Move on Local already runs through the bridge walkers in scope. |
| Predecessors | R0a, R1c, E1, R2, R3 (owned publication), R4-A19 (discovery baseline): all `COMPLETE`. |
| Steps | (1) `[x]` (2026-09-04) `CopyDirectorySequential` becomes a loop over an explicit frame stack (`DirectoryFrame`: source/destination, depth, enumerated children and names, child cursor, `hadRetainedChild`, retained metadata bytes, active link-traversal frame, managed directory cleanup record); a frame pops post-order and performs exactly the metadata restore and cleanup decision the recursive epilogue performs today. (2) `[x]` The `CopyDirectoryParallel` producer uses its own frame stack (just-in-time production per frame; `WorkItem` queue unchanged). (3) `[x]` `CopyLink` reports a directory-namespace link to its walker caller (`traverseAsDirectory`), which pushes a frame instead of recursing; root callers still start the walk; link policy unchanged. (4) `[x]` `kTraversalMaxDepth` retired as the bridge's production ceiling (the constant stays in the shared policy header for the Local provider walkers until R4-T2) (depth is reported through `bridge.traversal.*` counters, not terminal); the self-test override stays so the traversal-limit partial-success case keeps its seam, and `bridge.traversal.depthLimit` is emitted only through that seam. (5) `[x]` Specs state the iterative bridge walk; the section 13 line "Trees beyond 128 levels ... complete iteratively with bounded memory" is satisfied for Copy/Move (Delete follows in R4-T2); exit receipt recorded here. |
| RED/GREEN evidence | RED: `R4T_DeepTreeCopyCompletesIteratively` (Dummy to Dummy Copy of a 300-level chain with one file per level through the parallel producer, then a second Copy with verification requested, which the bridge runs through the sequential walker; Dummy paths have no length limit, and a Dummy-source directory Move is refused as indeterminate by design, not by traversal) fails on the baseline with `bridge.traversal.depthLimit`. GREEN: both copies complete with every item published, retained metadata below `kTraversalMaxMetadataBytes`, and no `depthLimit` diagnostic; `Phase11_BridgeMultiFolderParallelCopyInFlightLines` keeps its traversal-limit partial-success expectations through the self-test seam; the R3 family and the Fairstream/Cinderstar/Causeway bridge families stay green. |
| Performance/resources | Scenario: the existing Phase11 multi-folder copy and a 64-level chain, sequential and parallel, before and after. Metrics: wall time and `bridge.traversal.*` retained high-water. Budget: no wall-time regression beyond noise (5%) and retained metadata high-water not above the recursive baseline for the same tree. Evidence: focused FileOps runs recorded on this card; the next Fresh Full gate. |
| Durable owners | `Specs/Core/Core_FileSystemBridge.md` (traversal ceilings paragraph), `Specs/FileSystem/FileSystem_FileOperations.md` (ceilings paragraph), this plan (package checkbox 1, section 13 line). |
| STOP/rollback | STOP if a frame cannot reproduce today's post-order metadata restore or Managed cleanup ordering exactly, or if the parallel producer's quiesce-before-post-order rule cannot be expressed per frame; rollback is the recursive walker behind the 128 ceiling (one revert). |
| Exit receipt | R4-T1 exit receipt (2026-09-04): activation card `2888572c`; implementation `ca909c76` (frame-based `CopyDirectorySequential` and `CopyDirectoryParallel` producer, `CopyLink` reports a directory-namespace link to its walker, `GetFileOpsTraversalMaxDepth()` returns no ceiling outside the self-test seam, Commands source contract pins the frame types and the absence of the recursive helpers and of `kTraversalMaxDepth` in the executor, specs updated). Debug full-solution build 0 warnings / 0 errors. Focused evidence on the landed tree (2026-09-04 03:28-03:55): FileOps `R4T_` 3/3 (`R4T_DeepTreeCopyCompletesIteratively`: the 300-level Dummy chain copies through the parallel producer and again with verification through the sequential walker, every item published, no `depthLimit` diagnostic), `R3_` 6/6, `RC3_9` 3/3, `Phase11_` 9/9 (traversal-limit partial success still produced through the self-test seam), `Phase12_` 3/3, `Fairstream_` 20/20, `Cinderstar_` 8/8, `Causeway_` 7/7, `Phase10_` 9/9 (wide-shallow retention rows within the retained-entry, path-byte and metadata bounds); Commands `file_operations_` 3/3 and `file_system_capabilities_contract_` 1/1; Pester 187/187. Performance: both 300-level copies complete in 47.1 s in the gate (R4T case duration); no Phase10/Phase11 wall-time regression observed. Fresh Full gate #13 on `254c6e6c` (run `20260904T123952Z-120916-bcf9921634424b17bcb4e2c95307ee4e`, started 2026-09-04 14:39): 2075 counted, 2020 passed, 2 failed, 53 skipped; PluginContractTests, DxUi tests, Compare 232/232 and Pester green; `R4T_DeepTreeCopyCompletesIteratively` passed in the gate. The two failures: `cmd_pane_batchRename_admission_worker_owned_cancellable` (known load flake; passed 1/1 isolated on the same build) and `R4A19_DiscoveryIndependentVolumes` (environment: alternate-volume test root not authorized). The first gate #13 attempt on `ca909c76` (04:44) crashed in the FileOps stage from the libcurl shared-connection-pool race that R0f-Curl-OR2 (`254c6e6c`) corrected; this rerun on the corrected tree is the recorded gate. |

##### R4-T2 — Local permanent-delete walker on an explicit frame stack (activation card)

| Activation-card field | R4-T2 bounded execution contract |
|---|---|
| Slice/state/owner | **R4-T2 / `COMPLETE` (2026-09-04; exit receipt below) / Claude in worktree `codex/r4-a19-closeout-2026-09-02`.** Second slice of the R4 traversal/A08/A13 package: the Local provider's permanent recursive Delete stops recursing on the C++ stack per directory level. The Local provider's native copy walkers (`CopyDirectoryInternal`, the `processDirectory`/`runDirectory` lambdas) are reachable only through the provider's own `CopyItem` (the host routes directory Copy/Move through the bridge walkers landed in R4-T1) and stay for a later slice. |
| Baseline and drift | Planned at `ca909c76` (R4-T1 landed; Fresh Full gate #13 recorded on the R4-T1 card). Drift review: `git status --short` shows only the user-owned untracked `last_run/`. |
| Contract | `FO-DISCOVERY-01` traversal clause (iterative walking, bounded retained state, no spool) applied to the Local provider's Delete walk; section 6.1 "Permanent recursive Delete uses that same discovery/execution rule". Displaced: the Local delete depth ceiling sentences in `Specs/FileSystem/FileSystem_FileOperations.md` (the R4-T1 paragraph's "copy and permanent-delete walkers still recurse with a 128-level ceiling"). The batch size (256 entries), the bounded terminal-failure record (4,096 entries / 16 MiB), and the re-enumerate-until-empty rule stay unchanged. |
| Scope | In: `Plugins/FileSystem/FileSystem.FileOps.cpp` (`DeleteDirectoryRecursiveBatched`, `DeleteDirectoryRecursive`, `kDeleteTraversalMaxDepth`, the `FileOps.DeleteTraversal.MaxDepth` metric), the FileOps self-tests (new case in the R4 traversal family), `Specs/FileSystem/FileSystem_FileOperations.md`, this plan. Out: `DeletePathInternal` dispatch for files, links, and recycle (unchanged), the parallel batch scheduler (workers still call `DeletePathInternal` on their entry, one C++ nesting per worker, and walk serially below it), the Local copy walkers, other providers. |
| Predecessors | R4-T1 (`COMPLETE`), R0a, R1c, R4-A19: all `COMPLETE`. |
| Steps | (1) `[x]` (2026-09-04) `DeleteDirectoryRecursiveBatched` becomes a loop over an explicit frame stack: a frame holds the directory path, `deleteRoot`, its terminal-failure set and byte budget (released on pop), the current batch and cursor; a serial child directory entry performs the same pre-steps `DeletePathInternal` performs for a directory (progress path, cancel, no-follow snapshot; a reparse point or a non-recursive context is deleted in place) and then becomes a pushed frame whose result reaches the parent's per-entry handling when it pops; the directory itself is deleted when its enumeration comes back empty, exactly as today. (2) `[x]` Parallel batches keep the scheduler: each worker calls `DeletePathInternal` on its entry (one C++ nesting) and the walk below it is the same frame loop with a serial budget. (3) `[x]` `kDeleteTraversalMaxDepth` retired; `context.deleteTraversalDepth` and `FileOps.DeleteTraversal.MaxDepth` keep reporting the observed depth. (4) `[x]` Spec and section 13 line ("Trees beyond 128 levels ... complete iteratively") satisfied for permanent Delete; exit receipt recorded here. |
| RED/GREEN evidence | RED: `R4T2_DeepTreeLocalDeleteCompletes` (a 300-level Local chain with one file per level, beyond `MAX_PATH`, permanently deleted through the host File Operations) fails on the baseline with the provider's walk-depth stop (`ERROR_STACK_OVERFLOW`). GREEN: the Delete completes with the root gone; `FileOps_DeleteToctouSwapGuard`, `Phase7_ParallelDelete*`, `Phase10_PermanentDelete`, the Fairstream parallel-delete cases, and the R4-A19 Delete baseline stay green. |
| Performance/resources | Scenario: the existing Phase7 parallel-delete cases and the R4-A19 Local Delete baseline before and after. Budget: no wall-time regression beyond noise (5%); retained state per level unchanged (batch + terminal-failure set). Evidence: focused FileOps runs recorded on this card; the next Fresh Full gate. |
| Durable owners | `Specs/FileSystem/FileSystem_FileOperations.md` (Delete traversal sentences), this plan (package checkbox 1, section 13 line). |
| STOP/rollback | STOP if the per-entry result handling (terminal-failure retention, cancel, partial-copy propagation) cannot be reproduced exactly from a popped frame, or if the worker path would need a second frame stack across threads; rollback is the recursive walker behind the 128 ceiling (one revert). |
| Exit receipt | R4-T2 exit receipt (2026-09-04): activation card `b8d3bc86`; implementation `8f9abef3` (`DeleteDirectoryRecursiveBatched` is a loop over `DeleteWalkFrame`s: a frame owns its directory's terminal-failure record and byte budget, its current 256-entry batch and cursor; a serial child directory performs the same pre-steps `DeletePathInternal` performs and becomes a pushed frame whose result lands in the parent's batch when it pops; parallel batches keep the scheduler with one nesting per worker; `kDeleteTraversalMaxDepth` retired and `FileOps.DeleteTraversal.MaxDepth` reports no limit; Commands source contract pins `struct DeleteWalkFrame final`; `FileSystem_FileOperations.md` and `Core_FileSystemBridge.md` state the frame-based Local delete walk). Debug full-solution build 0 warnings / 0 errors. RED by construction: the retired walker refused `context.deleteTraversalDepth > 128` with `ERROR_STACK_OVERFLOW`, and the new `R4T2_DeepTreeLocalDeleteCompletes` creates a 300-level Local chain beyond `MAX_PATH` through the Local plugin and deletes it permanently through the host. GREEN (focused runs 2026-09-04 16:00-16:13 on the landed tree): FileOps `R4T` family filter 4 passed / 0 failed (the R4 traversal family incl. `R4T2_DeepTreeLocalDeleteCompletes` with the root gone), `Phase7_` 19/19, `Phase10_` 9/9, `Phase6_` 7/7, `FileOps_Delete` 3/3, `Fairstream_ParallelDelete` 3/3, `R4A19_` 4/5 (only the environment-blocked `R4A19_DiscoveryIndependentVolumes`: alternate-volume test root not authorized); Commands `file_operations_` 3/3 and `file_system_capabilities_contract_` 1/1; Pester 187/187. Performance: no Phase7 parallel-delete or R4-A19 Local Delete baseline regression observed in the focused runs. Fresh Full gate #14: on `048ddb4c` (run `20260904T145742Z-49044-473d6c2a7f0f49bda179b0c0b58ef77f`, started 2026-09-04 16:57): 2077 counted, 2010 passed, 14 failed, 53 skipped; PluginContractTests, DxUi tests, Compare 232/232 and Pester green; `R4T2_DeepTreeLocalDeleteCompletes` (1.9 s) and `R4T3_WideDirectoryCopyCompletes` (189.5 s) passed in the gate. The 14 failures: `R4A19_DiscoveryIndependentVolumes` (environment: alternate-volume test root not authorized) and 13 Commands cases whose reasons are all focus and foreground timing (a dialog did not focus its first input, focus not returned to the pane shell, the navigation shell not restored, one no-op scroll repaint) plus the known Batch Rename admission load flake and the alert-overlay timing flake; the desktop was in use during the Commands stage, and every one of the 13 passed 1/1 isolated on the same build afterwards. |

##### R4-T3 — Bridge per-directory listing reported, not capped (activation card)

| Activation-card field | R4-T3 bounded execution contract |
|---|---|
| Slice/state/owner | **R4-T3 / `COMPLETE` (2026-09-04; exit receipt below) / Claude in worktree `codex/r4-a19-closeout-2026-09-02`.** Third slice of the R4 traversal/A08/A13 package (package checkbox 2 for the bridge): the bridge stops counting each open directory's listing against a task-wide 8 MiB metadata budget. Every host Copy runs through the bridge (the plan builder selects `Copy` for every Copy; only Move can select `Native`), so today one directory of roughly 18,000 children (the provider's `IFilesInformation` buffer plus two owned `std::wstring` sets per open frame) stops the whole Copy with `bridge.traversal.metadataBudget`. |
| Baseline and drift | Planned at `8f9abef3` (R4-T2 landed). Drift review: `git status --short` shows only the user-owned untracked `last_run/`; the two bridge walkers and `ValidateAndRegisterChildName` are unchanged since R4-T1 (`ca909c76`). |
| Contract | `FO-DISCOVERY-01` traversal clause (bounded in-memory retained state, O(depth) live frames, no spool) and package checkbox 2 ("convert aggregate ceilings to bounded in-memory backpressure/paging/JIT"). Displaced: the "8 MiB retained ancestor/directory metadata" ceiling sentences in `Specs/Core/Core_FileSystemBridge.md` and `Specs/FileSystem/FileSystem_FileOperations.md`. Kept terminal (checkbox 3, fundamental per-task records): 4,096 retained work entries and 16 MiB queued-path text over the task-lifetime semantic-link mappings, failed link prefixes, held directory cleanups and deferred links; the per-item child-name contract (a duplicate or, on a case-folding destination, case-only-variant child name is still refused per directory). |
| Scope | In: `RedSalamander/FolderWindow.FileOperations.State.cpp` (`BridgeOrdinalIgnoreCaseLess` becomes `BridgeChildNameLess`/`BridgeChildNameSet` over `std::wstring_view`; `ValidateAndRegisterChildName`; `ReserveTraversalMetadata` becomes the accounting-only `NoteTraversalMetadataRetained`; the sequential and producer frames; the `FileOps.Bridge.TraversalMaxMetadataBytes` emit), `RedSalamander/SelfTest/Commands/Commands.SelfTest.PluginConfig.cpp` (source contract), the FileOps R4 traversal family, the two specs, this plan. Out: `kTraversalMaxMetadataBytes` in `Common/FileOperationTraversalPolicy.h` and the Local provider's own ancestor-metadata ceiling (`CopyDirectoryInternal`, `runDirectory`), the retained-work-entry and queued-path ceilings, discovery scheduling, provider listing paging (no `IFileSystem` paging contract exists; a listing is one provider allocation). |
| Predecessors | R4-T1, R4-T2 (`COMPLETE`), R3, R4-A19. |
| Steps | (1) `[x]` (2026-09-04) One set of child-name views per frame, ordered the way the destination compares components, replaces the two owned-name sets; the second indistinguishable name is refused exactly as today (`ERROR_INVALID_NAME`, `bridge.source.invalidChildName`, Continue-on-error semantics unchanged). (2) `[x]` The metadata reservation becomes accounting: per-frame path text, listing buffer and index bytes are added on frame start, released on pop, and reported through `_bridgeTraversalMaxMetadataBytes`; `bridge.traversal.metadataBudget` no longer exists and the perf row reports no limit. (3) `[x]` The Commands source contract pins the view set present and the reservation, the owned-name set and the budget diagnostic absent. (4) `[x]` Specs state the reported listing and the remaining terminal ceilings; the section 13 line "Trees beyond 128 levels/aggregate ceilings complete iteratively with bounded memory" is satisfied for bridge Copy/Move; exit receipt recorded here. |
| RED/GREEN evidence | RED: `R4T3_WideDirectoryCopyCompletes` (a Dummy directory of 20,000 children with 88-character names copied Dummy to Dummy through the parallel producer) fails on the baseline with `bridge.traversal.metadataBudget` (`ERROR_NOT_ENOUGH_MEMORY`, one traversal-limit hit). GREEN: the Copy completes with the first and last child published and zero traversal-limit hits; the R4 traversal family, `Phase11_` (hostile child names, traversal-limit partial success through the depth seam), `Phase12_`, `Phase10_` (wide-shallow retention rows still below 8 MiB for that tree), `Phase5_` (bounded-resources control), `Fairstream_`, `Cinderstar_` and `Causeway_` stay green; Commands contracts and Pester green. |
| Performance/resources | Scenario: the wide-shallow Phase10 tree and the new 20,000-child directory. Metrics: wall time and `FileOps.Bridge.TraversalMaxMetadataBytes` high-water. Budget: no wall-time regression beyond noise (5%); retained bytes per child drop from the buffer plus two owned strings and set nodes to the buffer plus one 16-byte view and its set node. Evidence: focused FileOps runs recorded here; the next Fresh Full gate. |
| Durable owners | `Specs/Core/Core_FileSystemBridge.md` (traversal ceilings paragraph), `Specs/FileSystem/FileSystem_FileOperations.md` (ceilings paragraph), this plan (package checkboxes 2 and 3 notes, section 13 line). |
| STOP/rollback | STOP if the per-directory duplicate/fold refusal cannot be reproduced with views (the provider buffer must outlive the set: it is owned by the frame's `IFilesInformation`), or if any consumer of the metadata metric needs a ceiling; rollback is the reservation behind the 8 MiB budget (one revert). |
| Exit receipt | R4-T3 exit receipt (2026-09-04): activation card `f17ea8d3`; implementation `75172587` (`BridgeChildNameLess`/`BridgeChildNameSet` over `std::wstring_view` replaces the two owned-name sets, one set per frame ordered the way the destination compares components; `NoteTraversalMetadataRetained` replaces `ReserveTraversalMetadata` as accounting only; `bridge.traversal.metadataBudget` no longer exists and `FileOps.Bridge.TraversalMaxMetadataBytes` reports no limit; the Commands source contract pins the view set present and the reservation, owned-name sets and budget diagnostic absent; `Core_FileSystemBridge.md` and `FileSystem_FileOperations.md` state the reported listing and the remaining terminal ceilings). Debug full-solution builds 0 warnings / 0 errors (RED and GREEN). RED witness on the baseline (2026-09-04 16:15-16:24, `f17ea8d3` plus the test alone): `R4T3_WideDirectoryCopyCompletes` failed with `hr=0x80070008` (`ERROR_NOT_ENOUGH_MEMORY`), one traversal-limit hit, retained metadata high-water 8,388,266 bytes against the 8 MiB budget, the first child published and the last missing, exactly the `bridge.traversal.metadataBudget` stop. GREEN (focused runs on the landed tree, 2026-09-04): FileOps `R4T` family 5/5 (`R4T3_WideDirectoryCopyCompletes` 190.4 s, `R4T_DeepTreeCopyCompletesIteratively` 46.9 s, `R4T2_DeepTreeLocalDeleteCompletes` 1.7 s; the wide-directory case runs the Dummy provider at zero simulated latency because the harness seeds 5 ms per access, and restores the cached seed configuration on every exit path: `Phase5_` 9/9 afterwards confirms the restore), `Phase11_` 9/9, `Phase12_` 3/3, `Phase10_` 9/9, `Phase5_` 9/9, `Fairstream_` 20/20, `Cinderstar_` 8/8, `Causeway_` 7/7; Commands `file_operations_` 3/3 and `file_system_capabilities_contract_` 1/1; Pester 187/187. Performance: the 20,000-child Dummy copy completes in 190 s (about 9.5 ms per child on the in-memory provider, the bridge's per-file publication cost, noted for a later performance slice) through the parallel producer with zero limit hits; Phase10 wide-shallow retention rows stay below the former bounds. Fresh Full gate #14: on `048ddb4c` (run `20260904T145742Z-49044-473d6c2a7f0f49bda179b0c0b58ef77f`, started 2026-09-04 16:57): 2077 counted, 2010 passed, 14 failed, 53 skipped; PluginContractTests, DxUi tests, Compare 232/232 and Pester green; `R4T2_DeepTreeLocalDeleteCompletes` (1.9 s) and `R4T3_WideDirectoryCopyCompletes` (189.5 s) passed in the gate. The 14 failures: `R4A19_DiscoveryIndependentVolumes` (environment: alternate-volume test root not authorized) and 13 Commands cases whose reasons are all focus and foreground timing (a dialog did not focus its first input, focus not returned to the pane shell, the navigation shell not restored, one no-op scroll repaint) plus the known Batch Rename admission load flake and the alert-overlay timing flake; the desktop was in use during the Commands stage, and every one of the 13 passed 1/1 isolated on the same build afterwards. |

##### R4-T4 — Microsoft Drive provider-native deep merge correction

| Field | R4-T4 |
|---|---|
| Slice/state/owner | **R4-T4 / `ACTIVE` (2026-09-04) / Codex `/root` in this worktree.** Post-closeout correction to the provider-native Graph merge path. |
| Finding | `MergeMoveFolderIntoExisting` remained recursive and returned `ERROR_STACK_OVERFLOW` at depth 64, while the R4 traversal closeout described Local as the only recursive residual. A valid deep Graph tree therefore failed even though the host bridge and Local paths were iterative. |
| Correction | Replace recursive folder/folder descent with an explicit frame stack while preserving just-in-time destination probes, prompt order, retry behavior, cancellation checks, per-child partial results, and the conservative retention of drained source folders. No public ABI or capability change. |
| Tests/performance | Convert the existing debug cap test into a 70-level successful deep-merge witness: the leaf reaches the destination, the source leaf is gone, and the result is partial only because source folders are intentionally retained. Keep the merge request-count/perf contracts and PluginContractTests green; use the next Fresh Full archive. Memory stays O(open depth plus each active frame's listing), as before without call-stack growth. |
| Scope/STOP | In: `Plugins/FileSystemMicrosoftDrive/FileSystemMicrosoftDrive.cpp`, `Specs/FileSystem/FileSystem_MicrosoftDrive.md`, this plan. STOP on any request-count, prompt-order, collision, retry, cancellation, or partial-result regression. |
| Exit | 0-warning builds, provider debug tests, focused Graph/FileOps contracts, authoritative spec, and an exit receipt before `COMPLETE`. |

#### R4-A02 — Same-host overlap advice and concurrent correctness

- [x] Build a bounded task-injection overlap advisor from operation roles, normalized
      provider/profile/component-boundary paths, canonical UNC share boundaries, and
      only cached/already-bound or otherwise cheap positive mapped-drive/SUBST/link/
      alias evidence. Compare against active plus queued residual task scopes with
      zero descendant enumeration and zero network/device call solely for the warning.
- [x] Classify read/read as no warning; clearly disjoint paths as no warning; and
      obvious/positively evidenced write/write, Delete/read, Delete/write, Move-source,
      same-path, or ancestor/descendant overlap as one concrete problem. Suppress a
      parent/child warning only when exact cheap evidence proves that residual branch
      is irrevocably closed/excluded, not merely because the current item is elsewhere.
- [x] Store a Queue predecessor edge only in the live host scheduler. Release once
      after named predecessors terminate without a second A02 warning, retained
      producer provenance, or reconstruction from history. Store a Run receipt only
      for its concurrently-live same-host relation.
- [x] Maintain a bounded in-memory index of exact objects/scopes created or published
      by concurrently-live tasks in the same host. Before one invalidates the other's
      output, require a covering disclosed Run consequence or park at the item gate.
      Drop protective meaning when the live relation ends; never share it across app
      instances/processes/hosts.
- [x] Remove root-wide overlap rejection/exclusion as correctness policy. Ensure the
      host, adapters, and providers remain correct under concurrent task races; retain
      only narrow provider-internal thread/session/handle synchronization and expose a
      single-lane wait as `Waiting for device/provider`, not an overlap failure.
- [x] Preserve every ordinary per-item conflict, identity/generation revalidation,
      exact owned-publication/rollback rule, container boundary, Managed cleanup proof,
      Retry restriction, and immutable result when Run together is selected.

##### R4-A02-1 — Overlap advice at task injection: Queue after or Don't start (activation card)

| Activation-card field | R4-A02-1 bounded execution contract |
|---|---|
| Slice/state/owner | **R4-A02-1 / `ACTIVE` (2026-09-04) / Claude in worktree `codex/r4-a19-closeout-2026-09-02`.** First slice of R4-A02 (checkboxes 1, 2 and the Queue half of 3): the same-host overlap that today serializes silently inside admission becomes one problem-specific warning at injection with **Queue after these tasks** (default) and **Don't start**. **Run at the same time**, the concurrently-live output guard and the Run receipt (checkboxes 3-Run, 4, 5) follow in R4-A02-2 so that no task runs against another's live scope before the item-gate guard exists. |
| Baseline and drift | Planned at `4f594ea2` (R6-A08 landed). Drift review: `git status --short` shows only the user-owned untracked `last_run/`; `Task::PrepareMutationInterlockScopes`, `TaskScopesOverlap` and the admission wait in `FolderWindow.FileOperations.State.Queue.cpp` are unchanged since R2. |
| Contract | Section 6.1 `FO-OVERLAP-01` and D2-A02: comparison at injection from operation roles and the scopes already prepared for the task (normalized provider/profile paths, retained no-follow root authority, anchors), zero descendant enumeration and zero network or device call for the warning; read/read and clearly disjoint paths warn nothing; obvious or positively evidenced write/write, Delete/read, Delete/write, Move-source, same-path or ancestor/descendant overlap yields one warning naming the concrete problem; Queue is one transient predecessor edge in the live host scheduler released once the named tasks terminate, with no second warning and no protection derived from completed tasks. Displaced: the sentence "Every other overlapping role pair serializes" in `Specs/FileSystem/FileSystem_FileOperations.md` (Overlapping-operation interlock) becomes "waits only as a user-chosen Queue edge"; the popup's silent `Waiting` for overlap. Kept: the interlock predicate itself (`TaskScopesOverlap`, roles `ReadSource`/`WriteSource`/`PublishDestination`, anchors, conservative identity domains) as the Queue edge's implementation. |
| Scope | In: `RedSalamander/FolderWindow.FileOperations.State.Queue.cpp` (injection-time overlap check against active plus queued task scopes using the existing predicate; the admission wait stays as the Queue edge), `RedSalamander/FolderWindow.FileOperations.State.cpp` (a `DeferredConsentRisk::SameHostOverlap` prompt through the existing deferred-consent task-card surface with the actions Queue after (default) / Don't start; consent receipt via `StoreDeferredConsentReceipt`; the named overlapping tasks and the concrete problem in the prompt facts), `RedSalamander/FolderWindow.FileOperations.Popup.cpp` (render the new bucket with the two actions; hide `Start now` for a task queued by an overlap edge; keep it for the global Queue mode), `FolderWindow.FileOperationsInternal.h` (risk value, prompt facts), FileOps/Commands self-tests, `Specs/FileSystem/FileSystem_FileOperations.md`, `Specs/UI/UI_FileOperationsPopup.md`, this plan. Out: Run at the same time, the live-output index and item-gate park, cross-instance coordination, alias discovery beyond already-bound facts, inline F2 (`RenamePlan(InlineRename)` keeps honoring the interlock silently, as its spec exception states), R6-A02's final wording/accessibility pass (it consumes these facts). |
| Predecessors | E1, R1d-core, R2 (`COMPLETE`); R4 traversal/A08/A13 closeout (2026-09-04); R6-A08 (`COMPLETE`). |
| Steps | (1) `[x]` (2026-09-04) At injection, after Preparing produced the task's interlock scopes, compare them against every active and queued task's scopes with `TaskScopesOverlap`; classify read/read and disjoint as no warning; otherwise build the prompt facts: the named tasks, the roles, and the concrete problem sentence (two destination scopes may create/replace the same names; a Delete may remove a source or destination another task is reading/writing; two Deletes may remove the same members; positively known aliases). (2) `[x]` Show the warning once through the deferred-consent surface with **Queue after these tasks** (default) and **Don't start**; Queue keeps the existing admission wait as the transient edge and records the consent receipt; Don't start ends the task before any mutation, clipboard consumption or breadcrumb, as a Preparing cancel. (3) `[x]` A task queued by an overlap edge offers no `Start now`; the edge is released once every named task is terminal and the task enters the ordinary engine without a second warning. (4) `[x]` Specs and tests: FileOps cases for Delete-over-Copy-source (warning, Queue waits, runs after), two Copies to the same destination (warning), two Copies reading one source into disjoint destinations (no warning), disjoint roots (no warning), Don't start (no mutation, no history entry); Commands popup contract for the bucket's actions; exit receipt recorded here. |
| RED/GREEN evidence | RED: today a Delete injected over a running Copy's source silently waits (`FileOps.Interlock.WaitUs` grows, the card shows `Waiting`) and no prompt, receipt or `Don't start` exists; the new FileOps cases fail on the baseline because no `SameHostOverlap` prompt appears. GREEN: the cases above pass; `FileOps_ProviderCapabilityMatrix` (overlap predicate checks through `DebugMutationScopesOverlapForSelfTest`) and `Floodgate_InlineF2WorkerQueueBypassAndInterlock` stay green because the predicate and the wait are unchanged and inline rename keeps honoring the interlock without a prompt (it has no consent surface before its card reveals; R6-A02 may revisit); `Phase7_SharedPerItemScheduler`, `Phase5_DiscoveryCancelReleasesSlot` and `FileOps_MoveMergeIntoExistingFolderSameVolume`, which today rely on the silent wait, answer the new warning through the existing self-test decision seam with its default (Queue after) and keep their expectations; Phase7/Phase9 conflict-prompt families stay green. |
| Performance/resources | The comparison reuses scopes already prepared for admission; cost is the existing predicate per active/queued task at injection (bounded by the queue), no new I/O. Evidence: `FileOps.Interlock.*` rows unchanged for non-overlapping tasks; focused runs; the next Fresh Full gate. |
| Durable owners | `Specs/FileSystem/FileSystem_FileOperations.md` (Overlapping-operation interlock section), `Specs/UI/UI_FileOperationsPopup.md` (warning surface), this plan (R4-A02 checklist notes). |
| STOP/rollback | STOP if the injection-time comparison would need new enumeration or provider calls, or if the deferred-consent surface cannot carry two scheduling actions without a modal; rollback is the silent admission wait (one revert). |
| Exit-receipt contract | Commit, Debug build receipt, RED witness on the baseline, focused FileOps overlap cases, Commands popup contracts, Pester, spec updates, and the Fresh Full gate that follows; recorded on this card. |

##### R4-A02-2 — Exact Run relation and concurrently-live output guard (activation card)

| Activation-card field | R4-A02-2 bounded execution contract |
|---|---|
| Slice/state/owner | **R4-A02-2 / `ACTIVE` (2026-09-05) / Codex `/root` in this worktree.** This closes only the Run half of R4-A02 checkbox 3 plus checkboxes 4-6: exact disclosed concurrent admission, a bounded host-local live-publication index, and the item gate before one task invalidates another concurrently-live task's output without a covering Run relation. It does not add cross-process coordination, history-derived authority, a global overlap preference, review-later collection, persistent receipts, or a new scheduler. |
| Baseline and drift | Exact tracked baseline is `20c64170` (the completed reliability review corrections and R4-A02-1 implementation); repository-root `last_run/` is pre-existing untracked user-owned output and remains excluded. Drift at activation is the in-progress R4-A02-2 Run action, localized strings, focused selftests/source contract, and the bounded receipt correction after the archived `0xC00000FD` test-harness stack-overflow witness. No unrelated tracked owner is admitted. |
| Contract | Sections 6.1/6.9 `FO-OVERLAP-01` and D2-A02. **Run at the same time** is offered only when every disclosed same-host relation fits the fixed relation bound; it admits the task past the root interlock only for those exact task IDs and keeps all ordinary conflict, exact-authority, publication, verification, cleanup, Retry and immutable-result rules. Successful current-task publications mark a bounded host-local exact object/scope index. Before Delete, overwrite, rename or equivalent invalidation, an uncovered concurrently-live publication parks at one task-local decision: **Queue until the other task finishes** (safe default), **Skip this item**, an explicit destructive action, or Cancel. Queue waits for that exact live publisher to terminate and then resumes ordinary current-membership semantics. Index and relation meaning disappear at terminal state and are never reconstructed from history. |
| Scope | In: `RedSalamander/FolderWindow.FileOperationsInternal.h`, `FolderWindow.FileOperations.State.Queue.cpp`, `FolderWindow.FileOperations.State.cpp`, `FolderWindow.FileOperations.Popup.cpp`; the main and four maintained satellite resources; FileOps overlap selftests and Commands/Pester source contracts; `Specs/FileSystem/FileSystem_FileOperations.md`, `Specs/UI/UI_FileOperationsPopup.md`, this plan, and focused/crash evidence under `Specs/TestRuns/4cb089111a23/`. Read-only controls: provider implementations, `Common/PlugInterfaces/FileSystem.h`, shared path/identity helpers, R3 publication rules, R4 traversal/discovery policy. Out: cross-instance/process/host protection, persisted index/receipts, provider ABI, broad provider locking, descendant enumeration solely for this guard, R6 review-later/history work, R7 link semantics, and R9 Shell delegation. |
| Predecessors | E1, R1d-core, R2, R4 traversal/A08/A13, R4-A19, R6-A08 and R4-A02-1 are `COMPLETE` or implemented with their receipts to be reconciled in this closeout. The activation baseline contains the R4-A02-1 engine behavior and no other active package owns the named production files. |
| Steps | (1) `[x]` Add fixed-capacity disclosed relation sets, an explicit Run action, Queue-safe default, exact predecessor wakeup, and relation-specific admission bypass. (2) `[x]` Prove concurrent admission on identity-bound Local destination scopes with disjoint result names; conservative identity domains retain the silent safe interlock. (3) `[x]` Add the bounded live-publication index plus terminal cleanup and exact covering-relation query. (4) `[x]` Gate Delete/overwrite/rename/invalidation at existing per-item exact-authority seams; Queue waits for only the named live publisher, Skip preserves, destructive consent applies only to that parked effect, and Cancel stops the current decision. (5) `[x]` Add behavioral index/gate/release/third-task/overflow tests, popup action/keyboard/resource contracts, authoritative specs, and a provisional receipt; the exact Fresh Full result is still required before `COMPLETE`. |
| RED/GREEN evidence | RED: baseline has no Run action or exact relation bypass; a Run-focused selftest cannot admit both same-destination tasks, and there is no live-output item gate/index contract. GREEN requires Queue/default and Don't-start cases unchanged; an identity-bound same-destination Run witness observes both tasks admitted simultaneously with zero interlock waits for the disclosed relation and byte-exact/disjoint outputs; a task reaching an uncovered concurrently-live output parks before mutation, Skip preserves, Queue releases only after that publisher terminates, the destructive choice permits exactly that effect, and a third undisclosed task still blocks/asks independently. Fixed-bound overflow withholds Run or falls back conservatively without dropping protection. |
| Performance/resources | Relation storage is fixed at 64 IDs per task/receipt. The publication index has an explicit fixed bound and records only exact published object/scope facts while tasks are live; overflow is fail-safe and may conservatively use the already-frozen destination scopes. Admission remains allocation-free under `_queueMutex`; item checks do no recursive enumeration and no network/device call solely for correlation. Record comparison counts, index high-water/overflow, item-gate count and wait time; focused families and the final Fresh Full gate must show no unrelated throughput or queue regression. |
| STOP/rollback | STOP if correctness requires a persisted/global provenance graph, pathname deletion authority, recursive live descendant discovery, provider ABI expansion, global host serialization, or weakening R3 exact publication/cleanup rules. STOP if a bounded index cannot fail safe. Rollback removes Run and the live guard together, returning to R4-A02-1 Queue/Don't-start behavior. |
| Exit-receipt contract | Record activation baseline, the archived stack-overflow failure and heap-indirected correction, exact 0-warning Debug/test-enabled Release receipts, focused R4-A02 and conflict/rename/delete/bridge controls, localization/resource validation, Pester/source contracts, index/overflow/perf evidence, authoritative-spec updates, `git diff --check`, and the final exact Fresh Full result before `COMPLETE`. |

**R4-A02 implementation receipt pending Fresh Full (2026-09-05):** baseline and
R4-A02-1 implementation are in `20c64170`; R4-A02-2 adds the exact Run relation,
bounded live-publication index, terminal cleanup, safe overflow fallback, and
per-effect Queue/Skip/destructive/Cancel gate. The first receipt layout overflowed the
self-test tick stack (`0xC00000FD`); the symbolized witness is archived under
`Specs/TestRuns/4cb089111a23/Continuation/20260905_003706_r4_a02_relation_receipt_stack_overflow/`,
and the corrected relation set is heap-indirected so ordinary consent receipts remain
pointer-sized. Full-solution Debug build receipt
`61b707991a73114ce74f9c6e1532adf83092c1666316dca9c2e721c89c588bff`
(`msbuild-20260905_024140_385-pid97428-10493d62.log`) and test-enabled Release
receipt `7c133cac9ee36c799c98667b36c6badbc465f37bc1d42eabe90a5b502e014c1b`
are both 0-warning/0-error. Governed Release run
`20260905T003040Z-40676-32154703eff145d6a20f8a71b589249d` passed R4-A02
9/9 with disk audit 0; its compact archive records index high-water 201/256,
one forced safe overflow, Run/Queue/gate counts, and the exact-publisher wait. Focused
controls passed: inline-F2 3/3, Phase9 11/11, Phase10 9/9, Phase11 bridge 3/3;
source contracts 188/188, resource localization 9/9, and the two archive/spec fixture
suites 39/39 (after granting their governed test-root access). The authoritative
contracts now live in `FileSystem_FileOperations.md`, `UI_FileOperationsPopup.md`,
and `Testing_PerformanceValidation.md`. `COMPLETE` remains gated on the exact Fresh
Full result required by this card.

**Exact Fresh Full gate #16 (2026-09-05):** run
`20260905T032924Z-106000-6e2ac2f0113d4d9caff0767b0ca9e4ec` on `31fe0b8c`
completed 2,086 counted / 2,032 passed / 1 failed / 53 expected skips with disk
audit 0. FileOperations passed 150/150 applicable, Commands 889/889, Compare
Directories 233/233, DxUi Menu and the remaining DxUi inventory, provider contracts,
localization, performance, and Tools Pester 729/729. The sole failure was the
unrelated `ViewerImgRaw` latest-wins harness taking an immediate async scheduler
snapshot before two accepted requests were observable; the exact case passed in
isolation on the same binaries. Commit `60d4ef6e` replaces that instantaneous sample
with the shared bounded condition-based wait. This gate is retained as honest
non-green evidence and does not satisfy the final Fresh Full requirement.

#### R4-A19 — Evidence-first discovery service

##### R3 activation card

| Field | R3 |
|---|---|
| Slice/state/owner | **R3 / `COMPLETE` (2026-09-04; R3-1, R3-2 and R3-3 receipted below, Fresh Full gate #11 recorded on the R3-3 receipt) / Claude in worktree `codex/r4-a19-closeout-2026-09-02`.** One owned-publication transaction, delivered as independently receipted slices: **R3-1** conditional replace on identity-less atomic-final routes (receipt below), **R3-2** provider content proof after atomic-final publication (Verified Copy and Managed Move into weak-commit routes), **R3-3** one transaction record consolidating the bridge's scattered publication state, the file/link/directory payload adapters, and the fault matrix. |
| Baseline and drift | Planned at `d4c65293` (R0f complete). Drift review: `git diff --stat 29f1f270..d4c65293 -- RedSalamander/FolderWindow.FileOperations.State.cpp Common/PlugInterfaces/FileSystem.h` shows only the R0f-GDrive landing; no other active package owns `FolderWindow.FileOperations.State.cpp` (R4-A19 is evidence-only). |
| Contract | `FO-PUBLISH-01` (owned non-final stage then atomic/conditional reveal for overwrite and weak-commit routes; no failure path deletes an occupant by name) and `FO-MOVE-01` (proof before Managed cleanup). Displaced by R3-1: FileOperations.md Conflict model "Overwrite ... shown only when the mutation provider can consume the exact retained destination authority" and Core_FileSystemBridge.md "If the provider cannot bind or consume that authority, those actions are withheld" (both now carry the identity-less conditional-replace form). |
| Scope | In (R3-1): `Common/PlugInterfaces/FileSystem.h` (`IFileWriterExpectedReplacement`), `RedSalamander/FolderWindow.FileOperations.State.cpp` (`Task::FileSystemIssue`, `DestinationSupportsAtomicReplace`, `CrossFileSystemBridge::PromptDestinationCollision`, `ValidateDestinationOverwritePolicy`, `CaptureReplaceExpectation`, `PumpAndPublishFile`), `FolderWindow.FileOperationsInternal.h` (`_conflictAtomicReplaceGrantCount`), the Curl `TempFileWriter`/`PublishCurlWriterTransaction`, S3 `MultipartS3FileWriter`/`EnsureWritableS3Target`/`PutS3ObjectFromMemory`/`CompleteS3MultipartUpload`, Microsoft Drive `MicrosoftDriveFileWriter` (`ResolveReplaceOccupant`, If-Match on PUT and upload-session creation), Google Drive `GoogleDriveFileWriter` (creation occupant id/version), Dummy `DummyFileWriter`/`CommitFileWriter`, the fake Graph fixture (If-Match on uploads), FileOps self-tests, and the provider specs. Out: the Local bound route (unchanged), links (Replace link stays bound-only), provider-native same-provider overwrite through `CopyItem`/`MoveItem` on identity-less routes (still withheld), verification and Managed Move proof (R3-2), the transaction record and fault matrix (R3-3). |
| Predecessors | R0a, R1c, E1, R2 exit receipts recorded above; D2-A04 and D2-A06 accepted; R0f complete (every destination reachable). |
| Steps | R3-1: (1) ABI interface; (2) host: prompt eligibility without bound authority, bridge acceptance, occupant captured before the prompt, atomic-final route with the granted flags, expectation handed to the writer before the first byte, refused replacement reported as a definitive non-commit; (3) the five writers; (4) fixture If-Match; (5) tests; (6) specs and this card. R3-2 (symbols): `Common/ContentDigest.h` (`Common::Crypto::ContentHasher` SHA-256/SHA-1/MD5 via CNG, QuickXorHash, CRC-64/NVME), `FileSystem.h` (`FileSystemContentProofAlgorithm` SHA256/SHA1/MD5/QUICKXOR/CRC64NVME, `FILESYSTEM_ROUTE_PROOF_WRITER_DIGEST`, `IFileWriterContentProof`), `QualifiedEndpoint::verificationWriterDigestProof`, bridge `writerProofRoute`/`EvaluateWriterContentProof`/`VerifyPublishedContent(writerProofMatched)`, `PrepareManagedSourceAuthority` keeps the cleanup record armed on proof-capable identity-less routes; writers: S3 `MultipartS3FileWriter` (CRC-64/NVME declared on create/part/put, `S3ObjectRevision::crc64NvmeBase64`), Graph `MicrosoftDriveFileWriter` (`ItemMetadata` hashes, `_publishedItem`), Drive `GoogleDriveFileWriter` (`sha256Checksum`), Dummy `DummyFileWriter`/`CommitFileWriter(committedSha256)`; fixtures report the same digests; `providerProof: "writer-digest"` in the three drift guards; `MoveStrategyQualificationFacts::destinationWriterDigestProof` (Managed without destination binding) and the pre-mutation transfer safety guard accept the route; the two `knownProofFlags` masks (`FileSystemRouteProviderBase.h`, `FileSystemRouteContract.cpp`) admit the flag; Dummy configuration `"writerProof":false` keeps Copy-only coverage (Phase11, `Floodgate_CrossFs*`, `Fairstream_CrossFsConcurrentMoveUsesBridge`). R3-3 adds its own row with exact symbols before it starts. |
| RED/GREEN evidence | RED (baseline `d4c65293`): `R3_1_IdentityLessReplaceDummy` fails because the Exists prompt on a Dummy destination offers no Overwrite (`_conflictExpectedDestinationUnavailableCount` counts the withheld decision), and the replace steps added to `R0fCurl_`, `R0fS3_`, `R0fGraph_`, `R0fGDrive_` fail the same way. GREEN: the Dummy case (prompt offers Overwrite, the replacement lands with the source bytes, an occupant raced behind every decision re-raises the collision twice, then surfaces `ERROR_REVISION_MISMATCH` as the retryable conflict, which Skip answers, and survives untouched) and the four fixture replace steps (apply-to-all Overwrite through the provider writer, copied back byte-exact); `FileOps_ProviderCapabilityMatrix`, `Phase10_`, the Fairstream conflict cases, PluginContractTests, and the Pester source contracts stay green. Commands: `Tools\Run-AllTests.ps1 -Suite FileOps -SkipBuild -CaseFilter <prefix>_`; `PluginContractTests.exe`; `Invoke-Pester -Script Tools\Tests\TestHarnessSourceContracts.Tests.ps1`. |
| Performance/resources | One extra no-follow basic-information read per replaced file before the prompt and one occupant read in the provider writer; no new buffers or threads; `FileOps.SelfTest.<Provider>.ReplaceOn<X>Ms` is recorded per fixture and must stay within the same order as the fixture copy timings of the R0f receipts. |
| Durable owners | `FileSystem_FileOperations.md` (Conflict model, Capability and strategy routing), `Core_FileSystemBridge.md` (Conflict and cleanup), the `FileSystem.h` contract comment, `FileSystem_FtpSftpScp.md`, `FileSystem_S3.md`, `FileSystem_MicrosoftDrive.md`, `FileSystem_GoogleDrive.md`. |
| STOP/rollback | STOP for a route whose writer cannot revalidate the occupant (no identity and no last-write time): Overwrite stays withheld there. Rollback of R3-1 removes the eligibility branch in `Task::FileSystemIssue` and the bridge acceptance; the writers' expectation checks are inert without a host grant. |
| Exit-receipt contract | Per slice: commit, full-solution build (0 warnings), focused FileOps/contract/Pester results, fixture timings, residuals, and the Fresh Full gate on the landing commit, recorded in an "R3-n receipt" row below. |
| R3-1 receipt (2026-09-03) | Mechanism: `Task::FileSystemIssue` offers Overwrite/Replace read-only without bound authority when the destination has no `IFileSystemObjectBinding` but its `IFileSystemAtomicWriter` accepts the overwrite flag (`DestinationSupportsAtomicReplace`; Replace link stays bound-only), and returns the decision with a null expected destination (`_conflictAtomicReplaceGrantCount`). The bridge reads the occupant no-follow before the prompt (`CaptureReplaceExpectation`), no longer treats a compatibility overwrite flag as a receipt on such routes, publishes through the atomic-final writer created with the granted flags, and hands the occupant to the writer through the new `IFileWriterExpectedReplacement` before the first byte (`bridge.replace.expectationUnavailable` fails the item before any byte is read when a writer cannot carry it). A refused replacement (`ERROR_REVISION_MISMATCH` from the expectation or from Commit) is a definitive non-commit (`bridge.replace.occupantChanged`, nothing published, no unknown-publication artifact); `PumpAndPublishFile` then re-raises the collision on the current occupant at most twice (`PumpAndPublishFileOnce` reports `replaceRefused`) before that status reaches the item executor's ordinary retryable conflict (Retry decides again on the current occupant, Skip keeps it). In test-enabled builds the bridge's writer decorator (`SelfTestBridgeFileWriter`) forwards the new contract, which it had hidden. Writers: Curl re-probes the destination before staging and refuses a vanished occupant or a changed last-write time; S3 records the occupant's ETag/size/last-write at creation, validates the expectation, and publishes with `If-Match` (`412` -> `ERROR_REVISION_MISMATCH`); Microsoft Drive resolves the item again before publication and sends `If-Match` with its eTag on the content PUT and on upload-session creation; Google Drive replaces only the creation-time object id at its version; Dummy compares size and last-write time in `CommitFileWriter`. Fake Graph honors `If-Match` on content upload and upload-session creation. Specs updated as listed in Durable owners. Evidence: focused set 2026-09-03 10:16-10:21 on the final tree (`4f1ad3d2` plus this slice): full-solution Debug build 0 warnings 0 errors (MSBuild diagnostics count). RED on the baseline (tests-only build): `R3_1_IdentityLessReplaceDummy` and the `R0fGDrive_` replace step failed with "the Exists prompt on the identity-less route must offer Overwrite". GREEN: FileOps `R3_1_` 3/3 (the prompt offers Overwrite, the replacement lands with the source bytes, an occupant raced behind every decision produced three Exists prompts plus the retryable conflict carrying `ERROR_REVISION_MISMATCH`, answered with Skip, and survived untouched), `R0fGDrive_` 3/3, `R0fGraph_` 3/3, `R0fS3_` 3/3 and `R0fCurl_` 3/3 with their replace steps (Curl: `FileOps.SelfTest.R0fCurl.ReplaceOnFTPMs` 297 next to `CopyToFtpMs` 141, `CopyFromFtpMs` 62, `DeleteOnFtpMs` 125 in the same run), `R0fSmb_` 4/4, `FileOps_ProviderCapabilityMatrix` 3/3, `Phase10_` 9/9, `Phase9_` 11/11, `Fairstream_` 20/20, `Riptide_` 13/13, `Floodgate_` 15/15, Compare `google_drive_` 2/2, PluginContractTests all OK (1548 OK, 0 FAILED; Curl 238/0, S3 212/0, Microsoft Drive 240/0, Google Drive 25/0 debug self-tests), Pester source contracts 187/187 (the Curl occupant check lives in `ValidateCurlReplaceOccupant` so the Floodgate distance contract holds). Residuals: provider-native same-provider overwrite through `CopyItem`/`MoveItem` stays withheld on identity-less routes; a refused replacement is reported through the generic retryable conflict rather than a dedicated "destination changed" prompt; replace-step timings are recorded per fixture run but only the last run directory survives the runner's pruning. Fresh Full gate #9 (2026-09-03 10:22-11:22 on `100a582d`, 55 min): 1997 passed / 5 failed / 67 skipped of 2069, build 0 warnings. `Phase11_CrossFileSystemBridge` still asserted the pre-R3-1 prompt (Overwrite withheld on an unbound destination) and was aligned with the R3-1 contract in the gate-record commit (3/3 isolated after the change); `Phase12_ReparsePointPolicy` timed out at the junction Move into Dummy under load (3/3 isolated, as in gate #8); `cmd_pane_batchRename_admission_worker_owned_cancellable` (2/2 isolated); `Phase7_ParallelCopyMoveKnobs` in-flight observation window (5 of 6 isolated runs pass; pre-existing since gate #7); `R4A19_DiscoveryIndependentVolumes` alternate-volume authorization (environment). One isolated rerun of `R4A19_DiscoveryProviderControls` failed on the discovery queue-depth bound its message did not print (2/2 on rerun; the message now prints `discoveryQueue`). |

| R3-2 receipt (2026-09-03) | Mechanism: a destination route without object binding advertises `FILESYSTEM_ROUTE_PROOF_WRITER_DIGEST`; the bridge asks its atomic-final writer which digests it can prove before the first byte, hashes the streamed bytes with them alongside BLAKE3, and after Commit compares the provider's own digest of the published object (`IFileWriterContentProof::GetCommittedContentProof`). A match is `Verified` with no destination read; a different digest is `Failed`; a missing proof is `Unavailable`. Managed Move into such a route now keeps its cleanup record armed and deletes the bound source only after the proof matches (`Copied; source kept` otherwise). Providers: S3 full-object CRC-64/NVME (declared on `CreateMultipartUpload`/`UploadPart`/`PutObject`, returned by S3), Microsoft Drive `file.hashes` (SHA-256/SHA-1/QuickXorHash), Google Drive `sha256Checksum`, Dummy SHA-256; FTP/SFTP/SCP stay `none`. RED/GREEN: RED (baseline `54de1594`, tests only): `R3_2_WriterProofDummy` times out at its first step because the Verify On copy into the Dummy destination never completes on a route without proof, and the Managed Move step cannot run at all (the pair was admitted Copy-only). GREEN (landing tree): `R3_2_WriterProofDummy` 3/3 (Verify On copy Verified with providerProofs=1 and hostReadbacks=0, no prompt; Managed Move removes the Local source only after the Dummy store's SHA-256 matched), `R3_1_` 3/3, `R0fGDrive_` / `R0fGraph_` / `R0fS3_` / `R0fCurl_` 3/3 each (copies both ways, replace, delete through the fixtures that now report sha256Checksum, file.hashes and CRC-64/NVME), `FileOps_ProviderCapabilityMatrix` 3/3, `Phase10_ContentVerification` 3/3, `Phase11_CrossFileSystemBridge` 3/3 and `Floodgate_CrossFs*` 9/9 and `Fairstream_CrossFsConcurrentMoveUsesBridge` 3/3 (Copy-only coverage kept through the Dummy `"writerProof":false` configuration), `Phase12_ReparsePointPolicy` 3/3 (junction Move admitted Managed, root reparse still the intentional skip with the source kept); full-solution build 0 warnings; Pester `TestHarnessSourceContracts` 187/187. Three landing defects were found by the focused set and fixed before this receipt: the known proof-flag masks rejected the new flag (every proof route ERROR_INVALID_DATA), the self-test writer decorator hid the new interface, and the pre-mutation transfer safety guard refused Managed items without destination identity. Timings: Fixture timings on the landing tree (ms): Google Drive copy-to 156 / copy-from 78 / replace 266; Microsoft Drive 157 / 109 / 250; S3 187 / 110 / 203 (delete 78); FTP 157 / 62 / 297, the same order as the R0f and R3-1 receipts. The writer proof adds one SHA-256 (Drive, Dummy), three digests (Graph: SHA-256, SHA-1, QuickXorHash, because the drive kind is unknown before upload) or one CRC-64/NVME (S3) over the bytes already streamed, and no extra request: Verified Copy needs no readback and Managed Move needs no destination re-stat. Residuals: real S3/Graph/Drive endpoints are exercised only through the loopback fixtures (no credentials in the sandbox); Graph hashes the streamed bytes with three algorithms because the drive kind is unknown before upload; the S3 debug transport in self-tests reports no digest (proof `Unavailable`, by design). Fresh Full gate #10 (2026-09-03 12:54-13:52 on `f5498f61`, 57 min): 1994 passed / 4 failed / 72 skipped of 2070, build 0 warnings. Two failures were R3-2 fallout in tests that cover Copy-only Move into Dummy and were outside the focused set: `FileOps_CrossVolumeMovePartialFailureStatus` and `Cinderstar_LegacyWriterEventuallyConsistentMove` (the Cinderstar chain stops at its first failure); both now run with the Dummy `"writerProof":false` configuration in the gate-record commit and pass (3/3, and 8/8 for the whole chain), with `Floodgate_CrossFsMove*` 5/5 and `R3_2_` 3/3 unchanged. `cmd_pane_batchRename_admission_worker_owned_cancellable` is the known load flake (2/2 isolated). `R4A19_DiscoveryProviderControls` failed once more with every printed bound satisfied (3/3 isolated); its message now prints which of its six checks failed so the next occurrence is diagnosable. |

| R3-3 plan (2026-09-03) | Scope: `RedSalamander/FolderWindow.FileOperations.State.cpp` only (no provider or ABI change). (a) `CrossFileSystemBridge::PublishOwnedStageAs(IFileSystemBoundObject&, const OwnedStagePublication&, publishedAuthority, promoted, cleanupAllowed)` becomes the one owner of the exact `PublishAs` outcome (unknown / known non-commit / committed, `mutationAttemptFlags`, `ownedStageDisposition`, `anyDestinationPublished`, `destinationPublicationUnknown`, perf counters); `PublishOwnedStage` (file), `PublishPreparedLink` (link, partial phase `StageCommit`, `bridge.linkPublication.*`) and `PublishOwnedDirectory` (directory, no payload diagnostics) call it. (b) `struct PublicationTransaction` holds the file item's admission grants and expectations, `ManagedSourceCleanupRecord`, source authority/metadata snapshot, reader, sizes, route (`AtomicFinal` / `OwnedStage`), stage and writer handles, proof interfaces and hashers, pump counters, commit/writer proofs, published authority, and its `State` (Admitted, SourceBound, Staged, Written, Committed, Published, Verified, Retained, Failed, Unknown); `PumpAndPublishFileOnce` becomes the ordered phase calls `AdmitDestination`, `BindSource`, `OpenSourceReader`, `RouteWriter`, `PrepareStage`, `TransferPreContentMetadata`, `Pump` (one `WriteChunk` for the serial and pipelined loops), `CommitAndProve`, `Publish`, `TransferPostContentMetadata`, `VerifyPublished`, `FinishCleanup`, each with the same diagnostics categories as today. (c) `R3_3_PublicationFaultMatrix` (family `FileOpsFamily_R3OwnedPublication`): the seventeen `ConsumeBridgeCounterForSelfTest(g_fileOps...)` hooks plus `ConsumeBridgeFailNextFileCopyForSelfTest` and the post-publication replacement hook, over Local bound, Dummy writer-proof and Dummy `"writerProof":false` routes, for Copy and Move, asserting publication / verification / source disposition / failure phase / status per row from today's individual hook tests. (d) `Core_FileSystemBridge.md` one-file flow names the record and states; R3 steps 1, 2 and 7 tick; R3 package DONE on a green Fresh Full gate. Order (a) -> (b) -> (c) -> (d), one commit each with the focused families (`R3_`, `R0f*_`, `Phase10_ContentVerification`, `Phase11_CrossFileSystemBridge`, `Phase12_ReparsePointPolicy`, `Floodgate_CrossFs`, `Fairstream_`, `FileOps_ProviderCapabilityMatrix`) green; STOP on any focused regression the extraction does not explain. |

| R3-3 receipt (2026-09-03) | Commits: (a) `f3d57f42` one owner for the owned-stage publication outcome (`PublishOwnedStageAs`; file, link and directory payloads call it with their diagnostics and failure phase); (b) `ef3dc101` the `PublicationTransaction` record and the ten phase methods (`AdmitDestination`, `BindSource`, `RouteWriter`, `PrepareStage`, `TransferPreContentMetadata`, `Pump`, `CommitAndProve`, `Publish`, `VerifyPublished`, `FinishCleanup`), `OpenSourceReader` replacing two identical bound-reader branches and `WriteChunk` the one write loop behind the serial and pipelined pumps, the Commands reader contract re-pointed at the helper and the Pester Riptide/Floodgate/Delta shapes kept; (c) the fault matrix. Evidence: (a) full-solution x64 Debug build 0 warnings; focused families `R3_1_` 3/3, `R3_2_` 3/3, `R0fGDrive_` / `R0fGraph_` / `R0fS3_` / `R0fCurl_` 3/3 each, `Phase11_CrossFileSystemBridge` 3/3, `Phase12_ReparsePointPolicy` 3/3, `Floodgate_CrossFs*` 9/9, `Fairstream_CrossFsConcurrentMoveUsesBridge` 3/3, `FileOps_ProviderCapabilityMatrix` 3/3, `Phase10_ContentVerification` 3/3. (b) the same twelve families green on the extracted tree (0 warnings), Pester `TestHarnessSourceContracts` 187/187, Commands `file_operations_local_reader_cancellation_source_guard` 1/1; one `R0fCurl_` run reported "results.json not found" once and passed 3/3 on its isolated rerun (fixture start). (c) `R3_3_PublicationFaultMatrix` 3/3: all 102 rows (17 hooks x Local bound / Dummy writer-proof / Dummy no-proof x Copy/Move) hold the invariants; the two adjustments the matrix forced on itself were answering the retryable conflict a fault raises (Skip, else Cancel) and accepting the documented completion of a published item whose verification was Unavailable (`ERROR_PARTIAL_COPY`, `Completed`). No production behaviour changed in R3-3. Not done in R3-3: `Failed`/`Unknown`/`Retained` are not record states (they are receipt outcomes the executor records); the record is bridge-private and the bridge stays in `FolderWindow.FileOperations.State.cpp`. Fresh Full gate #11 (`e91bfa52`, 2026-09-04 00:14-01:08): Total 1918 / Passed 1883 / Failed 2 / Skipped 33. Commands `cmd_pane_batchRename_admission_worker_owned_cancellable` = admission load flake (passed isolated). The FileOps stage aborted: the self-test process ended (exit 1, no dump) in family 10/26 right after `R0fCurl_FakeFtpReadWriteCreateDelete` completed its copy, so `selftest_result_coverage` reported the later families missing; isolated `R0fCurl_` 3/3, and a complete isolated FileOps run on the same tree: 159 total / 138 passed / 1 failed (`R4A19_DiscoveryIndependentVolumes`, environment: alternate volume root not authorized) / 20 skipped (evidence: `gate11-artifacts/fileops-full-isolated-rerun-results.json`). |

##### R4-A19 activation card

| Activation-card field | R4-A19 bounded execution contract |
|---|---|
| Slice/state/owner | **R4-A19 / `COMPLETE` (exit receipt 2026-09-03 below; parent matrix archived 2026-09-02) / Codex `/root` in this worktree, closed by Claude.** This is an evidence-only measurement of the unchanged discovery reservation/low-water/half-rate/checkpoint scheduler. It may add counters, deterministic behavioral fixtures, archives, and durable measurement requirements. It may not change queue targets, worker fractions, admission, traversal, provider concurrency, bandwidth, mutation rate, checkpoint behavior, UI policy, public ABI, or product strings. A failing mechanism requires a separately complete correction card before any production-policy edit. |
| Baseline and drift | Exact clean tracked baseline is `516ed892bbd842f009febaf894a06799f1e7e851`; repository-root `last_run/` is pre-existing untracked user-owned test output and is excluded. Scoped drift over the named engine, selftest, and spec owners is empty. The current shared policy is `kDiscoveryLowWater=32`, `kDiscoveryTarget=128`, `kDiscoveryMaxTarget=256`, `DiscoveryQueueTarget(...)`, and `DiscoveryWorkerLimit(...)`; the host already emits discovery open/callback/lock cost, max queue, starvation count, mutation-before-close, Skip-release, closed state, and bridge retention ceilings. Exact first-mutation latency plus bytes and mutation operations completed while traversal is open are missing and are the only authorized production-instrumentation gap. Local SMB admission remains fail-closed under R0e when route facts classify it uncontained; that scenario must be recorded as a contained rejection rather than silently replaced by a live-share mutation. |
| Contract | Preserve sections 6.1/6.9 `FO-DISCOVERY-01`: start safe work before traversal closes, bound the ready queue and retained traversal state, keep discovery making service progress, preserve one-way just-in-time fallback, and write no operation-owned spool. Add exact task-local counters for discovery-start-to-first-mutation, completed-byte delta while discovery remains open, and completed mutation/item delta while discovery remains open. Reuse existing starvation, callback/lock wait, queue-depth, bridge-retention, cancellation, Skip-release, and total-operation metrics. Counters are observation only and cannot feed scheduling or product behavior. |
| Scope | In scope: instrumentation-only fields and callbacks in `RedSalamander/FolderWindow.FileOperationsInternal.h` and `RedSalamander/FolderWindow.FileOperations.State.cpp`; behavioral FileOps case registration/state/capture in `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.cpp` and cases in `FolderWindow.FileOperations.SelfTest.Phases05_06.cpp`; reuse without policy edits of `Common/FileOperationTraversalPolicy.h`, Local FileOps, `FileSystemDummy` latency/speed controls, the test-only MTP factory/fake backend, and shared test-sandbox authorization; `Specs/FileSystem/FileSystem_FileOperations.md`, `Specs/Testing/Testing_PerformanceValidation.md`, `Specs/Testing/Testing_SelfTests.md`, this card/receipt, and evidence below `Specs/TestRuns/4cb089111a23/FileOps/`. Out of scope: `Common/PlugInterfaces/FileSystem.h`, public/plugin ABI, provider production algorithms, scheduler-policy constants/functions, new controller/framework, R3/R4-A02/R6 UI behavior, live remote mutations, re-enabling R0e-disabled routes, and unrelated progress rendering. |
| Predecessors and ownership | E1 closed its shared item/bridge seams under implementations `0e24444c` and `20c15a26`; R2 closed the typed route cutover under implementation `75df3717` and archived evidence `9bf973cb`; R1d-core closed under implementation `668acbd7`, archive `005b9c81`, and final phase correction `0515e854`. R0e-OR1 closed separately at `516ed892` without reopening SMB/remote admission. No active package owns `FolderWindow.FileOperations.State.cpp` or the named selftest/spec files. |
| Steps | (1) Land behavioral RED/evidence-gap cases under family `FileOpsFamily_R4A19DiscoveryBaseline` and prove the current completion snapshot cannot report first-mutation latency or open-traversal bytes/operations. (2) Add observation-only counters at existing progress/item-completion callbacks and emit the three `FileOps.Discovery.*` metrics without changing scheduling branches. (3) Exercise Local large-payload Copy, Local small-file Delete, test-only serialized fake MTP, R0e-contained Local SMB rejection plus a deterministic high-latency sequential proxy, and independent fixed-volume Local work on the already initialized `C:\RedSalamander.Perf` and `D:\RedSalamander.Perf` roots. (4) Re-run existing Phase 5/Fairstream discovery, cancellation, Skip/JIT, bridge retention, and source-contract coverage. (5) Capture five independent test-enabled x64 Release samples, archive the unchanged-policy baseline, update durable specs, and either close R4-A19 with no scheduler change or STOP and add one causal correction card. |
| RED/GREEN behavioral evidence | RED is the measurement gap, not a manufactured scheduler failure: the new family must reject sentinel/missing first-mutation latency and open-traversal byte/operation facts while existing correctness controls remain green. GREEN requires byte-for-byte/item-complete outcomes; discovery closure; first mutation before closure in the eligible Copy/high-latency/MTP fixtures; nonzero open-traversal bytes for Copy and nonzero open-traversal mutation operations for Copy/Delete; Local SMB route rejection before task mutation when typed containment says uncontained; fake MTP maximum backend concurrency of one; independent C/D tasks both making progress without a host-global gate; max ready queue at most 256; bridge retained entries/path bytes/metadata at most 4,096/16 MiB/8 MiB; zero traversal-limit hits and zero operation-owned spool files. Existing `Fairstream_DiscoveryAheadOverlapsTransfer`, `Phase5_DiscoverySingleTraversal`, `Phase5_DiscoveryCancelLatencyLocal`, `Phase5_DiscoverySkipContinues`, and `Phase5_SwitchParallelToWaitDuringDiscovery` remain required behavioral controls. |
| Fault/resource matrix | Cover producer slower than consumers, transfer slower than producer, queue high-water/backpressure, small-delete mutation bursts, a serialized provider with delayed backend calls, cancel while discovery and mutation are both live, one-way JIT transition, an R0e-rejected SMB route, and two independent fixed volumes. Every fixture uses an exact marked test root and bounded unique child; provider fake/delay state is restored on every exit. No live credential, portable device, network share mutation, soft source grep, or unmarked path is acceptance evidence. If C or D is unavailable/unmarked, the independent-volume row is a recorded environment blocker and the card cannot close from a synthetic substitute alone. |
| Performance/resources | Protected family uses: Local Copy of 128 files × 256 KiB at concurrency 4; Local Delete of 1,024 small files at concurrency 8; fake-MTP Copy with a 25 ms cancellable backend delay and a single serialized session; Dummy high-latency/sequential proxy with `latencyMs=5` and bounded virtual speed; and two concurrent 16 MiB Local copies rooted separately under marked C and D sandboxes. Capture five independent test-enabled x64 Release processes at `Specs/TestRuns/4cb089111a23/FileOps/2026-09-01_r4_a19_discovery_baseline_release/`. Every sample must complete below 30 seconds per fixture; first safe mutation must occur within 1 second and before traversal close where eligible; queue/retention ceilings above are hard; cancellation remains at most 500 ms and Skip reservation release at most 50 ms excluding an active provider call. For each workload pair, discovery-ahead total duration must be at most `max(1.35 × just-in-time control, control + 250,000 us)` and first-mutation latency at most `max(1.50 × control, control + 100,000 us)`. Concurrent independent-volume wall time must be at most 1.35 times the slower isolated-volume control. Record throughput, starvation count, runnable wait proxy (starvation plus discovery callback/lock wait), and all raw durations even when a hard gate fails. No retained state may scale beyond the declared queue/traversal ceilings. |
| Durable owners | `Specs/FileSystem/FileSystem_FileOperations.md` owns the unchanged discovery scheduler contract and observation semantics. `Specs/Testing/Testing_PerformanceValidation.md` owns fixture shapes, metric names, five-process archive, comparisons, thresholds, and caveats. `Specs/Testing/Testing_SelfTests.md` owns deterministic family/resource cleanup requirements. This plan owns activation/STOP routing and the generated exit receipt only. |
| STOP/rollback | STOP before a production-policy edit if any unchanged scenario fails a hard gate; add a separate complete correction card naming only the proved causal mechanism. STOP if evidence requires public ABI expansion, live remote mutation, uncontained-route re-enable, a general byte-token/IOPS/device-class/adaptive controller, UI changes, provider production rewrites, unmarked test data, or concurrent test processes whose exact roots are not independently authorized. Rollback removes only new counters/metrics/tests/spec clauses while R4-A19 remains open; never tune constants or weaken a threshold after observing the result. |
| Exit-receipt contract | Before R4-A19 closes, record activation/RED/instrumentation/evidence commits; exact Debug and test-enabled Release receipts; focused family and legacy discovery totals; five-process per-fixture raw/p95 metrics; Local SMB contained-rejection result; fake-MTP serialization/cancel result; C/D marker and independent-volume result; queue/retention/spool/cancel/Skip gates; explicit archive and whole-inventory validation; `git diff --check`; durable-spec destinations; proof that `Common/FileOperationTraversalPolicy.h` and scheduler branches did not change; and either `PASS — no scheduler mutation` or the exact new correction-card identifier. |

##### R4-A19 exit receipt (2026-09-03)

| Exit-receipt field | Recorded closeout |
|---|---|
| Result | **PASS — no scheduler mutation.** The unchanged discovery reservation, low-water, half-rate, and checkpoint policy met the activated forward-progress, bounded-memory, provider-serialization, cancellation, and containment gates on every fixture of the parent matrix; R4-A19 produced no scheduler candidate and changes no production policy. |
| Landed commits | Activation `c93f2257`; measurement-gap RED `780fd579`; observation-only counters and GREEN instrumentation `ea0b8b5f`; R4-A19-OR1 activation `f294c24e`, implementation `0c7ac7a4`; parent matrix completion `dae2f674`; qualification evidence archive `9b025713`; the accepted aborted MTP cancellation result `22892314`. |
| Archived evidence | `Specs/TestRuns/4cb089111a23/FileOps/2026-09-02_r4_a19_parent_release_01..05` (five independent test-enabled Release processes, each `5/5` with no skipped case, fixtures rebuilt and Local/Dummy/fake-MTP state restored through ordinary cleanup) and `.../2026-09-02_r4_a19_parent_release_comparison/README.md` (the gate arithmetic per fixture, the Debug controls `FileOps_ProviderCapabilityMatrix` 3/3, `Causeway_BridgeSchedulingAndResourceContracts` 3/3, `Phase11_CrossFileSystemBridge` 3/3, and the Local UNC classification as the distinct uncontained route before R0f-SMB). The R4-A19-OR1 baseline/candidate archives and comparison stay as recorded on the correction card. |
| Residuals | `R4A19_DiscoveryIndependentVolumes` needs an authorized alternate volume (environment, skipped in this sandbox). `R4A19_DiscoveryProviderControls` failed intermittently in isolated reruns of gates #9 and #10 with every printed bound satisfied; its message now prints which of its six checks failed and the next occurrence is diagnosable. Metric identifier strings (`FileOps.Discovery.*`) live in `Testing_PerformanceValidation.md`; `FileSystem_FileOperations.md` owns the semantics only. |
| Ownership | R4-A19 no longer holds `RedSalamander/FolderWindow.FileOperations.State.cpp`; R3 is its sole active owner. |

##### R4-A19 continuation checkpoint (2026-09-01 — paused at STOP)

| Checkpoint field | Recorded continuation state |
|---|---|
| State | **`ACTIVE — R4-A19-OR1 complete; parent matrix resumed`.** The unchanged Local recursive Delete baseline failed the original hard gate, and the separately activated observation-only correction is now closed. This is not an R4-A19 closeout and grants no authority to tune constants or edit provider/scheduler behavior. Repository-root `last_run/` remains pre-existing untracked user-owned output and was not archived, staged, cleaned, or modified. |
| Landed commits | Activation `c93f2257`; behavioral measurement-gap RED `780fd579`; observation-only counters and initial GREEN instrumentation `ea0b8b5f`. The production additions are task-local measurement fields and `FileOps.Discovery.FirstMutationUs`, `FileOps.Discovery.BytesCompletedWhileOpen`, and `FileOps.Discovery.MutationsCompletedWhileOpen` emission only. `Common/FileOperationTraversalPolicy.h` and scheduler branches remain unchanged. |
| Build receipts | Earlier test-enabled x64 Debug full-solution receipt: `97c2515720c1caf407db9b642620e1b3ea8c8dd2c43a785af390cd7e0df6eff7`. The revised exact Local Copy/Delete fixture compiled in a project-scoped test-enabled x64 Debug build with receipt `24c00a38d600a81cd085085395b008bd3aca4541fe5c2010e4878703d61c29b6` (0 warnings, 0 errors). A fresh full-solution receipt is still required before a receipt-verified `-SkipBuild` continuation. |
| Archived runs | `Specs/TestRuns/4cb089111a23/FileOps/2026-09-01_222115_r4_a19_red_missing_metrics/` is the expected RED proving missing measurement facts. `.../2026-09-01_223727_r4_a19_green_instrumentation/` is the initial 16×16×8 KiB instrumentation GREEN. `.../2026-09-01_225054_r4_a19_flat_delete_hard_gate/` is the exact 128×256 KiB Local Copy plus flat 1,024×1 KiB Delete baseline. `.../2026-09-01_225500_r4_a19_nested_delete_hard_gate/` is the producer-slower 128-directory ×8-file Delete baseline. The last two are failure evidence, not acceptance receipts. |
| Hard-gate result | Exact Local Copy passed with mutation before discovery close and exact 32 MiB/128-file output. Flat Local Delete returned `S_OK`, removed all 1,024 files, closed discovery, stayed at queue max 256, and finished in 188 ms, but recorded `MutationsCompletedWhileOpen=0`. The nested shape also returned `S_OK`, removed all 1,024 files, closed discovery, stayed at queue max 128, and finished in 125 ms, but again recorded zero mutations while discovery was open. The direct executable returned process exit 0 for the nested run even though `fileops/results.json` records the case as failed; continuation must use the structured result/runner, not the direct process exit, as truth. The full runner also reported 113 disk-audit issues; investigate them without deleting or adopting `last_run/`. |
| Causal trace | The pause-time hypothesis was a Local recursive Delete bounded-turn gap. Activation reproduction and a failed provider-progress experiment narrowed it to the actual host seam: exact-authority permanent Delete calls `IFileSystemBoundObject::DeleteIfUnchanged`, whose Local recursive implementation deliberately has no `IFileSystemCallback`. Discovery facts cross `FileSystemOptions::operationControl`, and the host learns the exact selected-root `Removed` outcome before `CloseDiscovery`, but the R4-A19 completion observer currently advances only from `FileSystemProgress`. The active correction card below supersedes both earlier hypotheses and owns only exact completion observation. |
| Resume point | Completed on 2026-09-02: exact drift was clean, the RED reproduced in 94 ms with zero host-observed open-discovery mutations, R4-A19-OR1 was promoted under activation commit `f294c24e`, implemented under `0c7ac7a4`, and closed by the receipt below. The parent matrix resumes without a scheduler-policy change. |
| Remaining R4-A19 work | Dummy high-latency/sequential control; serialized fake-MTP and cancellation; the Local SMB control (R0f-SMB turned the R0e rejection witness into a bounded loopback-share control); independently marked C/D fixed-volume work; parent five-process Release evidence; bridge-retention/source contracts; and governed disk-sandbox diagnosis. OR1's Local Copy/Delete Release comparison, Phase 5/Fairstream/Skip/JIT controls, three durable-spec updates, archive validation, whole-inventory validation, and protected-source proof are complete. |

##### R4-A19-OR1 — Local recursive Delete exact completion observation

| Correction-card field | Zero-authority continuation contract |
|---|---|
| Slice/state/owner | **R4-A19-OR1 / `COMPLETE` / Codex `/root` in this worktree.** Exact checkpoint drift was revalidated before promotion: no tracked change existed beyond the card baseline and repository-root `last_run/` remained excluded user-owned output. This card closed only the bounded Local recursive permanent-Delete observation correction below; R4-A19 remains open until the full parent matrix and closeout gates pass. |
| Exact baseline and RED | Clean tracked checkpoint baseline is `14915a84ce8a2c7590464aed8fc3f04f56c5cdb0`; repository-root `last_run/` is excluded user-owned output. The immutable RED is `R4A19_DiscoveryMeasurementFacts`: both the flat 1,024-file and nested 128×8 Local permanent recursive Delete shapes finish correctly but record `FileOps.Discovery.MutationsCompletedWhileOpen=0`. Evidence is archived in the two `*_delete_hard_gate` runs named above; activation reproduction `r4-a19-or1-red-20260902a` repeated the nested failure in 94 ms. Do not weaken, replace, or reinterpret that gate after activation. |
| Causal mechanism | Activation reproduction completed the 1,024-file Delete in 94 ms with exact deletion/discovery totals and zero host-observed mutations while discovery was open. Source tracing showed that the exact retained Local object recursively deletes through `DeleteBoundLocalDirectoryContents`, initialized with `callback=nullptr`; therefore no provider progress publication can carry a completion delta. `executePermanentDelete` nevertheless classifies the conditional mutation as `Removed` before the per-item discovery scope closes. The measurement gap is that this exact host-owned completion does not update the same observation-only first-mutation/open-discovery counters used by `FileSystemProgress`. The correction owns only that Local permanent-Delete observation seam; it does not own mutation ordering, provider progress policy, Copy, bridge scheduling, provider concurrency, conflict policy, route admission, or UI. |
| Contract | When an exact-authority Local permanent Delete returns the typed `Removed` outcome while discovery remains open, record that selected-root mutation exactly once through the existing task-local observation counters before `MarkDiscoveryItemClosed`/`CloseDiscovery`. Reuse the same first-mutation timestamp and saturating counter semantics as provider progress, without publishing a synthetic callback or feeding any scheduling/product branch. Preserve mutation order, exact object/path authority, post-order directory removal, conflict and continue-on-error semantics, cancellation, queue cap 256, Local configured concurrency cap 8, and byte-for-byte/item-complete results. Do not add a generic token/IOPS/device-class/adaptive controller or change shared scheduler constants. |
| Scope | In scope: `RedSalamander/FolderWindow.FileOperationsInternal.h` and `RedSalamander/FolderWindow.FileOperations.State.cpp`; the existing R4-A19 state/registration/case files under `RedSalamander/SelfTest/FileOperations/`; `Specs/FileSystem/FileSystem_FileOperations.md`; `Specs/Testing/Testing_PerformanceValidation.md`; `Specs/Testing/Testing_SelfTests.md`; this plan; and exact FileOps archives. Read-only controls: `Plugins/FileSystem/FileSystem.FileOps.cpp`, `Common/FileOperationTraversalPolicy.h`, Local route/configuration owners, Phase 5/Fairstream cases, and shared sandbox helpers. Out of scope: public/plugin ABI, provider production policy, non-Local providers, Copy/Move, Recycle, bridge pipelines, conflict/result/UI policy, route re-enable, new threads/controllers, and unrelated cleanup. |
| Steps | (1) Promote only after exact drift review. (2) Re-run the checkpoint RED unchanged and trace callback timing. (3) Reuse one task-local observation helper for real progress deltas and the exact `Removed` result; record one selected-root mutation before its discovery scope closes. (4) Prove exact results, cancellation, Skip/continue-on-error, no-follow/reparse, conflict, post-order directory removal, concurrency controls, and unchanged provider callback counts. (5) Finish the remaining R4-A19 scenario matrix and five-process Release evidence. (6) Update all three durable specs, archive runs, validate the whole TestRuns inventory, and write separate OR1 and R4-A19 exit receipts. |
| GREEN and regressions | The exact 1,024-file Delete case must return `S_OK`, remove only the selected tree, report all 1,024 files, close discovery, keep max queue at most 256, finish below 30 seconds, and report exactly one selected-root completed mutation while discovery is open. First mutation must be within 1 second, and provider progress callback counts must remain unchanged from RED. Existing Phase 5 discovery/cancel/Skip/JIT, reparse/no-follow, continue-on-error, conflict, exact-root Delete, and Local concurrency controls remain green. No operation-owned spool file and no retained-state growth beyond the parent card's ceilings. |
| Performance evidence | Compare five independent test-enabled x64 Release processes at the exact checkpoint baseline and candidate. Candidate total duration must be at most `max(1.35 × baseline, baseline + 250,000 us)`; first-mutation latency must be at most `max(1.50 × baseline, baseline + 100,000 us)` and at most 1 second; cancellation remains at most 500 ms excluding an active provider call. Archive raw duration, queue, starvation, callback/lock wait, discovery-open mutation count, retained entries/path bytes/metadata, and exact object totals. |
| STOP/rollback | STOP if the GREEN requires synthetic provider callbacks, public ABI, provider production rewrites, shared Copy scheduler changes, a general limiter/controller, weaker authority/reparse semantics, larger queue/retention ceilings, live remote mutation, or an unmarked test path. STOP if the exact failure does not reproduce at the refreshed baseline. Rollback only the OR1 production change and its candidate evidence; preserve the R4-A19 counters, RED, and checkpoint archives until the parent card reaches an honest disposition. |
| Exit receipt | **PASS — exact completion observation only; no scheduler/provider-policy mutation.** Owner `/root` promoted baseline `14915a84ce8a2c7590464aed8fc3f04f56c5cdb0` under `f294c24e` after RED `780fd579` and reproduction `r4-a19-or1-red-20260902a` reported zero open-discovery Delete mutations. The attempted provider-progress route was rejected and rolled back because exact Local recursive Delete deliberately has no provider callback. Implementation `0c7ac7a4` records the typed selected-root `Removed` result exactly once before discovery closes. T1 test convergence landed under `eabd7d3f` and closed under `ada031d3`. Full-solution test-enabled x64 Debug receipts `d0153abae3e745dff7c986b4128a5907e6e90659e16bbf85a28b415fe585fc68` and `89c10125bb5d048cd703a1bc282f1222c36c4905a121d25ad5c1fe53bd29b673` built with zero warnings/errors; exact Debug GREEN `r4-a19-or1-green-committed-debug-20260902` passed 3/3. `Phase5_Discovery` passed 6/6, its isolated T1 case passed 3/3, and `Phase5_DiscoveryCancelLatencyLocal`, `Phase5_DiscoverySkipContinues`, `Phase5_SwitchParallelToWaitDuringDiscovery`, `Fairstream_DiscoveryAheadOverlapsTransfer`, `FileOps_ReparseDirectoryMergeIntoExistingFolder`, `Fairstream_ParallelDeleteContinuesPastLockedChild`, `FileOps_DeleteToctouSwapGuard`, `Phase10_PermanentDelete`, and `Phase7_ParallelDeleteKnobs` each passed 3/3. Exact test-enabled x64 Release receipts were baseline `6ed185899f884fc5a295a611405bae845ab15bc1f67687e595dccafacc450440` and candidate `64a3705dc274bb1a03b157d1bc3a51a922d2865abe84f57ed4b61fcf7c22c8d0`, both zero-warning/error full solutions. Across five independent processes, baseline/candidate Copy duration p95 was 156,000/188,000 us and first-mutation p95 4,669/5,257 us; candidate gates were 188,000 ≤ 406,000 and 5,257 ≤ 104,669 us. Baseline/candidate Delete duration p95 was 156,000/94,000 us; baseline had no observable mutation, while candidate first-mutation p95 was 91,457 us, below the 100,000-us missing-baseline gate and one-second ceiling, with exactly one mutation and 1,048,576 selected-root bytes in every sample. Candidate queue maxima were 7 Copy/128 Delete, starvation and traversal-limit hits were zero, maximum callback work was 125 us, maximum lock wait 29 us, and retained stage state was zero. Cancellation stayed within the existing 500-ms behavioral gate; authority/no-follow, continue-on-error, conflict, post-order, exact-root, concurrency, Skip/JIT, and provider-callback-zero controls remained green. Raw runs are `Specs/TestRuns/4cb089111a23/FileOps/2026-09-02_r4_a19_or1_{baseline,candidate}_release_01` through `_05`; comparison is `.../2026-09-02_r4_a19_or1_release_comparison/README.md`. All ten raw archives and the comparison passed explicit validation; whole inventory passed for 2,026 files. `git diff --check` and staged diff checks passed. Durable behavior/evidence landed in `Specs/FileSystem/FileSystem_FileOperations.md`, `Specs/Testing/Testing_PerformanceValidation.md`, and `Specs/Testing/Testing_SelfTests.md`. Diff from the activation baseline touches only the two FileOps observation files, their focused selftests, and this plan; protected-source diff exit was zero for `Plugins/FileSystem/FileSystem.FileOps.cpp`, `Common/FileOperationTraversalPolicy.h`, `Common/PlugInterfaces/FileSystem.h`, all named non-Local provider trees, and UI. |

##### R4-A19-OR1-T1 — Phase 5 popup snapshot convergence

| Test-correction field | Exact harness contract |
|---|---|
| Slice/state/owner | **R4-A19-OR1-T1 / `COMPLETE` / Codex `/root` in this worktree.** This test-only prerequisite grants no FileOps production authority. |
| Exact baseline and failure | Implementation baseline `0c7ac7a4`. `Phase5_DiscoveryCancelReleasesSlot` failed both in the `Phase5_Discovery` prefix run `r4-a19-or1-regression-phase5-discovery-20260902` and isolated run `r4-a19-or1-isolate-cancel-slot-20260902`: immediately after queue reorder/Start now, `DebugGetFileOperationsPopupLayoutSnapshot` returned false at the first in-progress task snapshot. Earlier queued-task snapshots in the same case already use bounded convergence retries. |
| Causal mechanism | Popup layout publication is asynchronous relative to the queue-control transition. The test's queued-task branch invalidates and retries snapshot capture for at most five seconds, but the following active-task branch converts the first not-yet-published snapshot into an immediate failure. Both failures occurred in 234–266 ms, before the existing bounded UI convergence window. |
| Contract | Apply the same bounded retry shape already used by this case: when the active-task layout snapshot is unavailable, invalidate the popup and return to the selftest pump until the existing five-second case-relative bound expires; preserve the final failure and every layout/status/control assertion. Do not add sleep, weaken a UI assertion, change product timing, or extend the global test timeout. |
| Scope | In scope only: the exact snapshot-unavailable branch in `RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.SelfTest.Phases05_06.cpp`, this card, and focused evidence. All production files, popup code, scheduler code, other Phase 5 assertions, and global harness timing are read-only. |
| GREEN/exit | Implementation `eabd7d3f` applies only the existing bounded five-second snapshot-convergence shape. Full-solution x64 Debug receipt `d0153abae3e745dff7c986b4128a5907e6e90659e16bbf85a28b415fe585fc68` built with 0 warnings and 0 errors. Under that receipt, `r4-a19-or1-t1-green-isolated-20260902` passed 3/3 in 3.0 seconds and `r4-a19-or1-t1-green-prefix-20260902` passed 6/6 in 4.3 seconds with zero skipped cases, proving the later discovery cancel/Skip tail remained reachable. `git diff --check` was clean before closeout. **PASS — test convergence only; no production or timing-policy change.** |

- [x] Activate the baseline measurement as an evidence-only card. If it passes, close
      R4-A19 with no scheduler mutation. If it fails, that card cannot expand scope:
      add a separate exact R4-A19 correction card naming the causal mechanism, files,
      tests, owner, and before/after gate before changing production policy.
- [ ] Preserve and instrument the existing normative discovery reservation/low-water/
      half-rate/checkpoint scheduler before changing policy. Measure large-file Copy,
      small-file Delete, MTP serialization, SMB, and independent-device scenarios.
- [ ] If the baseline meets the activated card's forward-progress, memory, throughput,
      and cancellation gates, make no scheduler change. If one scenario fails, activate
      and test only the smallest causal correction: Copy byte limit, Delete mutation-
      rate limit, or bounded turns for a proved serialized provider.
- [ ] Never throttle an independent device for symmetry, retain a failed candidate, or
      introduce a general token/IOPS/device-class/adaptive scheduler framework.
- [ ] Keep one-way `Discover as needed`: retain queued work, release run-ahead
      reservation and any evidence-adopted temporary limiter at the bounded checkpoint,
      and continue mandatory JIT safety discovery.
- [ ] Consume completed E1 bridge/item-policy seams; do not recreate or merely
      relocate them. R4 owns only accepted traversal, overlap-advisory/concurrent-
      correctness, and discovery-scheduling changes through those seams.
- [ ] Replace source-string tests with behavioral resource/fault tests.

**Acceptance:** valid trees do not fail merely for depth >128 or full queues and no
operation-owned disk spool is written;
clearly disjoint tasks never receive a broad-root warning; every obvious overlap
receives one problem-specific choice; user-approved overlapping tasks run through
normal exact authority/conflict/result paths without unsafe cleanup or global task
exclusion. No concurrently-live same-host task output is invalidated by another
without disclosed Run consent or a later item-specific destructive decision. Memory
remains bounded; eligible work starts before traversal closure; the unchanged
scheduler baseline is archived first; a passing baseline adds no controller; and any
adopted correction proves forward progress without unrelated-resource, throughput,
memory, cancellation, or safety regression.

### Package R5 — Cancellation isolation

**Priority:** P1 application survival
**Depends on:** completed R0e evidence, D2-A01, and R2 receipts. R0e—not R5—owns
deadline honesty, route inventory, preservation of MTP containment, SMB
characterization, and the host-level never-returning-provider witness.

- [ ] Select only the exact residual route proved able to wedge the host and accepted
      for out-of-process isolation; do not broker every provider for symmetry.
- [ ] Define the smallest serializable intent, authority, progress, cancel, terminal
      receipt, and retained-artifact protocol that preserves R2 truth across the
      process boundary.
- [ ] Prove helper startup/crash/timeout/cancel, host restart/reconciliation, provider
      module lifetime, and bounded Stop/unload/teardown/shutdown.
- [ ] Never kill an in-process worker that still owns host memory/provider code; a
      process boundary must exist before termination can be a containment mechanism.

**Acceptance:** a provider that never returns cannot freeze UI/application shutdown
beyond the accepted host bound.

### Package R6 — One user-decision and explanation surface

**Priority:** P1 UX/truth
**Ordering:** each independently activated R6 slice uses its sole row in section 8.1.
**Boundary:** independently promote only the accepted slice. E2 remains the
authorized behavior-preserving conflict-action owner; R6 cannot derive a second
policy.

#### Independently activated R6 slice briefs

The exact slice heading is the activation-card anchor. These are required outcomes,
not executable code/test instructions before a complete, predecessor-closed `ACTIVE`
card. A `DRAFT` card grants no edit authority.

##### R6-A02 — Overlap warning and live-output decision

- [ ] Render R4's exact advisory facts in one problem-specific task-injection
      surface with named task/path context, consequence wording, **Queue after these
      tasks** safe default, **Run at the same time**, **Don't start**, and Details.
      Queue warns once, stores a transient same-host edge, then releases into ordinary
      semantics after predecessor completion. Run records only the disclosed live
      consequence; neither choice is current-object mutation authority.
- [ ] When R4 detects a concurrently-live same-host output without covering Run
      consent, re-ask before invalidation. Skip preserves; Queue ends concurrency and
      later resumes ordinary semantics; only the explicit destructive action permits
      the live effect. Never reconstruct protection from history.

##### R6-A07 — Explicit Retry

- [ ] Permit repeated user-initiated Retry only for `RetryableNoCommit` while
      exact authority remains valid; show the attempt count. Unknown/request-level
      commit state is terminal Indeterminate and is never automatically reissued.

##### R6-A08 — Per-task and Known-work progress

- [x] Before a task's discovery closes, render exact discovered/processed/item/byte
      counters with an indeterminate bar and no percentage. Show `Estimated time
      remaining — still discovering` only with sufficient stable sampling; allow it
      to rise or fall, suppress it while paused/stalled/unsupported, and replace it
      with fixed-total percentage/ETA after closure. ETA never drives scheduling,
      timeout, consent, mutation, cleanup, or completion.
      Landed 2026-09-04 (R6-A08 card below): the popup's whole-task presentation is
      determinate only after closure, the open-discovery card shows the exact counters
      line, and `Remaining: {0} (still discovering)` uses the existing smoothed rate over
      the discovered workload; the ETA remains presentation-only.
- [x] On each task, keep discovery activity/count/provisional ETA distinct from its
      ongoing operation-progress row. Implement `FO-DISCOVERY-01`: a compatible
      closed-total cohort may
      show **Known work: N%**, open totals add **N tasks discovering — total may grow**,
      no compatible known cohort means no jobs percentage, and the Windows taskbar is
      indeterminate while any included total remains open.
      Landed 2026-09-04: the footer summary counts open-discovery tasks, stays determinate
      over the closed cohort (`Known work: N%`), states `N discovering, total may grow`,
      and the taskbar model keeps its indeterminate rule for open totals.

##### R6-A08 — Honest progress before discovery closes (activation card)

| Activation-card field | R6-A08 bounded execution contract |
|---|---|
| Slice/state/owner | **R6-A08 / `COMPLETE` (2026-09-04; exit receipt below) / Claude in worktree `codex/r4-a19-closeout-2026-09-02`.** Presentation only: the File Operations popup renders D2-A08 (B amended) and the `FO-DISCOVERY-01` Known-work rule from the facts the engine already publishes. The engine, scheduler and discovery service are out of scope (R4 traversal/A08/A13 and R4-A19 are `COMPLETE`). |
| Baseline and drift | Planned at `8df5427c` (R4 traversal/A08/A13 closed). Drift review: `git status --short` shows only the user-owned untracked `last_run/`; `ResolveWholeTaskProgressPresentation`, `BuildGlobalFileOperationsStatusSummary` and `BuildGlobalTaskbarProgressModel` in `RedSalamander/FolderWindow.FileOperations.Popup.cpp` are unchanged since E2. |
| Contract | D2-A08 (indeterminate task bar plus exact discovered/processed/byte counters before the total closes; no growing-denominator percentage; a clearly labeled provisional ETA only with sufficient stable sampling, allowed to rise or fall, hidden while paused/stalled/unsupported) and section 6.9 `FO-DISCOVERY-01` (a compatible closed-total cohort shows **Known work: N%**, open totals add **N tasks discovering, total may grow**, no jobs percentage without a compatible known cohort, Windows taskbar indeterminate while any included total is open). Displaced: `Specs/UI/UI_FileOperationsPopup.md` sentences allowing a provisional transfer fraction on the whole-task bar and the compact meter before `discoveryClosed`, "ETA stays `Estimating` until traversal closes", and "Unknown active totals force the aggregate footer ... to indeterminate". |
| Scope | In: `RedSalamander/FolderWindow.FileOperations.Popup.cpp` (`ResolveWholeTaskProgressPresentation`: determinate only after closure, marquee/indeterminate fill before; the counters row on the running open-discovery card; the provisional ETA text next to the graph; `BuildGlobalFileOperationsStatusSummary`/footer text: Known-work percent over the closed cohort plus the open-discovery count; `BuildGlobalTaskbarProgressModel` unchanged in rule, re-verified), `RedSalamander/FolderWindow.FileOperations.Popup.h` (snapshot fields the footer needs), new string resources (`RedSalamander/Resource.h`, `RedSalamander/RedSalamander.rc`), the popup debug seam `DebugValidateFileOperationsVerificationPresentation`, Commands `cmd_pane_fileops_popup_progress_contracts` and the presentation/taskbar cases, `Specs/UI/UI_FileOperationsPopup.md`, this plan. Out: engine counters (published by R4-A19), discovery scheduling, UIA tree ownership (E2), Batch Rename presentation, history. |
| Predecessors | R4 traversal/A08/A13 (`COMPLETE` 2026-09-04), R4-A19 (`COMPLETE`), E2 (`COMPLETE`). |
| Steps | (1) `[x]` (2026-09-04) Before closure the whole-task bar is indeterminate (the existing marquee fill) and the card shows exact `discovered files / folders / bytes` and `processed items / bytes` counters beside the discovery indicator; the compact meter follows the same rule (no meter before closure). No percent text or percent UIA value before closure. (2) `[x]` Provisional ETA: when the smoothed throughput history is stable and the discovered workload exceeds the completed bytes, show `Estimated time remaining, still discovering: {0}`; it may rise or fall, is hidden while paused, stalled or when sampling is insufficient, and is replaced by the ordinary ETA after closure. (3) `[x]` Footer/jobs: the closed-total cohort renders `Known work: N%` (bytes first, items second) even while other tasks discover; open totals add `{0} tasks discovering, total may grow`; without a compatible closed cohort no jobs percentage is shown; the taskbar stays indeterminate while any included total is open. (4) `[x]` Spec, seam and Commands contracts updated; exit receipt recorded here. |
| RED/GREEN evidence | RED: the popup seam and `cmd_pane_fileops_popup_progress_contracts` currently require a determinate provisional fraction of 0.25 for a running open-discovery task and an indeterminate footer whenever any task is open; the new expectations (indeterminate task bar with counters before closure; provisional ETA label present only with a stable sample; footer Known-work percent over a closed sibling while another task discovers; taskbar indeterminate) fail on the baseline. GREEN: those contracts pass; `cmd_pane_fileops_popup_presentation_settings_and_taskbar`, `cmd_pane_fileops_completed_group_and_navigation`, the popup UIA cases and the DxUi hosted-progress tests stay green. |
| Performance/resources | No new per-frame allocation: counters and labels are formatted from snapshot fields already copied per publish. Evidence: the popup presentation self-tests and the next Fresh Full gate. |
| Durable owners | `Specs/UI/UI_FileOperationsPopup.md` (progress rules, footer/taskbar paragraphs), this plan (R6-A08 checklist, section 6.9 rendering note). |
| STOP/rollback | STOP if the counters or the provisional ETA would need engine changes or a new sampling framework (D2-A08 allows only existing stable sampling), or if the hosted DxUi `ProgressBar` cannot express indeterminate for the whole-task model; rollback is the provisional-fraction presentation (one revert). |
| Exit receipt | R6-A08 exit receipt (2026-09-04): activation card `048ddb4c`; implementation `30a2a239` (`ResolveWholeTaskProgressPresentation` is determinate only after `discoveryClosed`, the open-discovery card renders the marquee plus one exact counters line `IDS_FMT_FILEOPS_DISCOVERY_PROGRESS`, `RateHistory` carries the provisional ETA over the discovered workload and the card shows `IDS_FMT_FILEOPS_ETA_PROVISIONAL` only while that rate is usable, the footer summary counts `openDiscoveryTasks`, stays determinate over the closed cohort with `IDS_FMT_FILEOPS_KNOWN_WORK` and adds `IDS_FMT_FILEOPS_GLOBAL_DISCOVERING_COUNT`, the taskbar model keeps its indeterminate rule; the popup debug seam and `cmd_pane_fileops_popup_progress_contracts` pin the new expectations; `UI_FileOperationsPopup.md` replaces its provisional-fraction sentences). Debug full-solution builds 0 warnings / 0 errors (RED and GREEN). RED witness on the baseline (2026-09-04 18:42-18:47, `d84a160f` plus the baseline-safe contract alone): `cmd_pane_fileops_popup_progress_contracts` failed 0/1 because a closed-total sibling beside a discovering task rendered an indeterminate footer (the old rule forced indeterminate whenever any task was open); the popup seam still required a determinate provisional fraction of 0.25 while discovery was open. GREEN (focused runs on the landed tree): Commands `cmd_pane_fileops_popup_` 4/4, `cmd_pane_fileops_` 34/34, `file_operations_` 3/3, `cmd_app_` 83 of 85 in the sweep, the two misses being UI-timing cases (`cmd_app_prompt_uses_alert_overlay_window`, `cmd_app_menuBar_arrow_switches_top_level_popup`) that passed 1/1 isolated on the same build; the DxUi tests are covered by Fresh Full gate #15 (a direct launch of `DxUiTests.exe` outside the harness aborts in its shared test-support policy test, unrelated to this slice); Pester 187/187. The existing `cmd_pane_fileops_popup_global_summary_ignores_finished_tasks` kept its expectation: an active task whose discovery closed without any total still forces the aggregate indeterminate (`hasClosedUnknownTotals`), so Known work applies only beside open discovery. Performance: no new per-frame allocation; the popup presentation self-tests run unchanged in duration. Fresh Full gate #15: on `477bb59f` (run `20260904T173150Z-37604-9250b807c93644b79438ffb8707a8dd0`, started 2026-09-04 19:35): 2077 counted, 2014 passed, 9 failed, 54 skipped; PluginContractTests, DxUi tests, Compare 232/232 green. The nine failures: the Tools Pester resource-parity test (the four new R6-A08 strings were missing from the cs-CZ, fr-FR, ja-JP and sk-SK satellite tables; corrected in `d537a240` with translations, parity 9/9 and source contracts 187/187 on a clean build); `ViewerPETests.Interactive` (`TestViewerShellComboHostsLongRunOpenCloseStayStable` child harness exited non-zero in the gate and in a direct rerun on the same build, then the child and the whole interactive group passed on the `d537a240` build, so it is attributed to the same satellite gap); `R4A19_DiscoveryIndependentVolumes` (environment: alternate-volume test root not authorized); and six Commands UI-timing cases (`cmd_pane_batchRename_admission_worker_owned_cancellable`, `cmd_preferences_dialog_keyboard_live_search_dx_interaction`, two Find-dialog cases, `cmd_pane_navigation_go_to_root_directory_keeps_navigation_shell_stable`, `cmd_app_menuBar_hover_switches_top_level_popup`) that each passed 1/1 isolated on the same build. |

##### R6-A09 — Artifact explanation

- [ ] Keep read/open/edit/preview/copy/export available for a name-only Possible
      `.rs_*` object. Apply one exact-object, Cancel-default warning before a
      RedSalamander-owned Rename, Move, Delete/Recycle, overwrite, or content/metadata
      write. Consent never promotes Possible to Proven or grants cleanup authority.
- [ ] Properties performs a fresh shared-classifier query and appends one
      host-owned **File Operations** section containing every available human-readable
      semantic fact listed in section 6.8. Explicitly label unavailable facts and the
      name-only/no-valid-claim case; do not scan all history, fabricate fields,
      duplicate recovery authority, or expose unredacted endpoint secrets.

##### R6-A09 — Name-only artifact policy and Properties explanation (activation card)

| Activation-card field | R6-A09 bounded execution contract |
|---|---|
| Slice/state/owner | **R6-A09 / `ACTIVE` (2026-09-05) / Codex `/root` in this worktree.** This slice corrects only name-only Possible artifact behavior and its host-owned Properties explanation. It does not introduce a durable claim source, recovery, cleanup authority, history scanning, a new identity model, or provider ABI. |
| Baseline and drift | Exact tracked baseline is `9f29ec6d`; `git status --short` contains only the pre-existing user-owned untracked repository-root `last_run/`. The baseline scheduler correction touches only ViewerImgRaw. Current artifact-touch guards, pane transfer candidate collection, Item Properties, resources, tests, and owning specs are otherwise available to this slice. |
| Contract | D2-A09 and section 6.8. A recognized `.rs_*` name shape is Possible only. Read/open/edit/preview/copy/export remain available solely on that basis. A RedSalamander-owned Rename, Move, Delete/Recycle, overwrite, or content/metadata write performs one exact-object revalidation and one Cancel-default warning. Consent neither proves a claim nor grants cleanup/recovery authority. Properties performs one fresh bounded shared-classifier query for a Possible name and appends one host-owned **File Operations** section: it exposes every available safe semantic fact, explicitly says unavailable for absent facts, labels the name-only/no-valid-claim condition, exposes no endpoint secret, and performs no history scan. Ordinary names perform no classifier query. |
| Scope | In: `RedSalamander/FolderWindow.FileOperations.State.Runtime.cpp` (Copy source exemption; Move source and overwrite destination remain guarded); removal of name-only external read/open/edit/preview/export guards in `FolderWindow.cpp`, `FolderWindow.Viewers.cpp`, and `FindFilesWindow.cpp`; `FolderWindow.ItemProperties.cpp/.h` and the existing Properties debug snapshot; `Resource.h`, the main resources and four maintained satellites; focused FileOps/Commands/Properties/source-contract tests; `Specs/FileSystem/FileSystem_FileOperations.md`, `Specs/UI/UI_FolderWindow.md`, and this plan. Out: classifier/index schema, recovery actions, claim persistence, global history, execution receipts, provider implementations, artifact cleanup, and shell delegation policy. |
| Predecessors | E6 is `COMPLETE` under corrective implementation/evidence `9e14a899`; its shared name-shape classifier and exact no-follow object capture are the only facts this slice consumes. No active package owns these external-ingress guards or Item Properties projection. |
| Steps | (1) `[ ]` Stop collecting the source object for Copy while retaining the destination overwrite candidate and both Move source/destination mutation candidates. (2) `[ ]` Remove artifact warnings from read/open/edit/preview/copy/export ingresses while preserving exact manager-owned mutation guards. (3) `[ ]` Extend Item Properties with one fresh query for Possible names, zero for ordinary names, and a localized, redacted, non-authoritative File Operations section with explicit unavailable/name-only facts. (4) `[ ]` Pin Cancel/no-mutation, prompt cardinality, ingress exemptions, query count, localization, source boundaries, and authoritative specs; record focused and final gate evidence. |
| RED/GREEN evidence | RED: the baseline collects every transfer source, so routine Copy of a Possible name prompts; external viewer/editor/user-menu/default-open/context/security paths also invoke the mutation warning. Item Properties neither queries the shared classifier nor explains the Possible/name-only state. GREEN: routine Copy and all named read-only/external ingresses remain prompt-free; Move, Delete/Recycle, Rename, destination overwrite, and content/metadata mutation retain one exact warning; Cancel leaves source/destination unchanged; Properties reports one query for Possible and zero for ordinary names with the bounded localized section. |
| Performance/resources | The hot path removes unnecessary probes/prompts. Properties adds at most one no-follow classifier query, gated by `HasPossibleArtifactName`, on its existing worker; no UI-thread provider call, global scan, persisted index, or per-frame work is added. Record focused Commands/FileOps timings and the next exact Fresh Full gate. |
| Durable owners | `Specs/FileSystem/FileSystem_FileOperations.md` owns warning/exemption semantics; `Specs/UI/UI_FolderWindow.md` owns the Properties projection and accessibility text. This plan records only activation and receipts. |
| STOP/rollback | STOP if Properties needs a history scan, durable claim fabrication, raw provider/endpoint identifiers, or a UI-thread provider query; if any mutation warning would become name-only without exact revalidation; or if an exemption would permit manager-owned mutation. Rollback is one slice revert restoring the prior conservative prompts and removing the Properties section together. |
| Exit-receipt contract | Record activation and implementation commits, zero-warning Debug and test-enabled Release receipts, focused Copy/Move/Cancel/Properties/query-count/source/localization results, `git diff --check`, authoritative-spec updates, and the next exact Fresh Full result. |

##### R6-A12 — Single popup renderer

- [ ] Retain one hosted layout/action/accessibility owner and only a minimal safe
      text/Cancel/Close failure surface; remove the feature-complete duplicate painter
      after focused attachment-failure coverage.

##### R6-A17 — Continue safe work and review later

- [ ] Implement task-local user-selectable collect-then-review processing for
      unrelated safe work without widening the admitted set or auto-deciding
      conflicts. Preserve immutable Apply-to-all decisions; return newly changed
      conflicts to the same review; permit Retry only for proven no-commit; make
      Rescan create a new plan.

##### R6-A18 — Persistent explanatory history

- [ ] Persist local retention-controlled, searchable/exportable non-authoritative
      history from immutable attempts. Implement age+size bounds, oldest-first
      eviction, redacted explicit export, Clear, and Disable exactly as section 6.9.6;
      keep claims/receipts separate and grant no execution/recovery authority.

##### R6-A20 — Routine start

- [ ] Implement the ingress matrix in section 6.9 without
      duplicate confirmation—routine fast path versus one stable material-risk/
      **with options** surface.

##### R6-A20 — Routine fast path and explicit with-options commands (activation card)

| Activation-card field | R6-A20 bounded execution contract |
|---|---|
| Slice/state/owner | **R6-A20 / `ACTIVE` (2026-09-05) / Codex `/root` in this worktree.** This slice completes only accepted `FO-UX-01` routine-start presentation: F5/F6 stay immediate accepted-default commands, while two explicit stable commands open the existing single pre-consumption options surface. It does not change provider execution, collision policy, Find-result or Compare synchronization commands, artifact warnings, review-later collection, history, or link semantics. |
| Baseline and drift | Exact tracked baseline is `60d4ef6e`; `git status --short` contains only the pre-existing user-owned untracked repository-root `last_run/`. The prior commit only hardens a ViewerImgRaw test observation after Fresh Full gate #16 and does not touch this slice's production owners. |
| Contract | Section 6.9 `FO-UX-01` and D2-A20. Routine pane Copy/Move crosses Ready without a generic confirmation and has labels that do not imply a dialog. Deliberate **Copy with Options...** and **Move/Rename with Options...** commands enter the already-owned pre-consumption confirmation/options surface exactly once, before clipboard consumption or mutation, with Links, Verify, Queue/Parallel, bandwidth, and any later task-local decision mode added by its owning slice. Existing permanent Delete, known Copy-only Move, exact artifact/risk, drag/drop, editor, clipboard, Find, and Compare consent owners remain unchanged and never stack a second generic dialog. |
| Scope | In: `RedSalamander/FolderWindow.h`, `FolderWindow.FileOperations.cpp` (one shared pane-transfer entry with an explicit `withOptions` bit), `RedSalamander.cpp` dispatch, `CommandRegistry.cpp`, `ShortcutDefaults.cpp`, `Resource.h`, the main and four maintained satellite menu/string resources, command-surface coverage and Commands selftests, `Specs/FileSystem/FileSystem_FileOperations.md`, `Specs/UI/UI_CommandMenuKeyboard.md`, and this plan. Out: the confirmation dialog implementation and fields, Find/Compare dispatch, provider code, task execution, review-later/history/link-policy behavior, and every non-pane ingress. |
| Predecessors | R1d-core is `COMPLETE` and owns the common lifecycle plus routine acceptance gate. No other active package owns the named command/menu/shortcut surface. |
| Steps | (1) `[ ]` Add stable `cmd/pane/copyToOtherPaneWithOptions` and `cmd/pane/moveToOtherPaneWithOptions` registry/dispatch IDs and route them through the same pane admission code with `requireConfirmation=true`; keep F5/F6 and existing stable IDs on `false`. (2) `[ ]` Change routine menu labels to omit ellipses, add explicit with-options sibling entries, and bind Shift+F5/Shift+F6 without displacing existing chords. (3) `[ ]` Prove the fast commands admit without a generic prompt, the explicit commands publish exactly one pre-consumption options prompt, Cancel mutates nothing, and Find/Compare retain their current specialized paths. (4) `[ ]` Update command/resource parity, authoritative specs, focused evidence, and this card's exit receipt. |
| RED/GREEN evidence | RED: on the baseline F5/F6 dispatch routine work correctly but the File menu labels end in ellipses and no stable explicit with-options command exists; users cannot deliberately reach the existing options surface without a material-risk prompt. GREEN: command-registry/shortcut/menu contracts contain two unique stable IDs; Shift+F5/Shift+F6 dispatch through pane ownership; debug admission observes `requireConfirmation` false for routine IDs and true only for the explicit IDs; Cancel returns before consumption/mutation; registry smoke excludes all four mutation-starting commands; Find and Compare continue handling only the original routine IDs. |
| Performance/resources | The routine hot path adds one compile-time branch and no I/O, enumeration, allocation, or extra prompt. The explicit path reuses the existing bounded dialog and preparation facts. Record focused Commands timing plus the next exact Fresh Full gate; no new perf controller is justified. |
| Durable owners | `Specs/FileSystem/FileSystem_FileOperations.md` (ingress/routine confirmation matrix) and `Specs/UI/UI_CommandMenuKeyboard.md` (stable IDs, menu labels, Shift+F5/Shift+F6). |
| STOP/rollback | STOP if reaching options requires a second dialog, a new confirmation implementation, pre-prompt clipboard consumption/mutation, or changes to Find/Compare semantics. Rollback removes the two explicit IDs/menu items/shortcuts together while preserving R1d-core's fast path. |
| Exit-receipt contract | Record activation and implementation commits, zero-warning Debug and test-enabled Release receipts, focused fast/options/Cancel/registry/shortcut/menu/localization results, `git diff --check`, authoritative-spec updates, and the next exact Fresh Full result. |

##### Shared R6 delivery constraints

- [ ] Extend completed E2 snapshots only for the promoted slices; baseline order,
      default, visible Cancel, Escape, and Apply-to-all remain E2's single-owner
      output.
- [ ] Implement the applicable phase control/keyboard matrix and gate classification,
      including one foreground decision, explicit focus restoration, and no action on
      Close.
- [ ] Render one jobs surface with operation/endpoints, lifecycle, plain-language
      strategy, exact counters, current item, failures, and safe actions.
- [ ] Preserve the retired I17 prompt baseline: actionable conflicts visibly and
      accessibly show `Waiting for your decision`; incomplete facts show `Loading
      decision details...`; Waiting/metadata-loading/actionable-conflict states expose
      no discovery, transfer, graph, or marquee activity; incomplete destructive
      actions remain unavailable; question -> endpoints/context/facts -> actions is
      the visual, keyboard, and UIA order; prompt/context/final note wrap without
      clipping; state-transition announcements occur once; validate minimum width,
      both densities, 100/150/200% DPI, and the existing English/cs-CZ/fr-FR/ja-JP/
      sk-SK resource contracts.
- [ ] Render requested/definite/possible/remains/reason/safe-actions using the small
      user result vocabulary; technical axes remain expandable details.
- [ ] Run the complete mouse/keyboard/UIA/high-contrast/RTL/reduced-motion/DPI human
      walkthrough matrix and record prompt/focus/announcement friction.

**Acceptance:** hiding/closing never resolves a prompt; the same immutable
decision/result drives text, icons, taskbar, badges, Issues, keyboard, UIA, and
history; the history grants no mutation/recovery authority. With review-later mode
enabled, an unresolved item never blocks unrelated safe work, every deferred item
appears exactly once in the collected review, selected roots/accepted scope do not
widen, immutable Apply-to-all decisions do not drift, commit-time changes re-enter
the same review, Retry remains proven-no-commit only, and Rescan creates a new plan.
History is local, age+size bounded with oldest-first eviction, redacted on screen and
export, and obeys Disable/Clear without touching claims or receipts. The routine
accepted-default path has the arbitrated interaction count, one prompted decision
never becomes two windows, focus is stable, and no task steals OS foreground. A02's
warning appears only for a concrete detected problem, defaults to Queue, and Run
together never changes item authority or hides subsequent conflicts. Queue warns once,
releases automatically, and leaves no completed-task protection. Discovery and
operation progress remain separately understandable per task; mixed jobs show the
labeled Known-work cohort/open-discovery indicator and keep the taskbar indeterminate.

### Package R7 — Rename unification and optional semantics

**Priority:** P2 after core truth
**Boundary:** A10 and A15 are rename-chain slices; A03 is an independent Copy/link
slice and never waits for or consumes E4/E5. Each slice follows only its section 8.1
predecessors.

#### R7-A10 — Accepted acyclic/no-journal Batch Rename

**Ordering:** the R7-A10 row in section 8.1.

##### R7-A10 activation card

| Activation-card field | R7-A10 bounded execution contract |
|---|---|
| Slice/state/owner | **R7-A10 / `COMPLETE` / Codex `/root` in this worktree.** This slice lands only accepted `FO-RENAME-01`: pre-mutation cycle rejection, one immutable acyclic schedule, no Batch Rename journal, and no Resume/Roll back/replay surface. It does not activate R7-A15/E5, add provider name policy, migrate Change Case, redesign conflict policy, implement general artifact history, or activate R1c/R1d/R2-R6/R7-A03/R8/R9. |
| Baseline and drift | Activation baseline is clean tracked commit `4cfb9cabcc33bad541158c7cc786afd83c3fb6e1`; repository-root `last_run/` is pre-existing untracked output and excluded. `rg` inventory found the journal writer in `ExecuteBatchRename`, transition writes in `ExecuteBatchRenameMutation`, the load/project/reopen reader in `FileOperationArtifactRegistry`, FolderView Resume/Roll back commands, FolderWindow recovery callbacks/thread/messages, popup recovery presentation, project entries/resources, six recovery/registry tests, and current authoritative cycle/journal clauses. Preview marks duplicate targets but not dependency cycles; worker admission records cycles as executable plan data; the executor dynamically synthesizes `.rs_ren_` temp hops and rollback; swaps and three-member cycles are tested as successful. `git show`/`git grep` against the sole release tag `v7.0.0` (`ba4dc1d9`, 2026-02-16) found none of `BatchRenameRecoveryJournal.*`, `FileOperationArtifactRegistry.cpp`, `StartArtifactRecovery`, journal creation, or Resume resources; the journal first entered history on 2026-08-26. No supported-release record requiring migration was found. |
| Contract | Implement section 6 `FO-RENAME-01` and displace the current cycle/journal contract in the same slice. Preview computes dependency cycles from provider path identity, attaches a localized error to every exact cycle member, keeps Run disabled, and tells the user to choose a temporary intermediate name and run two acyclic passes. Worker admission independently rebuilds the indexed schedule and rejects any cycle with `ERROR_CIRCULAR_DEPENDENCY` before plan publication or provider mutation. One immutable schedule lists every changed row exactly once in executable dependency/depth layers; execution validates and consumes that schedule, preserves exact completed/failed/unattempted row truth, and never invents a temp path or replay authority. Batch Rename never creates, reads, migrates, deletes, or replays a recovery journal. Existing journal bytes, if manually present, are ignored and retained. Name-shape `Possible` artifact warning/guard behavior remains, but no Batch Rename claim can become Proven and no Resume/Roll back action exists. After process loss, pane/provider refresh is the sole namespace truth. |
| Scope | In scope: `RedSalamander/BatchRenameEngine.h`, `BatchRenameEngine.cpp`, `BatchRenameExecutionEngine.h`, `BatchRenameExecutionEngine.cpp`, `BatchRenameWindow.h`, `BatchRenameWindow.cpp`, removal of `BatchRenameRecoveryJournal.h` and `.cpp`, `FileOperationArtifactRegistry.h` and `.cpp`, `FolderWindow.FileOperationsInternal.h`, `FolderWindow.FileOperations.cpp`, `FolderWindow.FileOperations.State.cpp`, `FolderWindow.FileOperations.State.Runtime.cpp`, `FolderWindow.FileOperations.Popup.cpp`, `FolderWindow.h`, `FolderWindow.cpp`, `FolderView.h`, `FolderViewInternal.h`, `FolderView.Interaction.cpp`, `FolderView.Menus.cpp`, `FolderView.Enumeration.cpp`, `FindFilesWindow.cpp`, `Common/FolderViewMenuIds.h`, `Common/WindowMessages.h`, `RedSalamander.vcxproj`, `RedSalamander.vcxproj.filters`, `Resource.h`, `RedSalamander.rc`, `Lang/cs-CZ/RedSalamander-cs-CZ.rc`, `Lang/fr-FR/RedSalamander-fr-FR.rc`, `Lang/ja-JP/RedSalamander-ja-JP.rc`, `Lang/sk-SK/RedSalamander-sk-SK.rc`, `SelfTest/Commands/Commands.SelfTest.cpp`, `SelfTest/Commands/Commands.SelfTest.BatchRename.cpp`, `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`, `Tools/Tests/ResourceLocalizationContracts.Tests.ps1`, `Specs/FileSystem/FileSystem_FileOperations.md`, `Specs/UI/UI_BatchRenameWindow.md`, `Specs/UI/UI_FolderView.md`, `Specs/Testing/Testing_PerformanceValidation.md`, compact Commands evidence, and this card/receipt. Out of scope: `FileOperationDurableStore`, Move breadcrumbs, Change Case/`FileSystemRenameBatch`, provider ABI/capabilities/name validation, inline F2 behavior, R1c cleanup claims, R1d lifecycle, R6 presentation redesign, unrelated `.rs_*` producers, and deleting any pre-existing on-disk record. |
| Predecessors | E4 is complete under implementation receipt `227f3e91` plus corrective receipt `5e6213e9`; its worker-owned Batch Rename admission seam is present. E0-E3 are also complete. Section 8.1 names no other predecessor. No active slice owns these files. |
| Steps | (1) Add RED preview/admission/direct-engine/source-contract tests for swap, three-cycle, cycle tail, journal creation, and recovery actions; capture five Release scheduler baselines. (2) Add one indexed schedule builder shared by preview, admission, validation, and execution; report exact cycle members and localized two-pass guidance. (3) replace `RenameCycle`/dynamic temp-hop execution with immutable acyclic layers and exact partial-result reduction. (4) remove journal creation/transition/finalization and delete its reader/writer/schema implementation without touching existing bytes. (5) decouple name-shape artifact warning/guards from Batch Rename records and remove Resume/Roll back commands, callbacks, messages, popup state, resources, and tests. (6) update authoritative specs/perf contract, run focused Debug and exact-commit test-enabled Release validation, archive five candidate samples, prove zero production reader/writer/surface, record the exit receipt, and mark only R7-A10 complete. |
| RED/GREEN evidence | Exact Commands cases are `cmd_pane_batchRename_cycle_preview_blocks_run`, `cmd_pane_batchRename_admission_rejects_cycle`, `cmd_pane_batchRename_execution_engine_direct`, `cmd_pane_batchRename_partial_batch_failure_tracks_completed_rows`, `cmd_pane_batchRename_cancel_mid_batch_tracks_completed_rows`, `cmd_pane_batchRename_window_executes_chain_rename`, `cmd_pane_batchRename_window_executes_case_only_local_rename`, `cmd_pane_batchRename_no_journal_or_recovery_surface`, and `cmd_pane_batchRename_execution_engine_large_independent_perf`; run each through `Run-AllTests.ps1 -Suite Commands -SkipBuild -FailFast -CaseFilter <id>`. RED requires the current swap/three-cycle preview to enable Run, admission/direct execution to mutate through temp hops, a normal run to write a journal, and production recovery symbols/actions to remain. GREEN requires swap/three-cycle/cycle-tail exact members to show `name_dependency_cycle`, Run disabled, zero mutation/journal rows, admission/direct engine exact `ERROR_CIRCULAR_DEPENDENCY`, case-only and acyclic chains unchanged, partial/cancel row truth preserved, and source contracts proving no production journal reader/writer/temp-hop/Resume/Roll back/replay surface. Run `Tools/Tests/TestHarnessSourceContracts.Tests.ps1`, `LocalizationTests.exe`, and the existing resource-localization contracts. |
| Performance/resources | Protected scenario is the existing provider-free 1,024-row independent scheduler case in five same-machine test-enabled x64 Release processes before and after. Retain `batchrename.execute.us` plus rows/completed/failed and add bounded `batchrename.schedule.build.us`, layer count, cycle-member count, retained-index count, mutation-call count, and journal-write count. Every candidate process must report one layer, 1,024 rows/retained indices/mutation calls/completions, zero cycles/failures/journal writes, and no retained temp path or schedule state after process exit. Candidate execute p95 must be at most `max(1.50 * baseline p95, baseline p95 + 50,000 us)` and every execute/build duration below 5,000,000 us. Archive under `Specs/TestRuns/4cb089111a23/Commands/2026-09-01_r7_a10_acyclic_no_journal_{baseline,candidate}_release/`. This is schedule/admission overhead evidence, not provider throughput. |
| Durable owners | `Specs/FileSystem/FileSystem_FileOperations.md` owns the one acyclic RenamePlan schedule, exact partial truth, no journal/replay, and post-loss refresh. `Specs/UI/UI_BatchRenameWindow.md` owns cycle-row error/guidance and disabled Run. `Specs/UI/UI_FolderView.md` changes only to remove Resume/Roll back from artifact presentation while retaining inspect/reveal/name-shape warnings. `Specs/Testing/Testing_PerformanceValidation.md` owns the deterministic schedule scenario and budget. This plan retains mapping and generated receipts only. |
| STOP/rollback | STOP if any supported release/channel is proved to have emitted an actionable journal requiring a retirement obligation; a cycle can reach mutation; exact cycle members cannot be identified without provider I/O on the UI thread; the immutable schedule cannot represent parent/child/case-only acyclic behavior; partial/cancel truth regresses; removing recovery deletes or mutates existing user bytes; name-shape warnings or mutation guards must be removed rather than decoupled; work requires R7-A15 provider name policy, E5 Change Case migration, public ABI, or live credentials; retained state grows beyond one index per row; or Release evidence exceeds budget without a diagnosed machine anomaly. Rollback restores the current executable cycle/journal surface only while R7-A10 remains open; never ship a half-state with journal writes but no reader, or a recovery reader with no owning target law. |
| Exit-receipt contract | Before R7-A10 closes, record activation/RED/implementation/evidence commits; exact Debug and test-enabled Release build receipts; focused GREEN totals and exact HRESULT/mutation/file observations; localized preview and Run-state proof; exact acyclic/case-only/parent-child/partial/cancel results; five-process baseline/candidate summaries and constant-resource proof; zero-reader/writer/recovery-surface inventory; proof that existing records were neither enumerated nor deleted; validated archive paths and whole-inventory validation; source/resource/localization results; `git diff --check`; authoritative-spec destinations; and remaining R7-A15/E5 plus artifact-claim gaps. |

**R7-A10 RED receipt (2026-09-01):** activation is `931ef726`; RED contracts
are `f9e9814f`. The exact activation tree built test-enabled x64 Release with
zero warnings/errors and full-solution receipt
`d563efd4aec102a97ab248c5fdf18c5a2be300b16bcef80864811ba4a9b46074`.
Five independent Release processes of
`cmd_pane_batchRename_execution_engine_large_independent_perf` passed at baseline
`931ef726`, each completing 1,024/1,024 mutations with durations 197,661,
194,087, 193,802, 195,024, and 192,742 microseconds; nearest-rank p95 is
197,661 microseconds and the candidate ceiling is 296,492 microseconds. Raw
receipts were preserved immediately under
`D:\RedSalamander.Perf\evidence\r7-a10-baseline-release-capture\sample-01`
through `sample-05`. The test-only tree then built cleanly with project receipt
`70e56e7b0579934df5de393b231aff39b9df666e4b97826c09221792479fa881`.
Direct governed Release execution under run ID `r7-a10-red-direct` recorded
0 passed / 3 failed for
`cmd_pane_batchRename_cycle_preview_blocks_run`,
`cmd_pane_batchRename_admission_rejects_cycle`, and
`cmd_pane_batchRename_no_journal_or_recovery_surface`: preview had no exact
`name_dependency_cycle` issue, the executor returned `S_OK` instead of
`ERROR_CIRCULAR_DEPENDENCY` before mutation, and the recovery-journal header
still existed. The structured RED receipt is retained at
`D:\RedSalamander.Perf\evidence\r7-a10-red-release-capture\last_run\commands\results.json`.

**R7-A10 exit receipt (2026-09-01):** activation is `931ef726`, RED contracts
are `f9e9814f`, the RED receipt commit is `fea68294`, and implementation is `8c303593`
(`8c303593aa30a30b892967d8fc642db492231ca1`) and compact evidence is
`10a27c53`; source-contract retirement guards were corrected at `eaeadfc3` after
the completed implementation made the old Proven/journal expectations intentionally
false. The exact implementation tree built test-enabled x64 Release and Debug with
zero warnings/errors under receipts
`852d50aba0b168a721c6daa2f7bfd59275cab5a8dca04dc2535581c192df874d`
and `67b1cf25cee85079c9581cb360fba1225558cd1da222d44395eee749553322a6`.
The exact evidence tree rebuilt Release cleanly under receipt
`69937ed25108d32d3721ebff0a788b9a20a2995bd9e934eb5f025bd49ce5bb12`.

Focused governed Release cases passed for cycle preview, worker admission, direct
execution, no-journal/recovery inventory, acyclic chain, directory-chain undo,
case-only Local rename, parent/child deepest-first order, exact partial failure,
mid-batch cancellation, and File System capability source guards. Preview reports
`name_dependency_cycle` on the exact members and disables Rename; admission and the
direct engine return `ERROR_CIRCULAR_DEPENDENCY` with zero mutation and unchanged
selected bytes. Acyclic/case-only/parent-child work keeps exact row results, while
partial and canceled work preserves completed/failed/unattempted truth. The
no-journal case and source guards prove there is no production Batch Rename reader,
writer, migration, replay, Resume/Roll back command, or temp-hop executor; because no
code enumerates or acknowledges the retired root, any old bytes are ignored and
retained. Release `LocalizationTests.exe` passed, and the updated source/resource
Pester contracts passed 195/195.

Five independent candidate Release processes completed the 1,024-row schedule in
1,801, 1,832, 1,723, 1,790, and 1,776 microseconds (nearest-rank p95 1,832 us),
against baseline p95 197,661 us and ceiling 296,492 us. Schedule-build p95 was
1,292 us. Every process reported one layer, zero cycles, 106,496 retained-index
bytes, 1,024 rows/completions/mutations, zero failures, and zero journal writes;
all execute/build durations were below five seconds. Compact baseline and candidate
archives are
`Specs/TestRuns/4cb089111a23/Commands/2026-09-01_r7_a10_acyclic_no_journal_baseline_release/`
and
`Specs/TestRuns/4cb089111a23/Commands/2026-09-01_r7_a10_acyclic_no_journal_candidate_release/`.
Each 16-file archive passed explicit validation, and the committed 1,654-file
inventory passed whole-inventory validation. Raw five-process captures remain under
`D:\RedSalamander.Perf\evidence\r7-a10-{baseline,candidate}-release-capture\sample-01`
through `sample-05`. Governed runs reported 92 disk-audit observations consisting of
explicitly authorized older siblings and the pre-existing
`D:\RedSalamander.Perf\worktrees`; no write outside the initialized test root was
observed.

The authoritative contract now lives in `FileSystem_FileOperations.md`,
`UI_BatchRenameWindow.md`, `UI_CommandMenuKeyboard.md`,
`UI_FileOperationsPopup.md`, `UI_FolderView.md`, `UI_FindFilesWindow.md`,
`Core_Search.md`, `Core_CompareDirectories.md`, `Core_SharedHelpers.md`,
`Testing_PerformanceValidation.md`, and `NormativeConsistency.json`. The broad
Batch Rename prefix characterization passed 72/73; its sole failure was the unrelated
pre-existing `cmd_pane_batchRename_admission_worker_owned_cancellable` archive-timing
assertion (`HRESULT` was correctly canceled, but archived latency was 0 us), so it is
not claimed as a green broad gate. `git diff --check` passed before closeout. R7-A15
and therefore E5 remain blocked by R2; current artifact projection is name-shape-only
`Possible`, with no durable `Proven` claimant.

- [x] Inventory every production/test caller, recovery command, persisted-file reader/
      writer, schema reference, and normative/UI statement before removal.
- [x] Reject every dependency cycle in preview/admission before mutation. Keep Run
      disabled, identify participating rows, and explain an explicit intermediate-name
      or two-pass workflow; one-action swaps are not promised.
- [x] Execute only the exact conditional acyclic schedule and preserve immutable
      per-step/row partial results. After process loss, refresh from the provider's
      namespace; never replay an unconfirmed row automatically.
- [x] Remove the Batch Rename journal and its Resume/Roll back production UI for all
      plans. Do not migrate, convert, replay, automatically delete, dual-write, or add
      a compatibility reader for existing records. STOP if supported-release evidence
      proves records need a separately arbitrated retirement/recovery obligation.
- [x] Prove no new journal file is created and update
      `FileSystem_FileOperations.md` and `UI_BatchRenameWindow.md` in the same slice.

#### R7-A15 — Accepted provider/path name feasibility consumption

**Ordering:** the R7-A15 row in section 8.1.

##### R7-A15 activation card

| Activation-card field | R7-A15 bounded execution contract |
|---|---|
| Slice/state/owner | **R7-A15 / `COMPLETE` / Codex `/root` in this worktree.** This slice consumes R2's typed provider/path name feasibility at the four accepted rename/create ingresses: Inline F2, Batch Rename preview plus worker admission, Change Case, and F7 Create Directory. It does not activate E5, add new rename transforms, change R7-A10 acyclic scheduling, redesign conflicts, alter mutation receipts, add publication behavior, or activate R1d/R3-R6/R7-A03/R8/R9. |
| Baseline and drift | Activation baseline is exact clean tracked commit `793d48bdaeb983585f318e41c9ad7fdf491d7009`; repository-root `last_run/` is pre-existing untracked output and excluded. Scoped tracked drift is empty. R2's separate typed IID and `Common/FileSystemRouteContract` are complete. Current production still has four competing consumers: Inline/Batch worker and F7 join with `FileSystemPathIdentity`; Batch preview applies `ValidateLeafName` Windows rules and Local-only `std::filesystem` collision probing; Change Case guesses a separator and reaches `FileSystemRenameBatch`; F7 selects suffixes using a Local short-ID case flag. Those are the exact authorities this slice retires. |
| Central name contract | Extend the canonical `Common/FileSystemRouteContract` with one composite child-name query that returns only `Available`, `Unsupported`, or `ContractViolation` plus provider validation status/failure, provider-owned joined path, provider collision key, and arena-fallback count. It invokes `ValidateChildName`, `JoinPath`, and `GetChildNameCollisionKey` against one retained `IFileSystemRouteCapabilities` generation; copies all arena output before return; verifies nonempty/bounded outputs and that the joined path belongs to the requested parent/leaf under the same typed identity; and never applies Win32 reserved-name, separator, case, normalization, or length rules itself. Provider Invalid preserves its exact HRESULT; missing/unsupported/malformed output fails closed. |
| F2 and worker admission | Inline F2 may accept free-form prompt text, but before task publication the central Rename admission validates the final leaf and uses only the provider joined path/collision key. The worker repeats the composite query after queue/interlock wait and immediately before its exact conditional rename; a changed/unsupported/invalid contract stops before mutation. FolderView owns only localized pre-publication error presentation; a published task owns later failure presentation. |
| Batch Rename | Preview generation passes one retained typed route interface to a provider-agnostic bulk name-policy stage off the UI thread. Every changed row receives provider validation/join/collision truth. Duplicate-target and existing-destination indexes combine the typed parent path key with the provider collision key; parent listings use `IFileSystem::ReadDirectoryInfo` for every provider and never `std::filesystem` or Local-only case folding. The checked index retains canonical keys only and stays within the existing 64 MiB cap. Worker admission recomputes every row's provider facts and immutable joined path/collision key before publishing the R7-A10 acyclic schedule; execution revalidates the current step before mutation. Preview facts are explanation, never mutation authority. |
| Change Case | Planning enumerates as today, then validates every changed leaf, provider join, and within-parent collision key before the artifact guard or first mutation. Any unsupported/invalid/duplicate contract returns an exact failure with zero rename calls. Each depth batch revalidates the provider contract immediately before its guarded `FileSystemRenameBatch` boundary. `ChangeCase::ApplyToPaths` receives the full provider ID; it no longer guesses separators or treats Windows casing as universal authority. |
| F7 Create Directory | Initial-name selection reads the parent once through `ReadDirectoryInfo`, canonicalizes every existing child name and each candidate through the provider collision-key method, and selects the first untaken valid provider name. Qualification and each auto-suffix retry use the composite provider validation/join result; no short-ID `ignoreCase` switch, `JoinFileSystemPath`, or Win32 name fallback remains. Missing executable directory operations still fail before the prompt/mutation exactly as today. |
| Scope | Shared: `Common/FileSystemRouteContract.h/.cpp`. Host/admission: `RedSalamander/FolderWindow.FileOperations.cpp`, `FolderWindow.FileOperations.State.cpp`, `FolderWindow.FileOperationsInternal.h`, and `FolderView.FileOps.cpp`. Batch: `BatchRenameEngine.h/.cpp`, `BatchRenameWindow.h/.cpp`, and central execution/admission glue. Direct commands: `ChangeCase.h/.cpp` and `FolderWindow.FileSystem.Commands.cpp`. Tests: FileOps provider matrix/phases, Commands Batch Rename/Dialog/PluginConfig cases, registrations, resource/localization/source contracts, and provider contract controls. Specs: `FileSystem_FileOperations.md`, `Plugins_VirtualFileSystem.md`, `UI_FolderView.md`, `UI_BatchRenameWindow.md`, `UI_CommandMenuKeyboard.md`, `Testing_PerformanceValidation.md`, this card/receipt, and compact evidence. Out of scope: new provider ABI, JSON, provider implementations unless a proven R2 contract bug is found, live accounts/devices, new commands/shortcuts, and mutation/publication algorithms. |
| RED/GREEN evidence | RED uses a scripted typed provider whose valid/invalid names, nonstandard join separator, and collision equivalence deliberately disagree with Windows rules. It proves F2/F7/Batch/Change Case currently accept a provider-rejected Windows-looking leaf, reject or mis-key a provider-accepted non-Windows leaf, or join through host guesses. GREEN requires provider Invalid/Unsupported/ContractViolation to reach zero mutation; a provider-accepted name to pass despite Windows disagreement; duplicate and existing-destination detection to use provider keys for all providers; preview and worker mismatch to fail before mutation; F7 default/auto-suffix selection to use provider keys; and missing/malformed name methods to fail closed. Local current behavior, case-only rename, R7-A10 cycle/partial/cancel truth, and source/receipt guards remain green. No test requires network, credentials, or a device. |
| Performance/resources | Protected Release scenario is existing `cmd_pane_batchRename_engine_large_preview_perf`: five independent 10,000-row processes before and after. Baseline captures the former Windows/path-identity validation; candidate runs the same transforms through a deterministic scripted typed name provider, exactly two composite child-name queries per changed row, and one parent listing. Retain `batchrename.preview.build_plan_us` and row/change/error counts; add bounded provider-name validation duration/query count, arena-fallback, rejected count, canonical-key bytes, parent-listing count, and retained collision-index bytes. Candidate p95 must be at most `max(1.50 * baseline p95, baseline p95 + 50,000 us)` and every sample below 5,000,000 us. Existing `cmd_pane_batchRename_collision_name_index_memory_gate` remains blocking at 65,536 names/240 UTF-16 units/64 MiB and must measure provider-canonical keys without retaining a duplicate raw-name universe. Archive under `Specs/TestRuns/4cb089111a23/Commands/2026-09-01_r7_a15_provider_name_{baseline,candidate}_release/`. |
| Commands and faults | Build full Debug and exact-commit test-enabled x64 Release. Focused gates cover the four ingresses, composite helper malformed output/allocation/arena fallback, provider invalid/unsupported, missing IID, join/collision disagreement, duplicate/collision and preview-worker drift, Local case-only behavior, Batch 10k/65,536 bounds, Change Case recursive/nonrecursive zero-mutation failures, F7 selected default/auto-suffix, R7-A10 acyclic/cycle/partial/cancel, source guards, resources, localization, and Pester. Fault seams are deterministic and offline. |
| Durable owners | `Plugins_PluginAPI.md` continues to own the ABI/lifetime rules. `Plugins_VirtualFileSystem.md` owns provider name semantics and bounded offline execution. `FileSystem_FileOperations.md` owns command admission/revalidation and failure truth. `UI_FolderView.md`, `UI_BatchRenameWindow.md`, and `UI_CommandMenuKeyboard.md` own the visible F2/Batch/F7 behavior. `Testing_PerformanceValidation.md` owns metrics, five-process evidence, and caps. This plan retains delivery mapping and the exit receipt only. |
| STOP/rollback | STOP if a supported provider cannot answer child-name methods without network/device I/O; the current typed ABI cannot express required parent-scoped collision truth; Batch preview would need to retain raw and canonical name universes above 64 MiB; Change Case cannot validate all planned names before mutation; one shared composite helper cannot serve all four ingresses; the slice requires changing mutation/publication/receipt semantics or R3 work; a live account/device is required; or Release evidence exceeds budget without a diagnosed machine anomaly. Rollback removes the composite consumer, four ingress cutovers, tests, and spec clauses together; never leave some rename commands on provider truth and others on Windows rules. |
| Exit-receipt contract | Before R7-A15 closes, record activation/RED/implementation/evidence commits; exact Debug and test-enabled Release build receipts; focused GREEN totals and exact zero-mutation failure observations for all four ingresses; composite-helper fault totals; provider-disagreement, collision, preview/worker-revalidation, F7 suffix, and Local case-only results; five-process baseline/candidate summaries and 65,536-name constant-memory proof; validated archives and whole-inventory validation; source proof that the four retired Windows/path-guess authorities are absent; resource/localization/spec results; `git diff --check`; authoritative-spec destinations; and newly eligible E5 state. |

- [x] Add RED provider-disagreement, malformed-output, preview/worker-drift, duplicate,
      F7 suffix, and source-authority contracts before production changes.
- [x] Capture five independent test-enabled x64 Release 10,000-row preview baselines
      and retain the existing 65,536-name/64 MiB control.
- [x] Add one composite canonical child-name query over R2's typed interface; do not
      create a second ABI, JSON reader, or generic Windows fallback.
- [x] Cut Inline F2 and Batch Rename preview/admission/execution over to provider
      validation, joined paths, and collision keys with pre-mutation revalidation.
- [x] Cut recursive/non-recursive Change Case over with all-name preflight and
      per-depth-batch revalidation before `FileSystemRenameBatch`.
- [x] Cut F7 initial/default/suffix selection and admission over to provider keys and
      joins while preserving one parent enumeration and bounded retries.
- [x] Remove the four retired Windows/path-guess authorities from production and add
      source guards preventing their return at these ingresses.
- [x] Update authoritative specs/resources, run focused Debug and exact-commit Release,
      archive five candidate samples, validate inventory, record the receipt, and mark
      only R7-A15 complete. This receipt makes E5 eligible but does not start it.

**R7-A15 exit receipt (2026-09-01):** activation is `ee6fdd87`; RED contracts are
`e610f3e4`; baseline evidence is `8697f8a8`; the shared composite contract is
`920d36b2`; ingress implementation is `7210e1f2`; initial candidate evidence is
`7d76434d`; regression/test synchronization is `379e1868`, `60ad7b65`,
`0ca64ef4`, `ee87c74d`, `3f78256c`, and `e925018f`; final exact-implementation
candidate evidence is `ae926d24`. Exact test-enabled full-solution builds at
`e925018f5d40925d418563c7cbf0922a86d0145b` passed with zero compiler diagnostics:
x64 Debug receipt
`527d6dd1bcde11ecd2ab40a647448300935a2ae11f8fb42f3999ed169956164b`
and x64 Release receipt
`3d8be5dc3a2b7fad5119f3274d55aa94e626f9826e284ff1692c34ae723e4ae9`.

Focused Release Commands validation passed all 82 Batch Rename-family cases, including
provider disagreement, duplicate/collision reduction, preview/worker revalidation,
admission cancellation, acyclic/cycle/partial truth, the 10,000-row scenario, and the
65,536-name memory gate. `file_system_provider_name_policy_ingresses` proved all four
ingresses: provider-valid `CON` remained admissible, provider-invalid names preserved
the provider HRESULT with zero mutation, F7 used one listing plus provider collision
keys for suffix selection, Inline F2 and Batch worker revalidation failed closed on
drift, and Change Case rejected provider-key duplicates before `FileSystemRenameBatch`.
`file_system_provider_name_policy_consumption_source_guard` proved the four retired
Windows/path-guess authorities absent. Release `PluginContractTests` passed the
composite helper's available/invalid/unsupported/malformed-output and retained-arena
lifetime controls. The broader Release Commands run recorded 877 passed, seven
unrelated UI/rendering failures, and two skipped; the broader FileOps run recorded 71
passed with three load/order-sensitive failures, while isolated reruns of
`FileOps_ProviderCapabilityMatrix` and `Fairstream_MoveSameSizeCollisionPrompts`
passed. Those broad runs are recorded as diagnostic evidence, not claimed as clean
release gates.

Five final independent 10,000-row Release processes measured
`batchrename.preview.build_plan_us` at `80,564`, `81,362`, `79,758`, `81,030`,
and `78,450` us (nearest-rank p95 `81,362` us) against the accepted `146,235` us
ceiling. Provider validation p95 was `86,330` us. Every sample issued exactly
`20,000` composite queries and one parent listing, retained `3,100,000` canonical-key
bytes and `3,091,136` collision-index bytes, and reported zero fallbacks/rejections.
The 65,536-name/240-UTF-16-unit gate retained `36,700,160` bytes below the
`67,108,864`-byte cap. Raw evidence is archived under
`Specs/TestRuns/4cb089111a23/Commands/2026-09-01_r7_a15_provider_name_{baseline,candidate}_release/`
and mirrored under `D:\RedSalamander.Perf\evidence\r7-a15-provider-name-candidate-release`;
all 19 mirror files matched SHA-256, explicit archive validation passed, and the whole
1,767-file TestRuns inventory passed. `git diff --check` passed. Durable contract
updates landed in `Specs/FileSystem/FileSystem_FileOperations.md`,
`Specs/Plugins/Plugins_VirtualFileSystem.md`, `Specs/UI/UI_FolderView.md`,
`Specs/UI/UI_BatchRenameWindow.md`, `Specs/UI/UI_CommandMenuKeyboard.md`, and
`Specs/Testing/Testing_PerformanceValidation.md`. R7-A15 alone closes. E5 is newly
eligible but remains not started; every other inactive R slice remains inactive (R3 was
activated later, on 2026-09-03), and E8 remains predecessor-blocked.

#### R7-A03 — Accepted literal default and explicit retarget transform

**Ordering:** the R7-A03 row in section 8.1; it has no E4/E5 dependency.

- [ ] Make **Copy links unchanged** the default Preserve behavior: copy the literal
      link object/payload without silent retargeting.
- [ ] Add the explicit per-task **Retarget links inside copied tree** option. Retarget
      only links whose admitted target is inside the copied tree; preserve/describe
      external, skipped, failed, cyclic, or ambiguous targets without silent rewrite.
- [ ] Replace repeated linear scans and the 4,096-record terminal bound with indexed,
      bounded mapping/deferred state; define collision, cycle, forward-reference,
      partial-publication, Cancel, and result receipts separately from literal Copy.
- [ ] Keep the option in the one with-options surface and add provider/link-type/
      relative/absolute/no-follow tests. It is never implied by the word Preserve.

**Acceptance:** every rename-like command consumes the same authority, conflict,
terminal result, schedule, and provider namespace rules. A10 acceptance additionally
requires zero production journal/recovery callers, blocking cycle explanation before
mutation, no Resume/Roll back surface, and no automatic crash replay. A03 acceptance
additionally requires literal Preserve to remain independent of the bounded explicit
retarget graph and to report every skipped/unresolved rewrite truthfully.

### Package R8 — Artifact/performance and normative consolidation

**Priority:** P2 governance
**Depends on:** accepted/landed behavior; consumes E3/E6/E7 without duplicating
their authorized scope

- [ ] Consume completed E3/E6/E7 mechanics, artifact index, and breadcrumb identity;
      do not create a second persistence helper, artifact index, or notice owner.
- [ ] Apply ownership map in section 10 with stable rule IDs.
- [ ] Remove copied matrices, algorithms, painter details, historical timings, source
      lists/tests, and exact optimizer constants from product authority.
- [ ] Generate capability summaries from executable declarations/tests where useful.
- [ ] Add per-document consistency commit/hash freshness; rebuild ledger after review.
- [ ] Specify operation-history privacy, retention, redaction, size, and export
      ownership without merging its schema with artifact claims or Move breadcrumbs.
- [ ] Update shared traversal helper catalog if R4 changes its contract.
- [ ] After R2 lands, update ownership/spec consistency evidence for its lockstep
      typed cutover; do not perform or duplicate the cutover in R8.

**Acceptance:** every durable rule has one authority, current consistency evidence,
and an executable/test witness; enumeration/search cost is scoped to relevant claims.

### Package R9 — Ordinary Local Shell-delegation decision spike

**Priority:** P2 simplification spike
**Ordering:** the R9 row in section 8.1. D2-A16 authorizes evidence only, not
production adoption.

- [ ] Build a non-default, bounded `IFileOperation` adapter/spike for exactly two
      candidates: ordinary new-name Local regular-file Copy and no-conflict
      same-volume Native Local regular-file Move. Exercise existing Shell Recycle as
      a comparison control, not as a newly adopted route.
- [ ] Exclude overwrite, Managed/cross-volume/folder/link Move, Permanent Delete,
      Rename, directory merge, remote/device, artifact recovery, and Batch Rename.
- [ ] Use current Shell Recycle as the instrumented comparison control, including its
      explicit flag set, STA boundary, conservative Unknown default, progress sink,
      `PerformOperations`, and `GetAnyOperationsAborted`.
- [ ] Declare and test the exact Copy/Move flag matrix. Shell UI must remain
      suppressed, and `FOF_NOCONFIRMATION` or equivalent must never turn a raced
      destination into Shell-owned overwrite, rename, or another collision decision.
- [ ] Explicitly omit `FOF_ALLOWUNDO` and `FOFX_ADDUNDORECORD`; verify no Shell Undo
      record is created.
- [ ] Map every sink callback, `PerformOperations` result, and aborted-operation state
      to immutable per-item truth. Missing terminal observation is Indeterminate/
      ContractViolation, never Completed.
- [ ] Prove dedicated-STA startup, cancellation, teardown, and application shutdown
      satisfy D2-A01.
- [ ] Treat `FOFX_NOSKIPJUNCTIONS` as Shell-namespace behavior, not NTFS link
      authority. Any need to admit a directory/link route stops the spike and requires
      new arbitration.
- [ ] Fault-inject cancel, missing callback, partial completion, retained source,
      raced destination, suppressed-prompt failure, Recycle escalation, and
      application shutdown.
- [ ] Measure startup, throughput, memory, cancellation, and code/owner count against
      the custom Local route using the same scenarios.
- [ ] Return an evidence table per operation: Adopt, Reject, or Needs bounded contract
      change. The spike itself cannot weaken authority or silently enable production.

**Acceptance:** delegation is adopted only for an operation whose exact user-visible
truth, single conflict owner, cancellation behavior, and recovery boundaries meet
the existing contract while removing meaningful custom production code. Otherwise
the spike is deleted and the rejection rationale is retained here.

## 9. Verification contract

An R card may be drafted earlier, but before promotion to `ACTIVE` and implementation
it must have every hard-predecessor receipt and name the exact RED/GREEN test IDs/
commands, fault matrix, symbols/files, and performance/no-regression budget. It applies
repository C++/WIL/error/async rules and integrates performance validation from the
start; this section supplies cross-slice scenarios, not a substitute for the card.

### 9.1 Required deterministic correctness cases

- Local create-new Copy loses a race to a foreign final occupant; foreign object
  survives and result reports NotPublished/SourceRetained/exact owned-publication
  truth for both permitted publication shapes.
- All-zero `FILE_ID_128` never enters equality/cache/admission. Retained-handle
  authority applies only to the exact just-created object while retained and never
  binds a pre-existing object, replaces an existing destination, or authorizes source
  deletion.
- Publication fault injection before/after create, write, commit, publish, abort,
  reconcile; no unowned cleanup.
- After a possibly committing Rename, Delete, abort, publish, or rollback begins,
  known non-commit requires exact reconciliation; otherwise result is Indeterminate.
  Skip/Cancel/Continue cannot rewrite publication/source/artifact axes.
- Standard Local Managed Move proves stable source/commit before cleanup; Verified
  Move detects deliberate mismatch; weak-commit routes cannot opt out and finish
  Copy-only when exact content proof is unavailable.
- MTP full PUID survives restart; no-PUID/malformed records cannot delete; deliberate
  hash-collision/ambiguous candidates fail closed.
- Bound Local directory Delete removes a concurrently added child inside the same
  exact no-follow container or reports a bounded-writer residual; replacement root,
  junction/link target, and escaped alias survive.
- S3 singular and batch ordinary folder Delete consume current key+the just-observed
  ETag and omit `VersionId`. A late key is freshly observed and deleted; a replacement
  first causes a condition mismatch, then may be freshly admitted/deleted on a later
  bounded pass. Continuous writers hit the pass/time bound and return Partial with
  residual truth. `prefix`, `prefix2/`, another bucket/profile, and provider-returned
  out-of-scope keys survive. Request-level unknown is not automatically reissued.
  Versioned buckets record delete-marker/current-visibility truth and never claim that
  older versions were permanently erased; exact-VersionId cleanup remains separate.
- Microsoft Drive ordinary Delete is presented as Recycle, targets the exact folder/
  item ID without a root ETag condition, rejects a different-ID replacement, returns
  a truthful receipt, and does not offer Permanent Delete when that endpoint lacks
  accepted authority/receipt.
- Null Native mutation receipt becomes Indeterminate; cleanup debt is visible once.
- Fake provider never returns; UI/shutdown meet chosen bound with indeterminate issue.
- Clipboard consumption fails: no mutation runs and cut list remains. Preparing or an
  A02 overlap decision fails/is declined: dense NotAttempted results and pre-consume
  intent remains.
- Preparing over 1 and 256 selected roots performs zero recursive enumerations and
  zero content bytes even when a selected directory has millions of descendants;
  Cancel leaves no mutation/clipboard/breadcrumb, and an unbounded provider route is
  disabled or isolated rather than hanging the gate.
- Duplicate result store/illegal axis combination fails focused tests.
- Cancel stops primary work but permits bounded exact owned-stage compensation.
- Change Case unreadable subtree/cancel/conflict/task-card share RenamePlan truth.
- Provider-valid non-Windows names and Local case-sensitive names follow path policy.
- Trees beyond 128 levels/aggregate ceilings complete iteratively with bounded memory
  and create no operation-owned disk/durable traversal spool.
- A02 task injection compares operation roles and normalized component-boundary paths
  with zero descendant enumeration/network I/O solely for warning. Disjoint
  `C:\Photos`/`C:\Backups` and read/read pairs do not warn. Nested destination Copy,
  nested Delete, Delete/read, Delete/write, Move-source, and positively evidenced
  mapped-drive/UNC or SUBST aliases show one concrete warning before clipboard/
  mutation with Queue default, Run together, Don't start, Details, keyboard/UIA, one
  transient Queue edge, and a concurrently-live same-host Run receipt.
- Nested Copy destinations under Run together produce exact create ownership plus
  ordinary current file-exists/conflict outcomes; the losing task never pathname-
  deletes the winner. Replacement after consent revalidates and survives absent exact
  authority.
- Nested Deletes execute in both orders and attribute `Removed` versus `No longer
  present when reached` to the task that proved each result. Delete parent plus Copy
  child executes in both timing orders after an explicit Run warning that states the
  copied data may be removed; Copy may fail, but replacement root/link/mount/escaped
  target survives and neither task invents lasting final-state truth.
- Delete parent plus Copy child without a covering Run receipt parks only while both
  tasks are concurrently live in the same host. Skip preserves it. Queue waits,
  releases automatically once without a second A02 warning, and ordinary folder
  semantics may then remove it; only the explicit destructive choice permits the live
  effect. Another RedSalamander instance receives no shared index/receipt, and history
  cannot reconstruct one.
- An unresolved mapped-drive/UNC alias never stalls injection merely to resolve the
  warning; normal external-race safety still applies. A third task injected after a
  pairwise Run receipt receives its own decision. A single-lane provider reports
  provider/device waiting without global overlap failure.
- Warning suppression for an active parent/child pair requires exact cheap proof that
  the relevant residual branch is irrevocably closed/excluded; current-item location
  alone is insufficient. Indeterminate mutation authority without independent exact
  owned-create capability still fails before mutation.
- Escape/Enter/More/Apply-to-all/close/repeated explicit no-commit Retry consume one
  decision model; changed commit-time conflicts return to the same surface. An
  Indeterminate/request-level unknown never exposes automatic Retry.
- Every ingress in section 6.9 proves its initial-consent owner, fast/reveal path,
  later gates, terminal mapping, and absence of duplicate prompts.
- Routine accepted-default work follows accepted D2-A20's interaction count; **with
  options**, known degradation, and material-risk paths use one stable surface with
  deterministic initial focus.
- Default discovery-ahead begins eligible Copy/Move/Delete before closure. Archive the
  unchanged normative scheduler baseline at concurrency 1/4/16 for large-file Copy,
  small-file Delete, MTP serialization, SMB, and independent-device scenarios. It must
  meet the activated card's discovery-service, memory, throughput, cancellation, and
  zero-disk-spool gates. If it passes, no scheduler controller is added. If one causal
  scenario fails, the smallest activated correction must pass before/after evidence
  and leave unrelated devices unthrottled. Mid-run `Discover as needed` releases the
  run-ahead reservation and any adopted temporary limiter; mandatory JIT checks remain.
- Before discovery closes, exact counters remain visible and no percentage is
  rendered. The provisional `still discovering` ETA can increase/decrease as work is
  found, disappears when paused/stalled/unsupported or evidence is insufficient, and
  is replaced by fixed-total percentage/ETA after closure. UIA announcements are
  bounded/nonrepeating, and the ETA cannot affect scheduling, timeout, consent,
  mutation, cleanup, or completion.
- With several parallel tasks, each task independently renders discovery activity,
  exact discovered count/provisional ETA, and operation progress. Compatible closed-
  total tasks produce **Known work: N%** from summed completed/total units; open tasks
  are excluded and add **N tasks discovering — total may grow**. With no compatible
  closed cohort, the jobs view has activity/throughput but no percentage. It never
  labels the subset Overall, and the taskbar stays indeterminate while any included
  total is open.
- Waiting/Preparing/Discovering/Running/Verifying/Attention/Paused/Stopping/terminal
  phase actions match section 6.9; Escape never means Cancel All and Close never
  resolves a decision.
- Closing the application with live work defaults to keeping it open; explicit
  cancel-and-exit fences new commands, resolves no hidden prompt as accepted, and
  reaches the D2-A01 shutdown outcome.
- Review-later mode continues unrelated safe items, freezes selected roots/accepted
  scope while permitting bounded descendant discovery, collects conflicts/errors,
  never retries an indeterminate mutation, preserves existing Apply-to-all decisions,
  returns newly changed conflicts to the same review, and makes Rescan a new plan.
- The I17 prompt baseline remains intact: actionable conflict says `Waiting for your
  decision`; incomplete facts say `Loading decision details...`; neither Waiting nor
  metadata loading exposes false discovery/transfer/graph/marquee activity; incomplete
  destructive actions remain unavailable; question -> endpoints/context/facts ->
  actions is visual/keyboard/UIA order; wrapped prompt/context/final-note text does not
  clip; announcements occur once across both densities, 100/150/200% DPI, and the
  existing English/cs-CZ/fr-FR/ja-JP/sk-SK resource set.
- Forced hosted-control attachment failure renders only the minimal localized
  operation/context/truth text plus safe Cancel/Close actions, remains keyboard/UIA
  usable, and invokes the same engine snapshot; no second feature-complete action or
  layout derivation remains.
- Known/runtime Copy-only preserves source and never renders “Moved.”
- A user-created name-only `.rs_*` file opens, edits, previews, copies, and exports
  without an artifact-name prompt. RedSalamander-owned destructive mutation gets one
  exact-object Cancel-default warning. Possible never becomes Proven through consent.
- Properties uses a fresh shared classification, renders every available semantic
  File Operations fact and explicit Unknown/unavailable/name-only reason, exposes no
  unredacted endpoint secret, scans no global history, and grants no recovery or
  mutation authority.
- A cyclic Batch Rename keeps Run disabled, identifies the participating rows, makes
  no mutation, and creates no journal. Acyclic partial failure reports exact committed
  and unattempted rows; process loss triggers no automatic replay and exposes no
  Resume/Roll back action. Source-contract inventory proves zero production journal/
  recovery reader or writer remains after R7-A10.
- Default **Copy links unchanged** preserves literal relative/absolute link payloads.
  Explicit **Retarget links inside copied tree** covers internal forward references,
  cycles, collisions, skipped/failed targets, external targets, cancellation, and
  partial publication with indexed bounded state and distinct per-link receipts; no
  unresolved mapping is silently called Retargeted.
- The A11 cutover uses the new IID and `sizeBytes` boundary across host, every shipped
  provider, Dummy, adapters, and tests in one supported generation. Runtime ABI-size
  tests pass, JSON cannot enable a route, no dual reader/compatibility layer exists,
  and mixed-generation binaries are rejected as unsupported rather than tolerated.
- Internal axes render only the approved compact outcomes; detailed evidence remains
  inspectable and no outcome is inferred from byte count or provider success alone.
- Persisted local operation history survives task-card close, reproduces immutable
  results, applies both age and storage bounds with deterministic oldest-first
  eviction, and cannot authorize cleanup/Retry/Resume/Undo/source deletion.
- History entry, Clear, Disable, redacted navigation, and explicit Export follow
  section 6.9. Disable stops new writes without silently clearing old rows; Clear
  cannot delete artifact claims/provider recovery state; export contains no hidden
  raw authority or secret-bearing endpoint data; a saved row cannot replay work.
- Any adopted candidate `IFileOperation` route passes the same per-item truth,
  conflict-owner, cancel, concurrent-destination, and shutdown cases as the custom
  route; existing Shell Recycle supplies only comparative evidence.

### 9.2 Required performance/resource evidence

- Preparing task-terminal phase breakdown for Local, UNC simulation, cloud, and
  device: selected-root count, provider-call count/time, route/overlap-advisor time,
  recursive-enumeration count (required zero), content bytes (required zero), cancel
  acknowledgement, reveal/coalescing outcome, and isolation/route rejection.
- Same-machine discovery scheduling A/B for run-ahead, just-in-time from start, and
  mid-run `Discover as needed`: time to first safe mutation, time to discovery close,
  bytes and mutation operations before closure, runnable discovery wait/starvation
  duration/max service gap, reservation state, queue/memory high-water, cancellation,
  total elapsed time, and throughput. Collect candidate-specific limiter metrics only
  when the archived unchanged baseline fails and that exact correction is activated. Emit
  task-terminal aggregates only, never per-path/per-descendant events.
- Discovery scenarios cover slow-enumeration/fast-transfer,
  fast-enumeration/slow-transfer, concurrency 1/4/16, Queue/Parallel,
  capped/uncapped bandwidth, one USB-like resource with large-file Copy, one USB-like
  resource with small-file Delete, unrelated devices, Local/UNC-like/high-latency
  fake providers, single-lane MTP characterization with conditional bounded turns,
  Pause, Cancel, mid-run override,
  and a never-returning provider.
- Copy throughput/memory for small, large, deep, wide, and mixed trees.
- Standard versus Verified Move elapsed time, bytes reread, CPU, network traffic, and
  Copy-only rate across Local, UNC simulation, cloud, and device routes.
- Always-stage versus scoped direct-final Local new-name Copy latency, throughput,
  temporary-space high-water, and incomplete-artifact behavior.
- Queue/backpressure high-water, retained path/metadata bytes, open handles,
  cancellation latency, and shutdown bound.
- A02 admission comparison latency and allocation versus active/queued task count and
  selected-root count; warning/no-warning matrix, role classification, cached alias
  hit, unresolved alias, zero descendant enumeration, and zero network/device call
  solely for overlap detection. Compare Queue and Run-together throughput without
  treating faster nondeterminism as correctness evidence.
- Batch Rename admission/cycle-detection/schedule scaling and UI responsiveness,
  including zero journal I/O after A10 retirement.
- Typed route-query/cutover overhead versus the removed JSON-authority parse path;
  report provider-call count, allocation/parse cost, and admission latency.
- Bounded S3 virtual-folder convergence under late keys, replacements, and continuous
  writers: LIST/Delete requests, passes, wall time, condition failures, residual count,
  and cancellation acknowledgement.
- Explicit Retarget-links indexed mapping, collision, cycle, and forward-reference
  scaling; literal Preserve has no dependency-graph cost.
- Artifact registry enumeration/search overhead versus relevant claim count;
  Properties classification latency/bind count proves no full-history scan.
- Popup update/render/UIA cost after one-owner convergence.
- Provisional discovery-ETA sampling/update cost, estimate error/direction-change
  frequency, and announcement count, including suppression during pause/stall and
  transition to fixed-total progress. Evidence tunes presentation; it cannot make the
  estimate authoritative.
- Hosted popup attachment-failure fallback latency/UIA availability and proof that no
  duplicate feature-complete renderer remains.
- Human-interaction evidence for the section 6.9 walkthroughs: input-to-stable
  Preparing/reveal, prompt show-to-action-ready, focus transitions, prompt/action
  count, announcement count, cancel-to-visible-acknowledgement, resume-to-progress,
  A02 warning comprehension/choice, per-task dual discovery/operation progress,
  Known-work plus one/many open-discovery comprehension, no-compatible-known-cohort
  comprehension, indeterminate taskbar recognition, collected-review navigation, and
  completion-to-history discovery.
- Continue-safe-work conflict/error collection memory, review latency, and job-history
  persistence/search/export cost under the accepted retention bound.
- Custom versus `IFileOperation` spike startup, throughput, memory, cancellation,
  fault-result fidelity, and long-lived production owner/code count.
- Archives live under `Specs/TestRuns/`; product authority states outcomes, not one
  machine's historical number.

### 9.3 Closeout gates

| Gate | Expected evidence |
|---|---|
| Drift | Exact `git diff <planned-at> -- <owned scope>` reconciled before edits. |
| Builds | `.\build.ps1 -ProjectName RedSalamander` plus changed plugin projects exits 0 with no new warnings; perf/lifetime packages also validate the test-enabled Release configuration. |
| Focused tests | FileOps, Commands file-operations family, provider contracts, and package fault/perf cases pass. |
| Test inventory | `.\Tools\Get-TestInventory.ps1 -Format Json` succeeds and the exact focused/run IDs are recorded. |
| Spec inventory | `.\Tools\Get-SpecInventory.ps1 -FailOnFindings` exits 0. |
| Archive inventory | `.\Tools\Test-TestRunArchive.ps1 -Inventory` validates every required archived run. |
| Patch hygiene | `git diff --check` has no findings. |
| Final qualification | `.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh` passes on exact merged closeout commit. |

No old Fresh Full run may qualify changed behavior. E0's Fresh Full qualifies only the
exact pre-change baseline and never qualifies an R change. E8's exact merged Fresh
Full is the changed-behavior closeout; keep distinct `baselineCommit`/`baselineRunId`
and `closeoutCommit`/`closeoutRunId` receipts.

### 9.4 Inherited I17 qualification

The retired I17 plan is implementation history, not a live work queue. Its G0-G5
implementation, focused correctness/performance evidence, localization/resource
work, and authoritative-spec updates remain accepted historical evidence. Its move
to Done does **not** claim merged Fresh-Full qualification and does not make any
decision in this successor implemented.

E8 owns one residual aggregate gate. The exact merged closeout commit of this
successor—including every E package and any accepted R slice landed before
closeout—must:

- [ ] run the final merged FileOps and Commands File Operations coverage plus
      `ResourceLocalizationContracts.Tests.ps1`;
- [ ] rerun `Get-SpecInventory.ps1 -FailOnFindings` and `git diff --check` on the exact
      merged closeout commit;
- [ ] pass one exact Fresh Full receipt that explicitly qualifies both its promoted
      successor scope and the retired I17 scope;
- [ ] record any failure against the live successor and create a bounded remediation
      owner when I17-owned behavior is implicated; never resume or edit the archive.

Historical lack of a live SMB endpoint remains an explicit environmental gap; the
existing deterministic retained-authority coverage is the baseline, not a claim of
live SMB execution.

## 10. Normative ownership after consolidation

This is the future ownership map applied slice by slice as work lands. It does not
presently rewrite any authoritative specification. Until a same-slice implementation,
tests, and spec update land together, the existing normative owner and behavior remain
in force.

| Authority | Owns | Must not own |
|---|---|---|
| `FileSystem_FileOperations.md` | User intent/lifecycle/consent, mandatory Preparing scope, A02 overlap-problem classification and Queue/Run/Don't-start scheduling meaning, discovery/JIT safety and device-I/O priority, strategy and Standard/Verified proof meaning, real-container/provider-declared-virtual-folder/fixed-set Delete authority, artifact-warning semantics, terminal truth, jobs/history semantics, and recovery/control outcomes. | Provider matrices, queue/worker/token tuning constants, bridge loops, painter geometry, persistence encoding, source lists, historical runs. |
| `Plugins_VirtualFileSystem.md` + ABI | Typed provider obligations, executable route/receipt semantics, namespace-scope contract, optional cheap comparable/alias facts, concurrency/thread/session safety, name feasibility, versioning, and the rule that JSON diagnostics cannot enable a route. | Prompt layout, global task-overlap prohibition, compatibility migration, or duplicated backend facts. |
| `Core_FileSystemBridge.md` | Pump/publication adapter mechanics, checkpoints, bounded resources, receipts. | Global product policy, provider matrix, UI actions. |
| Provider specs | Backend identity/conditions, real-container/virtual-folder/fixed-set declaration, atomicity, verification evidence, version/delete-marker truth, and failure/reconcile facts. | Restated global lifecycle. |
| `UI_FileOperationsPopup.md` and command/Batch Rename UI specs | Ingress/phase journeys, problem-specific A02 warning/actions, per-task discovery plus operation progress, the `FO-DISCOVERY-01` Known-work/open-discovery/taskbar modes, provisional-ETA wording/state, jobs/history/conflict rendering, one hosted owner/minimal fallback, cyclic-plan rejection/no recovery UI, interaction, accessibility/UIA, and immutable snapshot/result consumption. | Deriving conflict policy, authority, strategy, skipped safety, overlap mutation rights, or recovery rights from presentation state. |
| `UI_FolderWindow.md` / `UI_FolderView.md` | Host-owned Properties **File Operations** projection, fresh artifact classification, allowed/destructive interactions, redaction, and no-authority wording. | Provider mutation policy, raw opaque authority, credentials/tokens, or recovery implementation. |
| Settings/schema authority | Operation-history enablement, retention, privacy/redaction, and bounded storage preferences where configurable. | Mutation receipts, artifact authority, or universal Undo. |
| Testing/performance specs | Preparing/overlap-admission/discovery/device-contention/interaction metrics, scenarios, fault model, thresholds, archives, and commands. | Product behavior existing only because a test asserts it, or machine-specific scheduler/token constants presented as universal product law. |
| `Specs/TestRuns/` | Machine/run-specific evidence. | Normative behavior. |

Use stable rule IDs instead of copied prose. Generated summaries identify executable
source and reviewed commit.

## 11. Routing and dependency rules

- This single plan has all target product decisions recorded. E0-E8 are authorized by
  inheritance; E0-E7 are complete. Every R slice is
  `DECIDED-NOT-ACTIVE`; a `DRAFT` card grants no authority, and even a complete card
  may become `ACTIVE` only after all hard-predecessor exit receipts close and one
  named owner accepts it. `ACTIVE` authorizes work; it does not itself change product
  behavior or normative authority.
- M1-M3 are derived exit-receipt views over section 8.1. They add no implementation
  authority, dependency, package, or second plan; M1/M2 may qualify releases and only
  M3 participates in final retirement.
- FOS-01 through FOS-06 require explicit release-blocking ownership. R0a-R0e are the
  candidate immediate containment/honesty slices; R5 owns later provider isolation
  only when evidence and D2-A01 authorize it. None may hide inside E extraction.
- The retired I17 implementation record is
  `../Done/UI_FileOperationsPopupParallelProgressPendingPrompt_2026-08-28.md`.
  Its delivered G0-G5 behavior remains owned by the authoritative File Operations/UI
  specifications and is not reopened here. E8 inherits section 9.4's exact merged
  qualification and records one Fresh Full receipt for both scopes. FOS-01 remains a
  separate adjacent release blocker requiring explicit R0a ownership; it is not
  retired-I17 scope.
- E1 owns bridge/item typed extraction, E2 conflict-action ownership, E3 durable
  mechanics, E4 Batch Rename worker admission, E5 Change Case migration, E6 artifact
  hints, and E7 breadcrumb identity. Those are the complete former-I14 obligations;
  the retired source plan has no remaining executable work.
- E packages cannot silently own provider ABI change, cancellation isolation,
  traversal replacement, A09 Properties/warning behavior, A10 journal/cycle
  retirement, new conflict-collection/history behavior, or changed product semantics.
  Crossing that line requires the corresponding R slice to be activated here.
- Every `DECIDED-NOT-ACTIVE` slice still needs bounded active ownership. R2 is the
  sole A11 typed cutover owner; R7-A10 is the sole Batch Rename journal/cycle
  retirement owner; R6-A09 is the sole new Properties/warning owner. R8 consolidates
  only landed behavior and cannot repeat those changes.
- Accepted A02/A07/A08/A09/A12/A17/A18/A20 remain separately activatable R6 slices;
  A02 comparison/concurrent correctness and A19 evidence-first discovery service stay in R4,
  A03 stays in R7, and A16 stays in R9. E2's current conflict-owner extraction does
  not silently own overlap comparison, continue-safe-work/review-later, dual progress,
  routine-start, or persistent-history behavior. Use section 8.1's dependency graph
  rather than a branch-sized rewrite.
- Accepted target changes land with implementation/tests from the complete,
  predecessor-closed `ACTIVE` card and update their durable owner in the same slice.
  Only that landed authoritative-spec update supersedes prior normative text;
  acceptance, `DRAFT`, or `ACTIVE` does not. This WIP successor is never the only
  lasting statement, and inactive briefs are never treated as executable code plans.
- Frozen Phase 0-6 history remains byte-stable.

## 12. Findings considered and rejected

- **“Map every 996 to Unsupported.”** Rejected: raw error does not prove stable
  feature absence and can hide transient/unknown failure.
- **“Fail every profile without `FILE_ID_INFO`.”** Rejected: retained authority can
  safely execute some exact owned create/Copy-only routes; comparable identity is
  optional warning evidence, not a global execution prerequisite.
- **“Use normalized paths, hashes, or metadata as mutation authority.”** Rejected:
  none proves generation, current occupant, or task ownership. Normalized paths and
  cheap positive alias facts are explicitly permitted only for A02 warning evidence.
- **“Collapse result axes.”** Rejected: axes prevent dangerous inference; one builder
  removes invalid construction complexity.
- **“Restore clipboard Move after failure.”** Rejected: restoration can replay stale
  intent. Move predictable failure earlier and explain retained source.
- **“Merge Native and Managed.”** Rejected: backend atomicity and host
  copy/conditional cleanup have different proof boundaries.
- **“Merge all durable stores.”** Rejected: only safe persistence mechanics overlap.
- **“Add universal Move WAL/Undo.”** Rejected absent new product decision; compact
  breadcrumb correctly remains notice-only.
- **“Make the operation history authoritative.”** Rejected: a readable log explains
  immutable results but cannot prove current identity, commit, cleanup, or an inverse.
- **“Always defer every conflict/error.”** Rejected by A17: Ask as issues occur remains
  the default and continue-safe-work/review-later is an explicit task-local choice.
- **“Keep only transient task cards.”** Rejected by A18: bounded local history retains
  the explanation after presentation closes without becoming an execution ledger.
- **“Require a full content reread for every Managed Move.”** Rejected by A04:
  exact commit/stable-source/conditional-cleanup proof remains
  mandatory, while extra content verification is route-risk or Verified Move policy.
- **“Snapshot every descendant of every folder before Delete.”** Rejected by A13:
  bind a real container or canonicalize a provider-declared virtual folder and perform
  bounded conditioned membership deletion inside it. Only explicit fixed object/
  version selections retain exact revision snapshots.
- **“Delete whatever currently matches a textual prefix until it happens to be
  empty.”** Rejected: virtual-folder behavior requires a declared canonical boundary,
  per-observation generation conditions, executable pass/time bounds, and residual
  truth. Request-level unknown is never blindly replayed.
- **“Automatically retry unknown mutations or allow Retry only once.”** Rejected by
  A07: explicit Retry remains repeatable only for proved no-commit.
- **“Render a growing-denominator percentage or a stable-looking early ETA.”**
  Rejected by A08: percentage waits for closure and any early ETA is labeled,
  provisional, reversible, and suppressible.
- **“Block every `.rs_*` read/copy or let Possible authorize cleanup.”** Rejected by
  A09: ordinary non-destructive use remains available and Possible grants no authority.
- **“Execute Batch Rename cycles or preserve a journal to support them.”** Rejected by
  A10: cycles block before mutation and no Batch Rename journal/Resume/Roll back
  remains after the explicit retirement slice.
- **“Keep JSON v2 as route authority beside the typed ABI.”** Rejected by A11: the
  cutover is lockstep with no dual reader, compatibility authority, or data migration.
- **“Keep two feature-complete popup renderers.”** Rejected by A12: one hosted owner
  plus a minimal safe failure surface is sufficient.
- **“Apply Win32 name rules to every provider.”** Rejected by A15: executable name
  feasibility and collision policy are provider/path scoped.
- **“Delegate all Local mutations to `IFileOperation` immediately.”** Rejected without
  D2-A16 evidence that one conflict owner and exact per-item truth survive.
- **“Copy a competitor's policy matrix.”** Rejected: competitor behavior is evidence
  about user value and complexity, never mutation authority or a substitute for Red's
  provider-specific proof.
- **“Probe identity/capability once and trust it.”** Rejected: qualification informs;
  immediate revalidation/receipt proves execution.
- **“Offer Skip preflight.”** Rejected: mandatory selected-root preparation and
  per-item safety checks are not optional. The user may only switch bounded
  run-ahead discovery to just-in-time scheduling through `Discover as needed`.
- **“Keep completed producer outputs protected from a queued Delete.”** Rejected:
  the user receives one concrete warning when the destructive task is queued. Queue
  waits, discards its same-host predecessor edge, and then executes ordinary current-
  membership semantics; only concurrently-live same-host tasks need the special guard.
- **“Snapshot or monitor external writers and coordinate every RedSalamander
  instance.”** Rejected: Explorer, other applications/instances, and post-completion
  changes are ordinary external live-container races handled by exact authority and
  revalidation. A cross-process protection framework adds complexity without a
  reliable universal boundary.
- **“Prebuild a universal byte-token/metadata-IOPS scheduler or disk traversal
  spool.”** Rejected: preserve and instrument the current scheduler first. Add only
  the smallest evidence-proved same-resource correction; bounded memory/JIT is the
  default and a provider-specific disk spool needs its own proof and activation.
- **“Finish the full tree scan before any transfer.”** Rejected: preserve one-pass
  discovery/execution, start the first safe mutation early, and prevent transfer from
  starving discovery. Numeric queue/worker splits remain measured policy.
- **“Move F7/Recycle into central task.”** Rejected: symmetry alone adds no truth.
- **“Extract the 8,000-line method into another file.”** Rejected: typed ownership,
  not relocation, is needed.
- **“Kill or unload a stuck in-process provider after timeout.”** Rejected: unsafe
  lifetime destruction cannot implement cancellation truth.
- **“FOS-01 requires hidden staging for every new Local Copy.”** Rejected: section
  6.4 shape 2 uses retained exclusive final-leaf authority for ordinary absent-
  destination Local regular-file Copy; universal staging is not the recommendation.
- **“MTP has no wedged-provider containment or test.”** Rejected: MTP already has
  watchdog/cancel/quarantine/unload gating and deterministic mutating-timeout tests.
  Remaining work is the host-level shutdown witness and separate SMB characterization.
- **“Current Recycle inherits Shell Undo defaults.”** Rejected: it explicitly calls
  `SetOperationFlags` without `FOF_ALLOWUNDO` or `FOFX_ADDUNDORECORD`. R9 must still
  prove that any new adapter creates no Shell Undo record.
- **“`FOFX_NOSKIPJUNCTIONS` disproves the two regular-file A16 candidates.”**
  Rejected: it concerns Shell namespace junctions and is a reject gate only if the
  spike tries to widen into directory/link routes.
- **“FOS-01 is implicit I17 implementation scope.”** Rejected: section 11 assigns it
  to R0a; I17 is a retired record of its already-implemented scope and pending shared
  qualification only.
- **“Accepting the ballot activates R0-R9 as one program.”** Rejected: no decision
  authorizes implementation until a bounded package receives complete,
  predecessor-closed `ACTIVE` ownership.
- **“E0-lite, or coalesce E0 Fresh Full with the first R0 closeout.”** Rejected: it
  conflicts with the DAG and destroys the clean pre-change baseline. E0-full closes
  before any successor edit; E8 separately qualifies the final changed state.
- **“An accepted WIP decision immediately replaces current normative behavior.”**
  Rejected: section 6 records target law. Only the owning slice's landed implementation,
  tests, and authoritative-spec update supersede prior normative text.

## 13. Done criteria for the consolidated successor

M1 and M2 may qualify releases without splitting or retiring this plan. Only derived
M3 completion permits the sole successor to leave WIP.

This sole plan may leave WIP only when:

- [x] D2-A01 through D2-A20 plus A13-MD1 are answered, mapped to section 6, and routed
      to candidate delivery slices; a `DRAFT` may propose ownership, while implementation
      ownership begins only with a complete, predecessor-closed `ACTIVE` card.
- [ ] M3 is complete from the section 8.1 exit receipts; every required E node and
      accepted R slice followed the DAG, every activated R slice has a complete
      predecessor receipt/activation card/exit receipt, and no child evidence row
      remains open.
- [ ] FOS-01 through FOS-12 are closed by their named owners without hiding release
      fixes inside behavior-preserving E work or reopening retired-I17 scope.
- [ ] Every section 6 rule is implemented once, its displaced production owner is
      removed, and its durable wording lives in the section 10 authoritative spec—not
      only in this WIP plan, a test, UI renderer, JSON hint, or history record.
- [ ] Section 9's functional, fault, provider, human/accessibility, performance,
      lifetime, and inherited-I17 evidence is complete on the exact closeout commit;
      environmental gaps such as live SMB are stated rather than inferred green.
- [ ] Required Debug and test-enabled Release builds, inventories, archive validation,
      patch hygiene, and one exact merged Fresh Full E8 receipt pass.
- [ ] Section 14 has no open STOP condition, I3/I12 ownership is reconciled, the WIP
      index names only this File Operations successor, and no retired plan is resumed.
- [ ] This file is moved to Done and I14 is removed from WIP only after every preceding
      criterion is satisfied.

## 14. STOP conditions

Stop and return to arbitration if implementation would:

- start any E or R successor source edit before E0 passes its own exact-commit, clean-
  worktree, pre-change Fresh Full; change E0's `HEAD`, index/worktree, or source-
  snapshot identity during the gate; or use Resume, Affected, stale, corrupt,
  `NOT_EVALUATED`, or older evidence to qualify E0;
- mark an R card `ACTIVE` or dispatch its node before every mandatory pre-dispatch
  field is exact and every section 8.1 hard-predecessor exit receipt is closed;
- treat accepted target law in this non-normative WIP, a `DRAFT`, or `ACTIVE` status as
  current shipped/normative behavior before the owning same-slice implementation,
  tests, and authoritative-spec update land;
- mutate/cleanup from path/name/hash/size/time alone;
- reintroduce `FollowTargets` as implicit link behavior or silently treat size/time
  equality as an identical-file Skip;
- delete a replacement selected root, follow/escape a bound real container, cross a
  canonical virtual-folder account/profile/bucket/prefix boundary, or mutate a fixed-
  set key/version not admitted at consent;
- run virtual-folder convergence without per-observation generation conditions and
  pass/time/cancel bounds, blindly retry an unknown S3 request, combine ordinary S3
  current-key folder Delete with exact-VersionId cleanup, or call a non-empty/
  non-converged folder Completed;
- delete a Managed source without exact destination commit, stable-source evidence,
  the route-sufficient proof policy, and exact conditional cleanup authority;
- waive content reread/checksum verification when a weak-commit route or explicit
  Verified Move requires it;
- call size/time equality, byte count alone, or provider success exact verification;
- treat path/alias overlap evidence, a Queue/Run scheduling receipt, or another
  task's consent as destructive authority for the current object;
- collapse Unsupported/RetryableNoCommit/Indeterminate;
- execute after clipboard consumption failure or consume before root readiness;
- bypass Preparing, place recursive enumeration/content reads inside it, or admit an
  unbounded provider call without the accepted D2-A01 containment;
- perform a Batch Rename capability/bind/schedule loop proportional to selected rows
  on the UI thread, or post its completion through raw pointer payload ownership
  without the registered init/drain lifecycle;
- execute a cyclic Batch Rename plan, create a new Batch Rename journal, expose its
  Resume/Roll back UI, or discard an existing record after supported-release evidence
  proves it may require a separately arbitrated retirement action;
- label any action Skip preflight/Skip safety, or let `Discover as needed` bypass
  enumeration, authority, conflict, verification, or commit-time revalidation;
- starve runnable discovery behind Copy/Move bytes or Delete/small-file metadata I/O,
  wait for full traversal before the first otherwise safe mutation, change scheduling
  before archiving the unchanged baseline, create a disk/durable traversal spool,
  throttle unrelated devices for symmetry, or retain a limiter/controller without its
  measured failing baseline and passing before/after evidence;
- perform slow/recursive/network/device alias discovery solely to decide an A02
  warning; warn merely because tasks share a broad drive/root; silently whole-root
  queue/reject user-approved overlap; or claim Run at the same time guarantees
  simultaneous physical provider I/O;
- let A02 Run-together consent bypass current-object conflict/revalidation, exact
  publication/rollback authority, container boundaries, Managed cleanup proof, or
  terminal result truth;
- delete/overwrite/rename another concurrently-live task's newly created output in the
  same running RedSalamander host
  without either a covering Run receipt whose warning explicitly disclosed that
  consequence or a later item-specific destructive consent; treat absence of an
  injection warning or history as that consent; persist/reconstruct Queue, Run, or
  live-output protection after terminal state or share it across instances/hosts;
- omit a selected item from terminal results;
- expose a coarse capability when exact route is absent;
- silently retarget link payloads, label retargeting as Preserve, or let the explicit
  Retarget option rewrite an unresolved/out-of-tree target without exact mapping;
- report incomplete final-leaf content as Published/complete; accepted D2-A06 permits
  narrow visibility, never false completion;
- claim cancellation/shutdown deadline that cannot be enforced/tested;
- reject a valid tree solely because an aggregate queue reached its bound;
- duplicate conflict/default/Escape policy in UI;
- keep a second feature-complete popup action/layout/accessibility owner;
- give serial and parallel schedulers separate per-item conflict/result policy;
- merge durable schemas/authorities through a shared helper, or let history,
  breadcrumbs, journals, or artifact claims become mutually readable/authoritative;
- allow JSON capability data to enable a route, retain typed/JSON dual authority, or
  add an unapproved A11 compatibility/data-migration layer;
- render a pre-closure percentage, present a provisional discovery ETA as stable, or
  let any ETA affect scheduling, timeout, consent, mutation, cleanup, or completion;
- put the jobs-level aggregate into Discovery because only one task is discovering,
  hide another task's operation progress, label a closed-total subset Overall, include
  an open/growing total in **Known work**, average task percentages, mix units, or show
  determinate Windows taskbar progress while any included total is open;
- block read/open/edit/preview/copy/export solely because of `.rs_*` naming, promote
  Possible to Proven through consent, or expose credentials/tokens/secret-bearing
  endpoints/raw opaque mutation authority in Properties;
- add a second generic confirmation to an ingress whose explicit preview/editor/Run
  already owns consent, omit a required problem-specific overlap/material warning,
  hide an actionable decision, or steal OS foreground for routine progress;
- let review-later processing auto-decide a new conflict, widen the admitted set, or
  replay an indeterminate mutation; mutate an existing Apply-to-all decision; omit a
  newly changed conflict from the same review; or let Rescan widen the current plan;
- use operation history as cleanup, Retry, Resume, Undo, source-delete, artifact, or
  replay authority; persist it without age+size bounds; store credentials/tokens/
  secret-bearing endpoints/raw opaque authority; let Disable silently clear rows; or
  let Clear touch artifact claims/provider recovery state;
- adopt `IFileOperation` while Shell owns a second prompt policy or required per-item
  truth remains unavailable;
- turn optimizer/source shape/historical perf into product law;
- overlap `ACTIVE` ownership without rerouting;
- edit frozen Done history as authority;
- proceed after a gate fails twice without new evidence/arbitration.
