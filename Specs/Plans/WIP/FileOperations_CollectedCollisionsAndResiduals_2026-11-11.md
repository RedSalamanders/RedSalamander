# I64 File Operations collected collisions and I63 residuals

Status: **ACTIVE, I64. Plan prepared on 2026-11-11; implementation not started.** This plan takes over
everything plan I63 did not finish, so that nothing is left only in a closed file. I63's owner decided on
2026-11-11: "Slice G1–G6: put in another detail plan for after". The history of I63 is in [the closed I63
plan](../Done/FileOperations_UxFidelityAndPaneActivation_2026-11-10.md) and is not repeated here.

Rank: this plan was written as `I63` and the closed plan as `I62`. Both were renumbered on 2026-11-11,
before their pull request, because the explicit state machines plan had taken `I62` on `main`. Commit
messages of the branch and the test-run archive folder `2026-11-11_i62-closeout` keep the old ranks.

Priority: **P1**. Effort: **L**. Part 1 (slice G, collected collisions) is about 2,000 to 2,400 changed
lines by the design's estimate, half of them tests, and is a multi-day change. Part 2 (I63 residuals) is
a list of independent items, most of them small. Implementation risk: **High** for slice G (the
worker-pause rule, preemption of a visible prompt, folder cleanup holds in a Managed Move), **Medium**
for the residuals of slices B and F, **Low** for the rest.

Planned on 2026-11-11 at the I63 closeout: the commit that adds this file is the planned-at commit (a
commit cannot name its own hash; the drift check below resolves it).

This is a **non-normative implementation plan**. Current contracts stay in force until each item lands
its code, tests and authoritative-spec edits together. Every "Target contract" block below is proposed
normative text for the named spec; it is not authority while it lives here.

## 0. Drift check

Run from the repository root before any edit, and again before each item:

```powershell
$plan = 'Specs/Plans/WIP/FileOperations_CollectedCollisionsAndResiduals_2026-11-11.md'
$plannedAt = git log --diff-filter=A -1 --format=%H -- $plan
git diff --stat "$plannedAt..HEAD" -- RedSalamander Common Plugins Specs docs Tools Tests
git status --short
```

Evidence in this plan has three kinds, and the text says which one applies:

- **Read on 2026-11-11.** A code anchor that a read-only survey of the I63 branch state recorded. Re-read
  each one before relying on it. The slice G survey records that `FolderWindow.FileOperations.State.cpp`
  changed while it was being read, so every line number in section 4.9 can be off by tens of lines.
  Re-anchor by function name.
- **Recorded by I63.** A statement from the implementation log in section 18 of the I63 plan. It says
  what that log claims was built or run. This plan has not repeated any of those runs.
- **Not established.** Stated as such.

No build or test was run for this plan.

## 1. Origin and outcomes

I63 delivered slices A, C, D and E, slices B1 to B8, G7, and most of F (I63 section 18). Its owner closed
it on 2026-11-11 and asked that slice G1 to G6 move to a new detail plan. The same closeout found
residuals in several slices. This plan has two parts.

| Part | Content | Section |
|---|---|---|
| 1 | Slice G: a name collision does not stop the rest of its task (G1 to G6) | 4 |
| 2 | I63 residuals: items not built or not tested, grouped by the I63 slice they came from, the P0 baselines, and one watch item | 5 |

Outcomes when this plan is done:

1. A name collision between files met by the bridge walk never stops the independent items of its task.
   The collisions are asked together, the task shows a live count, and each set-aside item is run later
   as its own transfer under the ordinary rules. The limits of the first version (section 4.2) are stated
   in the spec and the user documentation, and the owner has accepted them (G-D5).
2. Every item of Part 2 is either done with a test, declined by the owner with the reason recorded, or
   routed to a named owner.
3. No requirement of I63 lives only in the closed I63 file.

## 2. Ownership and coordination

- **I63** is closed. Its Done record is frozen. This plan does not edit it.
- **I18** (confirmation preferences and outcomes) owns start confirmation, the `fileOverwrite` and
  `directoryMerge` settings, metadata-outcome policy and the start-surface control for the task collision
  policy (I63 D-13). This plan adds no setting. I63 B6 exposed the policy value through
  `OperationAdmission::presetCollisionPolicy`; the control that sets it stays with I18.
- **I60 Wave 5** (`W5.2` queue lifecycle state, `W5.3` pause slot yield) edits the queue region that G5
  touches (section 4.5). Whichever lands second rebases. This plan introduces no second ordering key:
  `_queueOrderKey` is the ticket.
- **H3** (later features) held "Continue safe work and review conflicts later". I63 slice G took the row
  over on 2026-11-10. The I63 closeout moved H3's pointer to this plan.
- **DxUi** is an external pinned dependency. No item here requires a pin bump. If tabular figures on
  hosted graph labels (RF8) or any other item needs a DxUi change, stop and route it to the DxUi
  repository and **I31**.
- **I3**, **I29** keep their stress and qualification scopes. **I4** keeps the source-shape test
  migration. This plan adds behavior tests and updates only the source contracts its edits invalidate.
- Undo (I63 D-15) is not in this plan. It consumes the per-item typed results of B and G, so a separate
  plan opens after G lands.

## 3. Owner decisions

### 3.1 Recorded, and binding on this plan

- **D-4 (owner, 2026-11-10), authority ladder.** A location is asked for the level of safety it can
  provide and nothing more. Where file identity exists (more than the name), use it. Where it does not
  and rename or delete by name is atomic and cheap, use it with the expected item facts (name, size,
  dates, attributes). Otherwise behave as Windows File Explorer does. It governs replacement, the removal
  of a Move's source (B7) and a folder that already exists (B1.5). One extra check was added because the
  relaxed rule would otherwise permit a loss: one file reached by two spellings (B8).
- **D-16 (owner, 2026-11-10), rejected.** No policy compares file content, with Verify on or off. A copy
  never leaves a partial file under the final name, so size and dates are the right level for `Delete
  source for all with the same size and date`.
- **D-17 (owner, 2026-11-10).** Leftover recognition covers only the names the current build writes. A
  name an earlier version wrote is an ordinary file: no badge, no cleanup-dialog row.
- **Owner rule "right level of security, nothing more".** Use the strongest authority a location really
  provides, add no ceremony on top, and otherwise match Explorer. Every safeguard proposed in this plan
  states its cost next to its benefit. A safeguard with no stated cost is not ready for the owner.
- **D-6 (owner, 2026-11-11).** An action that would be alone under `More...` is a direct button. This
  includes `Delete all remaining permanently` and `Delete all remaining by name`. I63 delivered it at its
  closeout: a direct button immediately before Cancel, never the default and never the Escape action,
  with no added confirmation step. Nothing of it remains here.

### 3.2 Open for slice G

All five are OPEN. The first version of slice G (section 4) is written for the recommendations; the owner
may change any of them before G1 starts, and the closeout of slice G names the ones used. G-D2 and G-D4
need an answer before any code is written, because they decide whether the first version touches two of
the STOP conditions in section 9.

| ID | Question | Options | Recommendation and cost |
|---|---|---|---|
| G-D1 | Does one answer decide one file or every set-aside file of its class? | (a) A single answer decides only the file it was shown for; the next undecided file is then asked. Apply to all and the B6 policies decide every set-aside file they cover, and later ones as they are found. (b) A single answer is applied to every undecided set-aside file of the same class whose cached-decision eligibility holds, as Explorer does. | **(a).** Cost: more clicks when there are many collisions and the user chose neither Apply to all nor a policy. (b) costs a click that decides files whose own facts were never shown to the user. It is implementable: on a single answer, set the receipt on each undecided record with the same scope. I63's draft contract text said "an answer applies to the set-aside items of its class"; the proposed text in section 4.3 follows (a) and this row says what to change for (b). |
| G-D2 | What happens when an immediate problem (access denied, disk full, file in use) arises while a collected prompt is on screen? The arbiter holds one actionable prompt per task. | (a) The immediate problem preempts the collected prompt: its generation retires, a click in flight for it is refused, the immediate prompt is shown, and the collected prompt returns when that one is answered. (b) The immediate problem waits behind the collected prompt. (c) The collected question is rendered in a second card region with its own slot. | **(a).** Cost: the buttons on screen change without an action by the user. That is close to the I63 STOP "change the buttons of a prompt the user can already press", so (a) needs the owner's explicit acceptance, and the spec states it. No wrong action can be taken, because the retired generation is refused. (b) breaks D-12's "immediate": the waiting worker holds the bridge's callback mutex, which freezes that whole bridge until the answer. (c) is a large UI change. |
| G-D3 | Does a running task with only questions left give up its queue place in the first version? | (a) No. It keeps the queue as it does today, and the limit is documented. (b) Yes: the task leaves the active-operation count and its exclusive-transfer weight while keeping its scopes, and re-enters through the queue when answered. | **(a).** Cost: with only questions left, the next queued task does not start, so the I63 sentence "gives up its queue place" is not delivered in the first version. (b) needs `EnterOperation` to wait in the middle of a task and the keep-awake and Waiting presentation to change. That is not a contained change and touches the STOP about queue places left behind. |
| G-D4 | Does the first version include Managed Move, holding the parent folder's authority until the answer? | (a) Yes, for the parallel walker. The folder's cleanup authority stays held until its set-aside descendants are resolved. (b) Copy and Copy-only Move only; a Managed Move keeps immediate prompts, and G4 is not needed. | **(a).** Cost: the hold window is longer than an immediate prompt's, and the user cannot rename or move that source folder meanwhile. Today the same authority is held for the whole length of an immediate prompt, and the set-aside item itself holds nothing. (b) leaves the cross-volume Move, which the owner's scenario 10 names, asking at every collision. |
| G-D5 | Does the owner accept that a plain Local Copy (a provider-owned walk) and files selected directly (not nested in a selected folder) keep immediate prompts in the first version? | (a) Accept; plan the provider walks as a successor. (b) Do not count G as delivered until provider walks are covered. | **(a).** Cost: the most common case, a plain Local-to-Local Copy, still stops at each collision. That is the opposite of D-12's main case, and the owner should read it as a stated gap, not as completion. (b) needs a proof that the host can tell a provider's Skip for a collision from a Skip after a real failure (both set one failure flag in the Local provider), and a second pass inside the per-item loop before the root's single typed result is stored. |

### 3.3 Open for the residuals

| ID | Question | Options | Recommendation and cost |
|---|---|---|---|
| R-D1 | `Delete source` is not offered when the source location cannot identify items (a share without file IDs). Offer it there through the by-facts delete of B7? | (a) Offer it. (b) Leave such a source with the ordinary choices. | **(a)**, because D-4 allows level 2 and B7 already removes a Move's own source that way. Cost: the removal is by name, size and date instead of object identity, with B8's probe before it. (b) costs nothing and leaves a Move run again over such a source unable to finish by one choice. |

### 3.4 Decisions I63 implemented on its recommendation

I63 section 3 lists open decisions D-2, D-3, D-7, D-8, D-9, D-11, D-12, D-13, D-14 and D-15. The I63 log
records slices written to those recommendations. It records no confirmation from the owner for any of
them except D-6 (decided on 2026-11-11 and delivered). The closed I63 plan records each as "implemented
as recommended under the owner's order to implement; not confirmed one by one". Item RX1 asks the owner
for the explicit answer.

## 4. Part 1, slice G: collected collisions (G1 to G6)

### 4.1 Current behavior

Read on 2026-11-11 by the slice G survey; the survey changed no file.

- When one worker of a task waits at an ordinary prompt, the **whole task stops**.
  `Task::WaitWhilePaused` blocks every caller of the task while a prompt is active and owned by another
  thread (the conflict convergence wait). Its callers include the bridge's read loop once per chunk, the
  parallel worker loop, the producer, the sequential walker, `RelocateChild`, `Task::FileSystemIssue`,
  the task and verification progress callbacks, and the per-item loop. Pumps therefore stop between
  chunks with their reader, writer, stage and permits held, and provider walks stop at their next
  progress callback.
- The bridge holds its `callbackMutex` across the whole prompt, and so does the Local provider's parallel
  walk. A second collision of the same bridge or provider walk blocks on that mutex before it reaches the
  arbiter. A collision from another root of the same task waits in `BeginConflictPrompt` for the
  arbiter's single prompt slot.
- On a binding route, `Task::FileSystemIssue` binds the occupant before it waits, so the prompting worker
  holds a bound object on the occupant for the whole wait.
- The provider's issue callback accepts any `FileSystemIssueAction`. The Local provider handles
  Overwrite, ReplaceLink, ReplaceReadOnly, Retry, KeepBoth, Skip and Cancel. On Skip a file is marked
  skipped and a failure flag is set, the same flag a real continue-on-error failure sets.
- A host-owned selected root holds its `TransferMutationGuard` (bound source, bound destination,
  ancestors and the source parent) for the whole call, and the root's typed result is single-shot: a
  second terminal for the same root is rejected.
- A plain Local-to-Local Copy within one provider instance hands the task no destination file system, so
  it goes to the Local provider's `CopyItem`, a provider-owned walk. The bridge is used for a Managed
  Move, Verify, or an "independent new files" admission. Exactly when `independentNewFileTransfer` is set
  was not read, so the boundary of what the first version covers is approximate.

I63 section 11 states the intent: the owner promoted H3's "Continue safe work and review conflicts later"
on 2026-11-10, and D-12 recommends collecting name collisions by default, as Explorer does, with other
problems keeping their immediate prompts.

### 4.2 The contained first version, and what it does not cover

The only point where a collision can be set aside with nothing of the item held is the bridge's file
pre-flight (`ValidateDestinationOverwritePolicy`, reached from `AdmitDestination`). It runs before the
source is bound, before any permit and before the stage exists, as the code comment of the existing B4
hook states. The first version sets aside exactly these collisions:

- a nested regular file (a file inside a selected folder) met by the bridge's file pre-flight;
- of the plain class only: the incoming file onto a plain regular file, not read-only, not a link, not a
  folder, with a scope that may take Apply to all;
- in the parallel walker for Copy, Copy-only Move and Managed Move (G-D4), and in the sequential walker
  for Copy and Copy-only Move.

The first version does **not** cover, and keeps the immediate prompt for:

| Case | Reason |
|---|---|
| Plain Local-to-Local Copy (provider-owned walk), and every other provider-owned walk | The provider's Skip sets the same failure flag as a real failure, so the host cannot separate "set aside" from "failed" without a counting protocol that needs proof. The second pass would be a new nested call inside the per-item loop. This is a scope choice, not the I63 STOP: the Local walk can be answered Skip without an interface change (G-D5). |
| Files selected directly (selected roots) | The root's guard is held for the whole call and its typed result is single-shot. |
| Read-only occupants, links, type mismatches (file onto folder, folder onto file or link) | They hold a permit or the folder's cleanup authority, or are outside the B6 class. |
| A Managed Move in the sequential walker, a Managed Move with Verify on (it uses the sequential walker), and the rename merge | The sequential frame stack pops before the second pass, so the folder would lose its cleanup record. |
| Any problem other than a name collision, including a problem met in the second pass | D-12. |
| A running task with only questions left giving up its queue place | G-D3. |

The spec states each limit. The user documentation says that a plain Local Copy still asks at each
collision.

### 4.3 Target contract

For `Specs/FileSystem/FileSystem_FileOperations.md` (new section "Collected collisions" after "Conflict
model") and `Specs/UI/UI_FileOperationsPopup.md` ("Conflict Layout", card states, Issues). Proposed
normative text, written for the recommendations of section 3.2:

> A name collision between regular files met by the bridge walk does not stop independent work. The
> colliding item is set aside before any change to it: its source is kept, nothing is bound or staged for
> it, and the task continues with every item that does not depend on it. The card shows the question at
> once with a live count and up to three names, without taking focus, and says whether the rest of the
> task is still running or is finished. The user may answer at any time.
>
> One collected prompt is published at a time, for the oldest undecided item. Its action set is published
> with that item and does not change while the count changes: only the count, the listed names and the
> finished-or-running line change, under the same generation. A single answer decides the item the prompt
> was shown for. Apply to all and a B6 policy decide every set-aside item they cover, and later items as
> they are found.
>
> When the walk has finished everything else, each set-aside item is run as its own transfer, in original
> order, under the ordinary rules. Its facts are read again. An answer given for the item is used only
> when the item's paths, kind, size, last-write time and attributes are still the ones the user was shown
> and the answer is still offered; otherwise the item returns to the question as an immediate prompt on
> the current facts. An answer remembered by Apply to all covers its class by design and does not compare
> facts.
>
> A set-aside item holds no handle and no authority. A Managed Move does not remove a folder while one of
> its descendants is set aside; a folder that cannot be removed for that reason is named in the result.
> The set-aside list is bounded by count and by retained path text. At the bound, the next collision
> takes the immediate prompt, which waits for an answer as it does today, and the task records one
> Information note.
>
> A problem other than a name collision keeps its immediate prompt. An immediate prompt that arises while
> a collected prompt is shown replaces it; the collected prompt returns after the immediate one is
> answered. A click on the retired prompt is refused. [G-D2 (a)]
>
> Cancel on the card or on the collected prompt cancels the task. Each unresolved item is recorded as not
> attempted, with its source kept and its destination unchanged. Closing the popup, a timer, a focus
> change or navigation answers nothing.
>
> The computer is kept awake while the walk runs. When every remaining item waits for an answer, the rule
> of the keep-awake section applies unchanged: the request is released after two minutes without
> progress.
>
> **First-version limits (stated in the spec).** The cases listed in section 4.2 keep immediate prompts.
> A task with only questions left keeps its queue place. [G-D3 (a), G-D5 (a)]

If the owner chooses G-D1 (b), replace the second paragraph's third sentence with: "A single answer is
applied to every undecided set-aside item of the same class for which the answer is offered." If the
owner chooses G-D3 (b), replace the last limit and add the queue rule of I63 section 9 case (b) for a
running task.

### 4.4 Design in brief

The design was read on 2026-11-11; code anchors are in section 4.9.

- **State.** A bounded, ordered store of set-aside records lives on the `Task` under the arbiter's mutex,
  so publishing, counting and answering are one transition. A record holds source and destination path
  text, the two items' facts as values, the comparison, the decision scope, the bucket, the "offered"
  flags that the prompt needs, an optional receipt (a single answer), and its retained byte count. It
  holds no handle. The bridge keeps only side tables of values: the parallel directory holds and the
  sequential deferred folder metadata.
- **Prompt.** `ConflictPromptState` gains a `collected` flag, the record's ordinal, the undecided count,
  up to three names and a "rest finished" flag. A collected prompt has `ownerThreadId == 0`.
- **Workers.** `Task::WaitWhilePaused` does not wait for a collected prompt. This is the one deliberate
  change to an invariant. The survey found no code that needs "no progress during a prompt" for
  correctness; a full FileOps and Commands run is the only real evidence.
- **Answers.** `SubmitConflictDecision` consumes a collected answer under the arbiter lock through the
  same code the waiting path uses (cache store, B6 grant, accepted-decision diagnostic), sets the
  record's receipt when the answer is neither cached nor a policy, retires the generation and publishes
  the next record.
- **Second pass.** After the walk ends without failure or cancel, the bridge runs each of its own records
  in ordinal order through `PumpAndPublishFile`, after waiting for the record to be decided. The receipt
  is accepted in `BeginConflictPrompt` only on equal facts (section 4.3).
- **Folder holds.** In the parallel walker a folder with a set-aside descendant is not finished, and not
  removed, until the second pass resolves it. The hold propagates to the parent, and a parked folder's
  accounting is released when it is deferred.

### 4.5 Implementation

Anchors for each step are in section 4.9. The design suggests splitting delivery at least into (1) G2 and
the answer path, which the existing prompt tests can exercise, (2) the bridge steps G1, G3 and G4, and
(3) the card, resources, tests, specifications and documents.

**G0. Reading before code.** The slice G survey recorded these as not read. Read them first and record
the answer in the slice's closeout.

- [ ] G0.1 Whether `RequestActionablePromptPresentation` takes focus. The first version avoids calling it
  for a collected prompt and uses only `RequestPresentationReveal`; confirm that the reveal path takes no
  focus and add a debug counter to witness it.
- [ ] G0.2 Whether a `Terminal/Skipped` file-work report for a file that later publishes feeds any
  user-visible count (`Task::FileSystemReportFileWork`). If it does, the bridge needs a flag that frees
  the receipt slot without reporting Terminal.
- [ ] G0.3 Whether the Conflict card layout hides the stream rows. It decides whether the card can say
  "Running" with the question beside the streams. The first version keeps "Needs attention" from the
  first publication.
- [ ] G0.4 When `independentNewFileTransfer` makes a Local-to-Local Copy take the bridge, so the
  documentation states exactly which Local Copies the first version covers.
- [ ] G0.5 Whether `EnqueueTaskDiagnostic` may be called from the UI thread. The design avoids it by
  queuing diagnostics in the store and flushing them from a bridge thread.
- [ ] G0.6 The list of existing tests that answer a nested bridge collision while siblings are still in
  flight, for example the nested Keep Both regression and the pause point
  `g_fileOpsKeepBothNestedConflictPausePoint`. Search every test that builds a source folder with a
  pre-existing nested destination file on a bridge route. Each either sets the new test switch
  `g_fileOpsCollectCollisionsForSelfTest` off with a comment or has its expectations updated. Also run
  the Pester source contracts; the distance-bound contracts were not checked function by function.
- [ ] G0.7 P0.13 as a witness. I63 answered P0.13 by reading. Run one collision at a top-level item, at a
  bridge item with sixteen pumps, and inside a provider-owned walk, and record which other items progress
  while the prompt is unanswered.

**G1. Set aside at the bridge's file pre-flight.**

- [ ] G1.1 New header `RedSalamander/FileOperationsCollectedCollisions.h` in the style of
  `FileOperationsPowerRequest.h`, added to `RedSalamander.vcxproj` and `.vcxproj.filters`: the two bounds
  (section 4.6), `SetAsideRecordBytes`, `SetAsideFits` and compile-time contracts.
- [ ] G1.2 Records and store on the `Task` (section 4.4), the new `ConflictPromptState` fields, a
  per-bridge owner token that is a counter and never an address, and the `PerItemCallbackCookie` receipt
  pointer.
- [ ] G1.3 `Task::TrySetAsideCollision`. Read both items' facts with no lock held
  (`ReadConflictItemMetadata`, as `BeginConflictPrompt` does), refine the bucket, compare, build the
  decision scope, and return "not set aside" unless the bucket is a plain regular-file collision with an
  Apply-to-all-eligible scope. Under the arbiter lock: return "not set aside" when a policy grant or
  cached decision already answers the item (today's route answers it without a prompt); count and note
  the bound and return "not set aside" at the bound; otherwise append the record and publish. After
  unlocking, enqueue the flushed diagnostics and emit the counter. No Issues row is written per item.
- [ ] G1.4 Bridge hook. Add an out-parameter to `ValidateDestinationOverwritePolicy`; call the hook just
  before `PromptDestinationCollision` (before the callback mutex is taken and before
  `Task::FileSystemIssue` binds the occupant), gated on a non-null out-pointer, a collection walker being
  active, and neither a read-only nor a link collision. On "set aside" clear the expected destination and
  the granted flags and return `S_FALSE`. The `replaceOffered` estimate must match the conditions
  `Task::FileSystemIssue` uses; an estimate that is too generous is safe because the second pass checks
  the offered action against the fresh action policy.
- [ ] G1.5 Plumb the out-pointer through `AdmitDestination`, `PumpAndPublishFileOnce`,
  `PumpAndPublishFile` (its re-prompt loop stops on set-aside) and `CopyFile`. `AdmitDestination` must
  not record a skipped destination collision for a set-aside item. The selected root's `CopyFile` passes
  `nullptr`. New bridge members: the active walker kind (none, sequential, parallel; sequential only when
  the walk is not a rename merge and source cleanup is not permitted) and the owner token.
- [ ] G1.6 The release rule for "holds no handle": the set-aside path runs before any bind, permit or
  stage, so nothing of the item needs releasing. Report the item's file-work receipt as ended in the way
  G0.2 allows, and add a test that the item holds no handle (the source file can be renamed while it is
  set aside).

**G2. The arbiter: one collected prompt with a stable action set.**

- [ ] G2.1 `PublishCollectedPromptLocked`, next to `SetConflictPromptLocked`. An immediate prompt active:
  do nothing. A collected prompt active: refresh only the count, the names and the rest-finished flag
  under the same generation, and never touch the action arrays, the default action or the escape action
  (the stable action set holds by construction). No prompt active: build the policy for the oldest
  undecided record exactly as its own immediate prompt would, publish it with `collected = true` and
  `ownerThreadId = 0`, and request the reveal only.
- [ ] G2.2 Extract `ConsumeAcceptedDecisionLocked` from `WaitForConflictDecision` (cache store, B6 grant
  and counters, accepted-decision diagnostic) with no behavior change, and use it from both places.
- [ ] G2.3 `SubmitConflictDecision`, collected branch, under the lock after `TryAcceptLocked`: take the
  pending decision, consume it, set the published record's receipt when the answer is neither cached nor
  a policy, call `RequestCancel` after unlocking for Cancel, retire the generation, publish the next
  undecided record, notify. Put the accepted diagnostic in the store's pending list; a bridge thread
  flushes it.
- [ ] G2.4 The record is decided when its receipt is set, or a policy grant applies, or the cache holds
  an eligible decision for it. The decided test reuses the existing eligibility rules (for example
  `Delete source` only for the same size and date). Answer semantics follow G-D1.
- [ ] G2.5 `Task::WaitWhilePaused`: add `&& ! prompt.collected` to both places that compute
  `shouldWaitForConflict`, keeping the `ownerThreadId == 0` effect visible.
- [ ] G2.6 Preemption (G-D2 (a)). In `BeginConflictPrompt` the wait becomes "no prompt active, or the
  prompt is collected, or cancelled". A collected prompt that is active is retired with `ResetLocked`,
  and the ordinary flow continues. Republish the collected prompt after an immediate prompt retires, in
  `WaitForConflictDecision` right after `ResetLocked` and in `ClearConflictPrompt`; skip it when cancel
  is requested.

**G3. The second pass.**

- [ ] G3.1 Receipt. In `BeginConflictPrompt`, after the fresh metadata reads and the action policy and
  before the policy and cache block, accept the record's receipt only when the paths match
  (case-insensitive), the bucket is still the plain regular-file bucket, the facts are equal
  (availability, kinds, link flags, size known and equal, last-write time, attributes) and the action is
  still offered. Otherwise fall through to a fresh immediate prompt on the current facts. The bridge sets
  the receipt on the cookie only around one single-threaded second-pass call.
- [ ] G3.2 `RunSetAsideSecondPass`, called in the parallel walker after the job wait and before the
  abandoned-items release (with its own buffer lease), and in the sequential walker after the frame loop.
  It runs only when the walk ended without failure or cancel; otherwise every own record goes to cancel
  handling (G5.3). For each own record in ordinal order: wait until decided (a 50 ms wait loop that
  flushes pending diagnostics, checks cancel and refreshes the awaiting-answer state), erase the record
  and republish, call `PumpAndPublishFile` with the receipt and no further set-aside, then release the
  record's folder hold. A failure follows the worker's existing `recordWorkerFailure` and
  `continueOnError` rules.
- [ ] G3.3 Totals. The listing already counted the file; its bytes were never added to the completed
  bytes. The second pass adds them on publication, so progress never moves backwards and discovery still
  closes with no growth after close.
- [ ] G3.4 Later collisions after Apply to all or a policy. `TrySetAsideCollision` asks the cache and the
  grants first with the item's own facts; a match is not set aside and runs at once along today's route.
- [ ] G3.5 Typed result. The root's terminal is stored once, after the bridge returns, so in the first
  version there is no pending typed result to replace. A nested item that ends skipped counts exactly as
  today's nested skip.

**G4. Managed Move directory cleanup waits for set-aside descendants.**

- [ ] G4.1 Parallel walker. `DirectoryCompletion` gains a set-aside hold count and a deferred flag, and
  its finishing body is split into `finishDirectory`. The worker takes the hold before its normal release
  of the directory, so the zero test always sees it. When the last count is released while holds exist,
  the directory is marked deferred, its parked accounting is released (otherwise the producer's queue
  target never drains and a large Move hangs), and its parent count becomes a parent hold. After each
  second-pass record the hold is released; at zero a deferred directory finishes (restore, cleanup or
  keep, report) and the release cascades. Children finish before parents.
- [ ] G4.2 A failed second-pass item marks its directory incomplete, so the cleanup is skipped and the
  folder is kept. The result names kept folders through the existing retention diagnostics; a specific
  reason string for "kept because a file inside it was set aside and then skipped" is optional.
- [ ] G4.3 Sequential walker (Copy and Copy-only Move only). A frame with set-aside children moves its
  created-folder metadata and identity bytes (values) into a deferred entry charged to the byte bound, is
  marked restored, and is restored at zero holds, so the late file does not move the folder's last-write
  time after it was restored. The folder is bound again by path and accepted only with its recorded
  identity.
- [ ] G4.4 A Managed Move in the sequential walker later: moving the cleanup record out of the frame into
  a deferred entry is about 80 lines and the same hold logic. It is a candidate after the parallel
  version has evidence, not part of this plan unless the owner asks.

**G5. Card, queue, keep awake, cancel and close.**

- [ ] G5.1 Card. The collected prompt is an ordinary published prompt, so the status kind becomes
  Conflict ("Needs attention") at once. Copy the new fields into the popup snapshot. Add one line in
  "Conflict Layout" under the facts: the count and up to three names, then either `Other items continue
  meanwhile.` or `Everything else is finished.` Reserve the line's height from the first publication
  (slice F stability) so only text changes while the count changes. The pane decision mirror copies the
  same prompt and may show the line or omit it.
- [ ] G5.2 Issues. No row per item at set-aside time. Rows that appear: the existing accepted-decision
  rows, one Information row per item that ends unresolved at Cancel, and one Information row when the
  bound is first reached.
- [ ] G5.3 Cancel and withdrawal. For each own record: one Information diagnostic `Not attempted`,
  release the hold (the folder is kept; sequential: restore the deferred metadata), erase the record. At
  most the bound's count of rows is written. The bridge then returns `ERROR_CANCELLED` or the failure as
  today. Add a scope guard so `Task::WithdrawSetAside(ownerToken)` runs when the bridge's traversal ends
  by any path, so no record outlives its bridge and no question is shown that nobody can run.
- [ ] G5.4 Keep awake. A waiter counter drives `SetAwaitingUserDecision` (true on 0 to 1, false on 1 to
  0), because two roots may wait at once. The existing quiet period releases the request. A test asserts
  the flag is false during the walk and true in the second-pass wait.
- [ ] G5.5 Queue (G-D3 (a)): the first version does not yield the queue place. The spec states the limit.
  If the owner chooses (b), this item becomes the design of a running task leaving and re-entering
  admission, and it follows I60 `W5.2`.
- [ ] G5.6 Closing the popup answers nothing. A test closes and reopens the popup with a collected prompt
  shown and checks that nothing was submitted.

**G6. Bounds, instrumentation, registration and documents.**

- [ ] G6.1 The two bounds as named constants with compile-time contracts, and a test override
  (`g_fileOpsSetAsideBoundOverrideForSelfTest`) so the bound is reached with three items. Tests at the
  bound (section 4.6).
- [ ] G6.2 Counters `FileOps.Conflict.SetAside` and `FileOps.Conflict.SetAsideBoundReached`, plus task
  fields for tests. Optional: `FileOps.Conflict.SetAsideWaitUs` and
  `FileOps.Conflict.SetAsideRetainedBytesMax`, which feed the "retained bytes at the bound" evidence.
  Plugin counters go to ETW, so the FileOps runner archives them as it does the existing
  `FileOps.Conflict.*` counters.
- [ ] G6.3 Resource strings (section 4.7) in `RedSalamander.rc` and the four satellites, with
  `Resource.h` ids, marked unreviewed where machine translated. Run
  `ResourceLocalizationContracts.Tests.ps1`.
- [ ] G6.4 Family registration (section 4.8): new `SelfTestState::Step::FileOps_CollectedCollisions`, an
  entry `FileOpsFamilyDefinition{L"FileOpsFamily_CollectedCollisions", ...}` in
  `kFileOpsFamilyDefinitions`, the default phase order, a runner dispatch case, and
  `Specs/Testing/Testing_SelfTests.md`. The compile-time check beside the table enforces that every step
  is in a family.
- [ ] G6.5 Specifications and documents (section 6) and one `Changes/` fragment with one top-level
  bullet. Run `Tools/Test-Changelog.ps1` first. Remove the dead `BridgeCallback` wrapper if G confirms it
  is never instantiated; it is an S-item outside G and may be a separate change.

### 4.6 Bounds and instrumentation

Proposed values; the owner may change them.

- At most **512** set-aside records per task, and at most **4 MiB** of retained path text and facts per
  task. The byte bound is the effective limit for long paths: 512 records of 64 KiB path pairs would be
  32 MiB.
- At the bound the record is not created. The worker takes the immediate prompt, which preempts the
  collected one (G-D2 (a)); the convergence wait stops the task as today. One Information note is written
  per task and `SetAsideBoundReached` is counted.
- A set-aside item retains path text and fixed-size facts only. Nothing else grows with the number of
  collisions.
- Cost of a non-colliding item: nothing. Cost of a colliding item: its two metadata reads (the ones an
  immediate prompt makes), plus one bounded store insert.

### 4.7 Resource strings

The survey found the next free resource id after 50240 is 50241 and that 50228 to 50239 may be claimed by
pending work; re-read `Resource.h` before choosing ids. Proposed English text in the base
`RedSalamander.rc`, with positional placeholders, and the same keys in cs-CZ, fr-FR, ja-JP and sk-SK:

| Name | English |
|---|---|
| `IDS_FMT_FILEOPS_CONFLICT_SET_ASIDE_COUNT` | `Files already at the destination: {0:L}` |
| `IDS_FILEOPS_CONFLICT_SET_ASIDE_REST_RUNNING` | `Other items continue meanwhile.` |
| `IDS_FILEOPS_CONFLICT_SET_ASIDE_REST_DONE` | `Everything else is finished.` |
| `IDS_FILEOPS_NOTE_SET_ASIDE_NOT_ATTEMPTED` | `Not attempted: the task stopped while this file waited for a decision. The source is kept.` |
| `IDS_FILEOPS_NOTE_SET_ASIDE_BOUND_REACHED` | `Too many files wait for a decision; the task now asks about each new one at once.` |

The names listed under the count are item names and need no string.

### 4.8 Tests

New family `FileOps_CollectedCollisions`, registered in `kFileOpsFamilyDefinitions` (a case outside the
table is skipped silently by a family run). Style follows
`RunFileOperationBulkCollisionPoliciesSelfTest`: a `State::Task` built directly with a per-item mode, a
destination file system so the bridge is used, one source folder holding N files, and a responder thread
that answers through the prompt copy and `SubmitConflictDecision`. The walker is chosen by the per-item
concurrency budget (two or more for parallel, one for sequential) with Verify off. Test-only accessors:
the collected state (count, undecided, waiters, retained bytes, published ordinal), the two counters, a
trace of second-pass ordinals, the two test switches, and a counter of
`RequestActionablePromptPresentation` calls.

1. One hundred files in one folder, five collisions, parallel walker. Before any answer, with the
   responder waiting until a second-pass waiter exists, 95 destination files are published, the count is
   5 and the five sources are unchanged. This also proves the worker-pause change.
2. Answers Skip, Replace and Keep both given before the rest finished (a producer delay hook keeps the
   walk going) and after; results asserted per file.
3. Apply to all given at the second collision: later collisions are not set aside. The same for the B6
   Skip-same-size-and-date and Replace-older policies with mixed pairs.
4. An occupant changed while the item was set aside (answer Replace, then rewrite the destination to
   another size before the second pass): a new immediate prompt with the new facts, not collected.
5. Cancel with five unresolved: five "not attempted" Information rows, sources and destinations
   unchanged, `ERROR_CANCELLED`.
6. Bound: the override at 3. The fourth collision gets an immediate prompt while the count is 3,
   `SetAsideBoundReached` is 1, and after the answer the collected prompt returns.
7. Managed Move, parallel walker: a folder with one set-aside child stays at the source until the answer;
   Replace removes it; Skip keeps it and names it. Forty set-aside files in forty distinct subfolders do
   not stall the producer (the parked-folder regression).
8. Sequential walker, Copy, budget 1: the destination folder's last-write time equals the source's after
   a set-aside child was published late.
9. The action arrays, default and escape actions and generation are equal at every poll while the count
   grows. The popup automation IDs are checked in a Commands case.
10. No focus request: the counter of actionable-presentation requests stays 0. A real focus witness needs
    the desktop and the owner's agreement to the time.
11. Progress is non-decreasing and the final completed bytes equal the total.
12. The second pass runs in original order (trace).
13. An immediate problem preempts: inject a failure on a non-colliding file while a collected prompt is
    shown. The immediate prompt appears with a new generation, the old generation's submit returns false,
    and the collected prompt is republished afterwards. Check first that the existing failure injection
    reaches a prompt rather than a failure.
14. A source file of a set-aside item can be renamed while it is set aside (no handle).
15. Closing the popup answers nothing (G5.6).
16. Power: the awaiting flag is false during the walk and true in the second-pass wait; the quiet period
    stays covered by the existing pure-function tests.
17. With only questions left, the next queued task does **not** start in the first version (the limit of
    G-D3 (a) is pinned by a test, so a later change is deliberate). If G-D3 (b) is chosen, the case
    asserts that it starts.

Run each variant on the registered network test root where the environment provides it, and record skips
honestly. Update the existing cases found by G0.6. Test output stays under the approved
`X:\RedSalamander.Perf` root.

Performance: a non-colliding item pays nothing. Archive the 100-item scenario's elapsed time against the
same run answered at the first collision, and the retained bytes at the bound.

### 4.9 Code anchors

Read on 2026-11-11 on the I63 branch state; re-read before relying on any of them. The file changed
during the read, so numbers can be off by tens of lines. Use the function names. "State.cpp" is
`RedSalamander/FolderWindow.FileOperations.State.cpp`.

| Subject | Anchor |
|---|---|
| Convergence wait | `Task::WaitWhilePaused`, State.cpp:14949-15013 (predicate at about 14966 and 15000) |
| Arbiter single prompt slot | `BeginConflictPrompt`, State.cpp:6373 (wait at 6404-6407; policy and cache block at 6510; fresh metadata at 6423-6433 and 6497-6508) |
| Prompt retire | `WaitForConflictDecision` State.cpp:6643 (accepted-decision code 6729-6768; `ResetLocked` at 6772); `ClearConflictPrompt` State.cpp:6077 |
| Prompt publication | `SetConflictPromptLocked` State.cpp:6112 |
| Answer entry | `Task::SubmitConflictDecision`, `State.Conflict.cpp:19-35`; generation check in `FileOperationsConflictBroker.cpp:22` |
| Cache and policy eligibility | State.cpp:5965-6013 |
| Bridge file pre-flight | `ValidateDestinationOverwritePolicy` State.cpp:18995 (file branch 19080-19148, prompt call 19123); `AdmitDestination` 19658; `PumpAndPublishFile` 19610; `PumpAndPublishFileOnce` about 20983; `CopyFile` 21151 |
| Bridge callback mutex across the prompt | `PromptDestinationCollision` State.cpp:17745-17777 |
| Occupant bind before the wait | `Task::FileSystemIssue` State.cpp:10316-10341 |
| Walkers | `CopyDirectorySequential` about 22188; `CopyDirectoryParallel` 22322 (`DirectoryCompletion` 22346-22366, `releaseDirectory` 22410-22454, worker 22561-22572, job wait about 23159); `TraverseAndPublishPath` 23288 |
| Awaiting-answer state | `RefreshAwaitingAnswerQuiet` and `SetAwaitingUserDecision`, State.cpp:14632-14664; quiet period constant in `FileOperationsPowerRequest.h` |
| Queue gate | `State.Queue.cpp:2444-2445`; interlock checks 2329-2395 |
| Selected-root result | State.cpp:26097 and 27286; single-shot terminal check at 14567-14583 |
| Provider Skip | `Plugins/FileSystem/FileSystem.FileOps.cpp` file handling 8164-8212 (Skip sets the failure flag at 8196-8199); callback mutex across the prompt 4001-4031 |
| Card status kind | `Popup.cpp:1572-1575`; snapshot copy 5719-5760; decision mirror `DecisionMirror.cpp:297-302` |
| Test seams | `GetBridgeProducerDelayMsForSelfTest` about 23073; `ConsumeBridgeFailNextFileCopyForSelfTest` about 20992; pause point `g_fileOpsConflictConvergenceItemPausePoint`; family table in the runner at about 2499-2556 |

## 5. Part 2: I63 residuals

Each item comes from the I63 plan text and its section 18 log. "The I63 log records" means that log says
so; this plan has not repeated the run. IDs are `R` plus the I63 slice letter. Where the log is silent on
a thing the I63 text required, the item says "not confirmed in the I63 log".

### 5.1 Slice A, the activating press

- [ ] RA1 Tests. The I63 log records three `cmd_pane_folderView_activation_press_*` cases and states that
  the right-click and Quick Search cases of the test list are not written ("neither path reads the pane
  state"). Write both anyway, because they pin that statement: a right-click on an unselected item in the
  other pane makes it the sole target; Quick Search active in the other pane exits as today. Not
  confirmed in the I63 log: which of the other listed cases the three written cases cover (plain press on
  an unselected and on a selected item with three selected, background press, the Ctrl, Shift and
  Ctrl+Shift presses, the second plain press, double-click on a folder, middle press, drag from an
  unselected item under D-3, press with the same pane's NavigationView focused, and the Terminal-tab note
  of A5). List the covered ones, then write the missing ones.
- [ ] RA2 A7 inactive-selection visibility. The I63 log records that the renderer draws the inactive
  selection and a thinner current-item border and that this was read, not captured. Capture the pane that
  is not active with three selected items and a current item outside them through the visual gallery
  harness (invented fixture data), and record whether each is visible and distinguishable from the active
  style. If either is not, state the inactive style in `UI_FolderView.md` and fix the renderer; the
  first-click rule is not understandable without it.
- [ ] RA3 A5 debug assertion. Confirm the test-build assertion that the pane's focused flag equals the
  focused pane at button-down exists; I63 A5 required it and the log does not mention it.

### 5.2 Slice B, collisions as decisions

- [ ] RB1 B3.2, sizes on provider routes. The provider issue callback carries no sizes, so a collision
  met inside a provider walk shows no size. Make `ReadConflictItemMetadata` fall back from
  `GetFileBasicInformation` to the provider's item-information call that carries size, and name that call
  in the slice's closeout. Cost: one more metadata call on the routes that lack size; none on the Local
  route.
- [ ] RB2 Facts on the finished card of a one-item task (B3 target contract). The I63 log records the
  facts on the Issues row (B3.3) but not on the finished card.
- [ ] RB3 Exact bytes in a tooltip and in the accessible description (B3 target contract), with the full
  date and time. Not done per the I63 log.
- [ ] RB4 An item decided by a remembered answer (Apply to all) or by a policy carries no facts today.
  Decide with the owner whether the typed result should retain the facts read at that step, bounded by
  the existing reservoirs; B6.3 says a policy skip keeps its facts, so a policy skip that carries none is
  a gap against I63's own contract. Check the contract text first and either fix the code or amend the
  text.
- [ ] RB5 `Delete source` without item identity. Depends on R-D1. If (a): offer it for a source whose
  location cannot identify items, through the by-facts delete of B7 with B8's probe before it, and add
  the variants of the `DiscardSource` tests for an identity-less source. If (b): state the limit in
  `FileSystem_FileOperations.md` and in `docs/FileOperations.md`.
- [ ] RB6 `Delete source` limits recorded by the I63 log: non-Local sources and the native-rename Move
  route are not covered. State both in the spec, or cover them. A Move inside one location that is a
  native rename onto an existing name is the B1 decision; confirm what the user is offered there and
  write the test.
- [ ] RB7 Found, not fixed: the Issues "Problem" text of the no-Recycle-Bin question row reads "Recycle
  Bin delete failed." Give it a sentence that says the source's location has no Recycle Bin and the
  delete would be permanent. Also, if that question's bucket were ever `DestinationLink`, a Skip answer
  would be remembered for the task. The I63 log calls this an untested edge: add a test that the
  no-Recycle-Bin question never records a task-wide Skip, or change the bucket.
- [ ] RB8 B6 counts and policies. The replace counter counts decisions, not files replaced: count
  replaced files, and add the `replaced because older` sentence if missing. The two bulk policies are
  offered on a one-file task although they are task-wide choices; offer them only when the task has more
  than one item, as the B6 contract says.
- [ ] RB9 Translations. The new B-slice strings in cs-CZ, fr-FR, ja-JP and sk-SK are unreviewed (the log
  says so). Ask for a review, or record the owner's decision to ship them as they are. The same holds for
  the strings added by items below.
- [ ] RB10 B7 and B8 limits recorded by the I63 log:
  - A folder moved into itself under another spelling is not detected for a source without identity,
    because the guard's ancestor checks do not run without a bound source. Design a by-facts detection,
    or state the limit. Cost of detecting: extra opens on every identity-less folder Move; the benefit is
    avoiding a recursive copy into the source.
  - `BackupRead` on a real share is untested, and the log says that if it fails every source is kept.
    Test on the registered network test root where the environment provides one and record the result or
    the skip.
  - Three metadata opens per removed source. Measure it on a 2,000-item second run and decide whether to
    merge the opens.
  - B8.3: the location classes that do not enforce sharing should be a typed capability of the path
    profile instead of the probe's result alone. Collect the classes from the replacing-rename and
    two-spellings runs (RP3) and add the capability.
- [ ] RB11 B2 items, as the I63 closeout audit found them. B2.6: level 3 (a route with no replacing
  rename) is not built; the spec states there is no fallback to writing over the occupant, and the search
  for a Local location class without a replacing rename was not run (RP3). B2.7 is delivered (one
  Information note per task, a resource string). B2.8, the review checklist, was not recorded during the
  slice; the I63 closeout records what is known. Still without a named witness: an ancestor replaced by a
  link during a file replacement, the target turned into a directory or a link between the read and the
  rename, and that a reader never sees the final name missing (RP3). Write those tests, then record the
  checklist result in `FileSystem_FileOperations.md`.
- [ ] RB12 The performance evidence of slice B: `FileOps.Conflict.ExpectedItemUs`, a 500-collision
  Apply-to-all Replace beside the same run with Keep both, and the 2,000-item second runs under the skip
  policy and under the Delete source policy. The I63 log does not list them. Run them or record why not.

- [ ] RB13 B1.4 and P0.4, the refusal inventory. I63 turned only the identity-less existing-destination
  exits of `prepareExactTransferGuard` into decisions. List every exit of the guard and every
  `TransferSafetyState` value, classify each as *has a safe choice* or *has none*, give each of the first
  kind a decision with at least Skip and Cancel, and each of the second kind a specific sentence. Cost:
  reading only, until an exit is reclassified.
- [ ] RB14 B6.2, item-scoped policy-origin receipts. A bulk policy reads each item's facts at its own
  step but creates no receipt; `FileSystem_FileOperations.md` states they are not in this version. They
  exist for I18's metadata-outcome consent. Decide with I18's owner whether I18 needs them; if it does,
  the receipt belongs with I18's consent model and this item moves there.

### 5.3 Slice C, menus

Nothing remains. D-6 was decided on 2026-11-11 and delivered at the I63 closeout (section 3.1).

### 5.4 Slice D, staging names

- [ ] RD1 P0.9 baseline. The I63 log records that the staging-name perf baseline was not captured before
  the change. Recognition runs per row during FolderView enumeration, and a regression there is a STOP
  condition in I63. Capture the baseline from the I63 planned commit
  `99624c5385c6c0d42f8d1c6f8be54cbe69f506ef` in a separate worktree with
  `DebugMeasureArtifactNameShapeProjectionForTests`, capture the same measurement on the current build,
  and archive both under `Specs/TestRuns/`. If the owner decides the comparison is not worth a worktree
  build, record that decision.
- [ ] RD2 The I63 log records two departures from the D text: the case-only rename hop uses `<leaf>.<32
  hex>.rs-bak` because it carries the original (D2 said the whole-leaf `.rs-tmp` form), and S4 (the
  table-driven recognizer) is replaced by the smaller chain D-17 leaves. Confirm
  `FileSystem_FileOperations.md` and the documents named by D5 describe the shipped forms, and that the
  two stale statements D5 listed (the PID/TID/tick fallback in the two developer documents, the `.rsren-`
  claim in `docs/BatchRename.md` and `docs/DeveloperGuide.md`) are removed.

### 5.5 Slice E, queue order

- [ ] RE1 E8: emit `FileOps.Queue.HeadHoldUs` (the time the head of the queue was held while Preparing),
  as I63 E8 required. The I63 log records `FileOps.Queue.PassedOver` is emitted and `HeadHoldUs` is not.
- [ ] RE2 E4 (evaluate `ShouldQueueNewTask` inside the admission lock so two admissions issued together
  cannot both read "nothing active") is not changed in I63; the log says it was not confirmed as a
  defect. Write the failing-first case with two admissions issued together, then fix it or record that it
  is not a defect. E7 (the queue-paused ordering) is unchanged by choice and needs no work unless a case
  shows it contradicts the card order.
- [ ] RE3 Not confirmed in the I63 log: the registered family case `FileOps_QueueStartOrder` listed in
  the E tests, and the queue wait metrics for a twenty-task queue before and after (the head hold must
  equal the head's own Preparing time). The log records seven scenarios inside
  `FileOps_IdentityInterlockWaiting`. Compare them with the E test list (reversed Preparing durations,
  head waiting for a Permanent Delete answer, head cancelled or failing in Preparing, head queued after a
  named task, reorder, Start now, Queue to Parallel and back, constant card index, waiting ordinals, Move
  Up across an overlap edge) and add what is missing. Each case ends by asserting an empty queue and zero
  active operations.

### 5.6 Slice F, a quiet card list

- [ ] RF1 Budget case `FileOps_PopupLayoutStability` and the counter `FileOps.Popup.LayoutShiftPx` (the
  sum of card-top movement not covered by a running height motion or a user action). They need in-process
  counters and a fixed-tick clock for `Render`. Build both, run the twelve-second synthetic scenario (one
  to eight streams churning, rate text changing every tick), and enforce the budget: unanimated layout
  shift zero, at most four height changes per card, at most one scrollbar toggle, at most one window
  resize after the first second. Register the case. This also yields the P0.8 baseline in use (RP1).
- [ ] RF2 Reserved width for the widest form of each number in the three fact pairs (F2). The log records
  the sticky stacking and rate text hold, not the reserved width.
- [ ] RF3 F5, eased height reveal: layout at target height on the first frame, the bottom block and the
  cards below ease over 200 ms on whole device pixels, rows between are clipped at the top of the bottom
  block, decision content still snaps, reduced motion snaps, and an unsettled height motion selects the
  16 ms cadence for its bounded duration only. STOP conditions: body rows painting over the bottom block,
  the 16 ms cadence outside a bounded motion. Tests: reveal frames never paint a row below the bottom
  block's top; tick counts in `FileOps_PopupIdleCadence`.
- [ ] RF4 F6, viewport. Not done per the I63 log: a stable scrollbar gutter so card width is independent
  of the scrollbar, `AutoResizeWindow`'s `SetWindowPos` moved out of `Render` to a posted step with the
  two-second shrink hold, and the scroll anchor (the first visible card keeps its position by task ID and
  offset when content above changes height). The log also records that the 300 ms scrollbar delay of the
  target contract was dropped because it kept the scroll position at zero for cases that scroll right
  after the first paint. Decide with the owner: remove the 300 ms from the target contract text in
  `UI_FileOperationsPopup.md` (the shipped behavior), or restore it with a fix for those cases. Cost of
  restoring: a delay in showing the scrollbar; benefit: fewer toggles.
- [ ] RF5 F7, card enter and leave motion sharing the height animation state; removal keeps the entry
  until its collapse ends. Cost: a removed card stays for 200 ms.
- [ ] RF6 F9, deterministic *sequence* capture in the visual gallery (N frames of one synthetic scenario
  at fixed ticks), following the screenshot rules (invented fixture data, provenance, no desktop
  takeover). It is the review tool for RF1, RF3 and RF5, so build it first.
- [ ] RF7 Found, not fixed: the height pass and the painter measure a wrapping work summary at different
  widths. Make both use one width, or one measurement, and add a case with a summary that wraps at the
  boundary.
- [ ] RF8 F3: the I63 log records tabular figures on the three numeric fact rows the popup lays out
  itself; the hosted graph's labels are DxUi's and unchanged. If the owner wants tabular figures on the
  graph labels, route it to DxUi and **I31**; this plan does not change DxUi.
- [ ] RF9 Found, not fixed: `FileSystemReportFileWork` can match a stale entry by leaf after the engine
  compacted its rows. Since F1 gave each stream a slot for its lifetime, match by slot and receipt
  identity instead of by leaf, and add a test with a completed stream followed by a new stream of the
  same leaf name.
- [ ] RF10 F-slice specification and documents. The I63 log does not say whether
  `UI_FileOperationsPopup.md` has its "Layout stability" section, `Testing_PerformanceValidation.md` its
  stability scenario, `Testing_SelfTests.md` the sequence capture, or `docs/dev/TestingAndScreenshots.md`
  the capture rules. Confirm each, and amend what the landed behavior changed (the dropped 300 ms, the
  missing gutter).

### 5.7 Slice G7, keep awake

- [ ] RG1 Preparation before admission is not covered by the power request (the I63 log says so). Decide
  whether a long Preparing phase (a large discovery) should hold the request. If yes, extend the per-task
  lease to start at card creation; cost: a request held while a task only waits in the queue unless the
  lease follows the same quiet rule. If no, state the limit in the spec.
- [ ] RG2 No self-test drives real tasks through the keep-awake hooks; the log records only the fake API
  seam. Add a case that runs a real task (a small Copy with a pause point) and asserts one acquire at
  admission and one release at the end, with the seam counting.
- [ ] RG3 Modern Standby. No such machine was available. On one, record `powercfg /requests` while a long
  task runs and while it only waits for an answer, and whether `PowerRequestExecutionRequired` is needed.
  Record the skip where no such machine is available.
### 5.8 P0, characterization baselines

The I63 log records that P0.1 to P0.15 were not run as a separate slice, that each slice recorded what it
read instead of what it witnessed, and that P0.13 was answered by reading. Proposed disposition, for the
owner to confirm:

| P0 item | Disposition |
|---|---|
| P0.1, P0.2, P0.10 (pane press, parent-notify trace, test reliance audit) | Obsolete: slice A landed and its cases pin the behavior. |
| P0.3, P0.4, P0.5, P0.6, P0.7 (identity-less collisions, refusal inventory, Local replace route, menu census, queue order) | Obsolete as baselines: slices B, C and E landed with tests, and C4's flyout counter replaces the census. Still useful only as bug witnesses if B's tests prove too narrow. |
| P0.8 (layout stability baseline) | Has value: RF1 needs it. Item RP1. |
| P0.9 (staging-name perf baseline) | Has value only if the owner wants the recognition comparison. Item RD1. |
| P0.11 (source side and two spellings), P0.12 (the replacing rename) | Have value: they answer the classes B8.3 needs and the atomicity STOP condition. Items RP2 and RP3. |
| P0.13 (what continues during a prompt) | Answered by reading; a witness has value for slice G. Item G0.7. |
| P0.14 (inactive-pane selection) | Has value: item RA2. |
| P0.15 (power) | Obsolete: replaced by RG3. |

- [ ] RP1 Capture the P0.8 baseline with the F4 counters on the twelve-second scenario (card height
  changes, layout shift, scrollbar toggles, window resizes, existing paint timings), archive it under
  `Specs/TestRuns/`, and use it as the "before" for RF3 to RF5.
- [ ] RP2 P0.11 on the registered network test root: Move from the identity-less location to a Local
  disk, Move inside it, Copy from it, and one test folder reached by two spellings (the loopback alias
  the identity-less family uses) with the destination existing. Record the strategy the planner chose,
  what is detected, the share mode the source is opened with, whether a rename over the open source is
  refused, and whether the B8 probe distinguishes one file from two on each location class. Its output
  feeds RB10 (B8.3). Record the skip where the environment lacks the root.
- [ ] RP3 P0.12: run `RenameExactTo` with `replaceIfExists` over a regular file, a read-only file, a file
  another handle holds with and without delete sharing, a directory and a link; run a reader loop on the
  final name during 1,000 replacements and record whether it sees a missing name or partial content;
  record the requests per replaced file from an SMB client trace where one can be taken. The I63 log does
  not record that this was run. If a reader sees a missing name or partial content on the registered
  share, that is an I63 STOP condition (section 9): keep B1's decision and return level 2 to the owner.
  Record any Local location class without a replacing rename (D-11).

### 5.9 Watch item

- [ ] RW1 `DiscoveryScope_MultipleFoldersLiveProgress` failed once in a full FileOps run on 2026-11-11:
  the copied content was correct and one summary number of a plain Local Copy (result code, discovered
  files, discovered bytes or copied files) was not the expected one. It then passed 43 times: 36
  in-process repeats, three isolated attempts, three family-order attempts and one full run, as the I63
  log records. The cause is not established, and the case had not failed in the retained history. The
  case now logs `MultipleFolderLiveProgressSummary` with the four numbers when it fails. **First action
  on a recurrence: read that line**, before any change. No edit is planned until a recurrence gives data.

### 5.10 Decision record

- [ ] RX1 Record the owner's answer for each decision in section 3.4 (D-2, D-3, D-7, D-8, D-9, D-11,
  D-12, D-13, D-14, D-15) as accepted, changed or declined, in this plan's section 3.1 and in the specs
  the decision touches. A declined decision is removed from scope or routed to a named owner. The first
  ones to ask: D-12 (collect by default) and D-14 (always keep awake), because they are already in the
  code and the spec text for them is not confirmed in the I63 log.

## 6. Authoritative documents each item updates

| Item | Specs | User and developer documentation |
|---|---|---|
| G | `Specs/FileSystem/FileSystem_FileOperations.md` (new "Collected collisions" section after "Conflict model"; amend "Conflict model" for preemption and for collected prompts that do not stop workers; "Publication and managed Move" for folders held by set-aside descendants; the keep-awake section for second-pass waits; the queue rules for the first-version limit), `Specs/UI/UI_FileOperationsPopup.md` (Conflict Layout count line, card states, Issues rows), `Specs/Core/Core_FileSystemBridge.md` ("Conflict and cleanup": set-aside before binding, the second pass, holds), `Specs/Testing/Testing_SelfTests.md` (the family) | `docs/FileOperations.md`, `docs/Troubleshooting.md`, `docs/dev/FileOperationsEngine.md` |
| RA | `Specs/UI/UI_FolderView.md` if the inactive style changes | `docs/SelectionsAndLists.md` |
| RB | `FileSystem_FileOperations.md` (conflict model, action table, results), `UI_FileOperationsPopup.md` (conflict layout, finished card, Issues) | `docs/FileOperations.md`, `docs/dev/FileOperationsEngine.md` |
| RD | the documents D5 of I63 named | the same |
| RE | `FileSystem_FileOperations.md` (queue rules), `Testing_PerformanceValidation.md` if a counter is documented | none |
| RF | `UI_FileOperationsPopup.md` (Layout stability, Motion), `Testing_PerformanceValidation.md`, `Testing_SelfTests.md` | `docs/dev/TestingAndScreenshots.md`, `docs/VisualTour.md` captures |
| RG | the keep-awake section of `FileSystem_FileOperations.md` | `docs/FileOperations.md` |

Every item that adds a family or changes a pinned function also updates
`Specs/Testing/Testing_SelfTests.md` and the reviewed source-contract tests its edits invalidate. Each
delivered change carries one `Changes/` fragment with one top-level bullet; run
`Tools/Test-Changelog.ps1` first.

## 7. Order, slices and verification

Recommended order:

1. **Decisions first.** Ask G-D1 to G-D5 and R-D1 (section 3), and RX1. G-D2 and G-D4 must be answered
   before any slice G code.
2. **Residual quick fixes.** RA1, RA3, RB5 (after R-D1), RB7, RB8, RD2, RE1, RE2, RE3, RG2, and the
   stale-entry fix RF9. Each is small and independent.
3. **Slice G**, in the order G0, G2 with the answer path (testable with the existing prompt tests), then
   G1, G3, G4, G5, G6.
4. **The larger independent residuals**, which may run in parallel with or after G: the F slice in the
   order RF6 (sequence capture), RP1 (baseline), RF1 (budget case and counter), RF2, RF7, RF3, RF4, RF5;
   the B residuals RB1, RB2, RB3, RB4, RB10; and the evidence items RB12, RD1, RP2, RP3, RG1, RG3, RA2,
   RB9.

Each slice or item group is one change with its tests, authoritative-spec edits, documentation and one
`Changes/` fragment. Per change:

```powershell
.\Tools\Test-Changes.ps1 -Explain
.\Tools\Run-AllTests.ps1 -Suite FileOps -CaseFilter <family> -TestRoot <approved X:\RedSalamander.Perf root>
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter <prefix> -TestRoot <approved X:\RedSalamander.Perf root>
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
.\Tools\Test-Changelog.ps1
.\Tools\Test-Changes.ps1 -Mode PrePush
```

Rules that apply to every change:

- Build Debug for every project before calling a change mergeable; self-tests compile only in Debug. Zero
  warnings and zero errors. Build from PowerShell.
- Run whole FileOps families, not single cases, before PrePush. A family or affected pass is not
  repository qualification. A Full run happens only if the owner requests it or the changed behavior
  needs it. Slice G changes the worker-pause rule, so ask the owner for a Full FileOps and Full Commands
  run before G is called mergeable.
- Tests stay in the governed no-activation mode. Any focus-taking witness needs the owner's agreement to
  the time.
- Test output stays under the approved `X:\RedSalamander.Perf` root. A case that needs a second volume is
  run with `-AllowAlternateVolumeTestRoot`; the I63 log records that the second-volume case
  `R4A19_DiscoveryIndependentVolumes` fails without it, as on earlier runs.
- A FileOps case must be registered in `kFileOpsFamilyDefinitions`, otherwise a family run skips it
  silently.
- New user-visible strings go to `RedSalamander.rc` with positional placeholders and to the four
  satellites.
- Screenshots for specs and documentation come from the harness with invented fixture data.
- A failed run is recorded with its cause; it is not replaced by a later pass.
- Before judging a failing Fresh Full, check `git status`: another session's uncommitted edits can fail
  it.

## 8. Done

Move this plan to `Specs/Plans/Done/` only when:

- slice G is merged with its tests and its target contract is in the authoritative specs, with the
  first-version limits stated there, or the owner has chosen a wider scope (G-D3 (b), G-D5 (b)) and that
  scope is merged;
- every item of section 5 is done with a test or an archived record, declined by the owner with the
  reason recorded, or routed to a named owner;
- every open decision of section 3 is recorded as accepted, changed or declined;
- the after-evidence for slice G (the 100-item scenario's elapsed time against the same run answered at
  the first collision, the retained bytes at the bound, a Full FileOps and Full Commands run) and for
  RF1, RB12, RD1, RP2 and RP3 is archived under `Specs/TestRuns/` with its build receipt, including
  recorded skips for environments that were not available;
- a successor plan is named for the provider-owned walks (G-D5) and for undo (I63 D-15), or the owner has
  declined them;
- no requirement of this plan or of I63 lives only in this file;
- the exact Done path is admitted in the protected-history allowlist with its focused test.

## 9. STOP

Stop and return to the owner when any of these holds. The first group concerns slice G, the second slice
F, the third the B residuals and the rest of Part 2.

- Slice G would hold a handle or an authority for a set-aside item, pass its bound silently, or remove a
  folder with an unresolved descendant.
- Slice G would change the buttons of a prompt the user can already press, other than by the preemption
  of G-D2 (a) after the owner accepted it, or other than the count, the names and the rest-finished line
  under one generation.
- The power request would be held while every task is finished, paused by the user or only waiting for an
  answer.
- A queue change can leave a queue place behind a finished, failed or cancelled task, or lets Start now
  bypass a capacity or safety interlock. A second ordering key would be introduced beside
  `_queueOrderKey`, or the edit would conflict with an unmerged I60 Wave 5 change in the same function.
- Height motion would let body rows paint over the bottom block, or the 16 ms cadence would run outside a
  bounded motion.
- A fix needs a DxUi change or pin bump.
- A comparison of names, sizes or dates would be shown or recorded as identity or as equal content, or
  would authorize a step without a typed decision or a policy the user chose.
- Level 2 would hold a handle on the occupant across the transfer, create a set-aside copy, or fall back
  to writing over a file whose replacing rename was refused. That is more, or less, than D-4 asks.
- A link, a reparse point or a type mismatch would be replaced at level 2 or 3, or a folder would be
  replaced instead of merged.
- RP3 shows the replacing rename is not atomic on the registered share (a reader sees a missing name or
  partial content). Keep B1 and return level 2 to the owner.
- A destructive step would run on an item whose destination already existed with facts equal to the
  source's without the B8 probe, or after an inconclusive probe.
- The Permanent Delete by-name confirmation would be weakened. B7 removes only a Move's own source, after
  the publication proof; B4 removes a source only on its own explicit choice.
- A finding contradicts the root cause or the code state this plan states for its item. Correct the plan
  first.
- Slice G, once it lands, makes a previously green FileOps or Commands family fail for a reason other
  than a test that pinned an immediate nested prompt (G0.6).

## 10. Out of scope

Provider-owned walks and selected-root files for collected collisions (G-D5 successor); the queue yield
of a running task with only questions left, unless the owner chooses G-D3 (b); a Managed Move in the
sequential walker (G4.4); collection of problems other than name collisions (D-12); replacement without a
decision or a policy the user chose, and the `fileOverwrite` preference (I18); the start-surface control
for the task collision policy (I18); undo (I63 D-15); persistent history and H3's other rows; replacement
of directories, links or reparse points at identity-less locations; rename-the-existing-item choices;
content comparison of the two items, in the prompt or for any policy (D-16); conditional bulk policies
other than B6's two and B4.5's; legacy staging-name recognition (D-17); remote, S3 and MTP staging names;
a new scrollbar control; queue persistence across restarts; the exclusive-transfer starvation and pause
slot yield owned by I60 `W5.2` and `W5.3`; any change to DxUi.
