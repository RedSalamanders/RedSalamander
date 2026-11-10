# I62 File Operations UX fidelity and pane activation click

Status: **ACTIVE, I62 — plan prepared; implementation not started.**

Priority: **P1**. Effort: **L**, delivered as six independent slices (A–F) plus a shared
characterization slice (P0). Implementation risk: **High** for slice B (replacement authority on a
location without object identity) and slice E (queue admission), **Medium** for F, **Low** for A, C, D.

Planned at `99624c5385c6c0d42f8d1c6f8be54cbe69f506ef`, 2026-11-10.

This is a **non-normative implementation plan**. Current contracts stay in force until each slice lands
its code, tests and authoritative-spec edits together. Every "Target contract" block below is proposed
normative text for the named spec; it is not authority while it lives here.

## 0. Drift check

Run from the repository root before any edit, and again before each slice:

```powershell
git diff --stat 99624c5385c6c0d42f8d1c6f8be54cbe69f506ef..HEAD -- RedSalamander Common Plugins Specs docs Tools Tests
git status --short
```

Anchors marked `†` were read directly at the planned commit while this plan was written. The other
anchors come from a delegated read-only survey of the same commit; re-read each one before relying on it,
and cite the current line when it moved. No build or test was run for this plan.

## 1. Origin and outcomes

The owner reported six problems on 2026-11-10 (one with a screenshot of a failed Move card). Each maps to
one slice. The screenshot's file name is personal data and is not reproduced; examples below use invented
names.

| # | Report | Verified root cause | Slice |
|---|---|---|---|
| 1 | A click in the pane that is not active destroys the selection that pane had | `FolderView::OnLButtonDown` applies the active-pane selection rules to every press; a plain press replaces the selection and a background press clears it (`FolderView.Interaction.cpp:502-559` †) | A |
| 2 | A failed Move card gives no size or date and no Skip, Replace or source choice | A Managed Move onto an existing item at a location without object identity is rejected inside `prepareExactTransferGuard` before any conflict prompt exists (`FolderWindow.FileOperations.State.cpp:24091-24111` †, call site `:25464-25472` †) | B |
| 3 | A menu button opens a menu with one entry | Two menu buttons are shown whenever their overflow set is non-empty, with no count check: the completed card (`Popup.cpp:942-946` †) and the conflict prompt, whose production overflow set is always exactly one action | C |
| 4 | Staging names such as `clip.mp4.rs_tmp_<32 hex>_0` have no usable extension | One hand-rolled formatter appends `.rs_tmp_{:016X}{:016X}_{:X}` to the destination path (`State.cpp:16443-16478` †) | D |
| 5 | With many queued tasks, the task that starts is not the next card | A task joins the engine queue only inside `EnterOperation`, after its own concurrent Preparing phase (`State.Queue.cpp:2359-2374` †, `State.cpp:12695` and `:12827` †); the popup lists it as waiting from admission. The card sort also mixes two key domains (`Popup.cpp:5552-5564` †) | E |
| 6 | The card list flickers and changes size continuously | An expanded card's height snaps on every content change by rule (`Popup.cpp:4855-4857` †), row counts and stacked facts have no hysteresis, numerals are proportional, and the native scrollbar and window resize lag the snapped content (`Popup.cpp:9671-9684` †, `:6939-6943` †) | F |

Outcomes when this plan is done:

1. The first plain click in the other pane activates it, moves the cursor there and keeps the selection.
2. A name collision always becomes a decision that shows both items' facts and offers every choice the
   route can perform safely, including Replace on a location without object identity and a source-side
   choice for Move.
3. No menu button ever opens a menu with fewer than two commands.
4. Every staging artifact created on a Win32 path ends in `.rs-tmp` (incomplete) or `.rs-bak` (may hold
   an original).
5. Queued tasks start in the order of their cards, and a card that is passed over says why.
6. The card list changes geometry rarely, in bounded smooth steps, with measured stability budgets.

## 2. Ownership and coordination

This plan owns the six slices and the improvements listed in section 11. It does not reopen any Done
record.

- **I18** (confirmation preferences and outcomes) owns start confirmation, the `fileOverwrite` setting
  and metadata-outcome policy. Slice B adds decision actions and authority for one route class; it adds
  no setting and no automatic replacement. I18 section 5.4 ("reduced safe action set") must be
  reconciled with B2 when either lands second.
- **I60 Wave 5** (`W5.2` queue lifecycle state and `_queueTicket`, `W5.3` pause slot yield) edits the same
  `EnterOperation` region as slice E. Whichever lands second rebases. Slice E must not introduce a second
  ordering key: `_queueOrderKey` is the ticket. If `W5.2` lands first, implement E1 on its
  `QueueLifecycleState`; if E lands first, `W5.2` keeps E's admission-time queue place and tests.
- **H3** (later features) keeps "review conflicts later". Slice B does not park conflicts.
- **I57** keeps the interrupted-notice task-id lock. At the planned commit that increment is already
  under `_mutex` (survey: `State.Runtime.cpp:706-709`); slice E does not touch it.
- **I3**, **I29** keep their stress and qualification scopes. **I4** keeps the source-shape test
  migration; this plan adds behavior tests and updates only the source contracts its edits invalidate.
- DxUi is an external pinned dependency. No slice requires a DxUi pin bump. If tabular figures (F1) or a
  one-entry-menu guard needs a DxUi change, stop and route it to the DxUi repository and **I31**.

## 3. Owner decisions

Recorded:

- **D-1 (owner, 2026-11-10).** Scenario: the left pane is active, the right pane has three files
  selected, the user clicks another file X in the right pane once. Required result: the right pane
  becomes active, the cursor moves to X, and the three files stay selected. A second click selects
  normally. Keyboard modifiers keep their meaning (Ctrl adds or toggles, the others "as expected"). The
  behavior must be easy to understand and coherent. Observed today: the selection is lost.

Open, each with a recommendation. A slice that depends on an open decision implements the recommended
option only after the owner accepts it; everything else in the slice proceeds.

| ID | Question | Recommendation |
|---|---|---|
| D-2 | Is a press that activates the application window on the *already active* pane an activating press? | **No.** The rule depends only on which pane is active, so one sentence describes it. `_paneFocused` already tracks the active pane while the application is in the background (`GetPaneFromChild(nullptr)` returns `_activePane`, `FolderWindow.Interaction.cpp:797-802` †). |
| D-3 | An activating press lands on an unselected item while other items are selected, then the pointer leaves the drag threshold. | **Drag the pressed item alone**: replace the selection with it when the drag starts. The alternative (no drag, per current rule in `UI_FolderView.md:507` †) makes a cross-pane drag silently do nothing. |
| D-4 | May a regular file at an identity-less Local location be replaced after a typed Replace decision, using handle-verified corroboration facts and a set-aside backup (B2)? | **Yes.** It amends product invariant 14 and the replacement-authority rules; see B2 for the protocol and its limits. Without it, B1 alone still delivers Keep both, Skip and Cancel. |
| D-5 | Add a Move-only collision choice that keeps the existing destination and removes the source (B4)? | **Yes, files only in the first version**, Recycle Bin when the source location has one, otherwise through the existing Permanent Delete confirmation; never a default, never Apply to all. |
| D-6 | `Delete all remaining permanently` is today the only entry of a `More...` menu and the spec forbids a direct button. | **Direct button whose caption ends with an ellipsis and opens the existing permanent-delete confirmation step.** No one-entry menu, and the second deliberate step is kept. |
| D-7 | While the head of the queue waits for the user's answer on its own card, may the next queued task start? | **Yes.** Machine Preparing holds the queue place; waiting for a person does not. The passed-over card states the reason and starts next once answered. |
| D-8 | Card order in the popup. | **One stable key for every active card** (`_queueOrderKey`, then task id). Cards never move when a task starts; waiting cards show their queue position. |
| D-9 | Scope of the staging-name change. | **Bridge stage plus the Local provider's stage, backup and writer names** in one slice, so every Win32-visible artifact groups under two extensions. Remote, S3 and MTP names keep their dot-hidden forms. |
| D-10 | Add Explorer's "skip items with the same size and date" as a task-scoped bulk choice (B6)? | **Yes, as the last B step**, after B1–B5 are qualified. It may be deferred without blocking Done. |

## 4. Slice P0 — characterization and baselines (no behavior change)

Purpose: turn each survey claim into a witnessed fact and capture the "before" evidence the later
slices compare against. P0 changes tests, test hooks and evidence only.

- [ ] **P0.1 Pane press.** Add a failing-first case beside
  `cmd_pane_navigationView_unfocused_pane_click_focuses_target_pane`
  (`SelfTest/Commands/Commands.Tests.ViewCommands.cpp`) that focuses the left pane, selects three items
  in the right pane, sends a plain press to a fourth right-pane item with the existing
  `SendMouseClickToResolvedPointWindow` helper, pumps posted messages, and records today's result
  (selection collapses to the pressed item). Record the background-press result (selection cleared).
- [ ] **P0.2 Parent-notify trace.** Under `ENABLE_TESTS`, count `FolderWindow::OnParentNotify` entries
  whose `childId` matches a pane ID for a real button-down (`FolderWindow.Interaction.cpp:942-965` †,
  dispatch `FolderWindow.cpp:1363` †). `HIWORD(wParam)` carries a child ID only for `WM_CREATE` and
  `WM_DESTROY`, so the expected count is zero. This decides S1.
- [ ] **P0.3 Identity-less collisions.** With the existing injection
  `REDSALAMANDER_FILEOPS_BIND_IDENTITY_UNSUPPORTED_PATH`, and on the registered network test root
  `path.smb.primary` where the environment provides it, record for an existing destination: single-item
  Move, multi-item Move (does the task continue after the refused item, and what do the remaining items
  report), Copy (expected: a prompt without Replace, because `allowDestructiveReplacement` needs a bound
  object or an atomic-final writer, `State.cpp:9795-9865` †), folder onto folder, file onto folder, and a
  read-only occupant. Record the `Errors` count and both diagnostics for the single-item case.
- [ ] **P0.4 Refusal inventory.** List every exit of `prepareExactTransferGuard` and every
  `TransferSafetyState` that ends an item without a decision. Classify each as *must refuse* (same
  object, self-descendant, missing source, unsupported source binding) or *has a safe choice*.
- [ ] **P0.5 Local replace route.** Read how the bound Replace route publishes today (stage, backup
  `.rs_bak_`, publish, cleanup; `Plugins/FileSystem/FileSystem.cpp` near the backup at `:4670` and the
  exclusive writer at `:7626`) and state whether it can accept an occupant verified on its own handle
  instead of a `FileIdInfo` identity. This fixes the B2 implementation shape before any code.
- [ ] **P0.6 Menu census.** Under `ENABLE_TESTS`, record the entry count of every popup flyout opened by
  the visual gallery and the Commands popup cases. Expected today: one for every conflict `More...`,
  one for `Open destination`-only and `Open source`-only completed cards, zero for a completed card
  whose only overflow reason is a suppressed diagnostic.
- [ ] **P0.7 Queue order.** Add a deterministic Preparing pause point (same pattern as the existing
  `g_fileOps...PausePoint` hooks) and a failing-first case: six queued tasks, the later ones prepare
  faster; record the start order against the card order. Record the card index of each task across its
  waiting-to-running transition, and the result of Move Up across an overlap-queued neighbour.
- [ ] **P0.8 Layout stability baseline.** Add the stability counters of F4 (counters only) and a
  twelve-second synthetic scenario with one to eight streams churning and rate text changing every tick.
  Archive the baseline: card height changes, layout shift, scrollbar toggles, window resizes, and the
  existing paint timings.
- [ ] **P0.9 Staging-name perf baseline.** Run the existing artifact name-shape projection measurement
  (`DebugMeasureArtifactNameShapeProjectionForTests`) and archive the result; slice D compares to it.
- [ ] **P0.10 Test reliance audit.** List the cases that press an item in a pane without focusing it
  first, the three leftover scanners that search for the literal `.rs_tmp_`, and the cases that pin
  `conflictPromptCount == 0u`, `overflowActionCount == 1u` and the `More...` button; these are the
  expected edits of slices A, D, B and C.

Verify: the new cases are registered (a FileOps case must be in `kFileOpsFamilyDefinitions`, otherwise a
family run skips it silently), fail or record as stated on the planned commit, and the evidence is
archived under `Specs/TestRuns/` with its build receipt.

## 5. Slice A — the activating press

### Current behavior

`FolderView::OnLButtonDown` † calls `SetFocus` and then applies one rule set regardless of whether the
pane was active: plain press on an unselected item `SelectSingle`, on a selected item `FocusItem`, on
background `ClearSelection`. There is no handler for `WM_MOUSEACTIVATE` or `WM_MBUTTONDOWN`
(`FolderView.cpp:784-869` †). The pane's active state reaches the view through the posted
`kPaneFocusChanged` and `SetPaneFocused` (`FolderWindow.Interaction.cpp:36-62` †, `FolderView.cpp:55-67` †).
No spec sentence covers a press in the pane that is not active (`UI_FolderView.md:497-508` †).

### Target contract (for `UI_FolderView.md`, pointer target semantics)

> An **activating press** is a left-button-down delivered to a FolderView whose pane is not the active
> pane. It makes that pane active. A press that only brings the application window to the foreground on
> the pane that is already active is not an activating press.
>
> - A plain activating press on an item makes that item current and sets the range anchor to it. It MUST
>   NOT change the selection, whether or not the item is selected.
> - A plain activating press on empty background MUST NOT change the selection, the current item or the
>   range anchor.
> - Ctrl, Shift and Ctrl+Shift activating presses apply exactly the active-pane rules: they state a
>   selection intent explicitly.
> - A double-click whose first press is activating activates the item under the pointer.
> - Right-button and context-menu rules do not depend on the pane's prior state: a right-click always
>   targets what is under the pointer.
> - A middle-button press makes the pane active and changes nothing else.
> - When an activating press lands on an item outside the current selection and the pointer then leaves
>   the system drag threshold, the selection is replaced by the pressed item and that item is dragged.
>   Otherwise the existing drag-source rule applies. (D-3)

The one-sentence user rule, for `docs/SelectionsAndLists.md` and `docs/MainWindow.md`: *the first click
in the other pane takes you there and moves the cursor; it never throws away a selection. Ctrl, Shift
and right-click work the same in both panes.*

### Implementation

- [ ] A1. In `OnLButtonDown`, read `const bool activating = ! _paneFocused;` before `SetFocus`. Keep the
  alert-overlay early returns unchanged. Plain item press: `activating ? FocusItem(hit, false) : existing
  branch`; plain background press: skip `ClearSelection` and the anchor reset when `activating`.
  Modifier branches are untouched.
- [ ] A2. Drag (D-3): record `activating` and the pressed index in `_drag`. For an activating press only,
  arm the potential drag even when the pressed item is outside the selection; when the threshold is
  crossed in that state, call
  `SelectSingle(pressed)` and rebuild the immutable source snapshot before starting the drag. Release
  inside the threshold changes nothing further.
- [ ] A3. Add `WM_MBUTTONDOWN` to `FolderView::WndProc`: `SetFocus` and nothing else.
- [ ] A4. Activation ownership. FolderView and NavigationView already take focus themselves. If P0.2
  shows the pane-ID comparison in `OnParentNotify` never matches a mouse event, delete that branch and
  its dispatch (S1) instead of repairing it; otherwise keep it and document what it serves.
- [ ] A5. Staleness. `_paneFocused` is updated by a posted message that the queue delivers before the
  next input message. Add a debug assertion in the test build that `_paneFocused` equals
  `GetFocusedPane() == thisPane` at button-down, and cover the Terminal tab note in
  `FolderWindow.Layout.cpp:987` † (its `WM_SETFOCUS` posts no pane-focus change) with a case.
- [ ] A6. Focus-follows-pointer needs no change: with it enabled the pane is already active when the
  press arrives.

### Tests (Commands suite, prefix `cmd_pane_folderView_activation_press_`)

Plain press on an unselected item with three selected (three stay selected, pressed item current and
unselected, pane active, `GetActivePane` right); plain press on a selected item; background press;
Ctrl, Shift and Ctrl+Shift presses equal to the active-pane results; the second plain press collapses
the selection; double-click on a folder navigates; right-click on an unselected item makes it the sole
target; middle press activates only; drag from an unselected item (D-3) and from a selected item;
press while the same pane's NavigationView holds focus is *not* activating; Quick Search active in the
other pane exits as today. Update the cases found in P0.10. All cases use the governed no-activation
mode; a real-input witness of the application-activation ordering (D-2) is focus-taking and needs the
owner's agreement to the time.

Performance: the press path gains one branch and no I/O; the existing input-to-paint measurement
(`RecordInputToPaintStartIfViewportOrFocusChanged`) covers it. No new scenario.

## 6. Slice B — collisions become decisions with facts and real choices

### Current behavior

A Move from a Local disk to a share is a Managed Move. Its per-item guard binds the destination;
`BindObject` on a redirector without `FileIdInfo` returns `ERROR_NOT_SUPPORTED`, the guard state is
`Unsupported`, and none of the proceed-without-guard exceptions covers an existing destination †. The
item fails with `identity.destinationExistsUnverifiable` plus the trailing `task.result` row, which is
why a one-item Move shows `Errors: 2`. The conflict engine is never entered; the selftest
`RunFileOperationUnverifiableDestinationCollisionSelfTest` pins `conflictPromptCount == 0u`. This refusal
is normative today (`FileSystem_FileOperations.md:57-65` †, product invariant 14;
`UI_FileOperationsPopup.md:1600-1612`).

When a prompt *is* reached, the engine already offers Keep both, Replace, Skip, Cancel, an Apply-to-all
scope and Skip all, and the prompt shows `size • date` for both items. Gaps: Replace is withheld on any
identity-less Local destination for Copy as well, the two fact lines are not compared for the user, a
provider route that reports no size shows none, and Move has no source-side choice.

Windows Explorer parity, as the owner asked for:

| Explorer | RedSalamander at the planned commit | This plan |
|---|---|---|
| Replace the file in the destination | Only with a bound object or an atomic-final writer; never on an identity-less Local location | B2 |
| Skip this file | Yes, when a prompt exists | B1 makes the prompt exist |
| Compare info for both files, keep both | `Keep both` renames the incoming item; facts are two raw lines | B1, B3 |
| Do this for all current items | Apply-to-all scope; `Skip all` hidden in a one-entry menu | C2 |
| Skip items with the same date and size | Absent | B6 (D-10) |
| Size and date of both files visible | Only inside a live prompt; a failed card shows none | B3 |
| — (Move) keep the existing file and drop the source | Absent | B4 (D-5) |

### B1 — the refusal becomes a typed collision decision

Target contract (replaces product invariant 14 in `FileSystem_FileOperations.md`):

> A transfer that finds an existing item at a destination whose object binding is `Unsupported` raises
> the ordinary typed collision decision. The existence read establishes the collision only; it
> authorizes no overwrite, source cleanup, or identity or equality claim. `Keep both`, `Skip` and `Cancel`
> are always offered; a destructive action is offered only under B2. A context that cannot ask keeps the
> pre-mutation refusal and its `identity.destinationExistsUnverifiable` diagnostic. Missing destinations,
> failed probes, and unsupported source or ancestor bindings keep their original classification.

- [ ] B1.1 In `prepareExactTransferGuard`, replace the terminal return for *Move, `Unsupported`,
  destination identity unsupported, existence read succeeded* with a distinct result that the per-item
  loop turns into `BeginConflictPrompt` through the same policy path provider-issued collisions use. The
  bucket comes from `RefineTypedConflictBucket` on the read metadata (regular file, read-only file, type
  mismatch, link).
- [ ] B1.2 `Keep both` re-enters the loop through the existing `keepBothDestinationOverride`; the fresh
  sibling name is missing, so the guard is `Ready` and the ordinary new-file route runs. `Skip` records a
  user skip (typed result, source retained, not an Error). `Cancel` cancels the task.
- [ ] B1.3 Keep `FileOps.Conflict.DestinationIdentityUnavailable`; add a result dimension for the
  decision taken. Keep the resource string for the cannot-ask path.
- [ ] B1.4 Apply the P0.4 inventory: every *has a safe choice* exit raises a decision with at least
  Skip and Cancel; every *must refuse* exit keeps its refusal and gains a specific sentence where it
  still shows only a Windows status text.

### B2 — Replace at an identity-less Local location (D-4)

Principle: object identity proves that the item replaced is the item the user saw. Where the location
cannot provide it, the same proof is taken on a handle, and the original is kept until the new content
is proven.

Target contract (extends "Conflict model" and the identity vocabulary in `FileSystem_FileOperations.md`):

> **Corroborated replacement.** At an identity-less Local location, a regular file may be replaced after
> a typed Replace decision. The decision retains the occupant's corroboration facts as shown to the user:
> kind, compared attributes, size, last-write time, creation time, and the file index with the volume
> serial only when nonzero. The provider then (1) opens the occupant without following links, denying
> other writers and deleters; (2) re-reads the facts on that handle and stops with a fresh decision when
> any differs; (3) renames the occupant by that handle to a `.rs-bak` sibling; (4) publishes the incoming
> item through the ordinary new-item route, which refuses a name that reappeared; (5) removes the backup
> by its handle only after the publication proof the route already requires before source cleanup. A
> failure or cancellation after step 3 renames the backup back; if that fails, the result names the
> retained backup and the source is kept. A link, a directory, a reparse point of any kind, and a
> location that refuses rename by handle are never replaced this way. The route claims no identity and no
> content equality.

- [ ] B2.1 Host: a third eligibility beside `exactDestination` and `atomicReplaceEligible` in
  `Task::FileSystemIssue` and in the B1 prompt: Local endpoint, bind `Unsupported`, occupant a regular
  file. Capture the facts when the prompt reads its metadata.
- [ ] B2.2 Receipt: a distinct authority kind (corroborated occupant) carried by the decision; the
  existing rule "no destructive receipt without a typed user decision" is unchanged. A cached
  Apply-to-all Replace re-reads the facts for each item immediately before its own step 1.
- [ ] B2.3 Provider: implement steps 1–5 in the Local provider in the shape P0.5 selected, reusing its
  backup naming and the artifact registry's "possible original" class. Interfaces change in source-tree
  lockstep with size and version checks at the boundary.
- [ ] B2.4 One Information note per task: the location cannot identify files, so replacement was
  checked by name, size and dates.
- [ ] B2.5 Read-only occupant: only through the existing `Replace read-only` action and its wording.
- [ ] B2.6 Run the repository security review on the change before merge: occupant swap between prompt
  and handle, ancestor replaced by a link, backup-name pre-creation, crash between steps 3 and 5 (the
  backup is then a visible "possible original" leftover, never deleted automatically).

### B3 — decision facts

Target contract (for "Conflict Layout" in `UI_FileOperationsPopup.md`):

> The engine publishes, with the prompt, a comparison of the two items when both facts are known: date
> relation (newer, older, same within the coarser of the two locations' timestamp precision) and size
> relation (larger, smaller, same). Each role line appends its relation words; equal size and date render
> one line, `Same size and date`, and never the word identical. Unknown facts are omitted, not guessed.
> The exact byte count and the full date and time are in the tooltip and the accessible description. The
> same two fact lines are retained on the Issues row of an item that ended skipped or failed at a
> collision, and on the finished card of a single-item task.

- [ ] B3.1 Add the comparison to `ConflictPromptState` beside `ItemMetadata`; the popup and the pane
  decision mirror render it and derive nothing themselves.
- [ ] B3.2 Fill size on provider routes: `ReadConflictItemMetadata` falls back from
  `GetFileBasicInformation` to the provider's item-information call that carries size (name it in P0.3).
- [ ] B3.3 Retain the facts in the typed item result so the Issues pane and the finished card can show
  them; bounded by the existing reservoirs.
- [ ] B3.4 New `.rc` strings with positional placeholders, in all four satellites.

### B4 — Move: keep the existing destination, remove the source (D-5)

- [ ] B4.1 New `ConflictAction::DiscardSource`, offered only for a Move collision whose incoming and
  existing items are regular files. Caption `Delete source`; with D-6 styling rules: never the default,
  never the Enter or Escape action, not eligible for Apply to all.
- [ ] B4.2 Execution reuses the Delete route on the source's exact bound object: Recycle Bin where the
  source location has one; otherwise the existing Permanent Delete confirmation on the card. A changed
  source is refused and re-asks.
- [ ] B4.3 Typed result `source removed, existing destination kept`; the task is not Partial for it.
  When the two items differ in size, the prompt's problem sentence says so before the choice.
- [ ] B4.4 `FileSystemIssueAction` (the provider ABI) is unchanged: this action is host-executed.

### B5 — honest counts

- [ ] B5.1 The `Errors` chip counts item problems. The trailing `task.result` summary row is not counted
  when at least one item-level Error of that task explains the failure. Check the counter's contract in
  `UI_FileOperationsPopup.md` first; if it defines the chip as a diagnostic-row count, change the spec
  in the same edit.

### B6 — skip items with the same size and date (D-10)

- [ ] B6.1 A task-scoped choice in the collision prompt's scope controls: skip every later collision
  whose comparison is `Same size and date`. It is a Skip policy only; it never replaces and never
  deletes, and each skipped item is a typed user skip with its facts retained.

### Tests (FileOps suite; register every case in `kFileOpsFamilyDefinitions`)

Rewrite the five variants of the unverifiable-collision selftest: prompt count one; Keep both publishes
`name (2).ext` and removes the source; Skip retains both byte streams and reports a user skip; Cancel
mutates nothing; the cannot-ask path keeps the refusal. New `FileOps_CorroboratedReplace` family: happy
path for Copy and Move; occupant changed between prompt and handle (size, time, attributes, replaced by
a link, replaced by a directory); name reappears between steps 3 and 4; publication failure restores the
backup; restore failure retains the backup and the source; cancellation at each step; read-only
occupant; Apply-to-all over ten collisions with one changed occupant. `DiscardSource`: recyclable source,
non-recyclable source with confirmation, changed source, differing sizes. Comparison cues: newer, older,
same within two seconds, unknown size. Update the policy coverage selftest, the visual gallery fixtures
that build the failed card, and the Commands popup metadata cases. Where the environment provides
`path.smb.primary`, run the family there as well and record skips honestly.

Performance: the collision path adds one metadata read per collision and, for B2, one rename and one
delete per replaced file. Add `FileOps.Conflict.CorroboratedReplaceUs` and archive a 500-collision
Apply-to-all run next to the same run with Keep both.

## 7. Slice C — no menu with fewer than two commands

Target contract (for `UI_FileOperationsPopup.md`, "Button placement" and "Completed Card Actions"):

> A menu button is shown only when its menu will contain at least two enabled commands. A single command
> is shown as its own direct button with its own caption, standard variant, and the same command route.
> No command means no button. The button and its menu are built from one list.

- [ ] C1. Completed card. One builder returns the ordered completed-overflow command list for a task
  snapshot. `PrepareTerminalActions`, `ShowCompletedOverflowMenu`, the layout snapshot fields and the
  self-test hook consume that list; `CompletedTaskHasOverflowActions` is deleted. This also removes the
  zero-entry flyout that the two diverging predicates allow (`Popup.cpp:942-946` † against the menu
  builder's `hasDiagnostics`). A single command uses the ordinary 24-DIP caption padding.
- [ ] C2. Conflict prompt. Placement stays engine-owned in `ConflictActionPolicy::FinalizePresentation`:
  an overflow set of one is published as a primary action. Resulting direct orders: collision
  `Keep both, Replace, Skip, Skip all, Cancel` for Copy, with `Delete source` before `Skip` for Move
  (B4); `Delete all remaining permanently...` per D-6. The broker's placement validation follows the
  published placement, so the popup never submits an overflow action as primary.
- [ ] C3. The pane decision mirror consumes the same policy and needs no rule of its own.
- [ ] C4. Guard: in test builds every flyout records its entry count; a new case fails when any popup
  flyout opens with fewer than two entries across the visual gallery scenarios.
- [ ] C5. Accessibility: a direct button exposes its own name and Invoke; its automation ID keeps the
  `FileOperations.Action.<kind>.<taskId>.<data>` shape with a distinct kind for a direct completed
  action. Tab order stays visual order.

Spec lines to amend: `UI_FileOperationsPopup.md` 1834-1850 †, 2261-2265, 2314-2318, 2341-2342,
2379-2383, the mirror rows near 734-735, and the interrupted-Move card near 1741-1745 (its only
command, `Open destination`, becomes a direct button); `FileSystem_FileOperations.md` 1397-1408.
Tests to update: the cases listed by P0.10 (`overflowActionCount == 1u`, `conflictMoreVisible`,
`completedDiagnosticsMoreVisible`, the `More...` UIA lookups) and the visual gallery baselines.

## 8. Slice D — staging names with a real extension

### Target contract (for "Artifact warnings and retained state" in `FileSystem_FileOperations.md`)

> A staging artifact created on a Win32 path is named `<final name>.<32 hex>[_<stream>].rs-tmp` when it is
> incomplete content, and `<final name>.<32 hex>.rs-bak` when it may hold the only copy of an original.
> The token is 128 bits from the system generator; a failure of the generator fails the operation. The
> leaf never exceeds 255 UTF-16 units: only the readable prefix is shortened, never the token or the
> extension. Leftover recognition accepts these shapes and every shape an earlier version wrote. A name
> that merely ends in `.rs-tmp` or `.rs-bak` without the token is an ordinary file.

Example: `harbor-timelapse.mp4.60d51a8a47eba25f7f7b699fb5137228_0.rs-tmp`.

- [ ] D1. Bridge stage. `MakeTempDestinationPath` † builds the leaf with
  `Common::Paths::BuildUniqueSiblingName` (`Common/Helpers.h:354-393` †): prefix the final leaf, marker
  `.`, suffix `_<stream>.rs-tmp`, limit 255. This removes the hand-rolled formatter and fixes a present
  defect: the formatter adds at least 42 units with no bound, so a final leaf longer than 213 units
  yields a stage leaf above 255 and fails stage creation. Keep the self-test entropy-failure hook.
- [ ] D2. Local provider (D-9): copy stage `.rs_copy_tmp_<hex>` becomes `.<hex>.rs-tmp`, overwrite backup
  `.rs_bak_<hex>` becomes `.<hex>.rs-bak`, writer stage `.~rs-write-<hex>.tmp` becomes `.<hex>.rs-tmp`
  (its Hidden and Temporary attributes unchanged), through `MakeStageSiblingPath`. The case-only rename
  hop keeps its whole-leaf form and gains the `.rs-tmp` extension.
- [ ] D3. Recognition. Make `HasPossibleArtifactName` table-driven (S4): one table row per shape with
  its class. Add the new suffix-anchored shapes with a cheap `ends_with` pre-check. Keep every legacy
  row. Close two present gaps with tests: the current writer shape `<32 hex>.tmp` is not matched by
  `IsWriterTail` (`FileOperationArtifactRegistry.cpp:94-102` † against the generator at
  `FileSystem.cpp:6599` †), and the bridge shape used before June 2026 (`PID_TID_tick_stream`) is not
  matched by `IsBridgeTail`. Preserve the textual shapes the Pester source contracts pin (function
  signature, `.rs_bak_` before `.rs_ren_`).
- [ ] D4. Tests. Convert the three literal scanners found by P0.10 to the registry predicate; a scanner
  that searches for `.rs_tmp_` would pass silently once the name changes. Extend the shape table with
  new, legacy and near-miss names (`report.rs-tmp`, a 31-hex token, a non-hex token, upper-case hex).
  Update the planted fixtures. Add a bridge case with a 251-unit leaf (stage leaf exactly 255, publish
  succeeds) and a case with a surrogate pair at the cut. Cleanup-dialog cases cover both generations.
- [ ] D5. Documentation and specs: `FileSystem_FileOperations.md` (the recognized-shape list near 2883,
  the Local leaf bound near 1136-1146), `Core_Search.md` near 226, `Core_FileSystemBridge.md` near 290,
  `Plugins_VirtualFileSystem.md` near 1498, `docs/FileOperations.md`, `docs/Plugins.md`,
  `docs/dev/PluginHostModel.md`, `docs/dev/FileOperationsEngine.md`, `docs/release-note/vLast.md`. Remove
  two stale statements while there: the "falls back to PID/TID/tick" claim in the two developer documents
  and the `.rsren-` claim in `docs/BatchRename.md` and `docs/DeveloperGuide.md`.

Compatibility: no persisted store records a staging name, so no schema changes. An older build does not
badge the new names; state this in the release note. Performance: recognition runs per row during
FolderView enumeration; compare against P0.9 and archive the result. A regression there is a STOP.

## 9. Slice E — queued tasks start in card order

### Current behavior

Each task runs `PrepareForExecution` on its own thread with no ordering between tasks, then enters
`EnterOperation`, which is the only place that inserts it into `_queue` †. The readiness test admits
the first *present* queue entry without an overlap predecessor †. An earlier card that is still
Preparing has no place, so a later task that prepared faster starts first. Separately, the card sort
uses `_queueOrderKey` for waiting tasks and the task ID for the others †, so a card can move when its
task starts after a reorder or a late key; and Move Up or Down swaps with any adjacent waiting task,
including one held by an overlap edge (survey: `State.Runtime.cpp:2869-2937`, `Popup.cpp:5736-5740`).
Wakeups are `notify_all`, but the predicate names a single winner, so wake order is not the cause.

### Target contract (for "Overlapping-operation interlock" and the queue rules in both specs)

> In Queue mode a task takes its queue place when its card is created and keeps it through Preparing.
> Tasks start in queue order, which is the order of the waiting cards. A later task MAY start before an
> earlier one only when the earlier task (a) was cancelled, withdrawn or failed; (b) is waiting for the
> user's answer on its own card; (c) is queued behind a named task by an overlap decision; or (d) was
> passed by the user's own Start now or reorder. In cases (b) and (c) the earlier card states its reason,
> and it starts next once the reason clears. A task that is still Preparing holds the queue.
>
> Active cards are ordered by one key, the queue order key, then task ID. A card does not change position
> when its task starts. A waiting card shows its position (`Next`, then its ordinal). Move Up and Move
> Down are offered only for moves the engine will perform, and never across an overlap edge.

- [ ] E1. Queue place at admission. Where admission sets `_queueOrderKey` and `_waitForOthers` (survey:
  `State.Runtime.cpp:1448-1548`; batch rename `FolderWindow.FileOperations.cpp:3099-3137`), insert the
  task into `_queue` under `_queueMutex`. `EnterOperation` keeps the insertion as an idempotent
  fallback.
- [ ] E2. `firstUnblockedQueuedTaskId` walks `_queue` in order and classifies each entry: *holding* (no
  published prepared admission yet, not waiting for the user, not cancelled) stops the walk and returns
  no winner unless it is the caller; *waiting for the user* and *overlap predecessor* are skipped;
  *published and unblocked* wins.
- [ ] E3. Every exit removes the place and wakes waiters: prepare failure, withdrawal, cancellation,
  clipboard-Move admission refusal, mode switch to Parallel, Start now. Every transition into and out of
  *waiting for the user* wakes waiters. A missed removal stalls the whole queue; E-tests cover each exit.
- [ ] E4. Evaluate `ShouldQueueNewTask` inside the admission lock so two admissions issued together
  cannot both read "nothing active" (survey item; confirm in P0.7 before changing).
- [ ] E5. One comparator for `BuildSnapshot`: `(queue order key, task id)` for every active task; delete
  the dual-key branch. Publish the waiting ordinal in the snapshot and render it on the waiting chip.
- [ ] E6. Reorder. `MoveQueuedTask` and the popup's move-edge flags come from one engine function that
  skips neighbours that cannot move and refuses a swap that would invert an overlap edge.
- [ ] E7. Leave `UpdateQueuePausedTasks` and its entered-tick ordering unchanged unless P0.7 shows it
  contradicts the contract; if it does, order it by the same key.
- [ ] E8. Instrumentation: `FileOps.Queue.HeadHoldUs` (time the head held the queue while Preparing) and
  `FileOps.Queue.PassedOver` with its reason.

### Tests (new family case `FileOps_QueueStartOrder`, registered)

Six queued tasks with reversed Preparing durations start in card order; head waiting for a Permanent
Delete answer is passed and starts next after the answer; head cancelled while Preparing; head failing
Preparing; head queued after a named task; reorder followed by completion; Start now out of order; Queue
to Parallel and back with tasks Preparing; card index of every task constant across its start; waiting
ordinals correct after each event; Move Up across an overlap-queued neighbour refused. Keep the existing
reservation and interlock cases green (`DebugValidateSharedTransferAdmissionsForSelfTest`,
`ValidateQueuedAdmissionContracts`, `FileOps_DecisionMirrorQueueOrder`). Each case ends by asserting an
empty `_queue` and zero active operations.

Performance: archive queue wait metrics for a twenty-task queue before and after; the head hold must
equal the head's own Preparing time.

## 10. Slice F — a quiet card list

### Current behavior

One 100 ms timer invalidates the whole client and `Render` rebuilds the layout each time. Height comes
from the rows painted on that frame: `max(1, visible streams)` path rows, fact pairs that stack when a
width is crossed, and conditional notes †. `ResolveAnimatedTaskHeight` snaps an expanded card to every
new target † because `UI_FileOperationsPopup.md:2063-2067` † requires it, so that content never overlaps
while a shorter card interpolates. Numerals are proportional and the rate text changes every tick.
Stream rows are compacted in the engine, so a row re-binds to another file when an earlier stream ends.
The native scrollbar appears as soon as content exceeds the viewport † while the window grows only after
a debounce and an ease that is sampled at the 100 ms work cadence †; each toggle narrows every card.
There is no scroll anchoring, and card insertion and removal are instantaneous.

### Target contract (new section "Layout stability" in `UI_FileOperationsPopup.md`; amends 963-967, 1341, 1407-1413, 2063-2067 and "Motion")

> **Reservation.** Within one lifecycle phase a card reserves the rows it has needed. Stream rows are
> stable slots: a stream keeps its slot for its lifetime, and the reserved slot count is the high-water
> mark, released one row at a time after a slot has been unused for two seconds. A fact pair that stacked
> stays stacked until its facts have fitted with sixteen DIP to spare for two seconds. A conditional note
> stays reserved for two seconds after its condition ends. A phase change re-measures at once.
>
> **Numbers.** Numeric text uses tabular figures. Each numeric field reserves the width of its widest
> form. Rate and ETA text change at most twice per second; the graph keeps its cadence. A displayed file
> name stays at least 250 ms.
>
> **Height motion.** A card's content is laid out at its target height on the first frame. Its bottom
> block (progress and actions) and the cards below ease from the old position to the new one over 200 ms
> on whole device pixels; rows between are clipped at the top of the bottom block, so nothing overlaps.
> Decision content still snaps. Height motion is bounded UI motion and advances at the 16 ms cadence
> while it runs. Reduced motion snaps everything.
>
> **Viewport.** The list's scrollbar does not appear or disappear while an automatic resize is pending
> or running, and only after the overflow has lasted 300 ms; its appearance never changes card width.
> Automatic resize runs outside painting, grows at once, and shrinks only after two seconds of stable
> content. When content above the viewport changes height, the first visible card keeps its position.
> A card that appears expands in, and a card that leaves collapses, over 200 ms.
>
> **Budget.** In the stability scenario, unanimated layout shift is zero, a card changes height at most
> four times, the scrollbar toggles at most once, and the window resizes at most once after the first
> second.

### Implementation

- [ ] F1. Stream slots. `UpdateInFlightFileProgress` (survey: `State.cpp:2396-2516`) assigns the lowest
  free slot and never compacts; the snapshot carries slot indices; the popup reserves the high-water
  count with the two-second release. `FileProgressIsVisible` keeps its 300 ms linger for the row's
  content, not for its reservation.
- [ ] F2. Sticky facts. `PrepareFactPair` (survey: `Popup.cpp:7251-7297`) takes a reserved numeric width
  for all three pairs (progress, counts and ETA, throughput and rate), as the graph labels already do,
  and keeps `stacked` with the stated hysteresis. Conditional notes get the same hold.
- [ ] F3. Tabular figures on the numeric text layouts (`IDWriteTypography`, tabular-figures feature) in
  `EnsureTextFormats`; if the hosted DxUi typography owns the format, stop and route per section 2.
  Display hold for rate, ETA and the file name; `UpdateRates` runs once per 100 ms bucket even on 16 ms
  motion ticks (S10). The mini progress bar keeps its column when a row completes.
- [ ] F4. Counters and budget case. `FileOps.Popup.CardHeightChanges`, `FileOps.Popup.LayoutShiftPx`
  (sum of card-top movement not covered by a running height motion or a user action),
  `FileOps.Popup.ScrollBarToggles`, `FileOps.Popup.WindowResizes`. New case `FileOps_PopupLayoutStability`
  enforces the budget on the P0.8 scenario.
- [ ] F5. Height reveal. Replace the snap condition in `ResolveAnimatedTaskHeight` † with the reveal:
  layout at target, animated bottom block and following cards, row clip, integer pixels. Keep the snap
  for decision content (`Popup.cpp:9602-9605` †). In `RefreshTelemetryTimer` † an unsettled height motion
  selects the 16 ms cadence ahead of the 100 ms work cadence, for its bounded duration only.
- [ ] F6. Viewport. Scrollbar hysteresis and pending-resize awareness in the scroll pass †; a stable
  gutter so card width is independent of the scrollbar; `_scrollPos` is no longer reset when the bar
  hides unless the content fits. Move `AutoResizeWindow`'s `SetWindowPos` out of `Render` to a posted
  step, with the shrink hold. Anchor the first visible card by task ID and offset.
- [ ] F7. Card enter and leave motion, sharing the height animation state; removal keeps the entry until
  its collapse ends.
- [ ] F8. Cache `MeasureWrappedTextHeight` for the discovery summary beside the fact-pair cache (S7).
- [ ] F9. Capture. Extend the visual gallery with a deterministic *sequence* capture (N frames of one
  synthetic scenario at fixed ticks) so the stability scenario can be reviewed frame by frame. This is a
  harness capability and follows the screenshot rules (invented fixture data, provenance, no desktop
  takeover).

### Tests

`FileOps_PopupLayoutStability` (budget); stream-slot stability (a slot never re-binds while its stream
lives; release after two seconds); sticky stacking at a width boundary; tabular width equality of two
rate strings with different digits; reveal frames never paint a row below the bottom block's top;
reduced motion snaps; scroll anchor under height change above the viewport; scrollbar hysteresis;
window shrink hold; the existing `FileOps_PopupStableHostedTree`, `FileOps_PopupIdleCadence`,
`FileOps_StallLabelLayout` and collapse-settle cases stay green; visual gallery baselines regenerated
through the harness and reviewed.

Performance: the popup is a rendering hot path. Archive paint timings (the existing layout and paint
phase counters) for the stability scenario before and after; the median frame must not regress, and the
16 ms cadence must be active only during bounded motion (assert tick counts in `FileOps_PopupIdleCadence`).

## 11. Improvements and simplifications found on the way

Each is small, verify-first, and lands inside the slice named.

| ID | Finding | Action | Slice |
|---|---|---|---|
| S1 | `OnParentNotify` compares `HIWORD(wParam)` with pane IDs; that word is a child ID only for create and destroy † | Delete the branch if P0.2 confirms it never matches | A |
| S2 | The completed overflow button and its menu use two predicates; a retired `Failed items` action is still dispatched and a snapshot field is hard-wired false | One list builder; delete the retired action and field | C |
| S3 | `TypedConflictRecord::allowedActions` appears only at its declaration (survey) | Confirm, then delete or wire it | B |
| S4 | Artifact recognition is a chain of `HasTail` calls with two unrecognized current or historical shapes | One shape table; close both gaps | D |
| S5 | The bridge stage formatter duplicates the shared sibling-name builder and has no length bound | Use the shared builder | D |
| S6 | The card comparator switches key domain when a task starts | One key | E |
| S7 | `MeasureWrappedTextHeight` builds a text layout on every paint | Cache it | F |
| S8 | Four documents describe naming behavior that no longer exists | Correct them | D |
| S9 | `Errors: 2` for one failed item | Count item problems | B |
| S10 | `UpdateRates` recomputes on every 16 ms motion tick | Gate to the sample bucket | F |
| S11 | Move-edge availability is computed in the popup and again in the engine | One engine function | E |
| S12 | FolderView ignores the middle button, so that press never activates a pane | Handle it | A |

## 12. Order, slices and verification

Recommended order: **P0**, then **A**, **C**, **D** (small, independent), then **E** (coordinate with I60
Wave 5), **F**, and **B** last among equals only because B2 and B4 wait on D-4 and D-5; B1, B3 and B5 may
go as soon as P0.3–P0.5 are recorded. Each slice is one change with its tests, authoritative-spec edits,
user documentation and one `Changes/` fragment. No slice depends on another's code except: C2's button
orders include B4's action when B4 exists, B2's backup name uses D2's `.rs-bak` when D landed first, and
F's order stability assumes E5.

Per slice:

```powershell
.\Tools\Test-Changes.ps1 -Explain
.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter cmd_pane_folderView_activation_press_ -TestRoot <approved X:\RedSalamander.Perf root>
.\Tools\Run-AllTests.ps1 -Suite FileOps -CaseFilter <slice family> -TestRoot <approved X:\RedSalamander.Perf root>
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
.\Tools\Test-Changelog.ps1
.\Tools\Test-Changes.ps1 -Mode PrePush
```

Rules that apply to every slice:

- Build Debug for every project before calling a slice mergeable; self-tests compile only in Debug. Zero
  warnings and zero errors. Build from PowerShell.
- Run whole FileOps families, not single cases, before PrePush; a family or affected pass is not
  repository qualification. A Full run happens only if the owner requests it or the changed behavior's
  qualification requires it.
- Tests stay in the governed no-activation mode. Any focus-taking witness needs the owner's agreement to
  the time and uses the existing input warning and desktop lease.
- Test output stays under the approved `X:\RedSalamander.Perf` root. Evidence that a later run would
  delete is promoted at once.
- New user-visible strings go to `RedSalamander.rc` with positional placeholders and to the four
  satellites; captions follow the popup's wording policy.
- Screenshots for specs and documentation come from the harness with invented fixture data.
- A failed run is recorded with its cause; it is not replaced by a later pass.

## 13. Authoritative documents each slice updates

| Slice | Specs | User and developer documentation |
|---|---|---|
| A | `Specs/UI/UI_FolderView.md` (pointer target semantics, interaction summary), `Specs/UI/UI_FolderWindow.md` (active pane), `Specs/UI/UI_CommandMenuKeyboard.md` where it restates pane activation | `docs/SelectionsAndLists.md`, `docs/MainWindow.md`, `docs/dev/FolderView.md` |
| B | `Specs/FileSystem/FileSystem_FileOperations.md` (invariant 14, identity vocabulary, conflict model, action table, results), `Specs/UI/UI_FileOperationsPopup.md` (typed results, conflict layout, Issues), `Specs/Core/Core_FileSystemBridge.md`, `Specs/Plugins/Plugins_VirtualFileSystem.md` and `Specs/Plugins/Plugins_PluginAPI.md` if an interface changes | `docs/FileOperations.md`, `docs/Troubleshooting.md`, `docs/dev/FileOperationsEngine.md` |
| C | `Specs/UI/UI_FileOperationsPopup.md` (button placement, completed card actions, conflict layout), `Specs/FileSystem/FileSystem_FileOperations.md` (action placement) | `docs/FileOperations.md` |
| D | `Specs/FileSystem/FileSystem_FileOperations.md`, `Specs/Core/Core_Search.md`, `Specs/Core/Core_FileSystemBridge.md`, `Specs/Plugins/Plugins_VirtualFileSystem.md`, `Specs/Core/Core_SharedHelpers.md` if the helper's contract changes | section 8, D5 |
| E | `Specs/FileSystem/FileSystem_FileOperations.md` (interlock and queue), `Specs/UI/UI_FileOperationsPopup.md` (queue cards, waiting text) | `docs/FileOperations.md` |
| F | `Specs/UI/UI_FileOperationsPopup.md` (new Layout stability section, Motion, Progress Rules), `Specs/Testing/Testing_PerformanceValidation.md` (stability scenario and budget), `Specs/Testing/Testing_SelfTests.md` (sequence capture) | `docs/dev/TestingAndScreenshots.md`, `docs/VisualTour.md` captures |

Every slice also updates `Specs/Testing/Testing_SelfTests.md` when it adds a family, and the reviewed
source-contract tests its edits invalidate.

## 14. Done

Move this plan to `Specs/Plans/Done/` only when:

- every slice the owner accepted is merged with its tests, and each target contract is in its
  authoritative spec with no requirement left only in this file;
- D-2 to D-10 are each recorded as accepted, changed or declined, and a declined item is either removed
  from scope here or routed to a named owner;
- the P0 baselines and the after-evidence for slices B, D, E and F are archived under `Specs/TestRuns/`
  with their build receipts, including recorded skips for environments that were not available;
- the security review of B2 is recorded;
- the exact Done path is admitted in the protected-history allowlist with its focused test.

## 15. STOP

Stop and return to the owner when any of these holds:

- An activating press would change the selection without a modifier, or a modifier press would behave
  differently in the two panes.
- B1 would let an existence read authorize a replacement, a source cleanup, or an equality claim; B2
  would replace a link, a directory or a reparse point, would delete the set-aside original before the
  publication proof, or would run without a typed user decision; B4 would remove a source without the
  Delete route's own safeguards.
- The delete-by-name route (`IFileSystemIdentityDelete`) would be reused for replacement. Its contract
  forbids that; B2 is a separate route.
- A queue change can leave a queue place behind a finished, failed or cancelled task, or lets Start now
  bypass a capacity or safety interlock.
- A second ordering key would be introduced beside `_queueOrderKey`, or the slice-E edit would conflict
  with an unmerged I60 Wave 5 change in the same function.
- Legacy staging names would stop being recognized, a bare `.rs-tmp` or `.rs-bak` name would be
  classified as an artifact, or name recognition regresses FolderView enumeration.
- Height motion would let body rows paint over the bottom block, or the 16 ms cadence would run outside
  a bounded motion.
- A fix needs a DxUi change or pin bump.
- A characterization result in P0 contradicts the root cause this plan states for its slice. Correct the
  plan first.

## 16. Out of scope

Automatic replacement without a prompt and the `fileOverwrite` preference (I18); parking conflicts for
later review (H3); replacement of directories, links or reparse points at identity-less locations;
rename-the-existing-item choices; content-digest comparison in the prompt; conditional bulk policies
other than B6; remote, S3 and MTP staging names; a new scrollbar control; queue persistence across
restarts; the exclusive-transfer starvation and pause slot yield owned by I60 `W5.2` and `W5.3`.
