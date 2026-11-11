# I62 File Operations UX fidelity and pane activation click

Status: **DONE, I62 (closed 2026-11-11 by the owner).** Delivered: slices A, C, D and E, slice B (B1 to
B8, with the limits recorded in section 18), keep awake (G7) and the part of slice F named in section 18.
Not delivered and moved to [I63](../WIP/FileOperations_CollectedCollisionsAndResiduals_2026-11-11.md): collected
collisions (G1 to G6, by the owner's decision of 2026-11-11) and every item whose checkbox below is not
ticked. A ticked item was delivered; a closeout line under an item says what differs from its text. This
file is history and is not edited after the move to Done; the authoritative specs carry the contracts.

Revision 2, 2026-11-10: the owner set the replacement-authority rule (D-4), made "skip the same size and
date" required (D-10) and accepted the scenario review in section 1; slice B was rewritten and slice G
added. The same day the owner accepted a Move's `Delete source` choice, usable for all items already at
the destination and removing to the Recycle Bin where the source has one (D-5, B4).

Priority: **P1**. Effort: **L**, delivered as seven independent slices (A–G) plus a shared
characterization slice (P0). Implementation risk: **High** for slice B (replacement and source removal
at a location without object identity), slice E (queue admission) and slice G (collected collisions),
**Medium** for F, **Low** for A, C, D.

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
and cite the current line when it moved. Revision 2's anchors were read at `06fbdfb77`, which differs
from the planned commit only in planning files. No build or test was run for this plan.

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

A review on 2026-11-10 against the specs, the transfer guard and this plan found scenarios the six
reports do not cover. The owner accepted all of them into this plan.

| # | Scenario | State at the planned commit | Where |
|---|---|---|---|
| 7 | A folder already exists at a destination without object identity, as when an interrupted Move is run again | The guard's refusal exit tests existence, not kind (`State.cpp:24100-24111` †); every destination that is neither bound nor missing ends the same way (`FolderWindow.FileOperations.cpp:1683-1687` †) | B1.5 |
| 8 | Move from, or inside, such a location | A source that cannot be bound ends the guard with no source authority (`FolderWindow.FileOperations.cpp:1645-1649` †); a Managed Move then takes the generic refusal (`State.cpp:24112-24119` †). The planner's choice of strategy for this Move was not traced | B7, P0.11 |
| 9 | Source and destination are one file reached by two spellings | Same-object detection needs both sides bound (`FolderWindow.FileOperations.cpp:1659-1673` †); without identity only equal path text inside one endpoint is caught (`State.cpp:24043-24063` †) | B8 |
| 10 | A collision stops its task until it is answered | The per-item loop waits on its prompt and a task has one actionable prompt. What bridge pumps and provider walks already in flight do meanwhile was not traced | G, P0.13 |
| 11 | No choice before starting, and no "replace all older" | Neither spec names one; I18 owns the start surface | B6, D-13 |
| 12 | Running an interrupted operation again asks about every item already there, and a Move run again cannot drop the sources of those items | No bulk choice depends on the comparison of the two items | B6, B4.5 |
| 13 | The computer goes to sleep during a long operation | No power request exists in `RedSalamander`, `Common` or `Plugins` (searched for `SetThreadExecutionState` and `PowerCreateRequest` †) | G7 |
| 14 | Undo | `FileSystem_FileOperations.md` records "Undo/G2 remains deferred" † | D-15 |

Outcomes when this plan is done:

1. The first plain click in the other pane activates it, moves the cursor there and keeps the selection.
2. A name collision always becomes a decision that shows both items' facts and offers every choice the
   location can perform, at the level of safety the location provides and no other: Replace and source
   removal at a location without object identity, folder merge there, and a source-side choice for Move.
3. No menu button ever opens a menu with fewer than two commands.
4. Every staging artifact created on a Win32 path ends in `.rs-tmp` (incomplete) or `.rs-bak` (may hold
   an original).
5. Queued tasks start in the order of their cards, and a card that is passed over says why.
6. The card list changes geometry rarely, in bounded smooth steps, with measured stability budgets.
7. A collision never stops the rest of its task: independent items continue and the collisions are
   answered together.
8. An interrupted operation that is run again finishes without one question per item already there,
   and a Move run again can remove the sources of those items in one choice.
9. The computer does not go to sleep on its own while a task is running.

## 2. Ownership and coordination

This plan owns the seven slices and the improvements listed in section 12. It does not reopen any Done
record.

- **I18** (confirmation preferences and outcomes) owns start confirmation, the `fileOverwrite` and
  `directoryMerge` settings and metadata-outcome policy. Slice B adds decision actions, the authority
  ladder and two bulk choices in the collision prompt; it adds no setting and no replacement without a
  decision or a policy the user chose. The start-surface control proposed by D-13 belongs to I18 and
  consumes B6's task policy value. I18 section 5.4 ("reduced safe action set") must be re-read against
  the ladder when either lands second, and B6's policy-origin receipt is the one 5.4 describes.
- **I60 Wave 5** (`W5.2` queue lifecycle state and `_queueTicket`, `W5.3` pause slot yield) edits the same
  `EnterOperation` region as slice E. Whichever lands second rebases. Slice E must not introduce a second
  ordering key: `_queueOrderKey` is the ticket. If `W5.2` lands first, implement E1 on its
  `QueueLifecycleState`; if E lands first, `W5.2` keeps E's admission-time queue place and tests.
- **H3** (later features) held "Continue safe work and review conflicts later". The owner promoted that
  row on 2026-11-10: slice G owns it and H3's record points here. G starts from H3's accepted intent and
  keeps its constraints; the default is the one open point (D-12). H3 keeps its other five rows.
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
- **D-4 (owner, 2026-11-10).** A location is asked for the level of safety it can provide and nothing
  more. Where file identity exists (more than the name), use it. Where it does not and rename is atomic
  and cheap, use the rename. Otherwise behave as Windows File Explorer does, to keep a good user
  experience. Accepted the same day as consequences of that rule: it also governs the removal of a
  Move's source (B7) and a folder that already exists (B1.5), and one check is added because the relaxed
  rule would otherwise permit a loss: one file reached by two spellings (B8). Revision 1's design (a
  handle held on the occupant during the transfer and a set-aside backup) is withdrawn. D-4 supersedes
  the part of I35 D-7.3 that kept Move and replacement away from by-name authority
  (`FileSystem_FileOperations.md:1889-1891` †); Permanent Delete by name keeps its own confirmation.
- **D-5 (owner, 2026-11-10).** A Move collision offers a choice that keeps the existing destination and
  removes the source (B4), for regular files. It must be usable for all items, to finish a Move that was
  run before without keeping the sources: Apply to all and the task-wide form exist only for collisions
  with the same size and date, and an item whose facts differ always asks on its own. The source goes to
  the Recycle Bin when its location has one; otherwise the removal is permanent after one confirmation
  step on the card, asked once when the choice covers all items. Never a default. Known consequence:
  disk space is freed only when the Recycle Bin is emptied.
- **D-10 (owner, 2026-11-10).** "Skip items with the same size and date" is required, not optional: it
  is how an interrupted operation is run again (B6).
- **D-16 (owner, 2026-11-10): rejected.** The Delete source policy does not compare content, with
  Verify on or off. A copy is reliable and never leaves a partial file under the final name (the stage
  is published by one rename), so the same size and dates are the right level for B4.5.
- **Order to implement (owner, 2026-11-10).** Implementation of the whole plan is authorized: code,
  tests, normative specifications and documentation.
- **D-17 (owner, 2026-11-10).** Leftover recognition covers only the names the current build writes.
  A name an earlier version wrote is an ordinary file: no badge, no cleanup-dialog row. This removes
  the legacy rows of D3 and simplifies the recognizer and its tests.
- **Scenario review (owner, 2026-11-10).** Scenarios 7 to 14 of section 1 are in scope as listed there.
- **D-6 (owner, 2026-11-11).** An action that would be alone under `More...` must be a direct button.
  Delivered for `Delete all remaining permanently` and `Delete all remaining by name`: a direct button
  immediately before Cancel, never the default and never the Escape action, with focus still on Cancel
  when the prompt opens. No confirmation step was added: the one the recommendation below relied on is
  the task-start confirmation, which does not exist in the middle of a task. Cost, stated in the spec:
  one press deletes this item and every later item of the task that also cannot be recycled.
- **Slice G1 to G6 (owner, 2026-11-11).** Moved to another detailed plan, for later: I63.
- **Close (owner, 2026-11-11).** Move this plan to Done and open one pull request for all of it.
- **Second volume (owner, 2026-11-11).** The computer has several drives and every `X:\RedSalamander.Perf`
  folder is authorized; the second-volume cases run with `-AllowAlternateVolumeTestRoot`.

Open when implementation was ordered, each with a recommendation. The owner ordered implementation with
these still listed and changed none of them, so each slice implemented the recommendation. The table
keeps the original text; the closeout record follows it.

| ID | Question | Recommendation |
|---|---|---|
| D-2 | Is a press that activates the application window on the *already active* pane an activating press? | **No.** The rule depends only on which pane is active, so one sentence describes it. `_paneFocused` already tracks the active pane while the application is in the background (`GetPaneFromChild(nullptr)` returns `_activePane`, `FolderWindow.Interaction.cpp:797-802` †). |
| D-3 | An activating press lands on an unselected item while other items are selected, then the pointer leaves the drag threshold. | **Drag the pressed item alone**: replace the selection with it when the drag starts. The alternative (no drag, per current rule in `UI_FolderView.md:507` †) makes a cross-pane drag silently do nothing. |
| D-6 | `Delete all remaining permanently` is today the only entry of a `More...` menu and the spec forbids a direct button. | **Direct button whose caption ends with an ellipsis and opens the existing permanent-delete confirmation step.** No one-entry menu, and the second deliberate step is kept. |
| D-7 | While the head of the queue waits for the user's answer on its own card, may the next queued task start? | **Yes.** Machine Preparing holds the queue place; waiting for a person does not. The passed-over card states the reason and starts next once answered. |
| D-8 | Card order in the popup. | **One stable key for every active card** (`_queueOrderKey`, then task id). Cards never move when a task starts; waiting cards show their queue position. |
| D-9 | Scope of the staging-name change. | **Bridge stage plus the Local provider's stage, backup and writer names** in one slice, so every Win32-visible artifact groups under two extensions. Remote, S3 and MTP names keep their dot-hidden forms. |
| D-11 | The last level of D-4: what happens when the rename that replaces is refused, or the route has none. | **Refused because the occupant is in use or protected: show that error with Retry, Skip and Cancel, as Explorer does, and never fall back to writing over the file. No replacing rename on the route at all: write over the file after the Replace decision and, if the transfer fails, name the incomplete destination and keep the source.** No Local location is known to need the second case; P0.12 records whether one exists. |
| D-12 | Slice G's default. H3 planned "Ask as issues occur" as the default, with collection as a task-local choice. | **Collect name collisions by default**, as Explorer does, with no preference in the first version. Other problems (access denied, disk full, file in use) keep their immediate prompts. |
| D-13 | A choice before starting: "If a file already exists: Ask / Skip / Replace older / Keep both". | **Yes, as an immutable task option, default Ask.** This plan delivers the policy value and its engine behavior with B6; the control lands on the start surface with I18, which owns it. The Delete source policy of B4.5 is not offered before starting: it is chosen at a collision, where the facts are visible. |
| D-14 | Keep the computer awake while a task runs (G7). | **Yes, always, with no preference.** System-required only: the display may turn off, and closing the lid or choosing Sleep still sleeps. |
| D-15 | Undo for Copy, Move, Rename and Delete. | **Not in this plan.** Open a separate plan after B and G land, because undo consumes the per-item typed results they produce. The spec's "Undo/G2 remains deferred" stays. |

Closeout record of these decisions (2026-11-11). "As recommended" means the recommendation was built
under the order to implement; the owner did not confirm it one by one, and I63 item RX1 asks for that.

| ID | At closeout |
|---|---|
| D-2 | As recommended. |
| D-3 | As recommended. |
| D-6 | **Decided by the owner on 2026-11-11** and delivered as a plain direct button; see the recorded decision above. The recommendation's ellipsis caption and confirmation step were not built. |
| D-7 | As recommended. |
| D-8 | As recommended. |
| D-9 | As recommended, within D-17 (no legacy names). |
| D-11 | First half as recommended: a refused replacing rename shows its error and never falls back to writing over the file. The second half (a route with no replacing rename) was not built: no Local location is known to need it, and the search (P0.12) was not run; I63 item RP3. |
| D-12 | Not built: slice G1 to G6 moved to I63, which carries the question. |
| D-13 | Engine side as recommended (`OperationAdmission::presetCollisionPolicy`, default Ask); no control sets it. The control stays with I18. |
| D-14 | As recommended. |
| D-15 | Unchanged: undo is not in this plan. |

## 4. Slice P0 — characterization and baselines (no behavior change)

Purpose: turn each survey claim into a witnessed fact and capture the "before" evidence the later
slices compare against. P0 changes tests, test hooks and evidence only.

- [x] **P0.1 Pane press.** Add a failing-first case beside
  `cmd_pane_navigationView_unfocused_pane_click_focuses_target_pane`
  (`SelfTest/Commands/Commands.Tests.ViewCommands.cpp`) that focuses the left pane, selects three items
  in the right pane, sends a plain press to a fourth right-pane item with the existing
  `SendMouseClickToResolvedPointWindow` helper, pumps posted messages, and records today's result
  (selection collapses to the pressed item). Record the background-press result (selection cleared).
  *Closeout 2026-11-11: Delivered differently. The failing-first record of today's behavior was not
  captured; the equivalent witness is Commands case cmd_pane_folderView_activation_press_keeps_selection
  (and the _modifiers and _drag_and_double_click cases), which assert the new behavior with three items
  selected.*
- [x] **P0.2 Parent-notify trace.** Under `ENABLE_TESTS`, count `FolderWindow::OnParentNotify` entries
  whose `childId` matches a pane ID for a real button-down (`FolderWindow.Interaction.cpp:942-965` †,
  dispatch `FolderWindow.cpp:1363` †). `HIWORD(wParam)` carries a child ID only for `WM_CREATE` and
  `WM_DESTROY`, so the expected count is zero. This decides S1.
  *Closeout 2026-11-11: Not needed. No trace was added; the OnParentNotify handler and its dispatch were
  removed on the documented message contract (log slice A); no OnParentNotify remains in the tree.*
- [x] **P0.3 Identity-less collisions.** With the existing injection
  `REDSALAMANDER_FILEOPS_BIND_IDENTITY_UNSUPPORTED_PATH`, and on the registered network test root
  `path.smb.primary` where the environment provides it, record for an existing destination: single-item
  Move, multi-item Move (does the task continue after the refused item, and what do the remaining items
  report), Copy (expected: a prompt without Replace, because `allowDestructiveReplacement` needs a bound
  object or an atomic-final writer, `State.cpp:9795-9865` †), folder onto folder for Move and for Copy
  (expected for Move: the same refusal as a file, because the exit tests existence only), a folder Move
  run again after an injected interruption, file onto folder, and a read-only occupant. Record the
  `Errors` count and both diagnostics for the single-item case.
  *Closeout 2026-11-11: Delivered differently. No before-change record and no run on the network root;
  the 14-variant RunFileOperationIdentityLessCollisionDecisionSelfTest now witnesses single-item Move
  collisions at an identity-less destination (file, read-only occupant, changed and in-use occupant,
  whole-folder and per-item).*
- [ ] **P0.4 Refusal inventory.** List every exit of `prepareExactTransferGuard` and every
  `TransferSafetyState` that ends an item without a decision. Classify each as *must refuse* (same
  object, self-descendant, missing source, unsupported source binding) or *has a safe choice*.
  *Closeout 2026-11-11: Not delivered. No inventory of prepareExactTransferGuard exits or
  TransferSafetyState values is recorded in the plan, specs or tests; only the identity-less
  existing-destination exits were reclassified. Remainder: I63 item RB13.*
- [x] **P0.5 Local replace route.** Read how the Local provider publishes a stage at an identity-less
  location (the handle-bound stage with its lifetime nonce, `FileSystem_FileOperations.md:1124-1135` †;
  `RenameExactTo(destination, replaceIfExists)` on the stage's own handle,
  `Plugins/FileSystem/FileSystem.cpp:4376-4430` †) and how the remote providers implement
  `IFileWriterExpectedReplacement` (`Common/PlugInterfaces/FileSystem.h:1161-1176` †; survey:
  `FileSystemCurl.DirectoryOps.cpp:158-300`). State the smallest change that lets the Local route
  publish over an expected occupant: that writer interface as it is, or an expected-occupant argument on
  `PublishAs`. This fixes the B2 shape before any code.
  *Closeout 2026-11-11: Delivered differently. Read-only result: the shape chosen is the optional
  provider interface IFileSystemExpectedOccupantPublication (ReadExpectedOccupant,
  PublishOverExpectedOccupant), not the writer interface and not a PublishAs argument; it is documented
  in Common/PlugInterfaces/FileSystem.h and the spec.*
- [x] **P0.6 Menu census.** Under `ENABLE_TESTS`, record the entry count of every popup flyout opened by
  the visual gallery and the Commands popup cases. Expected today: one for every conflict `More...`,
  one for `Open destination`-only and `Open source`-only completed cards, zero for a completed card
  whose only overflow reason is a suppressed diagnostic.
  *Closeout 2026-11-11: Delivered differently. No census was recorded; the equivalent is the test-build
  flyout counter (NoteFileOpsFlyoutOpeningForSelfTest) that counts flyouts with fewer than two commands,
  asserted zero across the visual gallery.*
- [x] **P0.7 Queue order.** Add a deterministic Preparing pause point (same pattern as the existing
  `g_fileOps...PausePoint` hooks) and a failing-first case: six queued tasks, the later ones prepare
  faster; record the start order against the card order. Record the card index of each task across its
  waiting-to-running transition, and the result of Move Up across an overlap-queued neighbour.
  *Closeout 2026-11-11: Delivered differently. The Preparing pause point
  g_fileOpsPreparingBeforePublishPausePoint was added, but no before-change record of start order versus
  card order exists; the equivalent witness is DebugValidateQueueStartOrderForSelfTest (seven scenarios)
  plus the popup card-order cases.*
- [ ] **P0.8 Layout stability baseline.** Add the stability counters of F4 (counters only) and a
  twelve-second synthetic scenario with one to eight streams churning and rate text changing every tick.
  Archive the baseline: card height changes, layout shift, scrollbar toggles, window resizes, and the
  existing paint timings.
  *Closeout 2026-11-11: Not delivered. No twelve-second scenario and no archived baseline; only three of
  the four F4 counters exist and nothing was captured before the change. Remainder: I63 item RP1.*
- [ ] **P0.9 Staging-name perf baseline.** Run the existing artifact name-shape projection measurement
  (`DebugMeasureArtifactNameShapeProjectionForTests`) and archive the result; slice D compares to it.
  *Closeout 2026-11-11: Not delivered. The log states the baseline was not captured before the change;
  DebugMeasureArtifactNameShapeProjectionForTests exists but no result was archived for comparison.
  Remainder: I63 item RD1.*
- [x] **P0.10 Test reliance audit.** List the cases that press an item in a pane without focusing it
  first, the three leftover scanners that search for the literal `.rs_tmp_`, and the cases that pin
  `conflictPromptCount == 0u`, `overflowActionCount == 1u` and the `More...` button; these are the
  expected edits of slices A, D, B and C.
  *Closeout 2026-11-11: Delivered differently. No audit list was written; the affected cases were edited
  as part of slices A, C and D (activation press, flyout and More... lookups, and the .rs_tmp_ scanners
  now use the registry predicate or assert absence).*
- [x] **P0.11 Source side and one file by two spellings.** With the bind injection on the source root,
  and on `path.smb.primary`: Move from the identity-less location to a Local disk, Move to another
  folder inside it, and Copy from it. Record the strategy the planner chose (Native, Managed or
  Copy-only), the result and its text. Then reach one test folder by two spellings (the loopback alias
  the `FileOpsFamily_IdentityLess` cases use) and record a Copy and a Move from one spelling onto the
  other with the destination existing: what is detected, the share mode the source is opened with,
  whether a rename over the open source is refused, and whether the B8 probe distinguishes one file
  from two on each location class.
  *Closeout 2026-11-11: Delivered differently. Not run on a network root and no recorded table; the B8
  probe and the source-side Move are witnessed on Local by RunFileOperationIdentityLessSourceMoveSelfTest
  (ten cases including one file reached through a junction).*
- [ ] **P0.12 The replacing rename.** On `path.smb.primary` and on the loopback alias, run
  `RenameExactTo` with `replaceIfExists` over a regular file, a read-only file, a file another handle
  holds with and without delete sharing, a directory and a link. Record each status. Run a reader loop
  on the final name during 1,000 replacements and record whether it ever sees a missing name or partial
  content. Record the requests per replaced file from an SMB client trace where the environment allows
  one, otherwise record the skip. Record any Local location class where no replacing rename exists
  (D-11).
  *Closeout 2026-11-11: Partly delivered. Local read-only and in-use occupant outcomes are witnessed by
  the identity-less decision self-test; nothing was run on a network root, no 1,000-replacement reader
  loop, no client trace, and no search for a Local class without a replacing rename. Remainder: I63 item
  RP3.*
- [x] **P0.13 What continues during a prompt.** For a collision at a top-level item, at a bridge item
  with sixteen pumps, and inside a provider-owned recursive walk, record which other items progress
  while the prompt is unanswered and what the provider's issue callback accepts as an answer. This
  fixes the G shape.
  *Closeout 2026-11-11: Delivered differently. Answered by reading for the slice G design (a shown prompt
  stops every worker of the task through Task::WaitWhilePaused and the callback mutexes); no witness run
  was recorded.*
- [ ] **P0.14 Inactive-pane selection.** Capture through the harness the pane that is not active with
  three selected items and a current item outside them. Record whether the selection and the current
  item are each visible and distinguishable from the active style.
  *Closeout 2026-11-11: Not delivered. Not captured; the log says the renderer was read (inactive
  selection and thinner current-item border in FolderView.Rendering.cpp) but no harness capture was
  taken. Remainder: I63 item RA2.*
- [ ] **P0.15 Power.** Record that a running synthetic task leaves no entry in `powercfg /requests`.
  *Closeout 2026-11-11: Not delivered. powercfg /requests was not recorded; G7 is covered only by
  RunFileOperationKeepAwakeSelfTest through a fake power API seam. Remainder: I63 item RG3.*

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
> - An activating press on empty background, with or without modifiers, MUST NOT change the selection,
>   the current item or the range anchor. A miss has no toggle or range endpoint.
> - Ctrl, Shift and Ctrl+Shift activating presses on an item apply exactly the active-pane rules: they
>   state a selection intent explicitly.
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

- [x] A1. In `OnLButtonDown`, read `const bool activating = ! _paneFocused;` before `SetFocus`. Keep the
  alert-overlay early returns unchanged. Plain item press: `activating ? FocusItem(hit, false) : existing
  branch`; plain background press: skip `ClearSelection` and the anchor reset when `activating`.
  Modifier branches are untouched.
- [x] A2. Drag (D-3): record `activating` and the pressed index in `_drag`. For an activating press only,
  arm the potential drag even when the pressed item is outside the selection; when the threshold is
  crossed in that state, call
  `SelectSingle(pressed)` and rebuild the immutable source snapshot before starting the drag. Release
  inside the threshold changes nothing further.
- [x] A3. Add `WM_MBUTTONDOWN` to `FolderView::WndProc`: `SetFocus` and nothing else.
- [x] A4. Activation ownership. FolderView and NavigationView already take focus themselves. If P0.2
  shows the pane-ID comparison in `OnParentNotify` never matches a mouse event, delete that branch and
  its dispatch (S1) instead of repairing it; otherwise keep it and document what it serves.
- [ ] A5. Staleness. `_paneFocused` is updated by a posted message that the queue delivers before the
  next input message. Add a debug assertion in the test build that `_paneFocused` equals
  `GetFocusedPane() == thisPane` at button-down, and cover the Terminal tab note in
  `FolderWindow.Layout.cpp:987` † (its `WM_SETFOCUS` posts no pane-focus change) with a case.
  *Closeout 2026-11-11: Not delivered. No debug assertion comparing _paneFocused with the focused pane at
  button-down was added and no Terminal-tab focus case exists. Remainder: I63 item RA3.*
- [x] A6. Focus-follows-pointer needs no change: with it enabled the pane is already active when the
  press arrives.
  *Closeout 2026-11-11: Not needed. The item states no change is needed for focus-follows-pointer; no
  code was changed.*
- [x] A7. Visibility. The contract is understandable only when the pane that is not active still shows
  its selection and its current item; otherwise a first click that selects nothing looks like a defect.
  If P0.14 shows either is missing or indistinguishable, state the inactive style in `UI_FolderView.md`
  and fix the renderer in this slice, with a gallery capture. A1 does not ship without it.
  *Closeout 2026-11-11: Not needed. The renderer already draws the inactive selection and a thinner
  current-item border and UI_FolderView.md states the unfocused-pane style, so no renderer fix was
  needed; the P0.14 capture was not taken.*

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
`UI_FileOperationsPopup.md:1600-1612`). The exit tests existence only, so an existing folder is refused
the same way (`State.cpp:24100-24111` †).

The product already replaces without object identity on every other route. FTP, SFTP, SCP, S3,
Microsoft Drive, Google Drive and MTP offer Replace when their atomic-final writer accepts the overwrite
flag: the host reads the occupant before the prompt, hands it to the writer through
`IFileWriterExpectedReplacement`, and the writer replaces only that occupant, comparing the strongest
identity it has, down to size and last-write time (`FileSystem_FileOperations.md:1290-1301` †,
`Common/PlugInterfaces/FileSystem.h:1161-1176` †). The Local provider at an identity-less location is
the one route that refuses: `DestinationSupportsAtomicReplace` is false for it (`State.cpp:8038-8059` †,
used at `:9821` and `:9865` †). The same storage is replaceable today when it is reached over SFTP and
not when it is reached as a share. The Local provider already owns the needed primitive: its stage
renames itself over a final name by its own handle (`RenameExactTo(destination, replaceIfExists)`,
`Plugins/FileSystem/FileSystem.cpp:4376-4430` †).

On the source side, a source that cannot be bound ends the guard with no source authority
(`FolderWindow.FileOperations.cpp:1645-1649` †), and a Managed Move then takes the generic refusal
(`State.cpp:24112-24119` †). Same-object detection needs both sides bound
(`FolderWindow.FileOperations.cpp:1659-1673` †); without identity only equal path text inside one
endpoint is caught (`State.cpp:24043-24063` †).

When a prompt *is* reached, the engine already offers Keep both, Replace, Skip, Cancel, an Apply-to-all
scope and Skip all, and the prompt shows `size • date` for both items. Gaps: the two fact lines are not
compared for the user, a provider route that reports no size shows none, Move has no source-side choice,
and no bulk choice depends on the comparison.

Windows Explorer parity, as the owner asked for:

| Explorer | RedSalamander at the planned commit | This plan |
|---|---|---|
| Replace the file in the destination | With a bound object or an atomic-final writer; never on an identity-less Local location | B2 |
| Skip this file | Yes, when a prompt exists | B1 makes the prompt exist |
| Merge into a folder that already exists | Yes where the folder is identified; refused at an identity-less location | B1.5 |
| Compare info for both files, keep both | `Keep both` renames the incoming item; facts are two raw lines | B1, B3 |
| Do this for all current items | Apply-to-all scope; `Skip all` hidden in a one-entry menu | C2 |
| Skip items with the same date and size | Absent | B6 |
| Keeps copying the other items while the question waits | Absent | G |
| Size and date of both files visible | Only inside a live prompt; a failed card shows none | B3 |
| — (Move) keep the existing file and drop the source, for one item or for all with the same size and date | Absent | B4 (D-5) |
| — replace only the older files | Absent | B6 |

### The authority ladder (D-4)

Target contract (a product invariant in `FileSystem_FileOperations.md` that replaces invariant 14; it
also amends invariant 3, the identity vocabulary, "Publication and managed Move" 564-577 †, the conflict
model 1290-1313 † and `Core_FileSystemBridge.md` 261-269 †):

> **Authority is the strongest the location provides.** A destructive step on an existing item, which
> is replacing it or removing a Move's source, uses the first of these levels the location supports:
>
> 1. **Object identity.** The exact bound object, as today.
> 2. **Expected item and an atomic rename.** Where the location opens the item but cannot identify it,
>    the engine keeps the facts it read when the user decided (kind, compared attributes, size,
>    last-write time, creation time, and the file index when the location reports a nonzero one), reads
>    them again immediately before the step, and returns to the decision when any differs. A replacement
>    then renames the completed stage over the name in one operation; a source removal deletes the name.
> 3. **By name, as Windows File Explorer does.** Only where the route has no rename that replaces
>    (D-11).
>
> A location is not asked for more than it can give: no handle is held on the existing item during the
> transfer and no set-aside copy of it is made. The authority for the step is always the user's typed
> decision, a policy the user chose, or the confirmed Move itself. The facts only detect that the item
> changed since; they are never an identity or an equal-content claim, and no result says "identical".
>
> A collision at an identity-less location raises the ordinary typed decision. A link, a reparse point
> of any kind and a type mismatch are never replaced at levels 2 and 3. Linked ancestors remain hard
> stops. A context that cannot ask keeps the pre-mutation refusal and its
> `identity.destinationExistsUnverifiable` diagnostic. Missing destinations and failed probes keep their
> original classification.

Invariant 3 ("no hash, size, timestamp, filename, or pathname authorizes Skip, overwrite, artifact
cleanup, or source deletion") stays and gains one sentence: *a decision or a policy the user chose
authorizes; sizes and timestamps may only stop a step because the item changed, and select the items a
user's bulk policy applies to (B4.5, B6).* The "Corroboration facts" term of the identity vocabulary
covers these steps as well as delete by name.

### B1 — the refusal becomes a typed collision decision

- [x] B1.1 In `prepareExactTransferGuard`, replace the terminal return for *Move, `Unsupported`,
  destination identity unsupported, existence read succeeded* with a distinct result that the per-item
  loop turns into `BeginConflictPrompt` through the same policy path provider-issued collisions use. The
  bucket comes from `RefineTypedConflictBucket` on the read metadata (regular file, read-only file,
  directory, type mismatch, link).
  *Closeout 2026-11-11: Delivered differently. The guard returns a proceed result (S_FALSE) for an
  existing destination at an identity-less location for a bound regular-file source, and the bridge's
  file pre-flight raises the typed decision with the item facts, instead of the loop turning a distinct
  guard result into BeginConflictPrompt.*
- [x] B1.2 `Keep both` re-enters the loop through the existing `keepBothDestinationOverride`; the fresh
  sibling name is missing, so the guard is `Ready` and the ordinary new-file route runs. `Skip` records a
  user skip (typed result, source retained, not an Error). `Cancel` cancels the task.
- [x] B1.3 Keep `FileOps.Conflict.DestinationIdentityUnavailable`; add a result dimension for the
  decision taken. Keep the resource string for the cannot-ask path.
  *Closeout 2026-11-11: Delivered differently. FileOps.Conflict.DestinationIdentityUnavailable and the
  IDS_FILEOPS_DIAG_DESTINATION_EXISTS_UNVERIFIABLE string are kept; no separate result dimension was
  added, the spec says the one counter counts both the decision and the refusal.*
- [ ] B1.4 Apply the P0.4 inventory: every *has a safe choice* exit raises a decision with at least
  Skip and Cancel; every *must refuse* exit keeps its refusal and gains a specific sentence where it
  still shows only a Windows status text.
  *Closeout 2026-11-11: Partly delivered. Only the identity-less existing-destination exits were turned
  into decisions or kept as refusals with a specific sentence; no P0.4 inventory exists and the other
  guard exits were not reviewed for a Skip/Cancel decision. Remainder: I63 item RB13.*
- [x] B1.5 Folder onto an existing folder. No decision: the folders merge, as they do at an identified
  location (product invariant 7). The guard returns its proceed result for *source is a directory,
  destination exists and is a directory*, and the children collide one by one. A destination that is a
  link and a linked ancestor stay stops. Folder onto file and file onto folder are a type mismatch: Keep
  both, Skip, Cancel. For a Move, an emptied source folder is removed by the rule of B7; a source folder
  that still holds skipped or set-aside children is kept, and the result says how many. When I18's
  `directoryMerge` ask-once lands, it applies here unchanged.

### B2 — Replace at an identity-less Local location (level 2)

- [x] B2.1 Eligibility comes from the destination provider, per path, through one capability answer in
  the shape P0.5 selects; the host adds no route-specific test. Preferred shape: the Local provider
  answers the atomic-final overwrite question true only for a path whose bind is `Unsupported`, and its
  writer there is its existing handle-bound stage plus `RenameExactTo`. `Task::FileSystemIssue` and the
  B1 prompt then need no new branch: `atomicReplaceEligible` covers it. Replace link stays bound-only.
  *Closeout 2026-11-11: Delivered differently. Eligibility comes through the optional interface
  IFileSystemExpectedOccupantPublication queried from the destination provider (ReadExpectedOccupant
  answers only for an unidentifiable regular file), not through the atomic-final overwrite answer.*
- [x] B2.2 Expectation. The host hands the writer the occupant it showed the user, as it does for the
  remote routes. The Local route needs size and, when nonzero, the file index beside today's times and
  attributes; `FileSystemBasicInformation` carries neither. Extend the expectation through its
  size-versioned header or a sibling structure, in source-tree lockstep; the remote writers keep reading
  the fields they read today. A cached Apply-to-all Replace hands each item its own occupant.
  *Closeout 2026-11-11: Delivered differently. The expectation is a new struct FileSystemExpectedOccupant
  (attributes, times, size, file index) carried by the new interface rather than an extension of the
  existing expectation header; remote writers are untouched.*
- [x] B2.3 Publication. With the stage complete, the provider reads the occupant's facts by name
  without following links (one metadata read), compares them with the expectation, and calls
  `RenameExactTo(final, replaceIfExists)`. A difference returns `ERROR_REVISION_MISMATCH` and publishes
  nothing; the host raises the decision again on the current occupant under its existing bound of two
  more decisions per item. Any non-commit aborts the stage by its own handle. A changed kind (the name
  is now a directory or a link) is a difference.
- [x] B2.4 Read-only occupant: only through the existing `Replace read-only` action and its wording. The
  provider clears the attribute after the facts matched, then renames; if the rename fails it restores
  the attribute.
- [x] B2.5 Refused rename (D-11). A sharing or access refusal goes to the existing sharing and access
  decisions (Retry, Skip, Cancel). The stage is aborted and the occupant is untouched. There is no
  fallback to writing over the file.
- [x] B2.6 Level 3 (D-11). Only if P0.12 finds a Local location class with no replacing rename: after
  the Replace decision the route writes over the file; a failed transfer reports the incomplete
  destination by name, never as complete, and a Move keeps its source. If P0.12 finds none, level 3
  lands as contract text and a refusal test, with no Local implementation.
  *Closeout 2026-11-11: Delivered differently. Level 3 was not implemented; the spec states there is no
  fallback to writing over the occupant and the in-use variant is the refusal test. P0.12 was not run, so
  no Local class without a replacing rename was searched for.*
- [x] B2.7 One Information note per task: the location cannot identify files, so replaced items were
  checked by name, size and dates.
- [ ] B2.8 Review checklist, recorded with the slice as part of its ordinary review: the occupant is
  swapped between the read and the rename (the accepted limit of level 2; the spec states it); an
  ancestor is replaced by a link; the rename target has become a directory or a link; the stage is
  aborted after every refused rename; a reader of the final name never sees it missing (P0.12).
  *Closeout 2026-11-11: Partly delivered. No review checklist was recorded; the spec states the accepted
  swap-between-read-and-rename limit and tests cover a changed occupant and a refused rename, but
  ancestor-replaced-by-link and target-became-link-or-directory have no named witness. Remainder: I63
  item RB11.*

### B3 — decision facts

Target contract (for "Conflict Layout" in `UI_FileOperationsPopup.md`):

> The engine publishes, with the prompt, a comparison of the two items when both facts are known: date
> relation (newer, older, same within the coarser of the two locations' timestamp precision) and size
> relation (larger, smaller, same). Each role line appends its relation words; equal size and date render
> one line, `Same size and date`, and never the word identical. A last-write difference of one hour,
> within that precision, with equal sizes renders `One hour apart, same size`: volumes of the FAT family
> and some servers store local time and shift across a daylight-saving change. Such an item is neither
> newer nor older, no bulk policy applies to it, and its own decision offers Apply to all for that
> class. Unknown facts are omitted, not guessed. The exact byte count and the full date and time are in
> the tooltip and the accessible description. The same two fact lines are retained on the Issues row of
> an item that ended skipped or failed at a collision, and on the finished card of a single-item task.

- [x] B3.1 Add the comparison to `ConflictPromptState` beside `ItemMetadata`; the popup and the pane
  decision mirror render it and derive nothing themselves.
- [ ] B3.2 Fill size on provider routes: `ReadConflictItemMetadata` falls back from
  `GetFileBasicInformation` to the provider's item-information call that carries size (name it in P0.3).
  *Closeout 2026-11-11: Not delivered. ReadConflictItemMetadata still falls back from
  GetFileBasicInformation to GetAttributes with no size on provider routes; the log lists it as not done.
  Remainder: I63 item RB1.*
- [ ] B3.3 Retain the facts in the typed item result so the Issues pane and the finished card can show
  them; bounded by the existing reservoirs.
  *Closeout 2026-11-11: Partly delivered. The Issues pane retains and shows the comparison and both fact
  lines (IssuesProjection comparisonText); the finished card of a one-item task does not show the facts,
  and an item decided by policy or a remembered answer carries none. Remainder: I63 items RB2 to RB4.*
- [x] B3.4 New `.rc` strings with positional placeholders, in all four satellites.

### B4 — Move: keep the existing destination, remove the source (D-5)

Target contract (for the conflict model and the prompt's scope controls):

> For a Move, a regular-file collision offers **Delete source**: the existing destination is kept and
> the source is removed. When the task has more than one item and the two items have the same size and
> date, the prompt also offers **Delete source for all with the same size and date**, a task-wide policy
> the user chose. It is how a Move that was run before is finished: an item already at the destination
> loses its source, a missing item is moved, and a source folder left empty is removed. The policy
> applies only to regular-file collisions whose comparison is known and equal; an item whose facts
> differ or are unknown, a link and a type mismatch still ask. Equal size and dates are not a content
> comparison and no result says identical. The policy compares no content, with Verify on or off
> (D-16): a copy never leaves a partial file under the final name.

- [x] B4.1 New `ConflictAction::DiscardSource`, offered only for a Move collision whose incoming and
  existing items are regular files. Caption `Delete source`; with D-6 styling rules: never the default,
  never the Enter or Escape action. Apply to all is offered for it only on a `Same size and date`
  collision, and on a `One hour apart, same size` collision for that class alone (B3). On any other
  collision it is a single-item action: removing a source whose destination differs is decided item by
  item.
- [ ] B4.2 Execution reuses the Delete route on the source: its exact bound object where the source has
  identity, the re-check and removal of B7 where it has none. Recycle Bin where the source location has
  one; otherwise one confirmation step on the card, asked once when the action covers all items. A
  changed source is refused and asks again. B8 applies before the removal.
  *Closeout 2026-11-11: Partly delivered. Execution removes a bound Local source through its exact
  binding to the Recycle Bin, with one confirmation where there is no Recycle Bin
  (RunFileOperationDiscardSourceWithoutRecycleBinSelfTest); the B7 by-name removal for a source without
  identity is not used: Delete source is not offered there. Remainder: I63 item RB5 (owner decision
  R-D1).*
- [x] B4.3 Typed result `source removed, existing destination kept`, counted on the card as
  `Already there, source removed: N`; the task is not Partial for it. When the two items differ in
  size, the prompt's problem sentence says so before the choice.
- [x] B4.4 `FileSystemIssueAction` (the provider ABI) is unchanged: this action is host-executed.
- [x] B4.5 Task-wide form (owner, 2026-11-10). The policy value `DiscardSourceSameSizeAndDate` joins
  B6.1 for Move tasks, with B6.2's item-scoped receipts. At each collision the engine reads both items'
  facts at that step. Equal: B8 where it applies, then the removal of B4.2. Different or unknown: the
  item's own decision. A source that changed after its facts were read is kept and asks. Folders merge
  by B1.5, so a source folder the policy leaves empty is removed and the Move ends complete. The value
  is never set before the task starts (D-13).
  *Closeout 2026-11-11: Delivered differently. No separate DiscardSourceSameSizeAndDate policy value; the
  task-wide form is the Apply-to-all grant of Delete source restricted to Same size and date pairs (spec
  13b states a task-wide Delete source beyond 13a is not in this version).*
- [x] B4.6 The same folder reached twice. A Move whose source and destination are one folder under two
  spellings matches every item of this policy. With both sides bound, same-object detection ends each
  item as the same file, as today. At a destination without identity B8 runs before every removal, and
  where its answer is unknown `Delete source` and this policy are withheld. No source is removed on
  equal facts alone at such a destination.
  *Closeout 2026-11-11: Delivered differently. Delete source is offered only when the pre-flight can tell
  the destination is another file (B8 probe at an identity-less destination) and is withheld otherwise;
  there is no separate policy value to withhold.*
- [x] B4.7 Accepted limit, stated in the spec, the confirmation step and the documentation: the policy
  trusts size and dates, as the skip policy does, and reads no content (D-16).

### B5 — honest counts

- [x] B5.1 The `Errors` chip counts item problems. The trailing `task.result` summary row is not counted
  when at least one item-level Error of that task explains the failure. Check the counter's contract in
  `UI_FileOperationsPopup.md` first; if it defines the chip as a diagnostic-row count, change the spec
  in the same edit.

### B6 — bulk choices that depend on the comparison (D-10, D-13)

Target contract (for the conflict model and the prompt's scope controls):

> When a task has more than one item, a regular-file collision prompt offers two task-wide policies
> beside Apply to all: **Skip all with the same size and date** and **Replace all older**. Each is a
> policy the user chose. It applies only to regular-file collisions whose comparison is known and
> matches; every other collision still asks. `Replace all older` replaces only when the incoming item is
> newer by more than the timestamp precision. An item skipped by policy is a typed user skip with its
> facts retained, counted on the card as `Already there: N`; a Move keeps its source and says so. An
> item replaced by policy goes through the same level of the ladder as a single Replace. A Move has a
> third policy, `Delete source for all with the same size and date`, defined in B4.

- [x] B6.1 One task-scoped collision policy value (`Ask`, `SkipSameSizeAndDate`, `ReplaceOlder`, for a
  Move `DiscardSourceSameSizeAndDate` (B4.5), and for D-13 `Skip` and `KeepBoth`), recorded like an
  Apply-to-all grant: set once by a decision, immutable afterwards, exactly scoped to the task and the
  regular-file class.
  *Closeout 2026-11-11: Delivered differently. Arbiter grants SkipSameSizeAndDate and ReplaceOlder plus
  the PresetCollisionPolicy values Ask, Skip, ReplaceOlder, KeepBoth exist; there is no
  DiscardSourceSameSizeAndDate value (see B4.5) and the grants are not one single policy field.*
- [ ] B6.2 Each application reads the item's facts at its own step and creates an item-scoped,
  policy-origin receipt, the shape I18 section 5.4 describes, so `fileOverwrite=false` can reuse it.
  *Closeout 2026-11-11: Not delivered. Item-scoped policy-origin receipts for I18 are listed in spec 13b
  as not in this version; each application reads the item's facts at its step but creates no receipt.
  Remainder: I63 item RB14.*
- [x] B6.3 A policy skip opens no content and creates no stage: its cost is the collision's metadata
  read. This is what makes a second run of an interrupted operation cheap.
- [x] B6.4 Result counts and sentences for `Already there` and for `replaced because older`; `.rc`
  strings in all four satellites.
- [x] B6.5 D-13: the same value may be set before the task starts. The control belongs to I18's start
  surface; this slice exposes the value to it and tests the engine behavior through the test seam.
- [x] B6.6 Stated in the result and the documentation: under `Skip all with the same size and date` a
  second run of an interrupted **Move** leaves the sources of the items it skipped, and the result says
  how many. `Delete source for all with the same size and date` (B4.5) is the choice that finishes such
  a Move without keeping them.
  *Closeout 2026-11-11: Delivered differently. docs/FileOperations.md states that a Move keeps the source
  of a policy-skipped item and that Delete source finishes such a Move; the result counts items 'already
  there' but states no separate count of kept sources.*

### B7 — removing a Move's source without object identity (level 2 on the source side)

Target contract (for "Publication and managed Move" and the managed Move source-cleanup rules). It
amends `FileSystem_FileOperations.md:1889-1891` †, which today says "Move, Rename and
expected-destination replacement proofs never use this route" about delete by name (I35 D-7.3); D-4
supersedes that sentence for a Move's own source and for replacement, and leaves Rename, Pack and Unpack
as they are:

> A Managed Move from an identity-less Local location keeps the source's facts as read on the handle it
> copies from. After the destination's publication proof and every required consent it reads the facts
> again by name, without following links, and removes the source by name only when all match. Otherwise
> the source is kept and the item reports `Copied; source kept` with the reason. Nothing is asked beyond
> the Move itself: the Move is the decision. This is not Permanent Delete by name, which keeps its own
> confirmation. A source directory is removed only when empty, through the provider's empty-directory
> removal.

- [x] B7.1 Guard. For *Managed Move, source bind `Unsupported`, source opened* return a proceed result
  that carries the source facts instead of the generic refusal. A source whose open failed
  (`Indeterminate`) keeps its refusal.
- [x] B7.2 Cleanup. Where the bridge calls `DeleteIfUnchanged` on the bound source, the identity-less
  source uses a facts-conditional delete of one regular file in the Local provider. P0.11 decides
  whether that is the existing `IFileSystemIdentityDelete` file path (without its pinned tree and
  without the host's by-name confirmation card) or a separate small method; either way the Permanent
  Delete route and its confirmation are not changed.
  *Closeout 2026-11-11: Delivered differently. Cleanup uses the existing
  IFileSystemIdentityDelete::DeleteIfIdentity with the facts read before the copy
  (FinalizeByNameSourceCleanup), plus ReadUnidentifiedMetadataFeatures to keep sources whose metadata the
  pathname copy does not carry.*
- [x] B7.3 One Information note per task, shared with B2.7. A kept source is a typed result, not an
  Error.
- [ ] B7.4 A Move inside one such location stays a native rename. Onto an existing name it is the B1
  decision, and Replace is the same rename with the expectation check and no stage.
  *Closeout 2026-11-11: Not delivered. No code or test for a Move inside one identity-less location
  (native rename) onto an existing name was found; the log lists the native-rename Move route as not
  covered. Remainder: I63 item RB6.*

### B8 — one file reached by two spellings

Target contract (for the conflict model):

> At a destination without object identity, an existing item whose facts equal the source's may be the
> source itself under another spelling, such as a mapped drive letter and a server path. Before any
> destructive step on such an item the engine asks the location: it opens the source denying all
> sharing, then opens the destination item. A sharing refusal means one file; the item ends `Source and
> destination are the same file` and nothing changes. When both opens succeed they are two files and the
> ordinary rules apply. Where the first open fails, or the location class does not enforce sharing, the
> answer is unknown: Replace and Delete source are withheld for that item, no bulk policy replaces it
> or removes its source, and a Move keeps its source. Skip, Keep both and Cancel remain.

- [x] B8.1 The probe runs only for *destination without identity, facts equal, destructive step
  chosen*: Replace, a replacing policy, Delete source and its task-wide form. A skip never probes, so
  the second run of an interrupted operation under the skip policy pays nothing for it. Under the
  Delete source policy every item it applies to has equal facts, so each pays the two opens at such a
  destination.
- [x] B8.2 Why it is needed: a copy keeps size and dates, so after a replacement the source's re-check
  of B7 would pass on a location that also reports no file index, and the Move would remove the only
  copy. Where the location reports a file index the re-check already fails safely; the probe makes the
  outcome the same everywhere and says why.
- [ ] B8.3 The location classes that do not enforce sharing come from P0.11 and are a typed capability
  of the path profile, not a host list of server names.
  *Closeout 2026-11-11: Not delivered. The probe is not replaced by a provider capability of the path
  profile; the log lists it as not done. Remainder: I63 item RB10.*

### Tests (FileOps suite; register every case in `kFileOpsFamilyDefinitions`)

Rewrite the five variants of the unverifiable-collision selftest: prompt count one; Keep both publishes
`name (2).ext` and removes the source; Skip retains both byte streams and reports a user skip; Cancel
mutates nothing; the cannot-ask path keeps the refusal. New family `FileOps_IdentityLessTransfer`:

- Replace: Copy and Move; occupant changed between the prompt and publication (size, time, attributes,
  replaced by a link, replaced by a directory) returns to the decision with nothing published; read-only
  occupant; occupant in use gives Retry and Skip with the occupant untouched; Apply to all over ten
  collisions with one changed occupant; a reader loop never sees the final name missing.
- Folders: Move onto an existing folder merges; a folder Move run again after an injected interruption
  completes, removes emptied source folders and keeps the ones holding skipped children; folder onto
  file and file onto folder.
- Source side: Move from the identity-less location to a Local disk removes the source after the proof;
  a source changed during the transfer is kept; a source whose open fails keeps the refusal.
- One file by two spellings: Move with Replace ends `same file` with one file left intact; injected
  "sharing not enforced" withholds Replace and Delete source; two different files with equal facts are
  replaced after the probe.
- `DiscardSource`: recyclable source, source without Recycle Bin with its confirmation, changed source,
  differing sizes; Apply to all is absent on a collision whose facts differ.
- `DiscardSource` for all: a folder Move run again over 500 items already there, three of which differ
  (three prompts, 497 sources removed, missing items moved, emptied source folders removed, task
  complete); one confirmation for the whole task where the source has no Recycle Bin; one folder under
  two spellings leaves every file intact; injected "sharing not enforced" withholds the policy.
- Comparison cues: newer, older, same within two seconds, one hour apart, unknown size.
- Policies: skip-same over 500 collisions with three that differ (three prompts, no content opened for
  the others); replace-older with mixed and unknown facts; a policy never applies to a link, a folder or
  a one-hour item.

Update the policy coverage selftest, the visual gallery fixtures that build the failed card, and the
Commands popup metadata cases. Where the environment provides `path.smb.primary`, run the family there
as well and record skips honestly.

Performance: level 2 adds one metadata read per replaced file and nothing per new file; B7 adds one per
removed source; B8 adds two opens per destructive step on an equal-facts item. Add
`FileOps.Conflict.ExpectedItemUs`. Archive a 500-collision Apply-to-all Replace beside the same run with
Keep both, a 2,000-item second run under the skip-same policy, and the same 2,000-item second run of a
Move under the Delete source policy (one removal per item, plus B8's two opens at a destination without
identity).

## 7. Slice C — no menu with fewer than two commands

Target contract (for `UI_FileOperationsPopup.md`, "Button placement" and "Completed Card Actions"):

> A menu button is shown only when its menu will contain at least two enabled commands. A single command
> is shown as its own direct button with its own caption, standard variant, and the same command route.
> No command means no button. The button and its menu are built from one list.

- [x] C1. Completed card. One builder returns the ordered completed-overflow command list for a task
  snapshot. `PrepareTerminalActions`, `ShowCompletedOverflowMenu`, the layout snapshot fields and the
  self-test hook consume that list; `CompletedTaskHasOverflowActions` is deleted. This also removes the
  zero-entry flyout that the two diverging predicates allow (`Popup.cpp:942-946` † against the menu
  builder's `hasDiagnostics`). A single command uses the ordinary 24-DIP caption padding.
- [x] C2. Conflict prompt. Placement stays engine-owned in `ConflictActionPolicy::FinalizePresentation`:
  an overflow set of one is published as a primary action. Resulting direct orders: collision
  `Keep both, Replace, Skip, Skip all, Cancel` for Copy, with `Delete source` before `Skip` for Move
  (B4); `Delete all remaining permanently...` per D-6. The broker's placement validation follows the
  published placement, so the popup never submits an overflow action as primary.
- [x] C3. The pane decision mirror consumes the same policy and needs no rule of its own.
- [x] C4. Guard: in test builds every flyout records its entry count; a new case fails when any popup
  flyout opens with fewer than two entries across the visual gallery scenarios.
- [x] C5. Accessibility: a direct button exposes its own name and Invoke; its automation ID keeps the
  `FileOperations.Action.<kind>.<taskId>.<data>` shape with a distinct kind for a direct completed
  action. Tab order stays visual order.
  *Closeout 2026-11-11: Delivered differently. The direct completed action has its own hit kind
  TaskCompletedAction and the automation ID keeps the FileOperations.Action.<kind>.<taskId>.<data> shape;
  Invoke and tab order were not separately verified.*

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

- [x] D1. Bridge stage. `MakeTempDestinationPath` † builds the leaf with
  `Common::Paths::BuildUniqueSiblingName` (`Common/Helpers.h:354-393` †): prefix the final leaf, marker
  `.`, suffix `_<stream>.rs-tmp`, limit 255. This removes the hand-rolled formatter and fixes a present
  defect: the formatter adds at least 42 units with no bound, so a final leaf longer than 213 units
  yields a stage leaf above 255 and fails stage creation. Keep the self-test entropy-failure hook.
- [x] D2. Local provider (D-9): copy stage `.rs_copy_tmp_<hex>` becomes `.<hex>.rs-tmp`, overwrite backup
  `.rs_bak_<hex>` becomes `.<hex>.rs-bak`, writer stage `.~rs-write-<hex>.tmp` becomes `.<hex>.rs-tmp`
  (its Hidden and Temporary attributes unchanged), through `MakeStageSiblingPath`. The case-only rename
  hop keeps its whole-leaf form and gains the `.rs-tmp` extension.
  *Closeout 2026-11-11: Delivered differently. Local stage and backup names use BuildUniqueSiblingName
  with .rs-tmp and .rs-bak; the case-only rename hop is <leaf>.<32 hex>.rs-bak (it carries the original),
  not the whole-leaf .rs-tmp form.*
- [x] D3. Recognition. Make `HasPossibleArtifactName` table-driven (S4): one table row per shape with
  its class. Add the new suffix-anchored shapes with a cheap `ends_with` pre-check. Keep every legacy
  row. Close two present gaps with tests: the current writer shape `<32 hex>.tmp` is not matched by
  `IsWriterTail` (`FileOperationArtifactRegistry.cpp:94-102` † against the generator at
  `FileSystem.cpp:6599` †), and the bridge shape used before June 2026 (`PID_TID_tick_stream`) is not
  matched by `IsBridgeTail`. Preserve the textual shapes the Pester source contracts pin (function
  signature, `.rs_bak_` before `.rs_ren_`).
  *Closeout 2026-11-11: Delivered differently. HasPossibleArtifactName recognizes only names the current
  build writes (D-17 dropped legacy rows and the two legacy-gap fixes) and is a short chain, not a table;
  token-required .rs-tmp and .rs-bak checks are in FileOperationArtifactRegistry.cpp.*
- [x] D4. Tests. Convert the three literal scanners found by P0.10 to the registry predicate; a scanner
  that searches for `.rs_tmp_` would pass silently once the name changes. Extend the shape table with
  new, legacy and near-miss names (`report.rs-tmp`, a 31-hex token, a non-hex token, upper-case hex).
  Update the planted fixtures. Add a bridge case with a 251-unit leaf (stage leaf exactly 255, publish
  succeeds) and a case with a surrogate pair at the cut. Cleanup-dialog cases cover both generations.
  *Closeout 2026-11-11: Delivered differently. Scanners and the Pester contract use the registry
  predicate or assert absence of the old names; the shape table has new, legacy (now not recognized) and
  near-miss rows; the 251-unit leaf and emoji-cut cases exist; cleanup-dialog cases cover one generation
  only (D-17).*
- [x] D5. Documentation and specs: `FileSystem_FileOperations.md` (the recognized-shape list near 2883,
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

- [x] E1. Queue place at admission. Where admission sets `_queueOrderKey` and `_waitForOthers` (survey:
  `State.Runtime.cpp:1448-1548`; batch rename `FolderWindow.FileOperations.cpp:3099-3137`), insert the
  task into `_queue` under `_queueMutex`. `EnterOperation` keeps the insertion as an idempotent
  fallback.
  *Closeout 2026-11-11: Delivered differently. Admission sets Task::_holdsQueuePlace (Runtime.cpp,
  FileOperations.cpp batch rename) and the walk reads the task list; there is no second insertion into
  _queue.*
- [x] E2. `firstUnblockedQueuedTaskId` walks `_queue` in order and classifies each entry: *holding* (no
  published prepared admission yet, not waiting for the user, not cancelled) stops the walk and returns
  no winner unless it is the caller; *waiting for the user* and *overlap predecessor* are skipped;
  *published and unblocked* wins.
- [x] E3. Every exit removes the place and wakes waiters: prepare failure, withdrawal, cancellation,
  clipboard-Move admission refusal, mode switch to Parallel, Start now. Every transition into and out of
  *waiting for the user* wakes waiters. A missed removal stalls the whole queue; E-tests cover each exit.
  *Closeout 2026-11-11: Delivered differently. Exits clear _holdsQueuePlace and wake waiters, and
  transitions into and out of awaiting-user wake them; covered by DebugValidateQueueStartOrderForSelfTest
  (seven scenarios), not by one case per exit.*
- [x] E4. Evaluate `ShouldQueueNewTask` inside the admission lock so two admissions issued together
  cannot both read "nothing active" (survey item; confirm in P0.7 before changing).
  *Closeout 2026-11-11: Not needed. Not changed: the log records the double-read race was not confirmed
  as a defect (P0.7 not run).*
- [x] E5. One comparator for `BuildSnapshot`: `(queue order key, task id)` for every active task; delete
  the dual-key branch. Publish the waiting ordinal in the snapshot and render it on the waiting chip.
- [x] E6. Reorder. `MoveQueuedTask` and the popup's move-edge flags come from one engine function that
  skips neighbours that cannot move and refuses a swap that would invert an overlap edge.
- [x] E7. Leave `UpdateQueuePausedTasks` and its entered-tick ordering unchanged unless P0.7 shows it
  contradicts the contract; if it does, order it by the same key.
  *Closeout 2026-11-11: Not needed. UpdateQueuePausedTasks was left unchanged; P0.7 was not run and no
  contradiction was found (log).*
- [ ] E8. Instrumentation: `FileOps.Queue.HeadHoldUs` (time the head held the queue while Preparing) and
  `FileOps.Queue.PassedOver` with its reason.
  *Closeout 2026-11-11: Partly delivered. FileOps.Queue.PassedOver is emitted with a reason (awaiting
  user decision, decision answered); FileOps.Queue.HeadHoldUs is not emitted. Remainder: I63 item RE1.*

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

- [x] F1. Stream slots. `UpdateInFlightFileProgress` (survey: `State.cpp:2396-2516`) assigns the lowest
  free slot and never compacts; the snapshot carries slot indices; the popup reserves the high-water
  count with the two-second release. `FileProgressIsVisible` keeps its 300 ms linger for the row's
  content, not for its reservation.
- [ ] F2. Sticky facts. `PrepareFactPair` (survey: `Popup.cpp:7251-7297`) takes a reserved numeric width
  for all three pairs (progress, counts and ETA, throughput and rate), as the graph labels already do,
  and keeps `stacked` with the stated hysteresis. Conditional notes get the same hold.
  *Closeout 2026-11-11: Partly delivered. Sticky stacking with hysteresis (PairWantsStack, HeldCondition)
  and held note rows are done; a reserved numeric width for the widest form of each number is not.
  Remainder: I63 item RF2.*
- [x] F3. Tabular figures on the numeric text layouts (`IDWriteTypography`, tabular-figures feature) in
  `EnsureTextFormats`; if the hosted DxUi typography owns the format, stop and route per section 2.
  Display hold for rate, ETA and the file name; `UpdateRates` runs once per 100 ms bucket even on 16 ms
  motion ticks (S10). The mini progress bar keeps its column when a row completes.
  *Closeout 2026-11-11: Delivered differently. Tabular figures on the three numeric fact rows laid out by
  the popup (ApplyTabularFigures), rate/ETA/file name holds, UpdateRates once per 100 ms, kept mini bar
  column; the hosted graph labels (DxUi) are unchanged.*
- [ ] F4. Counters and budget case. `FileOps.Popup.CardHeightChanges`, `FileOps.Popup.LayoutShiftPx`
  (sum of card-top movement not covered by a running height motion or a user action),
  `FileOps.Popup.ScrollBarToggles`, `FileOps.Popup.WindowResizes`. New case `FileOps_PopupLayoutStability`
  enforces the budget on the P0.8 scenario.
  *Closeout 2026-11-11: Partly delivered. Counters CardHeightChanges, ScrollBarToggles and WindowResizes
  are emitted; LayoutShiftPx and the FileOps_PopupLayoutStability budget case do not exist. Remainder:
  I63 item RF1.*
- [ ] F5. Height reveal. Replace the snap condition in `ResolveAnimatedTaskHeight` † with the reveal:
  layout at target, animated bottom block and following cards, row clip, integer pixels. Keep the snap
  for decision content (`Popup.cpp:9602-9605` †). In `RefreshTelemetryTimer` † an unsettled height motion
  selects the 16 ms cadence ahead of the 100 ms work cadence, for its bounded duration only.
  *Closeout 2026-11-11: Not delivered. ResolveAnimatedTaskHeight still snaps for expanded content; no
  layout-at-target reveal and no 16 ms cadence rule were added. Only the scrollbar withholding during a
  pending resize is done (see F6). Remainder: I63 item RF3.*
- [ ] F6. Viewport. Scrollbar hysteresis and pending-resize awareness in the scroll pass †; a stable
  gutter so card width is independent of the scrollbar; `_scrollPos` is no longer reset when the bar
  hides unless the content fits. Move `AutoResizeWindow`'s `SetWindowPos` out of `Render` to a posted
  step, with the shrink hold. Anchor the first visible card by task ID and offset.
  *Closeout 2026-11-11: Partly delivered. The scrollbar is withheld while an automatic resize is pending
  (ScrollBarNeeded); no stable gutter, no posted resize out of Render, no anchored first visible card.
  Remainder: I63 item RF4.*
- [ ] F7. Card enter and leave motion, sharing the height animation state; removal keeps the entry until
  its collapse ends.
  *Closeout 2026-11-11: Not delivered. No card enter or leave motion; the log lists it as not done.
  Remainder: I63 item RF5.*
- [x] F8. Cache `MeasureWrappedTextHeight` for the discovery summary beside the fact-pair cache (S7).
- [ ] F9. Capture. Extend the visual gallery with a deterministic *sequence* capture (N frames of one
  synthetic scenario at fixed ticks) so the stability scenario can be reviewed frame by frame. This is a
  harness capability and follows the screenshot rules (invented fixture data, provenance, no desktop
  takeover).
  *Closeout 2026-11-11: Not delivered. No sequence capture was added to the visual gallery. Remainder:
  I63 item RF6.*

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

## 11. Slice G — a collision does not stop the task

### Current behavior

The per-item loop waits on its own prompt, and a task has one actionable prompt at a time. What bridge
pumps and provider walks already in flight do while it waits is recorded by P0.13. H3 held this feature
as "Continue safe work and review conflicts later"; the owner promoted it on 2026-11-10, and its
accepted intent in H3's record is the starting point. No power request is made while tasks run †.

### Target contract (new section "Collected collisions" in `FileSystem_FileOperations.md`; "Conflict Layout", card states and Issues in `UI_FileOperationsPopup.md`)

> A name collision does not stop independent work. The colliding item is set aside before any change to
> it: its source is kept, nothing is staged for it, and the task continues with every item that does not
> depend on it. The card shows the question at once with a live count (`3 files already exist; the rest
> is still copying`), without taking focus. The user may answer at any time. An answer applies to the
> set-aside items of its class and, when the user chose Apply to all or a policy (B6), to later ones as
> they are found. Each set-aside item is then run as its own transfer, in original order, under the
> ordinary rules: its facts are read again, and a change returns it to the question.
>
> The set-aside list is bounded by count and by retained path text. At the bound the task waits for an
> answer, as it does today. A set-aside item holds no handle and no authority. A Move does not remove a
> folder while one of its descendants is set aside. When everything else is finished and questions
> remain, the task is `Needs attention`, gives up its queue place (slice E, case b), and no longer keeps
> the computer awake. Cancel records each unresolved item as not attempted, source kept. Closing the
> popup answers nothing. Problems other than name collisions keep their immediate prompts (D-12).

### Implementation

- [ ] G1. Shape, from P0.13. Where the host owns the item loop (top-level items and bridge items), a
  collision becomes a typed pending record (source path, destination path, bucket, facts) instead of a
  wait in `BeginConflictPrompt`. Inside a provider-owned walk, the issue callback is answered with Skip
  and the host keeps the same record, to run the item later as a single transfer. If a provider walk
  cannot be answered that way without an interface change, that walk keeps its immediate prompt in the
  first version and the limit is documented; if that walk is the Local provider's, stop.
  *Closeout 2026-11-11: Not delivered. Designed only. Remainder: I63 Part 1.*
- [ ] G2. Arbiter. One collected prompt per class through the existing conflict arbiter and policy; no
  second action or default policy. Its action set is published once and stays stable: only the count and
  the listed names change, so a button under the pointer never changes meaning. The existing rule that
  buttons are never added to or removed from an actionable prompt is kept.
  *Closeout 2026-11-11: Not delivered. Designed only. Remainder: I63 Part 1.*
- [ ] G3. Second pass. A set-aside item runs through the ordinary per-item route with the decision as
  its receipt, after its facts were read again. Its typed result replaces the pending one. Totals count
  it from discovery, so progress never moves backwards.
  *Closeout 2026-11-11: Not delivered. Designed only. Remainder: I63 Part 1.*
- [ ] G4. Managed Move directory cleanup waits for the descendants that are set aside; the result names
  folders kept for that reason.
  *Closeout 2026-11-11: Not delivered. Designed only. Remainder: I63 Part 1.*
- [ ] G5. Card and queue. `Needs attention` with the count; Issues rows for unresolved items; the task
  yields its queue place through slice E's case (b) and takes it back when answered.
  *Closeout 2026-11-11: Not delivered. Designed only. Remainder: I63 Part 1.*
- [ ] G6. Bounds and instrumentation: the two bounds as named constants with tests at the bound;
  `FileOps.Conflict.SetAside` and `FileOps.Conflict.SetAsideBoundReached`.
  *Closeout 2026-11-11: Not delivered. Designed only. Remainder: I63 Part 1.*
- [x] G7. Keep awake (D-14). One process-wide power request (`PowerCreateRequest` with
  `PowerRequestSystemRequired` and a reason string) is held while at least one task is running or
  stopping, and released when every task is finished, paused by the user, or only waiting for an
  answer. The display is not kept on. A test seam counts acquire and release. On a Modern Standby
  machine record what `powercfg /requests` shows and whether `PowerRequestExecutionRequired` is needed;
  record the skip where no such machine is available.

### Tests (new family `FileOps_CollectedCollisions`, registered)

One hundred items with five collisions: ninety-five complete while the question is unanswered; answers
given before and after the rest finished (Skip, Replace, Keep both); Apply to all and a B6 policy catch
collisions found later; an occupant changed while set aside returns to the question; Cancel records the
unresolved items; the bound is reached and the task waits; a Move keeps a folder with a set-aside child
and removes it once resolved; with only questions left, the next queued task starts; the prompt's
buttons and automation IDs are unchanged while the count changes; no focus change. Power: acquired with
the first running task, released with the last, not held while a task only waits for an answer or is
paused.

Performance: a non-colliding item pays nothing. Archive the 100-item scenario's elapsed time against the
same run answered at the first collision, and the retained bytes at the bound.

## 12. Improvements and simplifications found on the way

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
| S13 | The Local route at an identity-less location refuses a replacement that every remote route performs with a size-and-date expectation | One authority ladder for every route (D-4) | B |
| S14 | The guard's refusal exit tests existence, not kind, and a source without identity gets the generic guard sentence | A decision by kind; a specific sentence | B |
| S15 | H3 and this plan would both describe collision collection | H3's row points to slice G | G |

## 13. Order, slices and verification

Recommended order: **P0**, then **A**, **C**, **D** (small, independent), then **E** (coordinate with I60
Wave 5), **F**, then **B**, then **G**. Inside B: B1, B3 and B5 as soon as P0.3 and P0.4 are recorded;
B2, B7 and B8 after P0.5, P0.11 and P0.12; B6 after B3; B4 after B3, with B4.5 after B6, B7 and B8. G
follows B1 and E because its second pass runs B's decisions and its waiting state is slice E's case
(b); G7 is independent and may
land as soon as D-14 is answered. Each slice is one change with its tests, authoritative-spec edits,
user documentation and one `Changes/` fragment. No slice depends on another's code except those and:
C2's button orders include B4's action when B4 exists, and F's order stability assumes E5.

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

## 14. Authoritative documents each slice updates

| Slice | Specs | User and developer documentation |
|---|---|---|
| A | `Specs/UI/UI_FolderView.md` (pointer target semantics, interaction summary), `Specs/UI/UI_FolderWindow.md` (active pane), `Specs/UI/UI_CommandMenuKeyboard.md` where it restates pane activation | `docs/SelectionsAndLists.md`, `docs/MainWindow.md`, `docs/dev/FolderView.md` |
| B | `Specs/FileSystem/FileSystem_FileOperations.md` (invariants 3 and 14, identity vocabulary, "Publication and managed Move", conflict model, action table, managed Move source cleanup, results), `Specs/UI/UI_FileOperationsPopup.md` (typed results, conflict layout and scope controls, Issues), `Specs/Core/Core_FileSystemBridge.md` ("Conflict and cleanup"), `Specs/Plugins/Plugins_VirtualFileSystem.md` and `Specs/Plugins/Plugins_PluginAPI.md` with the `IFileWriterExpectedReplacement` contract comment in `Common/PlugInterfaces/FileSystem.h` | `docs/FileOperations.md`, `docs/Troubleshooting.md`, `docs/dev/FileOperationsEngine.md` |
| C | `Specs/UI/UI_FileOperationsPopup.md` (button placement, completed card actions, conflict layout), `Specs/FileSystem/FileSystem_FileOperations.md` (action placement) | `docs/FileOperations.md` |
| D | `Specs/FileSystem/FileSystem_FileOperations.md`, `Specs/Core/Core_Search.md`, `Specs/Core/Core_FileSystemBridge.md`, `Specs/Plugins/Plugins_VirtualFileSystem.md`, `Specs/Core/Core_SharedHelpers.md` if the helper's contract changes | section 8, D5 |
| E | `Specs/FileSystem/FileSystem_FileOperations.md` (interlock and queue), `Specs/UI/UI_FileOperationsPopup.md` (queue cards, waiting text) | `docs/FileOperations.md` |
| F | `Specs/UI/UI_FileOperationsPopup.md` (new Layout stability section, Motion, Progress Rules), `Specs/Testing/Testing_PerformanceValidation.md` (stability scenario and budget), `Specs/Testing/Testing_SelfTests.md` (sequence capture) | `docs/dev/TestingAndScreenshots.md`, `docs/VisualTour.md` captures |
| G | `Specs/FileSystem/FileSystem_FileOperations.md` (new Collected collisions section, conflict model, managed Move directory cleanup, queue rules, task lifecycle for the power request), `Specs/UI/UI_FileOperationsPopup.md` (conflict layout, card states, Issues); H3's record keeps only its pointer | `docs/FileOperations.md`, `docs/Troubleshooting.md`, `docs/dev/FileOperationsEngine.md` |

Every slice also updates `Specs/Testing/Testing_SelfTests.md` when it adds a family, and the reviewed
source-contract tests its edits invalidate.

## 15. Done

Move this plan to `Specs/Plans/Done/` only when:

- every slice the owner accepted is merged with its tests, and each target contract is in its
  authoritative spec with no requirement left only in this file;
- every open decision of section 3 (D-2, D-3, D-6 to D-9, D-11 to D-15) is recorded as accepted, changed
  or declined, and a declined item is either removed from scope here or routed to a named owner;
- the P0 baselines and the after-evidence for slices B, D, E, F and G are archived under
  `Specs/TestRuns/` with their build receipts, including recorded skips for environments that were not
  available;
- the B2.8 review checklist is recorded with slice B;
- the exact Done path is admitted in the protected-history allowlist with its focused test.

**At closeout (2026-11-11).** The owner ordered the move to Done. Against the conditions above:

- Delivered slices have their tests and their contracts in the authoritative specs. Every requirement
  that was not delivered is an item of I63; none is left only in this file.
- The decisions are recorded in section 3. D-6 was decided; the others stand as built and I63 item RX1
  asks the owner to confirm them one by one.
- The P0 baselines were never captured, so this condition is **not met as written**. The ones that still
  have value are I63 items RP1, RD1, RP2 and RP3. The after-evidence (both full suites on the closeout
  build) is archived; section 18 names the folder.
- B2.8 was not recorded as a checklist during the slice. What is known: an occupant swapped between the
  last read and the rename is the accepted limit and is stated in the spec; a changed occupant publishes
  nothing and asks again, and a refused rename aborts the stage (both tested); a link or a directory at
  the target is never replaced at this level (tested for the link and the linked ancestor of a folder
  merge); that a reader never sees the final name missing was not witnessed. The remainder is I63 item
  RB11.
- The Done path is admitted in `Tools/Modules/Tooling/SpecInformationArchitecture.psm1` with the focused
  test `admits only the reviewed I62 File Operations UX closeout record`.

## 16. STOP

Stop and return to the owner when any of these holds:

- An activating press would change the selection without a modifier, or a modifier press would behave
  differently in the two panes.
- A comparison of names, sizes or dates would be shown or recorded as identity or as equal content, or
  would authorize a step without a typed decision or a policy the user chose.
- Level 2 would hold a handle on the occupant across the transfer, create a set-aside copy, or fall
  back to writing over a file whose replacing rename was refused. That is more, or less, than D-4 asks.
- A link, a reparse point or a type mismatch would be replaced at level 2 or 3, or a folder would be
  replaced instead of merged.
- P0.12 shows the replacing rename is not atomic on the registered share (a reader sees a missing name
  or partial content). Keep B1 and return level 2 to the owner.
- A destructive step would run on an item whose destination already existed with facts equal to the
  source's without the B8 probe, or after an inconclusive probe.
- The Permanent Delete by-name confirmation would be weakened. B7 removes only a Move's own source,
  after the publication proof; B4 removes a source only on its own explicit choice.
- Slice G would hold a handle or an authority for a set-aside item, pass its bound silently, remove a
  folder with an unresolved descendant, or change the buttons of a prompt the user can already press.
- The power request would be held while every task is finished, paused by the user or only waiting for
  an answer.
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

## 17. Out of scope

Replacement without a decision or a policy the user chose, and the `fileOverwrite` preference (I18);
the start-surface control for D-13 (I18); undo (D-15); persistent history and H3's other rows;
replacement of directories, links or reparse points at identity-less locations; rename-the-existing-item
choices; content comparison of the two items, in the prompt or for any policy (D-16); conditional bulk
policies other than B6's two and B4.5's; collection
of problems other than name collisions; remote, S3 and MTP staging names; a new scrollbar control; queue
persistence across restarts; the exclusive-transfer starvation and pause slot yield owned by I60 `W5.2`
and `W5.3`.

## 18. Implementation log

States follow the repository's evidence rules: *written* (code, tests, specification and documentation
edited), *built* (full Debug x64 build, zero warnings, attested receipt), *executed* (a named run with
its result). A slice is not done until its interactive cases have run.

Checkbox legend used at closeout: `[x]` delivered (a closeout line says where the delivery differs from
the item's text, or why the item was not needed); `[ ]` not delivered or partly delivered, with the I63
item that owns the remainder. Each state was checked against the working tree by a read-only audit on
2026-11-11, not by rerunning the item.

| Slice | State on 2026-11-11 | Evidence and remaining work |
|---|---|---|
| A | written, built, executed | Three Commands cases `cmd_pane_folderView_activation_press_*` pass in the full Commands suite. A4: the parent-notify handler is removed on the documented message contract (the high word is undefined for mouse buttons, the pane IDs are 1001 to 1004), not on a trace. A7: the renderer already draws the inactive selection and a thinner current-item border (`FolderView.Rendering.cpp`), read, not captured. The right-click and Quick Search cases of the test list are not written: neither path reads the pane state. The source contract in `TestHarnessSourceContracts.Tests.ps1` that pins `OnLButtonDown` is updated to the new text. |
| C | written, built, executed | Conflict policy, completed card, flyout counter, tests and gallery updated; the popup and gallery cases pass in both suites. A finished card with two commands keeps its `More...` menu; one command is a direct button. **D-6** (owner, 2026-11-11): Recycle escalation's `Delete all remaining permanently` and `Delete all remaining by name` are direct buttons immediately before Cancel, never the default and never the Escape action, with no added confirmation step; the policy contracts, the six submissions of `Phase10_*` and the identity-less family follow (`Phase10_` prefix 14/14 on the closeout build). |
| D | written, built, executed | New names and the D-17 recognizer (only names the current build writes; no legacy forms). The case-only rename hop uses `<leaf>.<32 hex>.rs-bak` (it carries the original), not the whole-leaf `.rs-tmp` form of D2. S4 (table-driven recognizer) is replaced by the smaller chain D-17 leaves. P0.9 baseline was not captured before the change. |
| E | written, built, executed | Engine: every admitted task holds its queue place (`_holdsQueuePlace`) until it enters or ends; the walk reads the task list instead of adding a second queue structure, so a missed removal cannot stall the queue. `DebugValidateQueueStartOrderForSelfTest` (seven scenarios) runs inside `FileOps_IdentityInterlockWaiting`. The popup cases (position text, move edges, card order) pass in the suites. A reading audit found no reverse lock nesting. E4 is not changed (not confirmed as a defect). E7 unchanged. E8: `FileOps.Queue.PassedOver` is emitted; `HeadHoldUs` is not. |
| B | B1 to B8 written, built, executed; see the limits | All self-tests below run inside `FileOps_ReviewRemediationContracts` (3/3 on the final build). **B1**: the refusal at a destination without item identity is a collision decision (`RunFileOperationIdentityLessCollisionDecisionSelfTest`, 14 variants). **B1.5**: a folder onto an existing folder there merges for a Managed Move too (Copy and the native Move already merged); one new guard exit, counter `FileOps.Conflict.DestinationIdentityUnavailable`; a link or a linked ancestor is still refused (`RunFileOperationIdentityLessFolderMergeSelfTest`, four cases). **B2.7**: the once-per-task note that items were checked by name, size and dates is the resource string `IDS_FILEOPS_NOTE_CHECKED_BY_FACTS` (it was an English literal until the closeout). **B2**: Replace there through the optional provider interface `IFileSystemExpectedOccupantPublication` (the bridge keeps its identity-owned stage; the provider renames it over the occupant only while the occupant's facts are the ones the user saw). A changed occupant publishes nothing and the decision is raised again; an occupant swapped between the last read and the rename is the accepted limit and is stated in the spec. **B3**: `CompareConflictItems` (two-second tolerance, one-hour class) published with the decision; the card, the decision mirror and now the Issues detail (B3.3) share one text builder. Not done: B3.2 (the provider issue callback carries no sizes), the facts on the finished card of a one-item task, exact bytes in a tooltip; an item decided by a remembered answer or a policy carries no facts. **B4** `Delete source` (`ConflictAction::DiscardSource`): offered by the bridge's file pre-flight for a Managed Move with a Local source when the existing destination is provably another file; the source goes to the Recycle Bin; Apply to all only for `Same size and date`. **D-5 rest**: where the source's location has no Recycle Bin the task asks once (`Permanently delete` / `Keep source` / Cancel, no Apply to all); the source is bound before the question and again with delete access after it, must be the same object with the same size and date, must still be another file than the destination, and is deleted through that bound object; any doubt keeps it (`RunFileOperationDiscardSourceWithoutRecycleBinSelfTest`, four cases, with the Local test seam `REDSALAMANDER_FILEOPS_RECYCLE_UNAVAILABLE_DIRECTORY`). Limit: `Delete source` is not offered when the source's location cannot identify items (a share without file IDs), so that source keeps the ordinary choices; non-Local sources and the native-rename Move route are not covered. **B5**: the trailing `task.result` Error is not counted when an Error of the task already explains the failure. **B6**: `Skip all with the same size and date` and `Replace all older` as arbiter grants (`RunFileOperationBulkCollisionPoliciesSelfTest`). **B6.5** (D-13): `OperationAdmission::presetCollisionPolicy` (Ask, Skip, Replace older, Keep both; default Ask) seeds the arbiter at start (`RunFileOperationPresetCollisionPolicySelfTest`); no control sets it (I18 owns the start surface). Limits: the replace counter counts decisions, the policies are offered on a one-file task too, translations are unreviewed. **B7**: a Managed Move whose source location cannot identify items is no longer refused. The provider reads the source's facts before the copy opens it; after the destination's proof the source is deleted by name only while the facts read on the deleting handle are the same (`IFileSystemIdentityDelete::DeleteIfIdentity`); an emptied folder is removed by name; anything else is "Copied; source kept", not an Error. A source carrying a Mark of the Web, another data stream, extended attributes, EFS encryption or a cloud placeholder is kept without a question, because the pathname reader copies none of them (new provider method `ReadUnidentifiedMetadataFeatures`). **B8**: after a replacement, the same-file probe (hold the destination denying all sharing, open the source by name) must answer "two files" before a by-name source is removed (`RunFileOperationIdentityLessSourceMoveSelfTest`, ten cases including one file reached through a junction). A defect found by that test and fixed: the transaction's own handles on the published file refused the probe, so every replaced Move kept its source; they are now released first. Limits of B7 and B8: a folder moved into itself under another spelling is not detected for such a source (the guard's ancestor checks do not run without a bound source); `BackupRead` on a real share is untested, and if it fails every source is kept; three metadata opens per removed source; B8.3 (a provider capability instead of the probe) is not done. |
| F | part written, built, executed | Rules as constexpr state machines with compile-time contracts in `RedSalamander/FileOperationsLayoutStability.h`, wired into the popup: reserved stream rows, sticky stacking and held note rows (F2), rate and ETA text hold and the kept per-file bar column (F3), `UpdateRates` once per 100 ms on motion ticks (S10), the scrollbar withheld while an automatic resize is pending or running unless the user asks to scroll (F5/F6 part). **F1**: the engine gives every transfer stream a row slot for its lifetime (lowest free slot, never compacted; the row array stays dense and the slot moves with its entry); the popup draws each stream in its slot and leaves an empty slot blank (`RunFileOperationStreamSlotSelfTest`). **F3**: tabular figures on the three numeric fact rows the popup lays out itself; the hosted graph's labels are DxUi's and are unchanged. **F4**: ETW counters `FileOps.Popup.CardHeightChanges`, `ScrollBarToggles`, `WindowResizes`. **F8**: the work summary is measured once per text, width and format. The 300 ms scrollbar delay of the target contract is dropped: it would keep the scroll position at zero for cases that scroll right after the first paint. Not done: the budget case `FileOps_PopupLayoutStability` and `LayoutShiftPx` (they need in-process counters and a fixed-tick clock for `Render`), a reserved width for the widest form of each number, the eased reveal (F5), the scrollbar gutter, posted resize and scroll anchor (F6), F7, F9. Found and not fixed: the height pass and the painter measure a wrapping work summary at different widths; `FileSystemReportFileWork` can match a stale entry by leaf after compaction. |
| G | G7 written, built, executed; G1 to G6 designed only | **G7** keep awake (D-14): one process-wide `PowerRequestSystemRequired` request while a task is admitted and neither finished, paused nor quietly waiting for an answer (`FileOperationsPowerRequest.h/.cpp`, per-task lease, pure predicate with compile-time contracts, `RunFileOperationKeepAwakeSelfTest` through a fake API seam). A wait for an answer stops counting only after the task reported no progress for two minutes, because other streams of the task may still transfer. Limits: preparation before admission is not covered; no self-test drives real tasks through the hooks; no Modern Standby machine was available, so `powercfg /requests` was not recorded there. **G1 to G6 are not implemented.** A read-only design found that a contained first version could set aside only nested file collisions of the bridge route: a plain Local Copy (a provider-owned walk) and selected-root files would keep their immediate prompts, the rule that stops every worker while a prompt is shown would be relaxed for collected prompts, and an immediate problem would have to take the prompt from a visible collected question. It needs five owner decisions (one answer per file or per class; preemption; deferring the queue yield; holding a parent folder of a Managed Move until the answer; accepting the coverage gap). Recommendation: a separate plan. |
| Verification before the closeout | executed 2026-11-11 | Debug x64, zero warnings. Full FileOps without the second-volume switch: 294 passed, 1 failed, 28 skipped; the failure was `R4A19_DiscoveryIndependentVolumes` (`alternate-volume-test-root-not-authorized`). That reason was reported as a machine limit and was not one: the run omitted `-AllowAlternateVolumeTestRoot`. Full Commands: 966 passed, 0 failed, 5 skipped. **One unexplained failure**: on the build before that one (identical except for one test log line) the full FileOps suite also failed `DiscoveryScope_MultipleFoldersLiveProgress` once: the copied content was correct and one summary number of the plain Local Copy (result code, discovered files, discovered bytes or copied files) was not the expected one. It did not reproduce in 36 in-process repeats, three isolated attempts, three family-order attempts or the next full run, and the case had never failed in the retained history. The copy route of that case is not changed by I62 except for the row slot of slice F, which does not touch the counts; the cause is not established. The case now logs `MultipleFolderLiveProgressSummary` with the four numbers when it fails (I63 item RW1). |
| Verification on the closeout build | recorded in the archive | The closeout build adds the D-6 direct button, the localized by-facts note and this closeout's planning and tooling edits. Its verification is one Fresh Full run with the interactive entries and `-AllowAlternateVolumeTestRoot`, started after this text was written, so its result is not stated here. The reviewed record, with per-case results and every skip reason, is `Specs/TestRuns/4cb089111a23/FileOps/2026-11-11_i62-closeout/README.md`. If that folder is absent from the commit that carries this file, the run was not archived and this plan claims no result for the closeout build. |
| P0 | folded into the slices | P0.1 to P0.15 were not run as a separate slice. Each slice records what was read instead of witnessed; P0.13 was answered by reading for the slice G design (a shown prompt stops every worker of the task through `Task::WaitWhilePaused` and the callback mutexes). |

Closeout (2026-11-11): the owner decided D-6, moved slice G1 to G6 to I63 and ordered the move to Done.
`Delete source` from a location without item identity is I63's open decision R-D1. Everything not
delivered here is listed in I63, Part 2.
