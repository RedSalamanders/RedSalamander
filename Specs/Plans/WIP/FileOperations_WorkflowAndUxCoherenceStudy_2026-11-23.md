# I66 File Operations Workflow and Plugin Contract

**Status: ACTIVE.** Final target specification and delivery plan.
**Priority: P1.** Restore ordinary Rename first. Land the small ingress fixes and the progress correction independently.
**Planned at:** `ac619f860` (2026-11-23). Run the [drift check](#drift-check) before each delivery.

## Implementation checklist

**Current work (2026-11-24).** Local Recycle batching reached context-menu, Find and Viewer admission. Find and Viewer
each used one Shell batch for twelve sibling files, and Viewer Cancel kept the files. The repaired same-folder Ctrl+drag
passed through its exact no-activation selection. The branch also contains the first exact Batch Rename handle-lifetime
and throughput changes. Implementation started at `bbb57a572`. Drift from the study baseline affected the ancestry-cleanup
test and its evidence contracts.

[PR #246](https://github.com/RedSalamanders/RedSalamander/pull/246) merged into `main` at `bcaa59d10`.
Its non-activating x64 Debug PrePush executed 14 entries with 1,799 passes, zero failures and 27 environment skips.
[The first main Release attempt](https://github.com/RedSalamanders/RedSalamander/actions/runs/45405656554/attempts/1)
timed out downloading a Ghostty build dependency on both architectures before native regression tests.
[Attempt 2](https://github.com/RedSalamanders/RedSalamander/actions/runs/45405656554/attempts/2) hit the same timeout.
Four desktop-lease entries and the qualification below remain open.

The follow-up x64 Debug PR selection on `842ac811a` executed 12 entries with 1,507 passes and zero failures.
Repository coverage remains `NOT_EVALUATED`. The current candidate also includes Queue/Parallel persistence,
bulk-pause ownership and a dedicated live SMB Rename case. Correcting the SMB child-family registration cleared the
first build failure. The selected x64 Debug build on `6df121e23` completed with zero warnings and zero errors, and
the non-activating popup and repeat-decision selections passed. A guarded Win32 check on both SMB endpoints
showed that a read-data watcher is required to reproduce the sharing violation. The earlier PrePush stopped after
its build so the corrected witness can run on the Windows endpoint before the next PrePush.

- [x] **Intake.** Preserve the source study, add the WIP index entry, and check source drift.
- [ ] **S0.** Finish ingress parity and qualification.
  - [x] Verify folder-picker Move, Viewer Recycle wording and safe cancel, and Cancel all in the focused x64 tests.
  - [x] Rerun the same-folder Ctrl+drag guard through its exact no-activation selection after the Windows provider instance fix.
  - [ ] Expand clipboard, blocked Retry and read-only case coverage as needed for the workflow walkthrough.
  - [ ] Verify the implemented short Copy-only Move question across mixed plans, changed endpoints, and Cut/queue
    preservation. The focused source-retention case passed.
  - [x] Route Local Recycle through the existing bulk executor. The focused Find and Viewer runs verified prompt counts,
    preserved files on Cancel, and compared Shell calls and settled resources against the per-item route.
  - [ ] Exercise the context-menu surface in the walkthrough and repeat the Local measurements on Release.
  - [ ] Implement typed per-item Recycle availability and direct permanent consent for known-unavailable routes.
    Prove mixed-group routing, consent scope and unchanged handling after an unknown result.
- [ ] **S1.** Qualify Rename by route.
  - [x] Release each exact Batch Rename handle after preflight and validate captured identity on every mutation
    rebind. The local exact-result test passed with a no-DELETE-sharing witness at the preflight boundary.
  - [ ] Qualify exact Local and SMB for nested rows, case-only hops, legacy API fallbacks and blocking watchers. Compare
    single- and 100-item Rename costs, including handle peaks and cancellation.
  - [x] Run the Local exact Rename refusal, legacy fallback, collision and lost-response controls.
    [The x64 Debug record](../../TestRuns/Local-x64/FileOps/2026-11-24_140134/README.md) retains their timings and receipts.
  - [x] Register the approved SMB test directory and confirm a read-only listing and mapping through `remote_file_smb`.
    The case passed in run `20261124T143343Z-95336-b6a780d1e7104d7092573afd510ba4d0`.
  - [x] Exercise the exact Rename eligibility check on the registered SMB endpoint. The Local provider opened the
    owned file and corroborated that exact file identity was unavailable. The Rename case skipped in run
    `20261124T145318Z-86896-cf65eb1fac8447518ff7509d6d3dfe51`.
  - [x] Prepare the separately approved Windows SMB test directory.
  - [x] Compare the namespace guard on both endpoints in diagnostic run `20261124T150306Z-85228-8ce4ce9864d743b0bd18f932d4bb5879`.
    Read-data handles denied competing DELETE opens on the owned parent and ancestor. Absolute no-replace Rename
    preserved the occupied target and moved the file to a free sibling. The first endpoint returned `ERROR_INVALID_PARAMETER`
    for `FileIdInfo`. Windows SMB returned an exact ID. These observations leave the host route and full namespace proof open.
  - [ ] Run live exact Rename on Windows SMB with the corrected watcher and implement and qualify the corroborated route.
  - [ ] Verify FTP and SCP with isolated test credentials. The dedicated profile selectors are registered and the
    scoped authentication setup was blocked before a live run.
  - [ ] Implement and qualify corroborated same-parent Local/SMB Rename without file IDs on Windows SMB and Samba,
    including changed parents and decoys.
  - [ ] Qualify new-name Rename on Curl, Drive, and Graph and its result receipts before enabling each protocol or
    profile. Verify F2, Batch Rename, Change Case and unsupported-route cases.
- [ ] **S2.** Complete progress projection and paired measurements.
  - [x] Preserve sampling through decision, live-output, pause and queue transitions. Freeze history while stopping
    and keep a running peer's ETA. The model and paced idle-cadence selections passed.
  - [x] Verify the Preparing elapsed clock at task creation and after directory Delete reservation. The focused x64
    Debug popup and admission cases passed without desktop activation.
  - [ ] Cover credential waits and align the popup and taskbar snapshots.
  - [ ] Collect paired allocation, timing, history and active-peer evidence.
- [ ] **S3.** One repeat control, coalesced cleanup and read-only decisions, overlap and sensitive-loss handling.
  Match card and mirror scopes and authority, including stage cleanup and post-publication results.
  - [x] Rerun the existing repeat-decision contract. The x64 Debug selection passed in
    `20261124T145409Z-86896-a97bc8a5503d4a98b153e3aba31f7647`.
- [ ] **S4–S5.** Queue controls and persistence, shared cap, completed grouping, mirror focus and Issues entry.
  Recheck I60/I64 ownership and verify interlocks, pause provenance, retention, UIA and focus.
  - [x] Verify that the popup stores Queue/Parallel and a new owner reads the saved default.
  - [x] Verify bulk pause ownership in the engine and popup model. Preparing and queued tasks are included.
    Resume all preserves the earlier individual pause and queue gate. The paired cadence measurements and
    save timings are in [the focused x64 Debug record](../../TestRuns/4cb089111a23/FileOps/2026-11-24_155152/README.md).
  - [ ] Verify the Start option's one-submission scope and the new pause behavior through the full interactive
    walkthrough, including focus, UIA, and a pending mutation.
- [ ] **S6–S7.** Options parity, asynchronous Create Directory and archives, and useful results. Preserve saved
  settings. Qualify worker cancellation, exit, refresh and cleanup.
- [ ] **S8.** Keep the owning specifications, consistency ledger and user guide current. The first ingress, exact
  Batch Rename, throughput, Move with Options, and Cut/Paste updates are in the branch.
- [ ] **Closeout.** Complete the walkthrough, native and live-provider gates, document checks and reviewed history
  admission. Move to Done after all required rows have evidence or an accepting owner for the exact transferred gate.

### Validation checkpoints

- The first exact Rename candidate `ab51279e6` stopped at compilation in
  `20261124T134444Z-100080-5b6b4d2e1e1046718998105aec11af0c`. WIL's final-path adapter requires its owned string type.
  The corrected candidate uses `wil::unique_cotaskmem_string` for the spelling proof.

- `dc30fb46b` built with zero warnings and zero errors using `build.ps1 -Configuration Debug -Platform x64
  -ProjectNames RedSalamander` with tests enabled. Build receipt
  `04985ac959f56c20a9f2a40d8f0613490866797615e6e138c911876fba3b7bd0`.
- Focused x64 Debug runs used the marked local sandbox. The runner reported `NOT_EVALUATED` at repository scope.
  The passed selections and the initial failed witness remain in its completed-run archive.

| Selection | Result | Run ID |
|---|---|---|
| `Commands_FileOpsFolderPickerMoveUsesHostQueue` | Passed | `20261124T080347Z-60876-a29df09ceaeb4316b8894cfd894b2c9e` |
| `folderView_drop_integrity_guards` | Failed on same-folder Ctrl+drag in the first candidate | `20261124T080442Z-60876-402869228c9c4feeb44ff233aad4da72` |
| `FileOpsFamily_BatchRenameExactResult` | Passed, including the post-preflight sharing witness | `20261124T080545Z-60876-bfc378cf17ad4c1ab786cf4527749906` |
| `FileOps_CopyOnlyMoveRetainsSourceWithoutDestinationProof` | Passed | `20261124T080633Z-60876-ff9be60a22e34d6eb6880fdfdf6e8987` |
| `Phase10_PermanentDelete` | Passed, including the Recycle confirmation snapshot | `20261124T080714Z-60876-d01ca01cc0aa43218ba40875855a2e8e` |
| `FileOps_RateWindowDisplay` | Passed | `20261124T080756Z-60876-09d1cdd6f4c049b79b05e52c6d9e390a` |
| `FileOps_CancelAllNeverStarted` | Passed | `20261124T080835Z-60876-ea661fa63ee5419c97544e2d3586433a` |

- The same-folder repair `56a5d9d08` also built cleanly in test-enabled x64 Debug. Its receipt
  is `ec35a64c2ee51216f7ffc47c35477e25558c67d6eb85c3b849788fd5a9544d56`.
  The approved first focus window ended before this build finished.

| Selection on `56a5d9d08` | Result | Run ID |
|---|---|---|
| `FileOps_RepeatDecisionContracts` | Passed without activation | `20261124T081614Z-64872-f5cb2a98b9c542b181b3b8a97ec8192d` |
| `FileOps_IdentityInterlockWaiting` | Passed without activation | `20261124T081715Z-64872-a7b1ace66f8f4fb58da7eb79d89dabd8` |
| `FileOps_PopupIdleCadence` | Failed because the tiny fixture reached publication before a new rate could be measured. The paced rerun passed on the candidate below. | `20261124T081817Z-64872-e3f005a81f5c48649ca3a0dd4aaff921` |

- The full solution built test-enabled x64 Debug on `0211122e8` with zero warnings and zero errors. Receipt
  `782884d4bec483414e035036c1b48f80cd71bd8b1ed4ab10e98f21625ff2a9c4`. The paced
  `FileOps_PopupIdleCadence` selection passed without activation in run
  `20261124T082752Z-59536-98f34ae1390f465989d44c3027b3b8b1`.
- Non-activating `Test-Changes.ps1 -Mode PrePush -Configuration Debug -Platform x64 -SkipBuild` on `0211122e8`
  executed 14 entries. Thirteen passed their entry gates. The run recorded 1,797 passed, two failed, and 27
  skipped cases. The Pester failures matched an obsolete transition tick and a fixed character distance
  in the live-output wait guard. Run `20261124T082910Z-59536-199bace9c93c4fffba708267a19c2e54`,
  final receipt `30e70fc85e83d409cf324d7d60cb2f17aaf93b1d65b04b0fa9ecda6bd23a09ad`.
  Commands, File Operations and both interactive viewer selections remain pending.
- After the guard repair, `TestHarnessSourceContracts.Tests.ps1` passed 243 cases with zero failures or skips
  in Pester 3.4.0. Run `20261124T092118Z-19224-3e7f22bb4d984e5fba1a4bdae44c663e` binds native sources at
  `0211122e8` and the updated test file SHA-256
  `aab6ae96a3e1b4179b64755fd50590d18604b37b13970ef4eb07f271f526b23b`.
  A current-head PrePush rerun remains open.
- Same-machine x64 Debug cadence comparison used baseline `bbb57a572`, verified receipt
  `aad8a6aefb2102d3f24a592ccb4ceb9c620f78f149bed0f39bb773a66376956a`,
  run `20261124T091748Z-75364-035371b12b9344f0a32412ced821a0ac`,
  and candidate `0211122e8`, final PrePush build receipt above,
  run `20261124T091911Z-69464-1f7f62f41fb24db186fcc04be6f8fa9d`.
  Both focused selections passed. In each unchanged synthetic interval of about 1.2 seconds, running
  recorded 11 ticks and 11 renders, paused, prompting and queued each recorded one tick and one render,
  and completed idle recorded zero. The changed live-copy workload requires a matched workload for a
  cost comparison. Allocation, retained history and concurrent active-peer measurements remain open.

- The first Recycle fixture candidate `fcbcea4ce` failed to compile because it passed an integer to the steady-clock
  timer. The corrected candidate `f5aa393f0` rebuilt with zero warnings and zero errors in x64 Debug. Receipt
  `9fb00c721b37714ca6f6f3dff919e0ddb520d01a822043a34ecd0954a3ba6804`.
  Run `20261124T102043Z-60344-b9aaba78a56e498da31ba7e65186af3b` passed the Local Recycle routes, one-prompt checks
  and exact purging. The [paired ingress evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-24_112639/README.md)
  retains the call counts, time, handles, private bytes and focused scope.
- The no-activation policy and all 243 source contracts passed in
  `20261124T102042Z-59168-e61c6ec3728c48cfae1399bc22ef9065`. The native Commands
  `folderView_drop_integrity_guards` selection used the same x64 Debug receipt, reported zero pending interactive
  entries, and passed in `20261124T102827Z-74852-3e73f64a46ac464cad850a29feed1a0f`. Both focused native runs
  reported `NOT_EVALUATED` at repository scope.

- Non-activating x64 Debug PrePush on `ee6b50125` built the full solution with zero warnings and zero errors.
  Run `20261124T103514Z-56072-2c2e7c56ea8d414db4ac0a1efb117d4c`, final receipt
  `0ebf0408ac73876d33fdccf4295efff3dd8917ee68a5b5415c53564d82d8fb6b`. The run recorded 1,798 passed, one failed, and 27 skipped.
  Thirteen of fourteen entries passed their gates. The unchanged Git metadata fixture found a `sleep` process
  when it looked up an exited PowerShell probe by numeric PID. Its focused group passed all three cases on the
  same head in run `20261124T111258Z-31232-3274593ad723444a99aecc4248262b88`. The broad verdict remains `FAILED`.
  Commands, File Operations, and both interactive viewer entries remain pending. The sandbox audit reported
  the same 13 existing findings as the earlier broad run.

- The test-enabled x64 Debug build at `9d4b49d71` used receipt
  `43a396a5e2b5167cea0ff4a6b0182faf99f8f0ab718727869e833ec1e1e24d20`. `FileOps_PopupIdleCadence` and
  `FileOps_ReviewRemediationContracts` passed without activation in runs
  `20261124T115346Z-19596-42d9378fb9594838b3962c7997e4d76a` and `20261124T120125Z-12688-ad787f3e6bed48438c508063879e91db`.
  The [paired Preparing evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-24_130014/README.md) records matching synthetic cadence and a 16-byte increase
  in the measured task object. Repository coverage remained `NOT_EVALUATED` for these selections.

- Non-activating x64 Debug PrePush on `b4bdf0086`, run `20261124T120545Z-65304-4582af4137ca4af0ae165551a14d61f7`,
  used final full-solution receipt `2b5b5c2dbda57d43a2e8e7f927e99782e335afbb224d75eee2899c208dbf8c92`.
  All 14 executed entries passed their gates. The run recorded 1,799 passed, zero failed, and 27 environment skips.
  Tools Pester passed all 1,496 cases, including the earlier Git metadata fixture. Change coverage was `PASSED`.
  Repository coverage remained `NOT_EVALUATED` with Commands, File Operations, and both interactive viewer entries
  pending. Release and live-provider evidence remain open.

Each delivered change updates the implementation, tests, and owning specification together. The current domain contracts
govern shipped behavior until those changes land.

## Target workflow

Submitted actions use the same File Operations task, with the existing delayed card reveal for quick success. Ask only when
the answer changes what happens. For task decisions, the card and Decision Mirror share one question and accept one answer.
Use ordinary progress for waiting and one result for a failure, partial operation, or uncertain outcome. Keep technical
reasons and optional diagnostic facts in the existing Issues details. Provider capabilities never create a new preference,
success notice, or confirmation.

| Action | What the user sees |
|---|---|
| F5, F6, Paste, drag/drop, or a destination picker with a supported route | Start with the captured defaults. A successful native Move on the same FTP/SSH endpoint remains a Move. `Ctrl`+drag within a folder copies under a different name. |
| Routine Move that can only copy | Before queueing, ask “Copy the items and keep the originals?” with **Copy** and **Cancel**. For a mixed selection, show the move count and the copy count with originals kept, with **Continue** and **Cancel**. Include the destination. |
| Shift+F5 / Shift+F6 or an ingress that explicitly requires options or approval | Keep one start surface. Show only relevant options and include any planned Copy-only items in that surface. Name Shift+F6 **Move with Options**. |
| F2, Batch Rename, or Change Case | F2 edits the name and commits on Enter. Batch Rename keeps its preview and Change Case its options. All three execute through the same task and use the same policy for a collision. A folder keeps its contents and is renamed as one item. |
| Delete at a location with Recycle support | Move to the Recycle Bin. When the ingress requires confirmation, as in ViewerSpace, say **Move to Recycle Bin?**. |
| Shift+Delete, or Delete where Recycle is known unavailable | Ask **Permanently delete?** once on the task after preparing the admitted items. Skip the Shell Recycle attempt when the provider has proved it cannot work. |
| A task cannot run yet | Show the reason on its waiting card and allow Cancel. Ask about overlap only in Parallel when the person can actually authorize a different safe outcome. |
| An issue needs a decision | Offer the actions the provider and task can fulfill. Cancel is the safe default where an affirmative action changes data. A decision that would be alone in `More…` is a direct action. |

### Start and delete behavior

- Use one lightweight host confirmation on the existing routine transfer start path for a known Copy-only Move. Ask
  before creating, publishing or enqueueing a task, and before consuming Cut. Bind acceptance to the captured selection,
  plans and endpoints, then revalidate. Accepted Copy-only items stay source-kept even if a later probe enables Move. A
  changed endpoint, selection or plan that would widen the accepted operation needs fresh authority. Cancel leaves the
  cut sequence, task queue and filesystem untouched.
  Explicit options and any already-known risk that can share the same receipt appear in the same confirmation. If the
  preparation or execution proof later downgrades a Move, keep the source and follow the per-item decision gate before
  affected publication where possible. Report the actual source and destination outcome if the downgrade is discovered later.
- Cut, Paste, pane menus, Find, Viewer and Compare keep the source and destination roles they captured. A local folder picker
  returns a local destination explicitly, even from a remote pane. Clipboard Copy/Cut offers Windows file-drop data only
  for representable paths. An unavailable command leaves clipboard contents intact and uses the ordinary pane feedback if
  invoked by shortcut. Same-folder drag Copy validates provider case rules, rejects source aliases and uses the existing
  sibling naming policy.
- Use typed per-item Recycle availability (available, unavailable, or unknown). Keep the qualified bulk Local Recycle
  route for pane, context menu, Find and Viewer when applicable. A mixed selection separates Recycle and permanent groups
  within one request and identifies the items covered by permanent consent before their removal. If Recycle is unknown, the ordinary guarded
  attempt may run. Escalate only on a proved non-commit with the original still present. A known removal or a disputed
  result never grants another delete. Apply this rule to ordinary Delete and the existing Move “Delete source” choice.
- Keep exact or corroborated Delete authority, including its existing by-name consent when required. Explain permanent
  removal and any material by-name limitation in the question itself. Keep method details in Issues. A read-only override
  is an explicit decision only after a provider proves the blocker. `Apply to all` remembers it only for matching typed
  scopes in the task, with a fresh per-item proof. Retire the obsolete blanket `noncached` wording in the File Operations
  spec. A generic access denial and an uncertain mutation never authorize the override.
- Fold Leftover Cleanup and archive delete-after disclosure into their existing confirmation surfaces where the entire
  set can be prepared. Preserve the exact artifact `TouchGuardReceipt`, capture the approved items, and revalidate on the
  worker after waiting and before mutation. Keep the existing by-name handoff rules for Leftover Cleanup and the exact-only
  rule for archive cleanup. Pack/Unpack deletes eligible sources only after archive success and recheck. A changed set needs
  its own current authority or a failed/kept result. UI coalescing grants no additional deletion rights.

### Collision and risk decisions

- Use one repeat checkbox, **Apply to all**, on both card and mirror. Checking it and choosing an eligible action applies
  that action to remaining matching conflicts in this task. When **Skip item** is the only repeatable action, place the
  checkbox with Skip and label it **Apply to all when skipping**. **Retry** attempts only the current item and creates no
  grant. A new question starts unchecked. Remove the redundant generic **Skip all** and omit the checkbox when no later
  eligible item or repeatable action exists. Reuse the existing typed task, class, item and endpoint scopes, including the
  non-mutating failure families, and check that every later action is still allowed.
- Keep genuinely different actions. **Skip items with the same size and date** and **Replace all older** apply their own
  comparison rule. **Delete source** is the existing qualified Move choice with its own scope and consent. A destination
  that is out of space can offer Retry and a scoped skip. Where the provider can only stop work for that destination,
  offer **Skip remaining at this destination** as a direct action. Existing explicit bulk permanent-delete escalation and
  by-name grants remain separately labeled and never default. They are separate operations or scopes with a different result.
- Offer **Keep both** on a name collision when a qualified unique sibling can be made using the provider's collision key.
  An observed file collision offers **Replace** only where the exact approved occupant can be rechecked and consumed,
  with read-only protection and artifact receipts as applicable. Rename never merges folders or offers folder replacement.
  Replacing a link requires the existing exact-bound **Replace link** action and its explicit authority.
  Remove Retry for a path-length failure with no actionable retry. Explain a read-only occupied destination as one fact.
  For F2 with no later item, a single Cancel already leaves the item's name unchanged.
- Combine known risk decisions for the same item before content is published. Keep explicit consent for known sensitive
  loss, including Mark-of-the-Web (`Zone.Identifier`) on Move, EFS, other named streams and extended attributes. I18 owns
  any change to which losses require consent. Use a qualified owned stage when metadata can only be assessed during
  transfer. A direct-final route needs proof that publication satisfies the captured policy. Apply the same consent
  rules to sparse inflation and hydration before either action starts.
  A real pre-publication **Skip item** leaves the source and existing destination intact and accounts for the stage. If sensitive loss
  surfaces after publication without a matching prior grant, keep the Move source, show **Copied; source kept** with the
  reason in details and continue only where safe. Preserve the actual committed destination and any uncertainty.
- Plain-language copy describes the choice or outcome, including cancellation of running and waiting tasks. Use existing
  contextual failure text and put HRESULTs and provider details in Technical details. Normal capability limits, a routine
  wait, and a corroborated successful rename add no separate alert.

### Queue, progress, and results

- In Queue mode or when concurrent overlap cannot be approved, wait behind the captured predecessors using the existing
  overlap edge and `_queueOrderKey`.
  In Parallel, ask only when `CanRunConcurrently()` exposes a choice that changes execution. Revalidate all interlocks and
  capacity independently of the answer. `Start now` can release queue ordering for that task, never a safety or capacity
  wait. If the state changes as it is pressed, the task stays in a recoverable queue position with a precise wait reason.
- Persist **Queue / Parallel** as the default for newly submitted work, with Queue on a fresh profile. Task start options
  override that choice for one submission. Preserve one process-wide transfer budget. Per-job Auto probes storage and the
  provider, and admission caps combined work at the process budget, default four, including internal file workers and
  opaque reservations. Explain this in the existing setting. A separate change to the numeric limit needs paired
  evidence and preserves the shared bound. Use the existing I60 pause, fairness and admission ownership.
- Show Pause while a transfer runs, including during discovery-ahead. Pause all applies to current live tasks, including
  those still preparing or waiting. Preparing may finish nonmutating checks while its mutation remains paused.
  Resume all releases the pauses set by that action and leaves independent prior pauses, decisions, and overlap edges
  in place. Per-task Pause/Resume remains available. Release transfer capacity only after explicit pause reaches its
  quiet point, and reacquire before I/O. A question retains its existing reservation under I60.
- Exclude decision, credential, queue/overlap and live-output wait time from throughput and stall sampling. An actual
  I/O stall during active work still hatches. Publish entry and exit of a blocking interval through a monotonic generation
  in the existing task snapshot, including when the popup is hidden or another task is running. Use one blocked-state
  projection for sampling and the taskbar. A collected collision that lets work continue does not block telemetry. Stopping retains the recorded graph
  while genuine counters can finish updating. Preparing reports real elapsed time and a paused task does not suppress or
  distort the running-only ETA.
- Keep failures, partial/uncertain outcomes, and completed tasks with warnings visible outside the clean Completed group.
  A clean task with only `Information` notes follows the ordinary grouping, auto-dismiss and clear rules. Show any new
  routine notes only in Issues when informational notes are enabled. `Clear completed` acts on only the clean
  completed/canceled IDs covered by its painted header. Failure details and interrupted-Move recovery notices keep their
  existing explicit acknowledgement. Use one Issues entry in the popup's existing Options menu and the existing per-task
  path to details.
- The Decision Mirror stays passive while an edit field, terminal, or another window has focus. Make its existing status
  entry and keyboard route usable for explicit focus. It shares controls, target inspection and scope help with the card.
  Preserve safe focus restore, no accidental acceptance by a held key, and one answer per question.

### Options and other entry points

- Use one bandwidth preset list and Custom parser for Preferences, explicit start options, and the running card. Display
  IEC units and preserve existing saved rates and accepted input equivalents. A File Operations Copy link-policy default
  can override its existing provider default at admission. Preserve inherited behavior on upgrade until a user sets it.
  An explicit per-task Copy choice wins. Native Move retains its provider semantics.
- Capture verification availability in the plan. On routine defaults, indicate unavailable verification in outcome details
  if requested by the default. Explicit Verify stays disabled on an unsupported route, and a newly unprovable required
  verification gets a truthful result, including source retention when the Move requires that proof.
- Create Directory uses one name prompt, then validates and executes asynchronously with the existing task, interlock,
  exit, cancellation and result rules. Refresh only the admitted location, and use the same duplicate-name policy and
  plain errors. Label Create Directory as that operation on the task and in Issues. Use the same worker ownership and
  completion rules when moving Pack/Unpack off the UI thread. Preserve
  the archive format/overwrite options and cleanup authority. `Copy to folder…` and `Move to folder…` can share the existing
  destination picker and typed request when exposed. Routine informational tasks use the existing reveal policy.

## Shared file-system plugin contract

The host owns intent, admission, queueing, consent, progress and presentation. Each provider owns its namespace rules,
metadata, mutation and result proof. Use the typed `IFileSystemRouteCapabilities` query for the exact provider instance,
root, profile, path and operation. Capability JSON and connection names remain diagnostic. Missing or malformed executable
facts disable that route. Use the existing task, plan and conflict broker for every admitted operation.

- For Copy/Move, qualify Copy, native Move, managed publication and source cleanup independently. A Copy-only request
  cannot reach source deletion. For Delete, Recycle, permanent removal, and any read-only or by-name decision are distinct.
  For Create Directory, use the same child-name and namespace contract as Rename. Provider verdicts describe the item and
  phase that occurred, regardless of the ingress or whether the provider implements object binding.
- Make executable Rename capabilities explicit per route. The base Rename claim requires a qualified new-name operation.
  The typed route distinguishes mutation-time refusal of an occupied name, a destination probe, and coexistence by
  item ID. A separate capability qualifies replacement of approved regular-file occupants on a new unbound route.
  Advertise supported scope, path and collision rules, case-only behavior, link kind and cancellation. A provider
  can support new names while declining Replace or recursive Change Case. Admit only a proved bounded or watchdog route.
- Keep the existing `RenamePlan`, worker and scheduler. Exact-object execution uses
  `IFileSystemBoundObject::RenameIfUnchanged`. Give qualified unbound providers one shared typed route contract and one
  versioned provider extension where the existing calls cannot carry the witnesses and result. A new interface uses a new
  IID. Keep existing vtables stable. Use size-versioned records and validate values, lifetimes and callback ownership.
  Rebuild and test the host, every shipped file-system provider and shared ABI consumer in source-tree lockstep when the
  contract changes. Each provider supplies a tested implementation or an explicit unsupported result. Align singular and
  bulk Rename entry points with the same provider rules. A legacy flag-only call never bypasses the checks.
- Reuse the canonical provider name validator, join, case comparison and collision key. Keep the typed source/destination
  and interlock scope through queueing. Revalidate the source, parent and root using the proof required by the admitted
  route. Keep exact source/destination proofs wherever available, and use conservative overlap scopes with unbound
  ancestors. Diagnose observed collisions before touching the occupant. Refuse unsafe aliases and linked parents.
- Require an authoritative no-follow kind for each selected root and discovered child. Recurse through proven directories
  only. Symlinks, name-surrogate reparse objects and Drive shortcuts are leaves. A provider can rename the leaf itself if
  qualified. Preserve the existing classification for proven non-name-surrogate reparses. When kind, parent or containment
  cannot be proved, refuse the affected recursive request before mutation.

### Rename authority and qualification

| Route | New-name Rename | Replacement |
|---|---|---|
| Local and identity-capable SMB | Retain exact no-follow authority and perform the conditional rename. | Preserve the existing exact-bound grants for regular files, read-only files and links when eligible, with the corresponding explicit collision decision. |
| Identity-less Local/SMB | After a successful no-follow open proves `FileIdInfo` unavailable, reuse the by-name Delete corroboration and bounded interlock. Check the captured facts again on the retained source handle. Use an exclusive rename within the admitted parent under the handle rules below. | Deliver new-name Rename first. A later Replace path uses the shared expected-occupant contract and requires its own qualification. |
| FTP, SFTP, and SCP (Curl) | Use the existing same-endpoint R0f-Curl destination probe followed by the native rename command. Qualify observed-collision protection on each shipped protocol route, including SCP's SFTP control path. The server can replace an occupant that appears after the probe. | Keep Replace unavailable while the provider cannot consume an approved occupant and finish cleanly. |
| Google Drive | Rename the resolved item ID without moving children, trashing an occupant, or deleting the source folder. Check actual sibling names as well as exposed names. An ambiguous raw target gets a collision. Preserve exact exposed-name resolution, literal suffix-like names and case. If an external creator adds the same name after the check, preserve both IDs and use Drive's existing disambiguation. | Withhold until its expected-occupant and cleanup path passes independently. |
| OneDrive / SharePoint (Graph) | Use the resolved source ID and prove the provider rejects a conflicting name without changing the occupant. Support exact final case where advertised. Keep `allowDirectoryMerge = false` for Rename. | Withhold until its expected-occupant and rollback/cleanup path passes independently. |
| Dummy, S3, S3 Tables, MTP, 7-Zip, IMAP, and any other loaded profile | Dummy uses the contract for deterministic tests. Other affected profiles advertise unsupported wherever the operation is absent or unqualified. Test those negative routes. Native Move, Copy+Delete or writer staging alone cannot advertise Rename. | Use the same eligibility rule if a future route is implemented. |

#### Curl destination check

Reuse the existing direct Rename path from [R0f-Curl](../../FileSystem/FileSystem_FtpSftpScp.md#conditional-publication-and-identity-safe-recovery)
for F2, Batch Rename and Change Case. Revalidate the admitted source and namespace using the facts this provider exposes.
On the worker, probe the destination fresh immediately before the server rename. An observed occupant returns a known
non-commit collision with no rename or backup command. A failed or ambiguous lookup stops execution. Keep both uses a
different validated sibling and repeats the check for that name. Preserve the existing connection, host-key, deadline,
cancel and unload rules. Standard FTP `RNTO` has no portable no-overwrite selector, and curl's libssh2 quote path supplies
overwrite flags. [FTP command specification](https://datatracker.ietf.org/doc/html/rfc959#section-4.1.3) ·
[curl SFTP quote implementation](https://github.com/curl/curl/blob/curl-8_23_0/lib/vssh/libssh2.c#L1883-L1893)

The Curl host exception covers ordinary same-parent Rename to a probed free name on the same endpoint. If an external
client creates the destination after the probe, the server can overwrite it. Host interlocks only coordinate
RedSalamander work. A probe grants no atomic no-replace, replacement, exclusive temporary-hop or rollback authority.
Qualify the outcome that
each protocol actually produces, including a lost server response. Document the limitation in the provider contract and
user guide. Routine success stays quiet. The same delivery reconciles the general no-replace wording and the R0f-Curl
provider exception while preserving the existing Copy/Move contract. Google Drive has a separate outcome: late duplicates
can coexist by item ID. [Google Drive Files resource](https://developers.google.com/workspace/drive/api/reference/rest/v3/files)

#### Local and SMB corroboration

Capture kind, relevant attributes, size where meaningful, timestamps and available nonzero file-index/volume facts.
The fallback starts only after a qualified successful no-follow open with unsupported `FileIdInfo`. Interface
unavailability, missing objects, access/sharing failures, unknown tags and malformed results
retain their own outcomes. A replacement with identical facts cannot be distinguished by this method.

Qualify the same-directory leaf form of `FileRenameInfoEx` or `FileRenameInfo` on the retained source handle, with
`RootDirectory = nullptr` and replacement disabled. A qualified parent-handle-relative form may be used where supported.
SMB2 requires a zero root on the wire, so test how Windows redirects the leaf on Windows SMB and Samba, including the
working directory, share root and ancestor changes. The existing absolute-path buffer is eligible for this route only
with a proved namespace guard through the call. Refuse when the admitted parent cannot be protected. Record the method
as diagnostic detail where requested. Successful Rename has no by-name alert, and permanent-delete consent stays
explicit. [Windows rename forms](https://learn.microsoft.com/windows-hardware/drivers/ddi/ntifs/ns-ntifs-_file_rename_information) ·
[SMB2 rename structure](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-fscc/52aa0b70-8094-4971-862d-79793f41e6a8)

### Replace and mutation results

- Deliver restoring Rename to a free name independently of adding any new Replace capability. New remote Replace is
  disabled by default. Enable it per provider only after the mutation can consume the approved regular-file occupant's
  provider identity/revision or qualified corroboration facts, immediately before replacement in the provider. Require
  explicit scoped consent, approved kind and read-only checks, and normal clean finalization. `ALLOW_OVERWRITE` carries
  intent only. A changed or missing expected occupant returns a known non-commit, so the same bounded collision question
  can refresh. Unknown results terminate recovery. If unavailable, omit Replace and keep valid alternatives.
- Validate the receipt even when the provider returns `S_OK`. Record the primary mutation outcome, known source
  disposition, and any independently observed rollback or cleanup artifact. Reuse `FileSystemItemMutationResult` where it
  conveys those facts and version it if a required fact has no representation. Provider dispatch publishes a single
  item result with the request identity and attempt it belongs to. A proven pre-mutation refusal can return a known
  non-commit. Missing or contradictory evidence after a mutation-capable call preserves uncertainty and never permits
  replay. Cancellation after dispatch follows the same rule. When a primary commit is proved, keep that commit even
  if cleanup later fails and show the actual cleanup problem once. Preserve unknown or retained artifacts and do not
  infer success, source presence, or clean rollback from HRESULT alone.
- Before enabling provider-native dispatch, qualify the receipt paths reached by new-name Rename. Drive must avoid the
  `DriveTransferTree` merge path in provider Rename and keep a verified primary commit if `TrashReplacedItem` later fails.
  Graph `BuildMoveMutationResult` must distinguish proven pre-mutation refusal, uncertain primary response, failed rollback
  and committed-with-cleanup-debt. Curl must carry `CurlPublicationResult` facts to `ReportItemCompleted`; its existing
  overwrite path always keeps the backup once one is created. Preserve that artifact fact even if the primary result is
  unknown, and cover current Copy/Move callers of the shared helper. Ordinary new-name Rename stays out of those
  overwrite paths. Test replacement-only cleanup with the existing affected callers or in the later replacement slice.
  It does not gate new-name Rename unless they share a changed producer.
- Batch Rename and Change Case validate the full requested mapping, provider collision keys and unrelated destinations
  before first mutation. Preview/preflight reports a known unrelated occupied target with the ordinary row/task result
  and lets the user correct the selection, without partial work. A proved later collision uses the existing task decision
  where the dependency schedule permits that action. Keep mappings immutable once accepted and revalidate each hop.
  Dependency cycles and case-only fallback use the existing exclusive temporary-sibling, rollback and truthful result
  rules where that transport safely supports them. Refuse a mapping that needs an unqualified temporary hop before the
  first mutation. If a later leg is unconfirmed, avoid replay or speculative cleanup.
  A provider may use one atomic case-only operation. Require the final spelling, preserve unchanged contents, and account
  for each skipped, failed, pending or committed row.
- For exact-bound folder work, avoid retaining parent and child mutation handles across incompatible dependency layers.
  Compare ingress snapshots when rebinding and keep required ancestry protection. Restrict legacy rename API fallback to
  qualified known non-commits. A `FileBasicInfo` fallback classifies only ordinary objects unless the same object has a
  trustworthy reparse tag. A watcher-caused failure is a hypothesis until verified per server. Where proved, suspend only
  owned watches that block the mutation, off the UI thread, release their handles before the call, and restore or refresh
  on every outcome against the panes' current navigation generation. Preserve pending I/O cancellation and external changes.

## Acceptance scenarios

Use this table for the product walkthrough and its deterministic self-test counterpart. Inspect the actual source and
destination, task result and prompt count. Simulated provider tests exercise the host path. Live qualification records
the exact server/profile tested.

| Scenario | Expected result |
|---|---|
| F2 on a folder on a qualified NAS, with and without file IDs | The folder has the new spelling and all children remain. Focus and pane contents stay coherent. Success produces no question or method note. Test a watched descendant, a changed parent, and decoy target names in the working directory and share root. |
| F2, Batch Rename and Change Case on qualified FTP, SFTP, SCP, Drive and Graph | A free name works and leaves no temporary artifacts. An observed occupied name preserves both objects. Use Keep both, Skip item where useful, or Cancel. Test provider case rules, folder collisions, and a Change Case link with a target outside the selected tree. |
| A name appears or changes during Rename | A true no-overwrite route preserves the late creator and reports the conflict. Drive preserves both item IDs. Inject a Curl destination just before the fresh probe and just after it. The first case preserves both objects without a rename command. In the second, verify the actual protocol outcome, including replacement on a fixture that permits it. A changed approved occupant invalidates a Replace decision. An unconfirmed mutation yields one uncertain result and no automatic retry. |
| A rename commits, then cleanup fails | Show the committed name and one cleanup issue, including any retained artifact. A lost primary response without proof stays uncertain. Test Drive cleanup, Graph refusal/rollback and cleanup, and the existing Curl retained-backup path in the providers. |
| F6 where only Copy is possible | The short Copy/Cancel question opens before a task exists. Copy submits once and keeps the originals. Cancel leaves the filesystem, queue and cut list untouched. Test mixed selections and endpoint changes while the question is open. Native Move starts directly. An explicit options invocation includes the disclosure in its existing dialog. |
| Context-menu Move from a remote pane to a local folder, or Ctrl+drag within a folder | The named source and destination are the ones captured by the command. Same-folder Copy makes a sibling with no self-overwrite. An unsupported Windows clipboard route writes no file-drop data. |
| Copy 1,000 files with three matching name clashes | The first shows one applicable set of actions. **Apply to all** plus **Keep both** or **Skip item** handles the other matching clashes with no questions. An unrelated read-only/link or different destination scope can still need its own choice. |
| Several files hit the same retryable error | Card and mirror label the checkbox **Apply to all when skipping**. With it checked, **Retry** retries only the current item and creates no grant. Choosing **Skip item** with it checked skips later matching errors within the task. |
| Delete on a NAS with Recycle known unavailable, and Viewer Recycle on Local | NAS shows one **Permanently delete?** decision and makes zero Shell Recycle calls. Viewer says **Move to Recycle Bin?**. Cancel leaves the items unchanged. Read-only override and unknown-Recycle outcomes keep their separate safety tests. |
| Copy or Move would lose known security metadata | A pre-publication Skip leaves the source and final destination unchanged and accounts for any stage. If unexpected loss is discovered after a copy commits, the Move source stays and the result names the actual outcome. |
| Queue overlap, then Start now and Pause all | An ineligible overlap waits with a reason and a cancel action. Start now respects the interlock and capacity. Pause all holds current preparation/queued mutation and active I/O, and Resume all restores only its own pauses. |
| Answer a question after a minute with the popup hidden or while typing in the terminal | One answer is accepted after explicit focus. The editor/terminal keeps its focus until then. Resume preserves the graph history without a false stall; a real provider stall still appears. |
| Clear completed with clean, warning and failed tasks present | Clean covered cards disappear, including one with an `Information` note. That note is available in Issues under the opt-in while the task is retained. Warning, failed, partial, uncertain and recovery records remain available through the existing card/Issues path. |
| F7 against a slow provider | The name is accepted once. The pane stays responsive while the operation runs. Confirm completion/refresh, cancel or exit while blocked, and work in another pane during the wait. |

The provider/FileOps tests must also cover changed source/parents, precise case-only spelling, mixed identity, alias
collisions, lost server responses, failed opens, unknown link kinds, temporary-hop rollback collisions, malformed and
duplicate receipts, failed cleanups with a successful HRESULT, and cancellation before and after dispatch. For provider
refusals assert the observable evidence appropriate to that route. Keep every fixture, including link targets, inside its
governed sandbox. Exercise published hosts and provider calls separately where the behavior changed in both.

## Delivery order

Each row can land with its own tests and owning spec updates. The first Rename delivery is complete for a route once its
free-name path, collision behavior and results are qualified. Replacement capability on another route does not gate
that delivery. The small ingress fixes and the telemetry correction can land independently while Rename work proceeds.

| Order | Work and exit |
|---|---|
| 1 | **S1, exact Rename.** Qualify handle retention, protocol fallbacks, case-only hops, and the watcher hypothesis with exact Local/SMB positives and negatives. Fix the watcher only on the routes where it blocks, with navigation/teardown tests. |
| 2 | **S1, new-name Rename.** Land corroborated SMB first. Then qualify Curl's existing destination-probe path, Drive, and Graph independently under the [route contract](#rename-authority-and-qualification). Wire all three Rename entry points to the same task and hold unqualified profiles disabled. Cover the existing callers of any changed receipt helper. |
| Independently, early | **S0 and the short Copy-only Move question.** Correct Viewer Recycle wording, remote context Move direction, clipboard and drag/drop guards, Cancel all, blocked Retry/read-only copy, and bulk Delete ingress parity. Add known-unavailable Recycle routing and the short Copy-only Move flow. Exercise affected ingresses and prove no extra questions. |
| 3 | **S3 and S2.** Use the single repeat control, coalesce artifact decisions and reconcile read-only consent, simplify overlap and sensitive-loss handling, and publish accurate progress on wait/resume. S2 can land as soon as its integration and rate evidence are ready. |
| 4 | **S4 and S5.** Close the queue controls and persistence, explain the shared cap, preserve problem results, and make the mirror and Issues entry consistent. Coordinate queue work with I60 and collected prompts with I64. |
| 5 | **S6 and S7.** Finish option parity, async Create Directory and archive work, and useful plain-language results. Keep async Create Directory and archive worker changes separately reviewable. |
| Throughout | **S8.** Update end-user docs and owning specs with each slice. Remove the obsolete queued destination selector, cover cut/paste Move, and synchronize options, conflicts, qualified provider behavior and outcome explanations. |

**Optional follow-up scope.** Expand Replace or strengthen Curl's rename primitive separately where a provider qualifies
it. Numerical tuning of the transfer budget needs its own paired evidence. Search, filtering,
export, bulk clear and reveal within Issues, a broader destination editor, and other unsupported-provider features
belong to separate proposals. The existing `Clear completed` action and Issues entry are part of I66.

### Specification and work ownership

| Owner | Required updates and coordination |
|---|---|
| Engine and UI | [FileSystem_FileOperations.md](../../FileSystem/FileSystem_FileOperations.md), [UI_FileOperationsPopup.md](../../UI/UI_FileOperationsPopup.md), [UI_FolderView.md](../../UI/UI_FolderView.md), [Core_FileSystemBridge.md](../../Core/Core_FileSystemBridge.md), and [docs/FileOperations.md](../../../docs/FileOperations.md). Record choices, result rules and user-visible workflows with the behavior change. |
| ABI and providers | [Plugins_PluginAPI.md](../../Plugins/Plugins_PluginAPI.md), [Plugins_VirtualFileSystem.md](../../Plugins/Plugins_VirtualFileSystem.md), [FileSystem.h](../../../Common/PlugInterfaces/FileSystem.h), [FTP/SFTP/SCP](../../FileSystem/FileSystem_FtpSftpScp.md), [Google Drive](../../FileSystem/FileSystem_GoogleDrive.md), [Microsoft Drive](../../FileSystem/FileSystem_MicrosoftDrive.md), and other provider specs whenever their capability truth changes. Reconcile stale provider declarations and contradictory general/provider-specific no-replace rules. |
| Settings and tests | [Core_SettingsStore.md](../../Core/Core_SettingsStore.md), [SettingsStore.schema.json](../../SettingsStore.schema.json), [UI_PreferencesDialog.md](../../UI/UI_PreferencesDialog.md), [Testing_SelfTests.md](../../Testing/Testing_SelfTests.md), and [Testing_PerformanceValidation.md](../../Testing/Testing_PerformanceValidation.md). |
| Existing plans | [I18](UI_ConfirmationPreferencesAndOperationOutcomes_2026-09-08.md) owns confirmation preferences and metadata policy. I66 keeps its sensitive-loss protections and advances the independent pre-publication work. [I64](FileOperations_CollectedCollisionsAndResiduals_2026-11-11.md) owns collected collisions and its popup budget. [I60](PreProduction_ReadinessAndSimplification_2026-11-07.md) owns admission fairness and pause-slot yielding. [I25](FileOperations_S3PathStyleCopyAndHarnessResiduals_2026-09-15.md) and [H3](Operation_FileOperations_LaterFeatures_2026-09-05.md) keep their provider and later-feature scope. [I29](FileOperations_PostMergeQualification_2026-10-05.md) and [I3](FileOperations_RemainingStressValidation_2026-08-04.md) keep their qualification scope. |

### Evidence and closeout

- Run `Test-Changes.ps1 -Explain` and the affected checks during implementation. Use focused FileOps, Commands,
  provider contract and native UI cases where required, in the approved marked sandbox. Use `Test-Changes.ps1 -Mode PrePush`
  before pushing and `Tools/Run-AllTests.ps1 -Suite Full -ValidationMode Fresh` when Full is required. Use zero-warning
  native profiles and the owning build/platform policy. Native ARM64 execution and live-server evidence have separate gates from cross-builds
  and owned fixtures. Keep unavailable gates open or transfer them to an accepting owner with the exact missing evidence.
- Check localized labels, focus, action equality, and UIA at the owning surface's minimum size, including the popup's
  480 DIP at 96/144 DPI and the resource satellites. Use bounded non-activating harness scenarios and invented fixture
  data for UI captures. Ask separately for the time of any test that requires real desktop focus.
- Pair performance runs when implementation affects a measured path, under the existing performance policy. Reuse a
  matching source/profile fixture and instrumentation. A wording or action-list correction uses behavior and layout
  assertions. A new route without a working baseline gets bounded-cost evidence and a positive outcome. Keep paired
  artifacts and source identities in the governed evidence archive.

| Path that can regress | Scenario and measurements |
|---|---|
| Bulk Recycle and Rename | Compare bulk and current Recycle ingress on the same batch, including Shell calls, elapsed time and retained resources. For the existing Rename path and changed batch/discovery work, compare repeatable single and 100-item runs, bind/I/O counts, handle peaks, and cancellation latency. Record the watcher/remote cost only when its path is exercised. |
| S2 telemetry | Compare fixed active, waiting, paused and stopping mixes. Measure timer cadence, `UpdateRates` cost and allocations per sample, and retained history. Require zero new allocations on the same steady-state scenario and zero false stall samples. |
| Queue and popup paths | Where a change affects hot admission predicates, redraws, layout or retention, reuse existing queue/popup metrics and I60/I64 fixtures. Measure wait/admission latency, fairness and permit counts, paint/sync and idle work where affected. Collect fresh paired performance evidence only for the paths changed. |
| Concurrency tuning, F7 and archive workers | For a numeric cap change, compare Release NVMe/SMB workloads with multiple jobs, latency, throughput, handles and memory. For async commands, measure UI dispatch responsiveness, completion and cancellation/exit latency with a stalled controlled provider or worker. |

Each slice includes its dated `Changes/` fragment and passes `Tools/Get-SpecInventory.ps1 -FailOnFindings -NoReports`
and `Tools/Test-Changelog.ps1`. Close I66 after the required delivery rows, workflow scenarios, provider gates and
authoritative updates are accounted for. Move it to `Specs/Plans/Done/`, admit the exact immutable path with its
positive and negative history test, and remove its WIP row. Record the receiving owner for any transferred evidence gate.

### STOP conditions

Stop the affected slice and record the failing case if any of these conditions holds.

- Exact `FileIdInfo` authority would be weakened, a retained-handle rename can escape the admitted parent, or recursive
  Change Case can follow a link.
- A route can modify an observed occupant without qualified Replace authority, `ALLOW_OVERWRITE` alone would grant
  replacement, or a new route relies on the Curl exception outside the documented R0f-Curl scope.
- A missing receipt or unknown mutation could trigger retry or cleanup, or a committed destination could be reported
  as `Skipped`.
- Cleanup would bypass the artifact `TouchGuardReceipt`, provider proof or the required worker-side revalidation.
- The queue change needs a second ordering key or permits `Start now` to bypass safety or capacity.
- A fix needs a DxUi dependency pin change. Route the producer fix and consumer qualification through their owning
  work. Keep the current pin in I66.

---

## Source study at ac619f860

The following sections describe the checkout before implementation. They retain the source anchors and original observations.
The target workflow and plugin contract above govern implementation choices. Server-specific behavior and UI timing that were
inferred by source inspection need their own reproduction and qualification.

### Drift check

Run from the repository root before starting any slice:

```powershell
git diff --stat ac619f860..HEAD -- Common RedSalamander Plugins Specs docs Tools Tests
git status --short
```

All line numbers and symbol anchors in this document were verified by read-only inspection of `ac619f860` (plus
RedSalamander's pinned DxUi revision `7fb2ab71852087d93b3a3e5b3ba504c37f56267c` in `Dependencies/DxUi.lock.json`) on
2026-11-23. Re-anchor by symbol name before editing.

---

### 1. Source findings

#### Q1. Is the File Operations workflow always good and coherent?

**Honest answer: No — not yet.**

RedSalamander's **safety and truth core is genuinely strong** and worth preserving verbatim:

- Immutable typed plans (`TransferPlan`, `RenamePlan`, `DeletePlan`), exact-object no-follow binding, atomic `.rs-tmp` /
  `.rs-bak` staging and conditional publication, and structured per-item outcome axes (`PublicationState`,
  `VerificationState`, `SourceDisposition`, `OwnedStageDisposition`, `ItemCompletion`).
- Durable interrupted-Move breadcrumbs that stay standalone across restarts and are never wiped by `Clear completed`.
- Non-activating File Operations popup, off-UI Preparing/discovery/shutdown drain, storage-medium-aware auto-concurrency
  (`HDD = 1`, `USB = 1..2`, `SSD = 4`, `NVMe / Network = 8`), and event-driven keep-awake with a 2-minute quiet release.
- Generation-tokened `FileOperationsConflictBroker`, FAT 2-second timestamp tolerance, DST 1-hour detection, in-pane
  conflict inspection with rollback, and the `I63` D-6 "never show a menu with only 1 item" polish rule.

However, the **end-user workflow built on top of that engine is NOT coherent yet**. Because the safety core, popup UI,
start confirmations, Decision Mirror, Issues Pane, Queue scheduler, Preferences page, and non-Local providers evolved in
separate waves (`I35`, `I63`, `I18`, `I64`, and `I60`), end users hit sharp seams:

1. **Three incompatible confirmation paradigms** coexist (blocking sync modals for `Shift+F5`/`Shift+F6`/Copy-only Move/Compare Move, async card prompts for `Shift+Delete`, and zero options/confirmation for routine `F5`/`F6`/Paste/Drag&Drop).
2. **Broken Rename on identity-less SMB/NAS AND across all non-binding providers (FTP, SSH [SFTP/SCP], Google Drive, OneDrive)**, plus folder-specific handle-retention defects and SMB protocol fallback gaps.
3. **Redundant and contradictory conflict controls** (`Skip item` + `Apply to all` checkbox + `Skip all`, dead checkboxes on 4(+2) failure buckets, `Apply to all` and `* all` offered on 1-file tasks and F2 inline rename, `Delete source` buried inside `More...`).
4. **Useless, double/triple, or technically opaque prompts** (Viewer Recycle lying *"Permanently delete"*, `MetadataLoss` firing *after* committing copies of browser-downloaded files with MOTW, `SameHostOverlap` interrupting even in Queue mode, `PathTooLong` offering `Retry`, double/triple prompts on Leftover Cleanup, Pack/Unpack Delete-After, and Shift+Delete of read-only files).
5. **Throughput graph grey-zone corruption** when waiting on or resuming from user questions (and during `Stopping` / `WaitForLiveOutputPublisher` / `PromptForConnectionSecret`).
6. **Hidden cliffs in Parallel mode, `Start now`, `Clear completed`, Decision Mirror focus-stealing, Issues Pane discoverability, Bandwidth/Preferences inconsistencies, sync UI-thread F7/Pack/Unpack, and cross-provider Move asymmetry.**

---

#### Q2. Why can't I rename a folder on SMB, FTP or SSH?

The current host requires object binding. Its coverage misses host Rename on SMB and non-binding remote providers.
Source inspection also identifies handle-lifetime and protocol paths that can fail. Server-specific behavior needs
reproduction on the affected profile before it establishes the cause of a reported failure.

##### Cause 1. The host requires bound authority
- **What `BindObjectAuthority` actually returns (`FolderWindow.FileOperations.cpp:1172-1195`)**:
  `ObjectBindingState::Unsupported` is returned in **two completely different situations** that current code conflates:
  1. **Local/SMB binding provider (`IFileSystemObjectBinding` IS implemented)**: `FileSystem::BindObject` (`Plugins/FileSystem/FileSystem.cpp:7272-7404`) opens the item no-follow, then calls `GetFileInformationByHandleEx(file, FileIdInfo, ...)` (`:2997, 7353`). On NAS/Samba shares that do not support `FileIdInfo` (or lack `FileIdInfo` on directories/share roots even when files return it — documented at `FileSystem.cpp:173-174` and `FolderWindow.FileOperations.Tests.IdentityLess.cpp:1493-1504`), `BindObject` returns `ERROR_NOT_SUPPORTED` (`0x80070032`). *(Note: a failed open MUST stay `Indeterminate` / `ERROR_OPEN_FAILED`, never `Unsupported`.)*
  2. **Non-binding providers (`QueryInterface(IFileSystemObjectBinding)` returns `E_NOINTERFACE` at `FolderWindow.FileOperations.cpp:1173-1178`)**: FTP, SFTP, SCP, Google Drive, and Microsoft Drive advertise Rename on supported profiles and implement `IFileSystem::RenameItem` (`FileSystemCurl.CopyMove.cpp:2988`, `FileSystemGoogleDrive.cpp:2684`, `FileSystemMicrosoftDrive.cpp:6214`). These providers lack object binding, so the host's binding request returns `ObjectBindingState::Unsupported` (`0x80070032`).
- **Why these production Rename requests fail closed (`0x80070032` `ERROR_NOT_SUPPORTED`)**:
  Unlike Copy (`bindingRequired = false` + retained-handle stage nonce + `I63` `IFileSystemExpectedOccupantPublication`),
  Delete (`I35` `IFileSystemIdentityDelete` for identity-less Local/SMB + `DeletePlan::nativeAuthority` for FTP/SSH/cloud),
  and Native Move (`MoveItem` with `bindingRequired = false`), **Rename was excluded from both identity-less Local/SMB
  fallbacks (`Common/PlugInterfaces/FileSystem.h:1129`, `FileSystem_FileOperations.md:2067`) AND `nativeAuthority`
  provider-native dispatch (`IFileSystem::RenameItem` has ZERO production callers; `FileSystemRenameBatch.cpp` is
  `#ifdef ENABLE_TESTS`-only via `ChangeCase.cpp:619`)**. Every rename hits up to five hard `Bound`-only gates:
  1. `Task::PrepareMutationInterlockScopes` (`FolderWindow.FileOperations.State.cpp:12171`): calls `builder.Add(..., parent, WriteSource, /*bindingRequired=*/true, false)`. `MutationInterlockScopeBuilder::Add` (`State.cpp:11220-11514`) binds **the parent AND every ancestor up to `\\server\share` (or remote root)**. If the provider lacks `IFileSystemObjectBinding` (FTP/SSH/Drive/OneDrive) OR ANY SMB directory/share root lacks `FileIdInfo`, Preparing aborts with `0x80070032` (*"This operation is not supported for this item or location."*).
  2. `Task::ExecuteInlineRename` (`State.cpp:15261-15266, 15331-15343, 15382-15399, 15434-15450`): requires `Bound` source, parent revalidation, `Bound`/`Missing` destination, `RenameIfUnchanged`, and post-rename `RevalidateObjectAuthority`.
  3. `Task::PrepareBatchRenameAdmission` (`FolderWindow.FileOperations.cpp:3123-3134`): requires `Bound` ingress snapshot for every row.
  4. `ChangeCase::BuildOrApplyToPaths` (`ChangeCase.cpp:277-284`): requires `Bound` roots when `includeSubdirs` is enabled.
  5. `Task::ExecuteBatchRename` / `ExecuteBatchRenameMutation` (`State.cpp:15995-16002, 15666-15775`): requires `Bound` for every source and destination.

##### Cause 2. Folder-specific handle retention
- **2A. `ExecuteBatchRename` retains folder and child handles** (`FolderWindow.FileOperations.State.cpp:15983-16053, 15841`):
  Pre-binds EVERY item in the batch simultaneously with `FILESYSTEM_BIND_RENAME` (`FILE_READ_ATTRIBUTES | DELETE`) into
  `authorities[]`, and after renaming each child executes `sourceAuthority = std::move(returned.authority)` (line `15841`),
  keeping that child's `DELETE` handle open while moving to the next dependency layer to rename its parent folder:
  - A live child handle can block renaming its parent. Measure which combinations and servers exhibit `ERROR_ACCESS_DENIED` or `ERROR_SHARING_VIOLATION` in this path.
  - *Hypothesis on Samba/strict SMB*: holding a `DELETE` handle on the parent folder while Layer 0 renames children can also cause `ERROR_SHARING_VIOLATION` (`0x80070020`), matching the Samba behavior previously recorded in `FileSystem_FileOperations.md:1186`.
- **2B. Opposite-pane `DirectoryWatch` holds an open `FILE_LIST_DIRECTORY` handle on the folder** (`Plugins/FileSystem/FileSystem.Watch.cpp:123-129`, `RedSalamander/DirectoryInfoCache.cpp`):
  If the opposite pane is displaying the target folder or any subfolder inside it, `DirectoryWatch` holds an open
  `FILE_LIST_DIRECTORY` handle with pending `ReadDirectoryChangesW`. The handle grants `FILE_SHARE_DELETE` and Rename does
  not suspend the watch before mutating. Whether this handle blocks the actual rename is an endpoint-specific hypothesis.
- **2C. Ancestor `FILE_READ_ATTRIBUTES` handles retained throughout rename**:
  `PrepareMutationInterlockScopes` retains open `FILE_READ_ATTRIBUTES` handles on the parent and every ancestor up to the
  share root across `ExecuteInlineRename` / `ExecuteBatchRename` (*hypothesis*: strict NAS lease setups may reject folder
  renames while ancestor handles remain open).
- *(Ruled out by static audit: `FileSystem.DirectoryOps.cpp` closes enumeration handles immediately; `DirectoryInfoCache` sets
  `allowOffscreenWatch = false`; `IconCache` and `NavigationView` retain no filesystem handles.)*

##### Cause 3. SMB protocol paths in bound Rename
- **3A. Incomplete `FileRenameInfoEx` -> `FileRenameInfo` fallback** (`Plugins/FileSystem/FileSystem.cpp:4642-4658`):
  Only falls back on `ERROR_INVALID_PARAMETER`, `ERROR_INVALID_FUNCTION`, or `ERROR_NOT_SUPPORTED` — missing
  `ERROR_CALL_NOT_IMPLEMENTED` and `ERROR_ACCESS_DENIED` that `Common/LocalFileTransaction.h:328-329` explicitly handles for
  filesystems/redirectors rejecting `FileRenameInfoEx`.
- **3B. Missing two-hop case-only rename fallback in `RenameIfUnchanged`** (`Plugins/FileSystem/FileSystem.cpp:4603-4700, 4948-5200`):
  Unlike `MovePathInternal` (`FileSystem.FileOps.cpp:8968-8976`, which calls `RenameCaseOnlyWithTemp`), `RenameExactTo`
  attempts only a single `SetFileInformationByHandle(..., replaceIfExists=FALSE)`. On SMB/Samba servers where direct
  case-only `FileRenameInfo` returns `ERROR_ALREADY_EXISTS` or `ERROR_ACCESS_DENIED`, Inline Rename enters a dead-end
  conflict prompt (`destinationIsSource == true`, Replace disabled) and Batch Rename / Change Case fail.
- **3C. Unconditional `FileAttributeTagInfo` requirement in `BindObject`** (`Plugins/FileSystem/FileSystem.cpp:7360-7379`):
  No fallback exists if an SMB server does not implement `FileAttributeTagInfo`. **Safety constraint**: `BindObject` at
  `:7367-7371` requires `ReparseTag` for `IsReparseTagNameSurrogate(...)`. `FILE_BASIC_INFO` / `GetFileInformationByHandle`
  supply attributes **without** `ReparseTag`, so any fallback is legal **only for non-reparse items
  (`(FileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) == 0`)**; reparse items MUST either resolve a trustworthy reparse tag
  separately or fail closed.
- **3D. Reconciliation failure escalates straight to `ERROR_IO_INCOMPLETE` (`0x800703E4`)** (`FileSystem.cpp:4670-4694`,
  `State.cpp:15453-15458, 15845-15849`):
  Reports `"Outcome unknown"` and aborts without offering Retry.

##### Test coverage
- The only SMB rename self-test (`R0fSmb_LoopbackReadWriteCreateDelete` in
  `FolderWindow.FileOperations.Tests.Phases05_06.cpp:7592-7604`) calls the **unused legacy `IFileSystem::RenameItem`** on a
  single file (`rename-me.bin`). Zero self-tests run `AdmitInlineRename`, `AdmitBatchRename`, or `AdmitScheduledRename`
  over SMB (file or folder), and zero host-level tests run Inline F2 / Batch Rename / Change Case over FTP or SSH (SFTP/SCP).

##### Baseline Rename and Move by provider

| Entry Point | Identity-less SMB/NAS (Local, no `FileIdInfo`) — File / Folder | Identity-capable SMB — File | Identity-capable SMB — Folder | FTP / SSH (SFTP, SCP) (`FileSystemCurl`) — File / Folder | Google Drive / OneDrive — File / Folder |
|---|---|---|---|---|---|
| **Inline F2** | Refused without required binding | Exact route, risks 3A–3D | Exact route, risks 2B and 3A–3D | Refused without binding | Refused without binding |
| **Batch Rename** | Refused without required binding | Exact route, risks 3A–3D | Exact route, risks 2A/2B and 3A–3D | Refused without binding | Refused without binding |
| **Change Case** | Refused without required binding | Exact route, risk 3B | Exact route, risks 2A/2B and 3B | Refused without binding | Refused without binding |
| **F6 Move to sibling folder on the same endpoint** | Native route can use `MoveFileWithProgressW` | Native route | Native route, in-place rename blocked by `SanityCheckBothPanes` | Native when admitted by the connection (`RNFR`/`RNTO` or SFTP `rename`) | Native when admitted by the provider (`files.update` / Graph `PATCH`) |

---

#### Q3. Why do some questions repeat or misstate the action?

Complete audit of every ingress dialog and all 26 `ConflictClass` buckets identified **10 distinct classes of defective or
low-value user questions**:

1. **BUG — ViewerSpace Recycle Bin delete lies "Permanently delete ... This cannot be undone."**
   - `FolderWindow::ExecuteViewerAction` (`FolderWindow.FileOperations.cpp:3883-3891`) passes `FILESYSTEM_FLAG_USE_RECYCLE_BIN` + `requireConfirmation = true`.
   - `StartOperation` (`State.Runtime.cpp:1317, 1417-1434`) takes the `!deleteBypassesRecycleBin` `else` branch and formats **`IDS_FMT_FILEOPS_CONFIRM_PERMANENT_DELETE`** (*"Permanently delete {0}: from: {1}\r\n\r\nThis cannot be undone."*) for a Recycle Bin operation.
2. **Single-file tasks and Inline F2 Rename offer `Apply to all`, `Skip all`, `Replace all older`, and `Skip all with the same size and date`.**
   - `BuildConflictDecisionScope` (`State.cpp:5289-5375`) and `RequestBridgeConsent` (`State.cpp:16976`) never check whether the task can possibly hit a second item (`_sourcePaths.size() > 1u` or directory root). Only `RecycleFailed`/`RecycleFailedByName` (`State.cpp:24453, 24885`) check `_sourcePaths.size() > 1u`.
   - Renaming **one file** with F2 onto an existing name shows the `Apply to all` checkbox AND `Skip all` (or `More...` containing `Skip all` + `Replace all older` / `Skip all with the same size and date`), plus both `Skip item` and `Cancel` (which do the exact same thing on a 1-file rename).
3. **Double and triple prompts where the user already explicitly stated intent (plus a normative spec contradiction on read-only delete):**
   - **Leftover Cleanup** (`FolderWindow.LeftoverCleanup.cpp:1232-1328`): User selects leftovers -> clicks `Delete Selected...` -> Prompt 2: `ConfirmArtifactTouch` modal (*"This request touches N item(s) that may be left over..."*, `State.Runtime.cpp:1121`) -> Prompt 3: `PermanentDeleteConfirmation` card (`State.cpp:12670`) if permanent. `ConfirmArtifactTouch` also mints the `TouchGuardReceipt` checked on the worker at `State.Runtime.cpp:2329`. See [Start and delete behavior](#start-and-delete-behavior) for the combined confirmation.
   - **Pack / Unpack "Delete after"** (`FolderWindow.FileSystem.Commands.cpp:12548-12566, 12748-12767`): User checks *"Delete files from disk after packing"* / *"Delete archive after unpacking"* -> clicks OK -> Prompt 2: `ConfirmPermanentDeletePaths` modal *before* packing/unpacking -> Prompt 3: `ConfirmArtifactTouch` *after* packing/unpacking if any path matches an artifact pattern (`DeleteHandoffForPlan` returns `None` at `State.Runtime.cpp:211`).
   - **Shift+Delete on Read-Only items (`State.cpp:12670` -> `State.cpp:27544, 27571, 5592`) & Spec Self-Contradiction**: After confirming `PermanentDeleteConfirmation`, hitting any read-only file (`.git/objects/`, etc.) pops a second `ReadOnlyRegularFile` confirmation (`Permanently delete read-only`). Furthermore, **`FileSystem_FileOperations.md` contradicts itself**: line `1537` states read-only permanent-delete consent is *"noncached"*, whereas line `3922` explicitly permits task-scoped `Apply to all` caching for provider-proven `ReadOnlyDeleteBlocked` (matching `IsCacheableConflictDecision(DeleteReadOnly) == true` at `State.cpp:5783`). [Start and delete behavior](#start-and-delete-behavior) spells out the two grants and the repeat scope.
4. **`MetadataLoss` on Move interrupts everyday downloaded files AFTER already writing and committing the copy.**
   - `State.cpp:21046, 21088-21112`: On any bridge Move that cannot preserve `MOTW` (`Zone.Identifier`), ADS, or EA (e.g., moving almost any browser-downloaded file from NTFS to FAT32, exFAT, SMB/NAS, SFTP, or S3), RedSalamander writes and **commits** the destination file first (`PublicationTransaction::State::Committed`), then stops to ask *"Moving this item would lose security-significant metadata such as Mark-of-the-Web, alternate streams, or extended attributes"* (`[Move anyway] [Copy and keep source] [Cancel]`).
   - Publication may already have replaced an existing file. Cancel leaves the published destination and stops the rest of the Move. A real Skip requires the decision before publication. See [Collision and risk decisions](#collision-and-risk-decisions).
5. **`SameHostOverlap` asks `[Queue after these tasks] [Don't start]` even in Queue mode and when concurrent execution is forbidden.**
   - `State.cpp:12854-12933`, `State.Queue.cpp:1722-1823`, `FileOperationsInternal.h:3044`: `CanRunConcurrently()` is `false` for all transfer-vs-transfer overlaps (`serializedTransfers`), leaving only `[Queue after these tasks]` and `[Don't start]` — and fires even when `New tasks: Queue` is already active!
6. **`PathTooLong` offers a useless `Retry` button** (`State.cpp:4987-4993` — `IsRetryableConflictBucket` fails to exclude `PathTooLong`).
7. **`ReadOnlyRegularFileExists` hides the fact that the destination file already exists** (`State.Diagnostics.cpp:475-476` renders only `IDS_FILEOPS_CONFLICT_READONLY`: *"Item is read-only."*).
8. **Pre-publication deferred consents (`EfsPlaintext`, `SparseInflation`, `PlaceholderHydration`) offer NO `Skip item` button** (`State.cpp:5755`), forcing all-or-nothing Cancel of a multi-file folder transfer (and leaving dead `RetainSource` branches at `State.cpp:20171, 20531`).
9. **Full options modal for a routine Copy-only Move** (`State.Runtime.cpp:1174, 1275-1289`). Any planned Copy-only item forces the transfer options surface even when the caller requested the defaults. This follows the admitted strategy. A supported native Move within the same endpoint does not take this branch just because it uses FTP or a cloud provider.
10. **Destructive default on Popup `Cancel all`, wall-of-text jargon, and raw hex `0x...` errors:**
    - Popup `Cancel all` (`Popup.cpp:13511`) sets **`OK` as the Enter default** and says *"pending"* (`IDS_MSG_FILEOPS_CANCEL_ALL_POPUP`) while canceling running tasks too — contradicting App/Compare Close (`Cancel` default at `FolderWindow.FileOperations.cpp:4436`).
    - `PermanentDeleteByNameConfirmation` / `RecycleFailedByName` append a **92-word internal-implementation essay** (`IDS_FILEOPS_CONSENT_BY_NAME_CHECKED`, `RedSalamander.rc:2079`) and label the button *"Delete by name"* instead of *"Permanently delete"*.
    - Pane/command error alerts (`RedSalamander.rc:1406-1407, 1469, 1473, 1647-1648, 1660, 1669, 1928`) print raw hex `Error: 0x{0:08X}`.

---

#### Q4. Why do Skip and Skip all overlap?

##### Equivalent decisions in the broker
- `ConflictAction::SkipAll` is host-only (`FolderWindow.FileOperationsInternal.h:661`); the plugin ABI `FileSystemIssueAction` has only `Skip` (`Common/PlugInterfaces/FileSystem.h:217-228`).
- In `BuildConflictActionPolicy` (`FolderWindow.FileOperations.State.cpp:5755-5762`), for all 10 non-deferred buckets (`RegularFileExists`, `ReadOnlyRegularFileExists`, `TypeMismatch`, `DestinationLink`, `ReadOnlyRegularFile` [non-delete], `NameNotRepresentable`, `SharingViolation`, `AccessDenied`, `PathTooLong`, `NetworkOffline`), whenever `applyToAllEligible` is `true`, the engine adds **both** `ConflictAction::Skip` AND `ConflictAction::SkipAll` while **also** keeping `applyToAllEligible = true` (which renders the checkbox `IDS_FILEOPS_CONFLICT_APPLY_TO_ALL`: *"Apply this choice to remaining similar conflicts in this task"*).
- In `FileOperations::ConflictBroker::TryAcceptLocked` (`RedSalamander/FileOperationsConflictBroker.cpp:53-70`):
  - Clicking **`Skip all`** produces `Decision{ConflictAction::Skip, /*applyToAll=*/true, generation}`.
  - Checking **`Apply to all`** and clicking **`Skip item`** produces the **exact same** `Decision{ConflictAction::Skip, /*applyToAll=*/true, generation}` — stored in the exact same `decisionCache` slot, logged identically, returning the exact same `FileSystemIssueAction::Skip`.

##### History of the duplicate action
1. **Original (`00b353e6e`)**: Only `Skip` + `Cancel` + `Apply to all` checkbox existed.
2. **Post-v7.0.48 (`d9536a480`)**: Added `SkipAll` as an overflow (`More...`) shortcut. For `DiskFull` (`State.cpp:5603-5613`), the author cleanly removed both the checkbox and one-item `Skip` (*"Skip All is its own explicit task grant; no checkbox or one-item Skip is needed."*), but kept BOTH `Skip` + checkbox AND `SkipAll` on all 10 other non-deferred buckets.
3. **`I63` Slice C (`ded63e43b`, `State.cpp:5530-5539`)**: Added single-overflow promotion (*"a menu needs two commands"*). Whenever `SkipAll` was the sole overflow action, it popped directly into the primary button row right beside `Skip item`: `[Skip item] [Skip all] [Cancel]` directly underneath `[ ] Apply this choice to remaining similar conflicts in this task`!
4. **`I63` B4 & B6 (`ded63e43b`, `State.cpp:5656-5675, 5767-5777`)**: Added `DiscardSource` (*"Delete source"*), `SkipSameSizeAndDate` (*"Skip all with the same size and date"*), and `ReplaceOlder` (*"Replace all older"*). Because redundant `SkipAll` already occupied an overflow slot, adding `DiscardSource` bumped `overflowActionCount` from 1 to 2 — **hiding the genuinely distinct `Delete source` action inside `More...` solely because redundant `Skip all` sat beside it** (`State.cpp:9265-9294`)!

##### Failure prompts with a duplicate repeat choice
On **`AccessDenied`**, **`SharingViolation`**, **`PathTooLong`**, and **`NetworkOffline`** (`State.cpp:5338-5342, 5750-5779`), the prompt displays:
- Checkbox: `[ ] Apply this choice to remaining similar conflicts in this task`
- Buttons: `[Retry] [Skip item] [Skip all] [Cancel]`

Checking the box does nothing for `Retry` (`ConflictBroker.cpp:68`), nothing for `Cancel`, nothing for `Skip all`, and makes `Skip item` 100% identical to `Skip all`. The exact same degeneracy happens on **`TypeMismatch`** and **`NameNotRepresentable`** whenever `Keep both` is not offered.

##### Bulk action inventory

| Action | Buckets | Coexists with Checkbox? | Distinct From `Apply to all`? | Verdict |
|---|---|---|---|---|
| `SkipAll` (*"Skip all"*) | 10 non-deferred buckets | **YES** | **NO** (identical to Checkbox + `Skip item`) | Use one repeat control |
| `SkipAll` (*"Skip all"*) | `DiskFull` (`State.cpp:5603`) | NO (no checkbox, no `Skip item`) | Sole skip/repeat action | **Keep** (clean pattern!) |
| `PermanentDeleteAll` (*"Delete all remaining permanently / by name"*) | `RecycleFailed`, `RecycleFailedByName` when `>1` item | NO (no checkbox) | Explicit task-scoped escalation grant | **Keep** (coordinate button adjacency with `I64` RU8) |
| `SkipSameSizeAndDate` (*"Skip all with the same size and date"*) | `RegularFileExists` when Same size & date | YES (checkbox silently ignored) | **YES** — conditional policy (skips only matching pairs) | **Keep**, fix placement/1-file gating & clarify vs checkbox |
| `ReplaceOlder` (*"Replace all older"*) | `RegularFileExists` when Newer + Replace | YES (checkbox silently ignored) | **YES** — conditional policy (replaces only older pairs) | **Keep**, fix placement/1-file gating & clarify vs checkbox |

*(Also note: `More...` is currently rendered **leftmost**, before all primary buttons, on both Card (`Popup.cpp:12616-12633`) and Decision Mirror (`DecisionMirror.cpp:668-679`), and `Delete source` silently ignores `Apply to all` when dates/sizes differ (`ConflictBroker.cpp:62-69`) with its caveat hidden only in the Card checkbox tooltip (`Popup.cpp:9237-9247`) and completely absent on the Decision Mirror.)*

---

#### Q5. Why can the throughput graph advance during a question?

##### Rate sampling and graph ownership
- `DxUi::ThroughputGraph` (`DxUi/src/Controls/DxUi.Controls.cpp:5696-6360`) holds a 180-sample (`18.0 s` at `100 ms/sample`) ring buffer with **no** internal ticker and **no** grey-zone renderer.
- The grey hatched zone is drawn by RedSalamander's `FileOperationsGraphStallBand` (`RedSalamander/FolderWindow.FileOperations.Popup.cpp:265-333`), which hatches `palette.disabledText` at `0.45` alpha over every contiguous run where `history.stalledSamples[i] != 0u`. It was designed solely for real I/O stalls (`I35 R10.2`), and `Specs/UI/UI_FileOperationsPopup.md:2034-2036, 2054-2056` explicitly forbids stall/rate progression during a user question.
- Samples are produced by `FileOperationsPopupState::OnTimer` (`Popup.cpp:14976-15026`) -> `UpdateRates()` (`:6350`) -> `BuildRateSnapshot()` (`:6229`) -> `AdvanceRateHistory()` (`:4340-4525`) -> `AppendResampledRateSamples()` (`:4186-4222`).

##### Sample and wait transitions to verify

1. **Root Cause A1 (CRITICAL — Missing `MarkRateSamplingStateChanged()` on ALL conflict/decision transitions + `OnTimer` skipping `UpdateRates()` while hidden/minimized):**
   - `Task::MarkRateSamplingStateChanged()` (`State.cpp:14495-14498`) is called **only** in `SetPaused` (`:14315`), `SetWaitingInQueue` (`:14420`), and `SetQueuePaused` (`:14483`).
   - **Zero** conflict/decision transitions call it: `SetConflictPromptMetadataLoadingLocked` (`State.cpp:6213`), `SetConflictPromptLocked` (`:6142`), `WaitForConflictDecision` (`:6689-6889`, including `ConflictBroker::ResetLocked` at `:6855`), `ClearConflictPrompt` (`:6107`), or `SetAwaitingUserDecision` (`:14793`).
   - `OnTimer` skips its regular `UpdateRates` call for a hidden or minimized popup (`Popup.cpp:15001-15005`). `RefreshTelemetryTimer` can still synchronize rates once on an aggregate waiting transition (`:7147-7163`), so the exact gap depends on which tasks are active and which transitions the sampler observes.
   - If an enter/exit pair is missed and the state-change tick is unchanged, the next `AdvanceRateHistory` can skip the reset (`Popup.cpp:4364-4390`). The maximum resample window is 18 seconds (`:4189-4190`), enough to displace the prior history. Test hidden and visible paths with both a single waiting task and a simultaneously active peer.

2. **Root Cause A2 (CRITICAL — Visible-popup 1 Hz waiting-timer gap on answer):**
   - Even with the popup visible the entire time, `RefreshTelemetryTimer` slows to `1000 ms` (1 Hz) during a loaded conflict (`Popup.cpp:7117-7213`).
   - When the user clicks an answer `600..999 ms` after the last 1 Hz tick, `prompt.active` clears without calling `MarkRateSamplingStateChanged()`.
   - At the next `100 ms` tick, `nowTick - stallReference` is already `700..1099 ms > 600 ms` (`RateStallThresholdMs`) -> immediately flashes false `"Speed: — · Waiting for the destination…"` and injects ~10 grey-hatched 0-rate buckets (~1.0 s of grey zone) for the user's decision time.

3. **Root Cause B (CRITICAL — Prompts & post-question waits that NEVER set `_conflictArbiter.prompt.active` keep the graph VISIBLE and LIVE-ADVANCING grey at 10 Hz!):**
   - **`WaitForLiveOutputPublisher`** (`State.cpp:12996-13004`, `State.Queue.cpp:2163-2191`): After the user answers `SameHostLiveOutput` with *"Queue until other task finishes"*, `WaitForConflictDecision` clears `prompt.active = false` and blocks on `_queueCv.wait(lock)` without setting `_waitingInQueue`, `_queuePaused`, `_waitForOthers`, or `_paused`. The throughput graph immediately reappears as `Running` and scrolls a grey hatched zone across the entire plot at 10 Hz, wiping 100% of real history in 18 seconds.
   - **`PromptForConnectionSecret`** (`Plugins/FileSystemCurl/FileSystemCurl.Shared.cpp:1181, 1292`): Modal FTP/SFTP/SCP password or SSH key passphrase prompt blocks the worker while `prompt.active == false` and `_awaitingUserDecision == false` -> same live 10 Hz grey-zone wipe of all 18 s of real history.
   - **`BatchRenameArtifactPrompt`** (`State.Runtime.cpp:2218`): Sets `_awaitingUserDecision`, which `BuildRateSnapshot` (`Popup.cpp:6318`) ignores.
   - **Pre-prompt slow `BindObject` I/O in `Task::FileSystemIssue`** (`State.cpp:10423-10448`): Runs before `prompt.active` becomes `true`, painting a grey stall band whenever remote metadata binding takes `> 600 ms`.
   - **`waitingForOthers` omitted** from `IsRateSamplingBlocked` (`Popup.cpp:3712`), `GetPopupTelemetryWork` (`State.Runtime.cpp:2781`), and `MarkRateSamplingStateChanged` (`State.cpp:14380`).

4. **Related Telemetry & Progress Bugs Discovered in the Same Audit:**
   - **`Stopping` erodes real throughput history to 0 over 18 s**: `task.stopping` is absent from `IsRateSamplingBlocked` (`Popup.cpp:3712`), so `AppendResampledRateSamples` keeps appending decaying-to-0 samples at 10 Hz — violating `UI_FileOperationsPopup.md:488` (*"Stopping retains actual counters, paths, result facts and throughput history"*). Also `taskMayAnimate` (`Popup.cpp:9304`) still allows marquee animation because `StatusCountsAsRunning(Stopping)` is true.
   - **Manually `Paused` Task B kills the Footer Aggregate ETA for concurrently running Task A**: `Paused` keeps `includeInAggregateCohort = true` (`Popup.cpp:1462`) with `transferBytesPerSec == 0.0`, which trips `etaInputsComplete = false` at `Popup.cpp:2330`.
   - **`Preparing... ({0} elapsed)` is stuck at `0s`**: `StatusTextForTask` (`Popup.cpp:1834-1836`) reads `task.operationStartTick`, which stays `0` until `ExecuteOperation` (`State.cpp:23856`).
   - **Why existing tests missed Root Cause A**: `CheckRateBlockedStates()` (`FolderWindow.FileOperations.Tests.Throughput.cpp:1016-1059`) tests `waitingForDecision = true` for 10 s and returns without ever transitioning `waitingForDecision` back to `false` or simulating a 1 Hz / hidden-popup gap.
   - **Forward coordination with `I64` Slice G**: When collected collisions (`prompt.active && prompt.collected`) land, the task keeps transferring while the collected prompt is shown — so `IsRateSamplingBlocked` will need `prompt.active && !prompt.collected` (and `ResolveWholeTaskProgressPresentation` will need to account for collected prompts).

---

#### Q6. What else is inconsistent?

Additional issues beyond Q2–Q5 are listed below. See [Specification and work ownership](#specification-and-work-ownership)
for related plans.

1. **[P0 BUG] Context-menu `Move...` (`FolderView::MoveSelectedItems`) inverts source/destination on non-local panes & lacks a Copy counterpart** (`FolderView.FileOps.cpp:1295-1379` omits `sourceContextSpecified = true` and uses local-only `IFileOpenDialog`; `FolderWindow.FileOperations.cpp:3541-3573` interprets `!sourceContextSpecified` on a non-local pane as an inbound local drop/paste).
2. **[P1 BUG] `FolderView::CopySelectionToClipboard()` writes raw non-local virtual paths into Win32 `CF_HDROP`** (`FolderView.FileOps.cpp:763-783` lacks the `SupportsFileDropCut()` guard present right below it in `CutSelectionToClipboard()` at `:789`).
3. **[P1 BUG] Ctrl+Drag Copy into the same folder is silently cancelled**, unlike `Ctrl+C` + `Ctrl+V` (`FolderView.DragDrop.cpp:701-711`), and uses `AreSameLocalVolume`/`EqualsNoCasePath` even for case-sensitive remote internal drops (`:696, 701-706`).
4. **[P1] Delete `ExecutionMode` mismatch degrades Shell Recycle Bin outside `Del`**: `CommandDelete` (`FolderWindow.FileOperations.cpp:4656`) uses `BulkItems` (and misses `EnsureFileOperations()` at `:4610`), while context-menu Delete (`:3713`), Find Results Delete, and Viewer Recycle (`:3963`) pass `PerItem` (1 STA thread + 1 shell call per item).
5. **[P1 UX Proposal] `Clear completed` permanently removes Failed, Partial, Warnings, and Indeterminate cards + their Issues diagnostics in one unconfirmed click** (`Popup.cpp:2394-2397, 2442-2445, 10365-10369, 15383-15386`, `State.Runtime.cpp:3303-3331`):
   - *Verified baseline*: `_completedGroupExpanded` starts **`true`** (`Popup.h:2115`), automatic collapse happens ONLY after an accepted `Cancel all` when every member ended clean (`UI_FileOperationsPopup.md:234-240`), and automatic dismissal (`IsAutoDismissableFileOperationCompletion`) vs explicit user `Clear completed` are separate contracts.
   - *UX friction*: Failed/Partial/Indeterminate cards are grouped inside `Completed (N)` alongside clean Done/Canceled tasks (unlike standalone interrupted-Move notices), so explicit `Clear completed` erases problem cards and their Issues rows without confirmation (and manual collapse folds problem cards behind counts).
6. **[P1] Parallel mode global serialization & broken `Start now` dead-end trap**: Any folder transfer, overwrite, or non-Local/identity-less transfer sets `activeExclusive = true` (`State.cpp:11563-11633`, `State.Queue.cpp:2415-2448`), serializing completely unrelated drives/providers. Worse, clicking `Start now` on a queued task (`State.Queue.cpp:2558`, `State.Runtime.cpp:2844-2864`) clears `_waitForOthers`, immediately slams into `activeExclusive`, **makes `Start now` disappear**, leaves the task stuck waiting, and removes it from the reorderable queue with no way back!
7. **[P1] No `[Pause]` button on cards during concurrent discovery-ahead** (`Popup.cpp:12641-12828`) even while files are actively copying.
8. **[P1 Contract Evolution — Owner Approved] `maxParallelTransfers` scope clarity vs storage-adaptive `Auto` budget**:
   - *Verified baseline*: `FileSystem_FileOperations.md:2295-2299` **intentionally** defines `fileOperations.maxParallelTransfers` (default `4`, range `1..16`) as one process-wide file budget covering both jobs AND their per-item workers (`State.cpp:14376, 25954`). It is a documented resource guarantee, not an accidental bug.
   - *UX/policy gap*: `IDS_PREFS_FILEOPS_PARALLEL_TRANSFERS_DESC` (*"Share this limit across file-operation jobs"*) obscures that it also caps a single job's storage-adaptive `Auto` budget (clamping NVMe `8` and SMB `8` to `4` in `FileSystem.cpp:1689, 1714`), and interacts confusingly with session-only `Queue` mode and per-plugin concurrency settings. See [Queue, progress, and results](#queue-progress-and-results) for persistence and the shared budget.
9. **[P1] Queue vs Parallel mode (`_queueNewTasks`) is NOT persisted** (`FolderWindow.FileOperationsInternal.h:3333` vs `Common/SettingsStore.h:349-380`), silently reverting to Queue on restart, and has inconsistent scope between Start Modal (per-task) and Popup Footer (session-global).
10. **[P1] Decision Mirror steals keyboard focus mid-keystroke** (`DecisionMirror.cpp:272-283, 727` checks only top-level window foreground, not active text inputs/filter/address bar/inline rename/embedded terminal) and lacks Popup Card parity (Off/On selector vs Checkbox, no `Show source`/`Show destination`/path copy, missing `Delete source` scope help).
11. **[P1] Issues Pane discoverability & missing basic tools**: No `Issues` button in Popup Footer or Options menu (`Popup.cpp:10220-10290, 13582-13642`); pane lacks search/filter, inline `Show informational notes` toggle, `Export`, `Clear`, and `Reveal in pane` (`IssuesPane.cpp:1509-1572`). Also `OperationStringId` maps `FILESYSTEM_CREATE_DIRECTORY` to `"Unknown"` (`IssuesPane.cpp:86`) vs `"Copy"` (`PromptText.cpp:294`, `Popup.cpp:11177`), and `FormatDiscoveryProgressText` (`Popup.cpp:529-532`) has a duplicate unreachable `FILESYSTEM_DELETE` branch.
12. **[P1] Bandwidth speed-limit presets disagree across all 3 surfaces + mixed IEC/SI copy**:
    - Preferences (`Preferences.FileOperations.cpp:44-52`): `1..500 MiB/s` + Custom (**missing `1 GiB/s`**).
    - Start Modal (`HostServices.cpp:2141-2150`): `1..500 MiB/s, 1 GiB/s` (**no Custom input**).
    - Running Card (`Popup.cpp:13686-13693`): `1..100 MiB/s, 1 GiB/s` + Custom (**missing `500 MiB/s`**).
    - Custom hints advertise `KB`/`MB`/`GB` while parsing binary IEC units (`TryParseBinaryThroughputText`).
13. **[P1] Sync UI-thread bypasses**: F7 Create Directory (`FolderWindow.FileSystem.Commands.cpp:10596-10760` blocks the UI thread on slow SMB/cloud, skips `RefuseFileOperationForPendingExit`, and reports invalid names as *"does not support creating folders"*) and Pack/Unpack (`:12548-12850` run synchronously on the UI thread, Local-only, ad-hoc Skip/Replace modal).
14. **[P2] Hidden/silent options & command-label lies**:
    - `IDS_CMD_MOVE_WITH_OPTIONS` (`RedSalamander.rc:2190`) says **"Move/Rename with Options"**, yet Shift+F6 has no editable destination/name field.
    - `Links` policy (`Preserve`/`Skip`) is prominent in Shift+F5 yet absent from `Preferences > File Operations` (stored only in hidden per-plugin JSON `reparsePointPolicy`, `State.cpp:1437-1475`).
    - Enabled `verifyAfterCopy` is **silently turned off** on routine F5 to unsupported destinations (`State.Runtime.cpp:1237-1244`) with zero note.
    - `InformationalTask` bypasses the 500 ms reveal delay (`State.Runtime.cpp:3630`); `Pause all` skips Preparing/Waiting tasks (`:3001-3015`); `IDS_FILEOPS_CHECK_AUTODISMISS_ON`/`_OFF` duplicate identical text.
15. **[P2] Cross-provider & SMB cliffs**:
    - The Local route advertises Recycle for SMB/UNC (`FileSystem.cpp:7102`). Where the location cannot recycle, that sends plain Delete through a Shell refusal before the escalation question.
    - Move FROM any non-Local provider (S3, Google Drive, OneDrive, FTP/SFTP/SCP, MTP) to Local always downgrades to Copy-only (while Local -> S3 is Managed Move!), S3 advertises `rename: false`, OneDrive advertises `delete: false`, MTP advertises `delete: false, rename: false`.
16. **[P2] End-user documentation drift (`docs/FileOperations.md`)**:
    - Lines `48-50` document a removed queued-card Destination selector (`UI_FileOperationsPopup.md:947` removed it).
    - Lines `53` vs `241` expose the mismatched bandwidth presets.
    - Lines `322, 328` describe `Ctrl+V` as copy-only, omitting `Ctrl+X` Move.
    - Lines `141-144` document the redundant `Skip all` + `Apply to all` overlap.

---

### 2. Baseline reference inventories

#### 2.1 Mutation ingress

| Ingress | Code Entry Point | Confirmation Surface | Options Exposed | Async / Interlock | Delete/Transfer Mode |
|---|---|---|---|---|---|
| `F5` Copy | `FolderWindow.FileOperations.cpp:4775-4918` | None | None | Yes | `PerItem` |
| `Shift+F5` Copy | `FolderWindow.FileOperations.cpp:4775-4918` | Sync modal (`HostServices.cpp:2070+`) | Links, Verify, Start, Bandwidth | Yes | `PerItem` |
| `F6` Move | `FolderWindow.FileOperations.cpp:4919-5065` | None (except Copy-only -> sync modal) | Only if Copy-only | Yes | `PerItem` |
| `Shift+F6` Move | `FolderWindow.FileOperations.cpp:4919-5065` | Sync modal (no rename field!) | Verify, Start, Bandwidth | Yes | `PerItem` |
| `Del` (pane shortcut) | `FolderWindow.FileOperations.cpp:4590-4670` | None (escalates on card if needed) | None | Yes (missing `EnsureFileOperations()`) | `BulkItems` |
| `Del` (context menu) | `FolderWindow.FileOperations.cpp:3713` | None | None | Yes | **`PerItem`** (regression) |
| `Del` (Find results) | `FolderWindow.FileOperations.cpp:3963` | None | None | Yes | **`PerItem`** (regression) |
| `Del` (ViewerSpace Recycle) | `FolderWindow.FileOperations.cpp:3883-3891` | **BUG: Sync modal saying "Permanently delete"** | None | Yes | **`PerItem`** |
| `Shift+Del` | `FolderWindow.FileOperations.cpp:4671-4770` | Async Task Card (after Preparing) | None | Yes | `PerItem` |
| `Ctrl+C` / `Ctrl+V` | `FolderView.FileOps.cpp:763` / Paste | None | None | Yes (`Copy` misses local guard!) | `PerItem` |
| `Ctrl+X` / `Ctrl+V` | `FolderView.FileOps.cpp:785` / Paste | None (except Copy-only) | None | Yes | `PerItem` |
| Drag & Drop | `FolderView.DragDrop.cpp:455-765` | None (except Copy-only) | None | Yes (same-folder Copy bug) | `PerItem` |
| Context `"Move..."` | `FolderView.FileOps.cpp:1295-1379` | Win32 Folder Picker | None | Yes (**inverts non-local!**) | `PerItem` |
| `F2` Inline Rename | `FolderWindow.FileOperations.cpp:2849` | Inline edit | None | Yes | Rename (fails on identity-less SMB & all non-binding providers!) |
| Batch Rename / Change Case | `FolderWindow.FileOperations.cpp:3350`, `FileSystem.cpp:4016` | Preview window / options dialog | Command-specific | Yes | Rename (same) |
| `F7` Create Directory | `FolderWindow.FileSystem.Commands.cpp:10596-10760` | Modal name prompt | None | **NO (Sync UI thread!)** | Sync |
| Find Copy / Move | `FolderWindow.FileOperations.cpp:3960+` | None | None | Yes | `PerItem` |
| Compare Sync Copy / Move | `FolderWindow.FileOperations.cpp:4070` | Copy = None; Move = Sync modal | Move only | Yes | `PerItem` |
| Pack / Unpack | `FolderWindow.FileSystem.Commands.cpp:12548-12850` | Own modal + pre-op Delete modal | Format, mask, Skip/Replace, Delete-after | **NO (Sync UI thread!)** | Sync + async Delete |

#### 2.2 Conflicts and deferred consent

| `ConflictClass` | Primary Buttons | `More...` Overflow (if `>=2`) | Default / Escape | Checkbox? | Key Defect(s) |
|---|---|---|---|---|---|
| `RegularFileExists` | `[Keep both] [Replace] [Skip item] ([Skip all] if 1 overflow) [Cancel]` | `>=2` of `Delete source`, `Skip all`, `Skip all with same size & date`, `Replace all older` | Cancel / Cancel | YES | Redundant `Skip all`; hides `Delete source`; shown on 1-file & F2 Rename |
| `ReadOnlyRegularFileExists` | `[Keep both] [Replace read-only] [Skip item] ([Skip all]) [Cancel]` | Managed Move: `[Delete source, Skip all]` | Cancel / Cancel | YES | Text says only *"Item is read-only."* (omits exists); redundant `Skip all`; 1-file leak |
| `ReadOnlyRegularFile` (Delete) | `[Permanently delete read-only] [Skip item] [Cancel]` | None | Cancel / Cancel | YES | Separate permanent-delete and read-only grants. See [Start and delete behavior](#start-and-delete-behavior) for scoped caching. |
| `ReadOnlyRegularFile` (Transfer/Rename) | `[Replace read-only] [Skip item] [Skip all] [Cancel]` | None | Cancel / Cancel | YES | Redundant `Skip all` + checkbox; 1-file leak |
| `TypeMismatch` | `[Keep both (if eligible)] [Skip item] [Skip all] [Cancel]` | None | Cancel / Cancel | YES | Redundant `Skip all`; dead checkbox when `Keep both` unavailable |
| `DestinationLink` | `[Keep both] [Replace link (if bound)] [Skip item] [Skip all] [Cancel]` | None | Cancel / Cancel | YES | Redundant `Skip all` + checkbox |
| `NameNotRepresentable` | `[Keep both (if eligible)] [Skip item] [Skip all] [Cancel]` | None | Cancel / Cancel | YES | Redundant `Skip all`; dead checkbox without `Keep both` |
| `TargetConflict` | `[Skip item] [Cancel]` | None | Cancel / Cancel | NO | No bulk skip for broken symlinks in a tree |
| `SharingViolation` | `[Retry] [Skip item] [Skip all] [Cancel]` | None | Cancel / Cancel | YES | **Zero-effect checkbox** (`Skip item`+box === `Skip all`) |
| `AccessDenied` | `[Retry] [Skip item] [Skip all] [Cancel]` | None | Cancel / Cancel | YES | **Zero-effect checkbox** |
| `DiskFull` | `[Retry] [Skip all (if qualified)] [Cancel]` | None | Cancel / Cancel | NO | Clean reference pattern (no checkbox, no redundant `Skip item`) |
| `PathTooLong` | `[Retry] [Skip item] [Skip all] [Cancel]` | None | Cancel / Cancel | YES | **Useless `Retry`** + **zero-effect checkbox** |
| `NetworkOffline` | `[Retry] [Skip item] [Skip all] [Cancel]` | None | Cancel / Cancel | YES | **Zero-effect checkbox** |
| `UnsupportedReparse` | `[Skip item] [Cancel]` | None | Cancel / Cancel | NO | Jargon (*"reparse point"*); no bulk skip |
| `Unknown` | `[Retry] [Skip item] [Cancel]` | None | Cancel / Cancel | NO | No bulk skip on recurring errors |
| `RecycleFailed` (Delete) | `[Permanently delete] [Skip item] ([Delete all remaining permanently]) [Cancel]` | None | Cancel / Cancel | NO | Clean (properly checks `_sourcePaths.size() > 1`) |
| `RecycleFailed` (Move `Delete source`) | `[Permanently delete] [Keep source] [Cancel]` | None | Cancel / Cancel | NO | Clean task-scoped escalation (second step after `Delete source`) |
| `RecycleFailedByName` | `[Delete by name] [Skip item] ([Delete all remaining by name]) [Cancel]` | None | Cancel / Cancel | NO | 92-word internal essay + *"Delete by name"* jargon |
| `PermanentDeleteConfirmation` | `[Permanently delete] [Cancel]` | None | Cancel / Cancel | NO | Followed by ReadOnly prompt on cases such as `.git` objects. See [Start and delete behavior](#start-and-delete-behavior). |
| `PermanentDeleteByNameConfirmation` | `[Delete by name] [Cancel]` | None | Cancel / Cancel | NO | 92-word internal essay + *"Delete by name"* jargon |
| `SameHostOverlap` | `[Queue after these tasks] ([Run at the same time]) [Don't start]` | None | Queue after / Don't start | NO | Interrupts even in Queue mode & when `CanRunConcurrently() == false` |
| `SameHostLiveOutput` | `[Queue until other task finishes] [Delete/Rename/Replace anyway] [Skip item] [Cancel]` | None | Queue until / Cancel | NO | Choosing Queue triggers live 10 Hz grey stall bug (`WaitForLiveOutputPublisher`) |
| `EfsPlaintext` | Copy: `[Copy without encryption] [Cancel]`; Move adds `[Copy and keep source]` | None | Cancel / Cancel | YES | Pre-publication: missing `Skip item`; checkbox shown on 1-file tasks |
| `SparseInflation` | `[Continue] [Cancel]` | None | Cancel / Cancel | YES | Pre-publication: jargon; missing `Skip item`; dead `RetainSource` branch |
| `PlaceholderHydration` | `[Download and continue] [Cancel]` | None | Cancel / Cancel | YES | Pre-publication: jargon (*"placeholder"*); missing `Skip item`; dead `RetainSource` branch |
| `MetadataLoss` | `[Move anyway] [Copy and keep source] [Cancel]` | None | Cancel / Cancel | YES | The question can occur after the destination commits (`State.cpp:21088`). See the target publication rules above. |
| `InsufficientSpace` / `SpaceUnknown` | `[Continue] [Cancel]` | None | Cancel / Cancel | YES | **Dead / unwired in production** (self-test only) |

#### 2.3 Throughput and progress

| State | Timer | `IsRateSamplingBlocked` | `MarkRateSamplingStateChanged` | Graph & Samples | Verified Bug(s) |
|---|---|---|---|---|---|
| **Paused** | 1000 ms | Yes | Yes (`SetPaused`) | Visible (`Paused` overlay); history preserved | Stays in footer cohort with rate `0` -> **kills aggregate ETA for running peers** (`Popup.cpp:2330`) |
| **Conflict (arbiter)** | 100 ms (loading) / 1000 ms (loaded) | Yes (`prompt.active`) | **NO** | Hidden while active | **Bug A1**: hidden/DecisionMirror wait backfills up to 180 grey samples on restore; **Bug A2**: 1 Hz gap injects ~1 s grey hatch on resume |
| **Non-arbiter Prompt / `WaitForLiveOutputPublisher`** | 100 ms | **NO** | **NO** | **Visible & live-advancing grey at 10 Hz** | **Bug B**: wipes 100% of 18 s real history into solid grey stall band |
| **Waiting** (`queuePaused`, `waitingInQueue`, `waitingForOthers`) | 1000 ms (except `waitingForOthers`: 100 ms) | Partial (`waitingForOthers` missing) | Partial (`SetWaitForOthers` missing) | Hidden | `waitingForOthers` omitted from `IsRateSamplingBlocked` & `GetPopupTelemetryWork` |
| **Preparing** | 100 ms | No (`started == false`) | No | Hidden; 0 samples | `Preparing... ({0} elapsed)` stuck at `0s` (`Popup.cpp:1834-1836`) |
| **Discovering** | 100 ms | No | No | Hidden pre-transfer; visible concurrent with marquee + provisional ETA | Missing `[Pause]` button on card while copying during discovery-ahead (`Popup.cpp:12641-12828`) |
| **Verifying** | 100 ms | No | No | Visible (`Verifying` overlay) + secondary series + 50/50 bar | Clean |
| **Stopping** | 100 ms | **NO** | No | Visible, unhatched | **Appends decaying-to-0 samples at 10 Hz**, wiping throughput history over 18 s (`UI_FileOperationsPopup.md:488`); marquee stays animated |
| **Stalled** (real I/O silence) | 100 ms | No | No | Visible + grey hatched `FileOperationsGraphStallBand` | Intended design (`I35 R10.2`) |

#### 2.4 Provider capabilities and workflow limits

| Provider | Object Binding (`IFileSystemObjectBinding`) | Copy / Move / Rename / Delete / Recycle | Cross-FS Move (`export/import`) | Verification | Notable Workflow Cliff |
|---|---|---|---|---|---|
| **Local NTFS/ReFS** | Yes (`FileSystem.h:120`) | All `true` | `["*"]` / `["*"]` | `hostReadback` + `blake3` | The default process budget caps single-job NVMe auto-concurrency from eight to four. See [Queue, progress, and results](#queue-progress-and-results). |
| **SMB / UNC** | Yes, but server may return `ERROR_NOT_SUPPORTED` for `FileIdInfo` | Advertises all `true` including Recycle | `["*"]` / `["*"]` | Bound-only | Rename fails closed on identity-less NAS. The Recycle route may fall through to escalation for locations without a bin. The global budget caps the effective Auto budget. |
| **FTP / SSH (SFTP, SCP)** (`FileSystemCurl`) | **No (`E_NOINTERFACE`)** | `copy/move/rename/delete: true`, `recycle: false` | `[]` / `[]` | `none` | **Inline F2, Batch Rename, and Change Case all fail (`0x80070032`)** even though `FileSystemCurl::RenameItem` is implemented; Cross-FS Move Copy-only; Verify silently dropped on F5; always serialized |
| **Google Drive** | **No (`E_NOINTERFACE`)** | All `true` | `[]` / `[]` | `writer-digest` | **Inline F2, Batch Rename, and Change Case all fail (`0x80070032`)** despite `RenameItem`; Cross-FS Move Copy-only; cross-root Copy fails (`I35` XP-09) |
| **Microsoft Drive (OneDrive)** | **No (`E_NOINTERFACE`)** | `copy: false`, **`delete: false`**, `rename/move/recycle: true` | `[]` / `[]` | `writer-digest` | **Inline F2, Batch Rename, and Change Case all fail (`0x80070032`)** despite `RenameItem`; same-endpoint Copy (`H3`) and Permanent Delete unsupported; Cross-FS Move Copy-only |
| **S3** | **No (`E_NOINTERFACE`)** | `rename: false`, `recycle: false` | `["*"]` / `["*"]` | `writer-digest` | F2 Rename unsupported; **asymmetric Move** (Local->S3 deletes source; S3->Local Copy-only!) |
| **MTP** | **No (`E_NOINTERFACE`)** | `nativeMove/delete/rename/recycle: false` | `[]` / `[]` | `none` | Cannot delete or rename; Move always Copy-only (`H3`); concurrency 1 |
| **7z / IMAP** | **No (`E_NOINTERFACE`)** | Read-only mount | Export copy only | `none` | Mounted 7z is read-only; Pack/Unpack are separate sync local-only commands |
