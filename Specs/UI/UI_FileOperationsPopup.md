# File Operations Popup UI Contract

Last updated: 2026-09-01

This spec covers the hybrid Direct2D/DirectWrite and hosted DxUi File Operations progress popup surface in:

- `RedSalamander/FolderWindow.FileOperations.Popup.h`
- `RedSalamander/FolderWindow.FileOperations.Popup.cpp`

Engine execution, conflict semantics, streaming discovery, plugin contracts, and performance-validation
rules remain owned by `Specs/FileSystem/FileSystem_FileOperations.md`.

## Purpose

The popup is an operational surface for long-running copy, move, delete, and informational tasks. It
MUST make the current state legible at a glance, keep active controls stable while the task list
changes, and expose enough debug state for deterministic selftests without pixel inspection.

## Reference Captures

Documentation screenshots MUST be captured from the real product popup, not from the HTML mockup.
Current Debug x64 product captures:

- Current conflict and minimum-width footer: `Specs/UI/Images/FileOperationsPopup_Product_Conflict_2026-07-10.png`.
- Active running/waiting popup: `Specs/UI/Images/FileOperationsPopup_Product_Active_2026-07-09.png`.
- Completed partial-result popup: `Specs/UI/Images/FileOperationsPopup_Product_Partial_2026-07-09.png`.

![Current conflict and minimum-width footer product capture](Images/FileOperationsPopup_Product_Conflict_2026-07-10.png)

![Active running/waiting File Operations popup product capture](Images/FileOperationsPopup_Product_Active_2026-07-09.png)

![Completed partial-result File Operations popup product capture](Images/FileOperationsPopup_Product_Partial_2026-07-09.png)

## Regions

The popup has three visible regions:

- Title bar: standard captioned tool window, independently minimizable/restorable.
- Operations list: scrollable stack of task cards.
- Global footer: always visible, including footer-only mode.

Footer-only mode hides the operations list, persists through `SettingsStore`, resizes the window to
the footer-only minimum height, and restores the previous captured rectangle when details are shown
again. The compact rectangle uses the normal `FileOperationsPopup` placement key; the last expanded
rectangle uses `FileOperationsPopupExpanded` and MUST NOT be overwritten while footer-only mode is
active.

When no valid persisted placement is restored, the popup is centered over its valid owner and clamped to the
owner monitor work area. Placement changes size neither the popup nor its z-order/activation state. This uses the
shared owner-centering geometry contract; monitor-work-area lookup failure leaves the existing position intact.

## Footer

The global footer MUST include:

- Aggregate progress bar across active file-operation work.
- Left transfer-action cluster: `Pause all` / `Resume all` first when applicable, then `Cancel all`
  while active work exists.
- Right settings/view cluster: one drop-down selector whose label is the current `Queue` or
  `Parallel` new-task mode, an `Options` overflow, then the details disclosure button.
- Status summary assembled from nonzero categories only (for example, `Running: 2`; omit
  zero-valued categories), optionally followed by aggregate ETA and throughput. Running,
  Waiting, and Needs-attention clauses MUST use separate count-neutral localized resources in
  the source file and all four satellite resource files; a combined singular/plural format or
  a retained dead combined-summary resource is forbidden.
- Right-aligned details chevron for footer-only collapse/expand.

Footer controls live outside the scrolled task-list viewport and MUST remain hit-testable there. The
`Options` overflow owns persistent preferences: auto-dismiss success/canceled and Compact/Expanded
density. Density MUST round-trip through canonical settings save even when it is the only non-default
file-operation setting. Auto-dismiss is a future-retention policy: changing it does not clear cards
already retained. `Clear completed` is an immediate bulk action for the current completed group and
MUST NOT be duplicated in the footer.

The footer is an 88-DIP band with separate aggregate-progress, status-summary, and command rows.
The supported minimum client width is 480 DIP. At that width, every required control MUST retain a
positive, non-overlapping hit target. The square Options and details buttons remain icon-only; the
current mode selector may narrow but MUST retain its label and drop-down affordance.

The Queue/Parallel selector MUST read `GetQueueNewTasks()` on every hosted-control sync so its label
always reflects engine state. Activating it opens a radio flyout with Queue and Parallel; choosing an
item calls `ApplyQueueMode(...)`, and choosing the current item is idempotent. It MUST NOT use an
animated sliding thumb or keep a second UI-local mode state. It uses the shared DxUi current-value
selector chrome: one whole click target, flat borderless rest state, quiet chevron, shared hover
animation across the complete target, and no split-button divider or nested chevron button.

A flyout opened from a popup button MUST use the invoking button's width as its minimum root-surface
width. Longer localized labels, shortcut columns, and other flyout content may expand the surface
beyond that minimum; the monitor work area remains the final size constraint.

## Task Cards

Every file-operation card MUST resolve exactly one `TaskSnapshot::StatusKind` before drawing. The
same resolved status drives:

- header text,
- status glyph,
- status stripe,
- status chip,
- graph overlay,
- caption severity,
- footer counters,
- Windows taskbar state.

The left stripe and header chip are required for every non-`None` file-operation status. Tone mapping:

| Status | Tone |
|--------|------|
| Running, Discovering, Preparing, Verifying | Accent |
| Waiting, Stopping, Canceled | Neutral |
| Paused | Muted |
| Conflict, Partial | Warning |
| Failed | Error |
| Done | Ok |

The chip uses localized short labels and MUST NOT be the only status signal. Status must remain
readable without color through glyph and/or text.

The card consumes the engine-owned `TaskLifecyclePhase`; it MUST NOT infer Preparing from missing
progress, zero totals, or a worker that has not emitted a transfer callback. `Preparing` and
`AwaitingAcceptance` render as Preparing. `Ready` is a transition boundary before Waiting/Running;
`Stopping` and `Terminal` remain explicit. Pause, conflict, verification, and discovery are
orthogonal facts and retain their existing status precedence without rewriting the lifecycle.

Publication does not synchronously create or show the popup. Every ordinary, clipboard, Inline
Rename, Batch Rename, and Change Case task arms the fixed 500-ms `kTaskCardRevealDelayMs` deadline.
Queue waiting and actionable conflict/failure/partial/canceled/indeterminate states reveal
immediately. Otherwise, a task that is still live at the deadline reveals its one stable card in
its current lifecycle state; a slow gate therefore appears as Preparing and later transitions in
place. Clean routine work that completes before the deadline keeps its completed-history receipt
without opening or focusing the popup. **Show File Operations** reveals a hidden live task
immediately. The delay is not a Preference, and deterministic tests use a fake clock rather than
sleeping.

### Confirmation surface

Routine accepted-default Copy/Move starts without a generic confirmation. This includes F5/F6,
destination-picker transfers, clipboard paste, and internal drag/drop when their ingress does not request a separate material
decision. The existing confirmation surface is shown only for an explicit `requireConfirmation`, a
known Copy-only Move, or an existing exact artifact/risk gate. Permanent Delete and command-owned
preview/editor/drop decisions retain their separate contracts.

When shown, Copy/Move confirmation captures the complete immutable task intent before task
publication. In addition to the operation summary and destination it exposes:

- for Copy, `Links`, with exactly `Preserve links` and `Skip links`; Follow is never displayed or
  accepted; a Move confirmation does not show the row;
- `Verify copied file contents`, initialized from `fileOperations.verifyAfterCopy` and changeable
  for this task without changing Preferences;
- the existing Queue/Parallel and bandwidth controls;
- for clipboard Move, the fixed warning that accepting the task consumes the captured cut list and
  that the list will not be restored if a source is retained.

`HostPromptRequest::fileOperationOptions` points to a caller-owned, synchronous
`HostFileOperationPromptOptions` in/out snapshot. The host renders its finite choices as four
keyboard-, mouse-, and UIA-invokable option rows and writes accepted values back only for the
affirmative action. Links exposes Preserve/Skip, Verify exposes Off/On, Start exposes Queue/Parallel,
and Bandwidth cycles Unlimited, the standard presets, plus the captured custom value when needed.
Flattening the choices into an unstructured message string or rereading Preferences/provider
configuration after acceptance is forbidden. Cancel leaves the caller snapshot unchanged, admits no
task, and consumes no clipboard sequence.

Verify also renders the caller-provided `verificationAvailability`: Supported exposes Off/On;
Unsupported and Not applicable force a disabled Off row with a localized reason; Check during
operation leaves Off/On enabled and explains that the exact route is resolved when work starts. The
prompt cannot infer support from provider name or method existence, and an accepted runtime-check
choice is not silently cleared when the per-item result later becomes Unavailable.

### Streaming discovery

There is no `Calculating` task state and no separate pre-calculation pass. After mandatory
Preparing/Ready, execution traversal may enter `Discovering` while the single traversal runs ahead
of execution. While discovery-ahead is active,
the card exposes `Skip discovery`; selecting it atomically switches the task to just-in-time
discovery, releases discovery-ahead reservations and scheduler throttles, removes the action, and
allows eligible transfer work to use the full task bandwidth/concurrency budget. The status becomes
`Discovering as needed` until traversal closes.

`Skip discovery` skips only the run-ahead user experience. It never skips recursive enumeration,
name validation, identity/containment checks, capability queries, conflicts, or revalidation before
an item is acted on. While traversal is open, the card presents discovery as a dedicated animated
three-dot activity signal and presents transfer independently: live source rows, transferred bytes,
throughput history, and the whole-task bar continue to update. Until traversal closes the whole-task
bar is indeterminate (a marquee once transfer is active) and the card shows the exact counters
instead: bytes and items processed so far against bytes and items discovered so far (D2-A08). No
growing-denominator fraction or percent is rendered. A provisional ETA, `Remaining: {0} (still
discovering)`, appears only while a usable smoothed transfer rate exists over the workload discovered
so far; it may rise or fall, is hidden while paused, stalled, or without progress callbacks, and is
replaced by the ordinary ETA once totals close. Aggregate progress reports **Known work** over the
closed-total cohort (see Progress Rules); the taskbar remains indeterminate while any included total
is open. After Skip, the ordinary operation rate/throughput remains
visible beside the secondary `Discovering as needed` signal; full transfer speed is not presented as
proof that totals are known. Once traversal closes, totals become final, determinate progress and ETA
may begin, and the action can never reappear for that task.

### Same-host overlap warning

When Preparing finds that the new task's scopes overlap another published-prepared or active task in this host
(R4-A02-1), the task card shows one deferred-consent prompt titled "Another task is working in the
same place" whose question names the concrete problem and the first overlapping task by number, for
example "This task may remove items that task #3 is still reading." Its actions are **Queue after
these tasks** (the default and Enter action), **Run at the same time**, and **Don't start** (Escape).
Run is present only when every disclosed relation fits the fixed bound and applies only to the named
concurrently-live task IDs; another undisclosed overlap still asks independently. Queue keeps the task
as `Waiting` behind the named tasks without a `Start now` control (the wait is the user's scheduling
choice, not the queue mode); Don't start ends the task as canceled before any mutation and leaves no
completed-history success entry. The warning appears at most once per task, and a task that prepared
later than a peer still learns about that peer once the peer has decided its own warning; when it
queues after a younger peer it takes its place behind every task admitted so far, which the queue
order on the cards shows; queue-mode
waits and the interlock wait itself remain silent. Read/read overlaps and clearly disjoint paths never warn, and
neither does the conservative wait of a provider without bound objects (its interlock cannot prove
which objects two paths name, so it serializes without naming a problem).

A Permanent Delete asks its confirmation on the card, not in a modal at ingress: once Preparing has
pinned every selected root the card shows "Permanent delete" with the question the modal used
("Permanently delete {what} from {where}. This cannot be undone."), and its actions are **Cancel**
(first, the default and Escape action) and **Permanently delete**. Cancel ends the task as canceled
before any mutation and leaves no completed-history success entry; Permanently delete lets the task
re-check the confirmed pathnames against the pinned objects and then run. Archive cleanup deletes
whose consent the archive prompt already captured do not ask again.

When a task without a covering Run relation is about to delete, rename, move, or replace an exact
item that another task is still publishing, the same card changes to a deferred-consent prompt titled
**Another task is still publishing here**. The question names the publisher task and explains that
changing the item can invalidate its result. Actions are **Queue until the other task finishes**
(default and Enter), **Skip**, the operation-specific **Delete it anyway**, **Rename it anyway**,
**Move anyway**, or **Replace it anyway** action, and **Cancel** (Escape). Queue waits only for that
publisher and then revalidates; Skip completes the item as skipped; the destructive action applies
only to the parked effect and has no Apply-to-all/cache semantics. If the bounded exact-output index
overflows, the destructive action is withheld and the conservative Queue/Skip/Cancel surface remains.
The prompt and its task relation disappear when the publisher terminates and are never reconstructed
from completed history or shared with another app instance.

### Inline-F2 completion exception

A one-step `RenamePlan(InlineRename)` uses workers and the normal safety/interlock engine. It follows
the common deferred presentation rules, with one narrower clean-completion exception:

- waiting on an interlock, conflict, failure, partial, canceled, or indeterminate state reveals the
  popup and card immediately;
- a clean-running task reveals its card only if it is still running at the common fixed 500-ms deadline;
- clean success before that deadline creates no card, no completed row, and does not activate the
  popup, while telemetry, pane refresh, and identity-based focus retention still occur;
- once a card has been revealed it never becomes retroactively silent and follows ordinary
  completion/auto-dismiss behavior.

Batch Rename, `RenamePlan(ChangeCase)`, and Copy/Move/Delete keep their ordinary completed-history
receipts after a clean hidden completion; only Inline F2 suppresses the completed row. Change Case
may show a threshold-delayed informational card during command-owned discovery, but once immutable
mappings are admitted the central File Operations task is the sole mutation/progress/result card
owner.

In normal themes, `Ok` status MUST resolve through the scoped `fileOperations.successText` theme
color and MUST be visually distinct from the active `Accent` tone. High-contrast themes may map
`Ok` to the system/menu text color when that is required for contrast.

In high-contrast themes, the card-level stripe/chip/tone semantics remain present, and warning/error/
ok statuses retain a glyph plus text signal. The non-client caption status glyph is suppressed in
high contrast so system caption contrast is not compromised.

Queued file-operation cards MUST expose `Start now` before the task begins. The action releases only
that task's wait gate and MUST NOT flip the global footer Queue/Parallel setting. The footer MUST
expose a bulk `Pause all` command while any active started task is running and a bulk `Resume all`
command when started tasks are paused with no unpaused running tasks. Bulk pause/resume toggles only
the per-task manual pause state for started live tasks and MUST NOT clear queue gating or change the
global Queue/Parallel mode. Discovering/preparing tasks that have not entered execution are not bulk-command
eligible and MUST NOT change a `Resume all` decision back to `Pause all`.

Queued tasks have a stable engine-owned queue-order key. The popup sorts queued cards by that key and
exposes Move Up/Move Down only when an adjacent queued task exists in the requested direction. A move
swaps the adjacent keys and the engine queue order under the queue mutex. Worker-thread start does not
make a task ineligible: queue actions remain valid until the task enters operation. `Start now` releases
only the selected task and queue reordering does not change the global Queue/Parallel mode.

Completed cards auto-collapse once when they first resolve to a finished state, unless auto-dismiss
removes them or the user already has a manual expanded/collapsed override for that task. The footer
Compact/Expanded density setting applies to every visible task card as a default display state, not
as a forced stored collapse; the per-card chevron can expand a compact-density row and restore the
card's normal actions without disabling later completed-card auto-collapse. Compact rows keep the
status signal and task name on one line. A final published byte/item denominator permits a mini
progress meter plus localized percent text. While discovery is open and transfer is active, the
compact row shows a marquee and no meter or percent text. A row with completed
work but no denominator MUST NOT render a misleading `0%` meter. When at least two completed file-operation cards remain visible, the
popup groups them under `Completed (N)`, expanded by default. `Clear completed` is the first control
in the group header and the disclosure chevron occupies the same trailing 18-DIP lane as task-card
chevrons. The chevron hides/shows completed rows without dismissing them. While collapsed, the header
shows separate nonzero badges for Completed, Partial, Failed, and Canceled result counts. Task-card
height and completed-group visibility animate together with bounded smoothstep easing over 260 ms.
User-triggered disclosure motion temporarily advances at a 16-ms cadence, then returns to the
100-ms telemetry timer; reduced motion snaps both transitions to their target state.

Each nonzero collapsed-group badge has a retained DxUi status region whose tooltip explains the
category and count (successful, partial/warning, failed, or canceled); the same text is exposed as
its accessible name and HelpText. `Clear completed` and other text buttons MUST NOT show a tooltip
that only repeats their visible label. Glyph-only disclosure and overflow actions retain descriptive
tooltips because their action name is not visible.

All popup disclosure affordances use the shared Fluent disclosure-chevron path. A leading-edge
collapsed disclosure points right; a trailing-edge collapsed disclosure points left. File Operations
task, completed-group, and footer details chevrons occupy a trailing lane and therefore use left/down,
while trees keep right/down. The transition takes the shortest 90-degree path with the shared
`PointToPoint` easing, uses native untransformed directional glyphs at rest, and snaps under reduced motion. Task-card
and completed-group chevrons are card-integrated disclosure targets without standard button fill/border
chrome; keyboard focus remains visible. Hand-drawn stroke chevrons are forbidden.

## Hosted DxUi Control And Accessibility Contract

The popup retains one top-level HWND and its existing hand-drawn card/chrome layout, but every
interactive hit target MUST be mirrored by a real `DxUi::Button` hosted by a `DxUi::WindowHost`
attached to that same HWND. Determinate progress is exposed by hosted `DxUi::ProgressBar` controls
and bandwidth history by hosted `DxUi::ThroughputGraph` controls. The legacy renderer remains
responsible for non-interactive text, status decoration, and card geometry; hosted controls use the
same calculated rectangles and action descriptors. Each progress rectangle and each graph has one
pixel owner: the hosted control. There is no second, hand-painted progress or graph route and no
legacy mouse hit-testing route; when the host is unavailable the popup shows the failure surface
below instead. File Operations progress descriptors set the hosted control's explicit track height
to the complete calculated lane height; the generic 2/4-DIP ProgressBar default is too subtle for the
popup's aggregate and whole-task marquee lanes.

## Failure Surface

When the DxUi host cannot attach to the popup HWND, or the popup's Direct2D target or device
resources cannot be created, the popup switches to a failure surface that needs no Direct2D,
DirectWrite, or DxUi: native `STATIC` text and two native `BUTTON` children owned by the same HWND.
The text states that the window's graphics could not start, that operations continue, and how to
stop or hide them; it carries the live global status summary and one line per operation (its kind
and status) and refreshes on the popup timer. The buttons are **Cancel all** (the same confirmation
and cancel path as the hosted footer control) and **Close** (hide-only, as the caption Close). No
task graph, queue or speed menu, Replace or Retry action, conflict decision, or other action is
reachable while the failure surface is shown; a debug or hosted-path invoke of any other action is
refused. Waiting tasks keep waiting and running tasks keep running unless Cancel all is confirmed.

`Tab` and `Shift+Tab` move between the two buttons, `Enter` invokes the focused button, and `Escape`
hides the popup. UIA names and Invoke come from the native control classes. Colors are system colors,
so high contrast applies automatically; a DPI change re-lays out the children with the system message
font for the new DPI. Close and reopen re-attempts the host; when it attaches, the hosted popup
returns and no failure-surface child remains. If a native child cannot be created, the popup reports
through the host error surface and Cancel all stays reachable from the main window's File Operations
command. Deterministic runtime coverage MUST be able to force the next host attachment and the next
Direct2D target creation to fail separately, without the production attach-error diagnostic, and
prove the contract above through a real popup fixture.

The host MUST receive the original `WM_SIZE` size type so minimize/restore lifecycle handling stays
truthful. A zero-area or minimized client MUST NOT resize the legacy Direct2D target or force an
immediate paint. Before rebuilding the hosted control tree, the popup MUST reset WindowHost hover,
capture, and focus pointers; focus may then be restored by the stable action identity. This prevents
retained interaction pointers from outliving controls removed during a layout rebuild.

Hosted action automation IDs are stable and encode action kind, task ID, and action data as
`FileOperations.Action.<kind>.<task>.<data>`. Aggregate/task progress and throughput graphs use stable
role/task IDs. Progress controls expose a read-only UIA RangeValue contract with minimum `0`, maximum
`100`, and the current displayed percentage. Throughput graphs expose a semantic Custom control.
The UIA fragment count MUST equal the popup's hosted-control debug counts; hidden or scrolled-out
legacy hit targets MUST NOT leave duplicate semantic controls.

Keyboard behavior is:

- `Tab` / `Shift+Tab` traverse hosted interactive controls in visual order.
- `Enter` invokes the focused action or the actionable decision snapshot's explicit default.
- Left/Right/Up/Down move between actions belonging to the same task when a neighbor exists.
- `Space` pauses or resumes the focused live task.
- `Delete` cancels the focused cancellable task.
- For an actionable conflict/deferred-consent decision, `Escape` invokes that immutable snapshot's
  Cancel action and never Cancel All. The bound decision is the focused task's decision; when focus
  is not inside a task with an actionable decision, `Enter`/`Escape` bind to a decision only if
  exactly one is actionable, so a key press never resolves a prompt the user is not looking at.
  Outside a bound decision, `Escape` hides the popup. Window Close always hides without resolving a
  decision.
- Window Close is the caption Close of the global File Operations popup, not a task action. It MUST
  remain hide-only and MUST NOT cancel the focused task or all tasks. The hidden surface remains
  reachable through **View > File Operations**, `cmd/app/showFileOperations`, and the default
  application shortcut `Ctrl+Shift+J`. Publishing a newly actionable decision automatically shows
  the same popup; metadata-loading attention alone does not. Explicit show and automatic show never
  submit, focus, or otherwise alter the published decision.

Hosted-control creation order is not the navigation contract. After layout, the popup
MUST derive one focus/UIA traversal list from final visible geometry: top-to-bottom,
then logical leading-to-trailing within a row (mirrored in RTL), with a stable action
identity tie-breaker. Hidden or removed controls are excluded. Forward and reverse
keyboard traversal and the UIA sibling chain MUST consume that same list, so rebuilding
cards, footer controls, or conflict actions cannot make semantic order disagree with the
painted surface.

Conflict appearance and terminal transitions to Done, Partial, Failed, or Conflict raise a bounded
WindowHost accessibility notification when the OS provider accepts it. Notifications supplement,
and do not replace, the stable Name/AutomationId/pattern state exposed by the control tree.
The per-task notification generation/history map MUST be pruned whenever a task is no
longer in the live popup snapshot. Teardown clears it completely. Reusing a numeric task
ID after removal starts a new generation; stale history must not suppress or retain the
new task's announcement.

Conflict semantic text is exposed as stable sibling regions in visual reading order: State,
Question, Facts, From/Deleting, To, Last note, then actions. Loading-to-actionable is a real state
transition and is announced once; repaint alone never repeats it. Hidden discovery, transfer, and
graph elements do not remain in the conflict card's UIA tree.

## Progress Rules

Per-file progress has one home: the current file line. Copy/Move cards MUST NOT render a second
under-graph current-item bar. The under-graph progress region is reserved for exactly one whole-task
bar. During streaming discovery it is indeterminate: a visible marquee once transfer is active, never
a fraction over the growing discovered-so-far denominator (D2-A08). The separate discovery indicator
and the exact counters line make the open scope explicit. Compact cards show no meter or percent text
until `discoveryClosed`; percent, ordinary ETA, and the hosted progress value begin only after closure.

The whole-task bar has equal semantic halves when verification is requested: the left half is
transfer/publication and the right half is verification. Verification uses the
`fileOps.progressVerify` fill plus diagonal hatching and a divider, so it remains distinct without
color; the ordinary single bar survives unchanged when verification is Off. A file row currently
being verified uses the same verification fill on its mini bar. The aggregate graph keeps transfer
as the primary series and exposes verification readback as a separate warning-tone series;
provider-proof completion advances the verification bar but creates no fake read throughput sample.
The current and footer bandwidth values include both series, while pump-only Copy/Move throughput
excludes verification. Aggregate ETA is remaining transfer time plus remaining verification time and
is hidden whenever either required denominator/rate is unknown.

The hosted DxUi `ProgressBar` remains the exclusive determinate owner for this whole-task model. For
Verify On, the popup supplies transfer and proof values to its two equal semantic halves and supplies
the themed verification color; the control owns the divider, diagonal hatch, invalidation, and
accessible state. The popup does not drop the hosted descriptor merely because verification is
enabled.

The footer aggregate bar uses live byte totals first and live item totals second. Finished cards may
remain visible when auto-dismiss is off, but all finished cards are excluded from live footer
counters, aggregate totals, need-attention totals, and Windows taskbar progress. Their completed
status remains on the individual card. The footer aggregate is **Known work** (`FO-DISCOVERY-01`): a
compatible closed-total cohort renders determinate progress and `Known work: N%` while other included
tasks still discover, and the status line adds `N discovering, total may grow`; without a closed
cohort the footer is indeterminate and shows no percentage, and an active task whose discovery
closed without any total keeps it indeterminate (nothing is known about it and it will not close
later). The Windows taskbar model stays
indeterminate while any included total is open, even when other active tasks have known totals.

One whole-task presentation decision supplies paint, hosted controls, debug snapshots, UIA,
footer, and taskbar. Unstarted Queue waiting, interlock waiting, conflict metadata loading, and
actionable conflict are excluded from the active aggregate cohort and expose no discovery activity,
transfer bar/marquee, or throughput graph. A running open-discovery card renders the marquee plus
exact counters and contributes only its open-discovery count to the footer; the footer's Known-work
denominator is the closed cohort alone and the taskbar stays indeterminate while that card is
included. Cohort membership and denominator rebasing are deterministic when tasks enter or leave
waiting/attention/running states.

Needs-attention remains visible through the taskbar's error state without publishing a progress
value or re-admitting the blocked task to the active aggregate cohort. Completed compact cards may
retain their final determinate meter as static history; they never contribute that value to the
live footer or taskbar aggregate.

All unsigned aggregate byte/item totals MUST use saturating addition. Extreme task snapshots may
clamp a displayed aggregate to `uint64_t::max`; they MUST NOT wrap backward and expose a smaller
progress value, ETA input, or taskbar total.

Aggregate throughput includes only live Copy/Move rates. Aggregate ETA uses the matching determinate
task/rate set and MUST be hidden when any live Copy/Move task has unknown byte totals or lacks a
usable matching rate; throughput remains visible. Taskbar progress MUST continue to update from the
popup timer while the popup is hidden or minimized, without requiring a paint.

## Typed results, clipboard consumption, and Issues

Cards and completed-group badges derive from the typed per-item axes rather than one HRESULT:
destination publication, verification, source disposition, and item completion. `Copied; source
kept` and `Moved; source folder kept` are Partial/warning results, never successful Move, and keep a
distinct source-retained badge.
Unknown source disposition keeps a separate localized `Source outcome unknown` badge; when one task
contains both retained and unknown sources, both source badges remain visible alongside the independent
verification badge. Unknown/indeterminate axes never remove source rows optimistically.
The canonical result line is localized as `Copied`, `Moved`, `Copied; source kept`, `Moved; source
folder kept` (a rename merge whose emptied source directory could not be removed because a child was
skipped, conflicted, or appeared late), `Skipped`, or `Outcome unknown`; it is carried from the engine result reducer and is not reconstructed from the
aggregate HRESULT. Unknown publication/source uses its distinct indeterminate warning state and
completed-group badge, while a
planned Copy-only or exact retained-source downgrade uses the warning/partial state.

Owned-stage cleanup is a separate axis from primary completion. When a provider proves the requested
mutation committed but reports the cleanup object `Retained`, the card remains Completed/Published,
adds exactly one localized cleanup warning, and uses the canonical result line `Completed; cleanup
item retained.` No automatic retry is attempted. `Unknown` or `RetainedIncomplete` cleanup still
renders as Indeterminate/Possible artifact; a failed or canceled primary is never rewritten to
Completed merely because its cleanup disposition is known.

Verification is a distinct badge/detail axis: Verified uses the success treatment; Failed uses the
error treatment; Unavailable and Canceled use warning/neutral detail without claiming byte equality;
NotApplicable is shown only where the detail surface needs to explain Native/directory behavior;
NotRequested adds no badge. A published destination with Failed/Unavailable/Canceled verification
never renders as a clean success. For Move it also carries `Copied; source kept`, because the source
cleanup gate was not crossed.

For a clipboard Move, accepted admission publishes a non-running card and a gated worker before the
clipboard barrier. Successful consumption records `Move queued — cut list cleared.` immediately.
If selected-root readiness or sequence-matched clearing fails, the accepted task terminates before
discovery/mutation. The card says that the Move did not start and the source remains unchanged; it
must not claim that the cut list was consumed. The host's bounded accepted-sequence ledger still
rejects another admission of that exact sequence. The terminal item rows are dense Failed/Canceled
`NotAttempted + Retained` truth rather than Unknown. After successful clear, the card and Issues view
retain the consumption fact after cancel, failure, Copy-only
downgrade, or `Copied; source kept`; neither Dismiss nor Retry recreates the old clipboard sequence.

When typed results contain retained sources, the completed overflow menu offers `Open source` and
`Select retained` where locations are known. `Cut retained items again` is offered only when every
retained item has a captured no-follow identity and no source disposition is Unknown. Select and
Cut-again rebind the complete set on click and compare object identities; one missing/replaced item
aborts the action and shows a warning. The current `CF_HDROP` Cut-again action is shown only for the
Local file provider. Cut-again publishes a new cut payload as an explicit user action; it never
simulates restoration of the old clipboard sequence.

If any source disposition is Unknown, the card says that some source states are unknown and that a
successfully cleared cut list was not restored. `Open source`/destination remains available where a
qualified path is known, but Select retained and Cut-again are withheld until a later operation has
an exact complete retained set.

Every skipped, warned, failed, retained-source, verification-mismatch, and indeterminate item has an
Issues row with its canonical bucket, qualified source/destination, selected conflict action or
grant, and typed axes. `Skip All` is represented as Skip plus its explicit same-bucket task grant;
Recycle escalation never gains Apply-to-all.

### Interrupted Move restart card

On startup, each valid non-terminal Move breadcrumb projects as a completed warning card with
`Outcome unknown`. The card identifies the last durable `Admitted` or `Executing` phase, admitted
Native/Managed/Copy-only counts, and total qualified source-root count. Its message explicitly says
that RedSalamander was interrupted, the Move outcome
cannot be inferred, and source and destination objects may both exist.

The card's live `taskId` is a newly allocated in-session card/action key. If the historical
breadcrumb identifier is shown, it is carried as `interruptedOperationId` and labeled
**Interrupted operation ID**. UI text, accessibility, commands, telemetry, and dismissal must not
call both values “task ID” or imply that the historical identifier names an active or resumable
task. Every card action continues to address the live session key; Dismiss resolves that card to and
acknowledges only its exact breadcrumb file.

The card can open a representative qualified source location or the qualified destination location
in the source/destination pane roles captured at admission; those actions perform normal pane
navigation/enumeration and do not infer item disposition. The persisted File Operations instance ID
is comparison-only canonical identity: restart navigation must decode `host/default` to the empty
pane context and `host/context/<opaque-context>` to that opaque context, and must reject any other
form rather than navigating with canonical identity text. A pane hint is presentation only and does
not weaken normal navigation/provider validation. The user
may then use the ordinary Refresh command. The card never offers Resume, Retry, Roll back, Delete
source, Delete destination, or automatic cleanup. It is distinct from a Possible name-shape warning
because its breadcrumb contains no object identity or mutation receipt.

The card uses the ordinary indeterminate warning tone without inventing an Issues payload.
`warningCount`, `errorCount`, and issue rows remain zero unless a real diagnostic exists, so Failed
items, Show log, and Export issues are absent on the breadcrumb-only notice. Open destination targets
the recorded destination folder; Reveal item is absent because the breadcrumb deliberately stores no
single destination leaf mapping.

`Dismiss` acknowledges only the card's regular non-reparse JSON breadcrumb directly beneath the
application breadcrumb directory. The host reopens the exact child no-follow beneath a still-identical
regular directory root and retires that file by handle; path text alone is not deletion authority.
It does not touch either provider namespace. Malformed, unknown-
schema, reparse, or out-of-directory records produce no actionable card and are never deleted by this
surface. If terminal retirement previously failed, the same conservative notice may reappear until
the user dismisses it; this is preferable to claiming a known outcome.

### Possible-artifact touch warning

The shared classifier projects only Ordinary or Possible name-shape state. Possible objects remain
visible and badged in FolderView, Find, Search, and Compare. The popup is not a provider-wide scanner,
does not list a recovery inventory, and exposes no Resume, Roll back, cleanup, or automatic-recovery
action. A stale badge, filename, task card, or search-index hint has no ownership authority.

When a user action or another task attempts to Copy/export, externally Open/Edit, Move, Rename,
Delete/Recycle, overwrite, write, or otherwise use a Possible item, reveal the popup/conflict surface
immediately with a prominent name-shape warning. Copy states that bytes may be incomplete;
destructive actions also state that the data may be altered or removed. Cancel is the safe default.
Continue is a one-request grant for the exact displayed identities, has no Apply-to-all/persistence,
and is revalidated immediately before the touch. It never upgrades the item to app-owned or grants
recovery/cleanup authority.

The shared host prompt uses the artifact-touch presentation: `Continue` is the affirmative label,
`Cancel` is the initial/default and Escape action, and there is no Apply-to-all control. If any
candidate cannot supply a current exact identity, the surface explains that the request is blocked
and exposes only acknowledgment; it never turns an unavailable identity into a Continue receipt.
Accepted typed File Operations requests are re-queried on their worker after queue/interlock waiting
and before I/O. FolderView/Find default Shell Open/Open With, configured external viewer/editor
launches, and user-menu programs use the same warning and immediately revalidate the exact referenced
set before process launch; internal viewer/VFS inspection remains prompt-free. Browsing the
current-directory Shell context menu is inspection. After a verb is chosen, the warning appears
before invocation. Opening the Shell Security page also warns first because that page can change
ACLs. Non-recursive Change Attributes and Change Case use this same Cancel-default warning after
their options are accepted and before mutation/worker admission. Change Case must additionally
retain the accepted receipt on its worker and revalidate immediately before each rename batch; that
worker-side step is not implied by its admission prompt. Their recursive forms must pause their
existing discovery/apply worker before the first guarded mutation and marshal the exact-set warning
to the UI thread. The popup does not invent a global artifact inventory.

Displayed speed and ETA values are display estimates and MUST be safe for extreme callback-silence
decay. Before converting floating-point rates or ETA seconds to integer display values, the popup
MUST clamp non-finite/negative values to zero and saturate values above `uint64_t::max`. ETA MUST
not be computed from decayed byte rates below 1 B/s; those rates are treated as zero so the UI
returns to estimating instead of showing a phantom multi-century ETA.

## Completed Card Actions

Completed cards keep `Dismiss` as the primary flat action and expose recovery/navigation commands
from the `More...` menu. The current completed-card action set includes
`Open destination`, `Reveal item`, `Failed items`, `Show log`, and `Export issues`.
`Open destination` and `Reveal item` are available for completed copy/move tasks with resolved
destination data; `Open destination` navigates the destination pane to the completed destination
folder, while `Reveal item` navigates to the parent folder and selects the completed item when a
single revealable destination item is known. `Failed items` is available only when the completed
task has warnings or errors, and MUST open the real File Operations Failed Items pane so the user
can inspect skipped or failed entries without dismissing the card.

The task captures destination plugin ID, plugin short ID, and instance context when the operation is
accepted and the supplied destination filesystem matches that pane. Explicit provider filesystem
objects that do not belong to the pane MUST NOT inherit its identity. Completed navigation resolves
the qualified identity even if the pane later changes provider. Missing, stale, or mismatched
identity MUST fail recoverably and restore the pane's prior provider/context/path instead of
executing against the current provider. Provider changes navigate directly to the qualified target;
the UI MUST NOT visit or enumerate an intermediate provider root.

When a conflict needs metadata, the task card appears immediately as **Needs attention** and its
detail area reads **Reading details...**. The loading state contains no decision buttons, Apply-to-all
control, default/Escape action, Skip All eligibility, button-publishability grant, or accepted
decision. After metadata resolves, the loading state
is replaced in place by one stable, complete action set. For built-in local-to-local **Exists**
conflicts, a file-on-directory collision therefore never publishes **Overwrite**, while a replaceable
file collision includes **Overwrite** in the first actionable layout. Cross-provider prompts retain
their provider-defined action set without depending on local Win32 metadata.

## Custom Speed-Limit Grammar

The custom speed-limit prompt uses the shared binary-throughput edit grammar. Empty input means
unlimited (`0`). A bare number means KiB/s. `B`, `K`/`KB`/`KiB`, `M`/`MB`/`MiB`,
`G`/`GB`/`GiB`, `T`/`TB`/`TiB`, and `P`/`PB`/`PiB` are accepted case-insensitively, with an
optional case-insensitive `/s` suffix. Both `.` and `,` are accepted as locale-independent decimal
separators; results round to the nearest byte and saturate at `uint64_t` maximum. Signed values,
multiple decimal separators, and unknown units are invalid.

The prompt's modal message drain is part of File Operations teardown. It checks the shared host
prompt-shutdown fence before waiting and immediately after every dispatched message. Shutdown clears
the pending edit result and unwinds the prompt; it never continues pumping arbitrary thread messages
after owner teardown begins. The containing File Operations prompt-dispatch scope keeps engine state
alive until that nested modal frame returns. The prompt never retains a task pointer across the pump:
it captures the task ID, re-finds the task after a successful answer, and silently discards the edit
when shutdown or task removal won the race.

The popup preserves its existing boundary policy and trims only the six ASCII whitespace characters
space, tab, carriage return, line feed, form feed, and vertical tab. Text emitted by the paired
throughput formatter (`GiB/s`, `MiB/s`, `KiB/s`, or `B/s`) MUST round-trip through the parser.

## Graph

The bandwidth graph samples smoothed display throughput at popup timer cadence. It MUST not visualize
callback silence as a trough when bytes later arrive in a burst.

The popup graph uses the shared `DxUi::ThroughputGraph` primitive. It shows:

- recent throughput history,
- per-stream colored bands when concurrent transfer streams are present,
- the current effective smoothed bandwidth as a labeled marker,
- the smoothed remaining time right-aligned on the same top label row as the left-aligned bandwidth.

An active Copy or Move card MUST NOT repeat bandwidth or remaining time as standalone detail rows
above the graph. Its expanded base height is 244 DIP before in-flight-row, conflict, and animation
adjustments—36 DIP shorter than the generic 280-DIP active card. The graph's accessible help text
MUST retain both visible values so this compaction does not remove speed or ETA from assistive
technology.

The configured speed limit belongs in the shared current-value selector/flyout, not as the graph
marker. Its visible value is `Limit: Unlimited` or the localized compact configured rate; the whole
selector opens the choices and it has no separate primary action.

In Rainbow theme, graph history keeps per-sample colors stable as samples scroll. It MUST NOT blink or
cycle historic samples by time.

Each active file row's mini progress bar MUST use the exact assigned color of that transfer stream's
band in the graph. The association uses the stream cookie and progress-stream ID, with the source path
as an identity guard; a path-derived hue is only a fallback before the graph has assigned a live stream
hue. Row bars and graph bands use the shared `DxUi::ThroughputGraphColorFromHue` conversion so matching
hues cannot render with different saturation or brightness.

The shared graph supports accent history, stable per-sample rainbow hue weights, per-stream bands,
the configured-limit reference, current-value overlay, and latest-value easing. High contrast uses
system-safe colors and reduced motion disables latest-value easing without suppressing the data.
History is bounded to the most recent 180 samples. Per-stream colors use at most 16 stable reusable
slots (`0..15`), hue-weight deduplication is by slot, and aggregate-only samples MUST NOT allocate a
slot or expand the palette. The shared semantic gate disables per-stream bands in high contrast;
Rainbow may show one admitted stream; the normal theme requires at least two concurrent admitted
slots. The hosted graph consumes this same gate and slot assignment.
Continuing streams MUST occupy their previously assigned slots before leftover
slots are given to newly appearing streams, so unique live-stream colors hold
when in-flight callback order permutes (same cookies, different array order).

The area and every per-stream band are translucent (18% alpha in light themes, 22% in dark themes;
High Contrast uses its explicit system-safe treatment). Hue selection MUST preserve the caller's fill
alpha instead of replacing it with an opaque hue. Discovery activity is rendered outside the graph;
status motion may draw lightweight strokes/dots over an otherwise empty plot but MUST NOT fill the
graph surface or obscure live history.

Band construction is bounded by `O(samples * activeSlots)` with a maximum of 180 by 16 logical
quads. The accepted painter rasterizes into one fixed 180x64 premultiplied-BGRA buffer, uploads one
`ID2D1Bitmap`, clips it with one covered-area geometry/layer, and issues one bitmap draw. It MUST NOT
reintroduce one geometry or draw submission per logical quad. The instrumentation contract reports
sample count, active slot count, logical quad count, rendered geometry count, and painter duration.
Per-stream bands form a complete stacked area from the graph baseline to the interpolated throughput
envelope. Adjacent band geometries MUST share exact boundaries and the area MUST retain a covered base
fill, so descending segments expose neither baseline gaps nor black seams between streams. The stack
is clipped to the envelope, so its upper edge follows and remains below the history stroke instead of
exposing square sample-column tops.

## Motion

The popup resolves reduced-motion through the app theme's DxUi palette. When reduced motion is
enabled:

- Automatic popup resize snaps to the target rectangle instead of debouncing/easing. Footer-only
  collapse still snaps to the footer minimum, and expand still restores the captured rectangle.
- Task, completed-group, and footer disclosure chevrons snap directly between their contextual
  collapsed direction and down.
- Dedicated discovery activity and graph status animation are static; neither suppresses throughput data.
- The graph latest-point easing is disabled; throughput history, rainbow/per-stream bands, and the
  current effective-bandwidth marker remain visible as static state.
- Task-card height, completed-group visibility, and hosted progress-value transitions snap to their
  target values.
- Every indeterminate global/task/file/conflict bar renders a stable centered segment instead of a
  moving marquee.

## Conflict Layout

Conflict prompts are inline on the affected card. Source and destination paths MUST be stacked in
full-width rows rather than side-by-side columns so long names keep useful width. The primary action
row exposes at most three buttons plus a `More...` menu for overflow actions.

The popup is a presentation-only consumer of the immutable engine conflict policy. It copies and
renders the supplied complete action order, primary/`More...` placement, default action, Escape
action, Apply-to-all and Skip All eligibility, metadata-loading state, and button-publishability
flag. It MUST NOT switch on the conflict bucket to add, remove, reorder, place, or select a default
action. Primary clicks, `More...` selections, `Enter`, and `Escape` are accepted only when they match
the currently published snapshot placement/action. A stale or unpublishable action fails closed.

Every active prompt begins with a persistent, nonanimated localized state line. While facts are
being decorated it reads `Loading decision details...`, exposes no action buttons, and contains no
stale discovery/progress/graph affordance. Once actionable it reads `Waiting for your decision`,
followed by the decision question, facts and object context needed to answer it, and only then the
actions. Waiting is not represented by a marquee, ellipsis animation, pulsing, or a fake transfer
percentage.

Question, facts, From/Deleting, To, and Last note use measured DirectWrite wrapping; their measured
height contributes to the card height. At the supported minimum width and all supported density,
locale, and DPI combinations, these decision-critical regions must remain unclipped and must not
overlap the actions or footer.

Checked File Operations controls, including conflict `Apply to all`, render the checkmark through the
existing DirectWrite icon format (`Segoe Fluent Icons` preferred, Unicode fallback) instead of custom
line geometry. The box fill, border, theme colors, hit target, and state semantics remain unchanged.
The Recycle-failed escalation prompt never renders `Apply to all`; its Permanent delete grant is
one-shot for the exact revalidated item. Cancel is the first keyboard/default action, followed by
Delete permanently and Skip, so Enter or initial focus cannot authorize deletion.

Every other actionable conflict also keeps the engine-published Cancel action directly visible. If
two consequence actions already occupy the first two primary slots, Skip moves to `More...`; Cancel
does not. The hosted default and cancel control identities both map to the published Cancel action.
Closing the popup hides it without changing the prompt; Escape invokes Cancel for that prompt.
If Close hid the popup before the prompt became actionable, publication of the completed immutable
action set restores the popup automatically. If Close hid an already-actionable prompt, the decision
remains parked and **View > File Operations** / `Ctrl+Shift+J` restores the exact same surface.

Deferred-consent prompts reuse this inline card surface but remain distinct from ordinary retryable
conflicts. The card shows the risk-specific message and, when known, a localized fact line for item
count and/or bytes; unknown facts are omitted rather than estimated. The published action order is:

| Risk | Primary actions | Final/default action |
|---|---|---|
| Insufficient space / space unknown / sparse inflation | Continue | Cancel |
| Placeholder hydration | Download and continue | Cancel |
| EFS plaintext | Continue as plaintext / Keep source | Cancel |
| Metadata loss on Move | Move anyway / Keep source | Cancel |
| Recycle failure | Cancel / Delete permanently / Skip | Cancel |

Eligible non-Recycle risks may expose `Apply to all` only for the exact task/risk/destination-root
scope. Recycle never exposes it. The prompt snapshot includes the deferred-consent marker, known
item/byte facts, and optional source/destination identity so tests and assistive presentation can
distinguish a consent gate from an ordinary conflict without reading path text. An action not present
in that immutable snapshot is rejected as Cancel. Closing the popup does not grant consent, persist a
choice, or make the prompt session-sticky.

For eligible non-Recycle Apply-to-all, one visible acceptance creates one receipt for the complete
task/risk/destination-root scope. Later matching gates reuse that scoped grant; the UI and receipt log
do not claim a separate object-specific acceptance. Recycle remains exact-item only.

When metadata is available from the source/destination `IFileSystemIO`, each stacked conflict row
shows compact size and modified-time metadata on the right side of the label row while keeping the
full-width ellipsized path on its own row. Modified times MUST convert the provider's UTC FILETIME to
local time and use the user's short-date/time formats. Provider metadata reads call
`GetFileBasicInformation` first and call `GetAttributes` only as a fallback when basic information
fails. Missing metadata is omitted rather than guessed. A Copy/Move collision exposes `Keep Both`
when the operation can retry that exact colliding item with a unique sibling destination. The
action preserves the existing destination and retries the incoming item as `name (2).ext`,
`name (3).ext`, and so on (or `folder (2)` for directories), probing existence through the
destination provider. Nested Local recursive Copy collisions are renamed in their own parent; they
never cause a second copy of the selected top-level root. Keep Both is eligible for Apply to all
only when the exact typed conflict scope permits it, and is available to same-filesystem and
cross-filesystem transfers.

The prompt action set is driven by the typed conflict class. Regular file Exists offers Overwrite;
read-only Exists offers Replace read-only; type mismatch never offers Overwrite; an exact final
destination link offers Replace link and never follows or replaces its target. Apply to all is
rendered only for an eligible exact scope. Destination-link scope includes the known link kind
(file symlink, directory symlink, or junction); an unknown kind is one-shot. `Skip All` is a
separate More-menu action that records the scoped Skip grant. Ordinary Skip never checks or stores
Apply to all. Recycle escalation and semantic target-conflict retention never expose Apply to all.
Overwrite, Replace read-only, and Replace link are omitted when the provider cannot return and
consume an exact no-follow destination receipt; the UI never presents a destructive button that is
known to end in `ERROR_NOT_SUPPORTED`. Target conflict exposes no mutation action. Replacing a
selected-root link does not suppress a later child collision prompt.

Prompt identity and actions are published under the conflict-arbiter lock before decoration.
Metadata calls and diagnostic logging run outside that lock, and late results merge only when the
same owner/bucket/status/paths are still active. Decoration MUST NOT create an `IFileReader` or read
content. The built-in local provider may use `GetFileAttributesExW`; provider paths use metadata-only
`IFileSystemIO` calls and may leave size unknown.

The Failed Items/Issues pane has one close contract regardless of whether close originates from its
window chrome or the popup toggle: save view state and placement, hide (do not destroy) the pane, then
restore active-pane FolderView focus when focus belonged to the hidden pane.

## Debug Snapshot Contract

`PopupLayoutDebugSnapshot` is the non-pixel acceptance surface. New UI controls or durable visual
contracts MUST extend it when feasible. Current closeout fields cover:

- footer control count and overlap detection,
- selected-task button overlap detection,
- single Queue/Parallel selector visibility, current mode, and live hit target,
- Options overflow visibility and live hit target,
- hidden legacy segmented-mode, auto-dismiss, and density footer controls,
- footer bulk Pause all / Resume all visibility and target action,
- compact-density persisted state,
- footer-only details toggle placement,
- high-contrast state and color-blind-safe status signal coverage,
- reduced-motion state, active layout-motion state, and animation enablement for auto-resize,
  disclosure controls, and graph status overlays; normal-motion geometry acceptance waits for a
  synchronous settled snapshot before applying minimum-width overlap gates,
- aggregate footer progress totals,
- aggregate ETA/throughput,
- taskbar progress state/value, timer-update count, taskbar-button readiness, interface availability,
  initialization attempts, and retry-pending delay,
- status stripe/chip visibility, tone, and resolved theme color,
- queued Start now visibility,
- queued Move Up/Move Down edge visibility and queue-order keys,
- confirmation Links/Verify/clipboard-consumption control state and immutable accepted snapshot,
- discovery mode, traversal-closed state, Skip-discovery visibility/release generation, and truthful
  provisional/final denominator state,
- Inline-F2 origin, reveal reason/deadline, ever-revealed state, and silent-clean-success suppression,
- typed publication/verification/source-disposition/item-completion axes, retained-source badge, and
  clipboard-consumed detail state,
- completed Open destination / Reveal item / Failed Items menu-action visibility,
- task collapsed/compact-row state, completed auto-collapse state, and compact progress visibility,
- completed-group visibility, expanded state, total and per-result badge counts, group toggle/clear
  visibility, and per-task hidden-by-completed-group state,
- duplicate progress-bar absence,
- conflict stacked-path layout,
- conflict source/destination metadata visibility and size/date compare coverage,
- conflict waiting/loading-state visibility, prompt-before-actions order, wrapped prompt/context/
  last-note fit, and absence of conflict discovery/transfer indicators,
- asynchronous popup-setting save generation/thread ownership and bounded flush coverage,
- graph current-bandwidth marker and label, right-aligned ETA label, absence of duplicate inline
  speed/ETA rows, and the operation-specific expanded base height,
- failure-surface state, native child count, and text, plus actual active-row canonical-color
  matches and mismatches,
- same-HWND WindowHost use, hosted action/progress/graph counts, focused automation ID, and accepted
  accessibility-notification count,
- task-card, completed-group, hosted-progress, and shared-graph motion state under normal and reduced
  motion.

Taskbar retry selftests drive the same `UpdateTaskbarProgress` path synchronously after injecting the
failure, capture the pending snapshot before pumping unrelated window traffic, wait the published
retry delay, and drive the retry explicitly. They do not depend on observing a one-second transient
state through a bulk message-queue drain.

The focused runtime guards `cmd_pane_fileops_popup_failure_surface_on_host_attach_failure` and
`cmd_pane_fileops_popup_failure_surface_on_d2d_target_failure` force each failure separately and
prove the failure-surface contract (native text with the live summary, Cancel all and Close, Tab,
Enter, Escape, UIA names and Invoke, high contrast and DPI, refused actions, hide and reopen, and a
hosted popup after close and reopen). `TestThroughputGraphBandsStayBelowHistoryLine` preserves
the band/history pixel contract, and `TestThroughputGraphHueChurnPerformanceScenario` records the
deterministic bounded-work metrics.

Accepted same-machine Release evidence is archived at
`Specs/TestRuns/4cb089111a23/FileOps/2026-08-15_175346_graph_before/` and
`Specs/TestRuns/4cb089111a23/FileOps/2026-08-16_064500_graph_after/`. Across five captures, median
frame time improved from 18,493 us to 16,583 us (10.3%) and mean from 18,768.4 us to 16,593.2 us
(11.6%). The accepted after captures report painter median 81 us, 180 samples, 16 slots, 2,864
logical quads, and one rendered geometry on every measured row.

The translucent-fill restoration candidate is archived separately at
`Specs/TestRuns/7d3a1247382a/FileOps/2026-08-27_165050_graph_transparency_release/`. Its five
Release captures retain the bounded 180-sample, 16-slot, 2,864-quad, one-geometry workload and
report a 111 us median band-painter cost. This is candidate-only evidence because the accepted
before/after archive belongs to a different machine profile; it does not establish a cross-profile
performance comparison.
