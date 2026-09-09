# File Operations features deferred after the reliable core

> **HOLD / NON-NORMATIVE PLAN.** Owner: product owner for scheduling, then one named
> executor per promoted feature. Planned at `58f22d0e`, 2026-09-05. The owner explicitly
> moved the original four features below out of the reliability milestone. On
> 2026-09-08 the owner also deferred Microsoft Drive same-endpoint Copy and a
> qualified MTP native Move route after correcting their inaccurate descriptions.
> All six features retain their
> accepted intent but are not authorized for implementation by this HOLD record and
> do not block [the core plan](../Done/Operation_FileOperations_ReliabilityFirstSuccessorPlan_2026-08-29.md).

This is a later-feature backlog, not a second reliability engine. Before starting,
review the live core's exit receipt and current domain specs, reconcile drift with
`git diff 58f22d0e -- Common RedSalamander Plugins Tests Tools Specs`, choose one
feature, and name its implementation scope and tests. No feature may weaken the
core's exact authority, immutable result, cancellation, consent, or privacy rules.
The original accepted decisions and historical evidence remain in
[the execution record](../Done/Operation_FileOperations_ReliabilityFirst_ExecutionRecord_2026-09-05.md).

## Deferred scope

The 2026-09-08 request for a **Confirmation** Preferences page and accurate outcomes
when metadata cannot be preserved has its own active owner:
[I18 — Confirmation preferences and accurate Copy/Move outcomes](UI_ConfirmationPreferencesAndOperationOutcomes_2026-09-08.md).
That specification reviews every supplied checkbox/message and default, includes
Copy/Move, paste and drag/drop start confirmation, and retains unsupported archive
editing/update and NTFS attribute-operation confirmation rows as deferred capability
work. It does not add a seventh feature to the six HOLD rows below or reopen the
completed reliability execution record.

| Feature | Former decision/slice | Dependency | Why it is later |
|---|---|---|---|
| Persistent searchable operation history | D2-A18 / R6-A18 | Immutable results, core final qualification | Useful explanation after cards close, but adds persistence, privacy, retention, settings, search, and export. |
| Continue safe work and review conflicts later | D2-A17 / R6-A17 | Core conflict arbiter, safe Retry, responsive lifecycle | Adds deferred authority/state and a review journey; ordinary immediate conflict prompts cover the first version. |
| Retarget links inside the copied tree | Optional part of D2-A03 / R7-A03 | Literal Preserve, owned publication, options surface | Adds a bounded dependency/mapping graph and partial-publication rules. Literal link copying belongs to the core. |
| Ordinary Local Shell Copy/Move experiment | D2-A16 / R9 | Final custom route/result/cancellation contracts | Alternative engine experiment must demonstrate a real reduction in complexity. No production adoption is implied. |
| Microsoft Drive same-endpoint/server-side Copy | I16 BR-1, owner deferral 2026-09-08 | Typed provider capability, exact asynchronous mutation outcome | Same-qualified-endpoint Copy is currently unsupported; only cross-endpoint bridge import/export Copy exists. |
| Qualified MTP native Move | I16 BR-2, owner deferral 2026-09-08 | Isolated WPD move route and typed non-commit/uncertain outcomes | Host Move is currently Copy-only/source kept. Neither native Move nor destructive host-managed Move is qualified. |

All six rows are HOLD under this document's status. Do not duplicate current core
statuses here. Promotion requires scheduling by the owner; dependencies are technical
prerequisites, not a reason to delay core retirement.

## Persistent history

Consume immutable completed attempts (`CompletedTaskSummary` and typed item results)
without retaining Task/provider lifetime or opaque mutation authority. Use the
existing safe durable-file mechanics from `FileOperationDurableStore.h`, with a
separate versioned non-authoritative history schema. Never read artifact claims or
Move breadcrumbs as history, or history as recovery/execution authority.

Entry points: View in history from completed cards and Operation history from the
jobs options menu. Support result/status/date/endpoint filters, search, navigation
to still-resolvable locations, explicit selected-field Export, Clear, and Disable.
Partial, failed, retained-source, cleanup-debt, and uncertain results must be as
discoverable as successes. Opening history does not replay work or steal focus.

Use both age and storage bounds, oldest-first eviction, and measured numeric
defaults in the settings/schema authority. History is local and enabled when this
feature lands. Disable stops new writes without silently clearing old entries;
Clear removes only explanatory history. Export contains the selected visible
redacted fields. Never persist credentials, tokens, secret-bearing URL fragments,
or raw opaque authority. A saved row cannot Retry, Resume, Undo, or delete source;
an intended new operation goes through ordinary admission.

Likely owners: `FolderWindow.FileOperations.State.Diagnostics.cpp`, popup/Issues
surface, `Preferences.FileOperations.*`, `Common/SettingsStore.*`, settings schema,
File Operations/UI/settings specs, and relevant Commands tests. Re-read the actual
result lifetime and settings page patterns before implementation. Tests must cover
restart, malformed/truncated records, storage failure, age/size eviction, Disable,
Clear isolation, redacted export, and search/cancellation costs at the retention bound.

## Review conflicts later

Default stays **Ask as issues occur**. Offer **Continue safe work and review later**
as a task-local choice in the existing options surface and eligible task menu;
do not turn it into a persistent Preference without a further product decision.

Only independent, unambiguous work continues. Park an affected item before its
guarded mutation, retain/revalidate exact authority under the normal rules, and
collect it once in original task/item order. Show `N items need review` without
stealing focus. Selected roots/scope remain fixed; bounded descendant discovery
inside already admitted containers remains permitted. Rescan creates a new plan.
Existing Apply-to-all decisions stay immutable and exactly scoped. New or changed
conflicts return to the same review. Cancel records unresolved items as unattempted
with retained source; closing the review does nothing. Only proved no-commit attempts
may Retry; uncertain commit is terminal explanation, not a deferred mutation action.
Injection overlap decisions occur before acceptance and cannot be collected.

Likely owners: existing task/decision snapshot in `FolderWindow.FileOperationsInternal.h`,
`FolderWindow.FileOperations.State.cpp`, popup/Issues controls and tests. Reuse the
single conflict arbiter; no secondary action/default policy. Tests cover independent
progress, dependent work parking, exact ordering/cardinality, stale authority,
bounded retained memory, cancel/close, and keyboard/UIA. Design backpressure before
holding arbitrary item/handle sets; no disk work spool is implied by this feature.

## Explicit link retargeting

Core **Copy links unchanged** remains the default and has no retarget graph cost.
Only the separate per-task **Retarget links inside copied tree** option may rewrite
a target, using an exact admitted mapping inside the copied tree. Preserve and
explain external, unresolved, skipped, failed, cyclic, or ambiguous target cases;
never silently guess a destination or claim Retargeted when no rewrite happened.

Define bounded indexed source/destination mapping, collision and forward-reference
handling, cycle behavior, partial publication, cancellation, and per-link result
receipts. Do not restore repeated linear scans or a 4,096-record terminal failure
as a limit on otherwise valid literal copying. Specify a truthful limit/refusal or
bounded work policy for the optional transform itself before coding. It must remain
independent of literal Preserve and must not authorize source deletion.

Likely owners: Local link payload adapter in `Plugins/FileSystem/FileSystem.FileOps.cpp`,
bridge link handling in `FolderWindow.FileOperations.State.cpp`, the existing options
surface, provider typed contracts only where necessary, and link tests/specs. Test
relative/absolute symlinks, junctions, provider support/refusal, Keep Both names,
forward references/cycles, failure/cancel, external links and retained source.
Measure mapping lookups, retained bytes/handles, and scale beyond the old bound.

## Shell Copy/Move experiment

This is evidence only. Test exactly two non-default candidates: ordinary new-name
Local regular-file Copy, and no-conflict same-volume Native Local regular-file Move.
Existing Shell Recycle is the comparison control. Exclude overwrite, folders/links,
cross-volume or Managed Move, Permanent Delete, Rename, merges, remote/device,
recovery, and Batch Rename. Use an isolated experimental adapter; no production
route or capability changes before an explicit adoption decision.

Document/test the complete `IFileOperation` flag matrix, dedicated STA ownership,
progress sink, `PerformOperations`, and `GetAnyOperationsAborted`. Suppress Shell
prompts; raced destinations must not authorize Shell overwrite/rename/conflict
decisions. Omit `FOF_ALLOWUNDO` and `FOFX_ADDUNDORECORD` and verify no Undo record.
`FOFX_NOSKIPJUNCTIONS` is not proof of the core's literal link semantics.

Missing terminal callbacks, ambiguous HRESULTs, and aborted operations map to
uncertain/contract-violation truth unless exact per-item facts prove otherwise.
Test destination races, partial completion, missing callbacks, cancel, retained
source, helper/STA startup failure, shutdown, and callback lifetime. Measure startup,
throughput, memory, cancellation, and actual removed production owners/code against
the final custom routes. Return Adopt/Reject/Needs bounded proof separately for
Copy and Move. Reject an experiment that creates a second policy owner or does not
materially simplify the code. Experimental code is removed when rejected.

## Microsoft Drive same-endpoint Copy

Current `CopyItem` returns `ERROR_NOT_SUPPORTED` and the typed descriptor does not
advertise same-endpoint Copy. Cross-endpoint bridge Copy does not establish a
same-drive route. I16 BR-1 closes the inaccurate specification claim; this HOLD
row owns the future implementation after promotion.

Define the Graph copy action, asynchronous monitor lifetime, terminal result and
reconciliation contract before advertising the capability. Preserve destination
conflict rules and exact endpoint/item identity. Unknown monitor/commit outcome
must remain indeterminate, with no automatic replay or false success. Resolve any
additional permission requirement with the product owner before expanding consent.

Likely owners: `Plugins/FileSystemMicrosoftDrive/`, its fake Graph backend, the
provider capability spec and File Operations provider table. Test accepted/pending/
completed/failed monitor states, cancellation and disconnect, destination races,
stale authority, malformed replies, uncertain completion, and callback teardown.
Prove zero host payload download/upload on the server-side route; measure poll
count, retained resources, cancellation latency and terminal reconciliation.

## Qualified MTP native Move

Current `nativeMove=false` is intentional: the legacy Move entry point can choose
WPD move or stream-copy/delete, so it is not proof of an isolated native route.
The provider also lacks bound/conditional source-deletion authority. Host strategy
selection therefore keeps the source, including for a same-device request.
I16 BR-2 closes the description/capability-truth item; this HOLD row owns the future
route after promotion. Do not describe current behavior as copy-then-delete.

Isolate a WPD move path whose capability is truthful for each device/session and
object kind. Advertise native Move only with tested success, cancellation,
non-commit and uncertain outcomes; never silently fall through to copy/delete.
Keep device serialization, MTA ownership, watchdog/cancel/quarantine and unload
quiet-point contracts. Unknown device outcome retains source-deletion restrictions
and requires reconciliation before any retry. A destructive host-managed Move
would need separately qualified source-deletion authority and is not implied here.

Likely owners: `Plugins/FileSystemMtp/`, fake WPD backend, provider and File
Operations specs. Cover unsupported/read-only/disconnected devices, raced identity,
device failure during mutation, cancellation, no fallback deletion, and teardown.
Require deterministic fake-backend evidence plus a separately recorded physical
device qualification; measure watchdog, cancellation and retained-session bounds.

## Remaining audit work and execution owners

This table is the entry point for all remaining work identified by the 2026-09-08
File Operations audit. It routes existing work without opening duplicate execution
queues. Historical Done records remain frozen; missing qualification is not itself
proof of a missing implementation. Closure evidence belongs with the active owner.

| Work | Execution owner / promotion route | Remaining proof or decision |
|---|---|---|
| P1 cancellation, P2 test races, and Phase 7 reliability | Complete in [I16 BR-3](../Done/FileOperations_BeelineResidualsAndProviderRoutes_2026-09-08.md), 2026-09-08 | FileOperations passed 178/0/20 in Fresh Full, including 78-ms queued cancellation, both Phase 7 cases and tasks-quiet/popup-host shutdown. Preserve these regression guards; no remaining implementation is queued here. |
| Placeholder, Curl directory Move, popup timing | Complete in [I16 BR-4/BR-5/BR-7](../Done/FileOperations_BeelineResidualsAndProviderRoutes_2026-09-08.md) | Real NTFS Cloud-placeholder worker Copy, one host FTP directory rename/refusal pair, and post-destruction popup focus restoration passed in Fresh Full. No remaining implementation is queued here. |
| Fresh Full Preferences Keyboard follow-up, 2026-09-08 | Complete in [I16 BR-8](../Done/FileOperations_BeelineResidualsAndProviderRoutes_2026-09-08.md#br-8-preferences-keyboard-search-input-identity) | Corrected the shared wait to require the exact native input HWND. Controlled RED/GREEN, 72 predecessor-context passes and canonical Commands 896/0/2 are archived; final Fresh Full passed all 35 entries. |
| Five stress-evidence areas | [I3 stress validation](FileOperations_RemainingStressValidation_2026-08-04.md) | Reconcile historical case/behavior vocabulary against current code before executing the remaining stress rows; do not restore superseded contracts. |
| Completion, reaper, shutdown and focus follow-up | [I12 bounded review follow-up](../Done/Operation_ReviewFollowup_FOTerminalFocus_2026-08-19.md) | Completed September 8 with repaired terminal entry points, repeated Commands proof and green Fresh Full. New completion/lifecycle changes preserve the authoritative contracts and need a current active owner; I3 retains stress evidence. |
| Human interaction qualification (retired core C5) | HOLD here until a named validation owner is promoted | Mouse, keyboard, UIA, RTL/translations and DPI observations across the operation/conflict/result journey. Automated passes do not replace these observations. |
| Live provider/environment qualification (retired core C0) | HOLD here; coordinate Google Drive scope with [I0](FileSystem_GoogleDrivePluginPlan.md) | FTP/SSH/quotas, IMAP TLS, physical MTP and Google Drive mutations, using owned fixtures and available credentials/devices; retain exact skip reasons when unavailable. |
| Bounded batches and IMAP discovery buffering | HOLD here for product decision, then a named implementation owner | Decide the bounded-batches requirement; assess whole-UID `SEARCH` buffering and define a truthful bounded discovery/backpressure contract before implementation. |
| Curl fixture bound (retired core C11) | HOLD here until a named diagnostic owner is promoted | Explain and qualify the recorded PASV-bound residual with fixture/transport evidence; do not increase the bound without a cause. |

The unowned rows retain the residuals recorded in the frozen
[core successor plan](../Done/Operation_FileOperations_ReliabilityFirstSuccessorPlan_2026-08-29.md)
and [execution record](../Done/Operation_FileOperations_ReliabilityFirst_ExecutionRecord_2026-09-05.md).
Their retirement did not certify every environment, interaction or later feature.

The 2026-09-08 BR-3 Full archive is
`Specs/TestRuns/4cb089111a23/FileOps/2026-09-08_073251_br3_fresh_full/`;
the isolated Preferences follow-up is
`Specs/TestRuns/4cb089111a23/Commands/2026-09-08_084447_br3_full_preferences_followup/`.
That historical gate failed on Preferences and its three isolated passes did
not clear it. I16's completed repair now has a new
[Fresh Full closeout](../../TestRuns/4cb089111a23/FileOps/2026-09-08_104319_br4_br5_br7_keyboard_fresh_full/README.md):
2,113 passes, zero failures, 52 declared skips, all 35 entries promoted.
The remaining rows above and all six deferred features retain their own scope;
this bounded closeout does not qualify every human interaction or live environment.

## Delivery and verification for any promoted feature

Use repository C++/WIL/threading/resource/localization conventions and the shared
helper catalog. Update the owning domain specs with implementation/tests, and archive
measured evidence under `Specs/TestRuns/`. Read the applicable project skills before
editing source, Preferences, public ABI, or tooling. Tool changes require the existing
inventory/help/caller contract; this plan does not authorize a general tooling rewrite.

Build through `.\build.ps1 -ProjectName RedSalamander -Configuration Debug` and
test-enabled Release (`$env:RSBuildEnableTests='true'`, restoring the prior process
environment afterward). Discover exact cases from the matching test-enabled
`RedSalamander.exe --selftest-list-cases --commands-selftest --fileops-selftest`:
wait for exit and capture its JSON stdout beneath the owned test root.
`Get-TestInventory.ps1 -Format Json` exposes harnesses/counts, not all native case names.
Use the governed runner with its build enabled for focused tests; `-SkipBuild` requires
a compatible full-solution receipt and byte-identical artifacts, not a project-only
build. Require nonzero expected case counts, zero failures,
source/resource policy checks, `.\Tools\Get-SpecInventory.ps1 -FailOnFindings`,
`.\Tools\Test-TestRunArchive.ps1 -Inventory`, `git diff --check`, and an exact final
`.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh` for the delivered feature
set. Runtime data stays in initialized `X:\RedSalamander.Perf`; interactive tests run
sequentially. No old core receipt qualifies later source changes.

A promoted feature is done only when its required behavior, privacy/lifetime/fault
tests, measured resource bounds, authoritative specs, and final qualification are
complete. Keep unfinished features explicit. Stop only the affected work when proof
would be weakened or a new product decision is required; continue independent work.

## Routed risk retained from the original review

The earlier audit recorded legacy plaintext password/passphrase support in
`Specs/FileSystem/FileSystem_Imap.md`, `Specs/FileSystem/FileSystem_FtpSftpScp.md`,
and `Plugins/FileSystemCurl/FileSystemCurl.Shared.cpp`. This is a separate security
review candidate, not an extra File Operations feature or a claim that a live secret
has been found. Before any migration, recheck current code and supported stored-data
behavior and name a credential-storage owner. Reuse Connection Manager/WinCred;
do not reproduce secret values or automatically clear user settings. Any confirmed
secret disclosure in File Operations diagnostics must be fixed in the core's privacy
boundary rather than deferred under this note.
