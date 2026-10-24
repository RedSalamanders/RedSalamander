# Terminal Pane Tabs, Profiles, And View Lifecycle

> **Status — Completed 2026-08-03.** The v1 pane command, folder Edit, opposite-pane reuse, Windows/WSL shell
> selection, path insertion, source-follow, profile/settings, and lifecycle behavior are implemented. Durable behavior
> is in `../../Terminal/Terminal_EmbeddedPlugin.md`; multi-terminal-item and asynchronous host-callback machinery are
> outside v1. Later WIP/checkpoint language below is historical.

## First-implementation checkpoint (2026-08-01)

The first pane slice is implemented: distinct Folder/Preview/Terminal tabs,
opposite-pane open/reuse, plugin-owned child HWND lifetime, Windows/UNC/WSL
launch mapping, directory Edit, the stable insertion commands, and synchronous
terminal close. Exact-root-PID PowerShell cwd/idle authentication, unsafe-paste
confirmation, passive UIA Text exposure, history-free follow, and deterministic
lifecycle coverage are green. The multi-tab profile chooser, production
accepted-command/history and semantic-close adapter, advanced UIA mutation and
geometry, refresh-retention, and full lifecycle matrix below remain WIP.

The working slice accepts insertion only when the request still matches the
target terminal's original source key, latest source generation, namespace, and
source path. The independent initiating-source behavior specified below for a
compatible selected terminal from another logical source is still release-open
with the multi-tab/SwapPanes work; the temporary stricter rejection is
fail-closed.

> **Authority:** Execution plan 4 of 6 for `Terminal_EmbeddedConPtyLibGhosttyVt_IntegrationPlan_2026-07-13.md`. The master owns behavior; plans 1–3 must be green first.

## Status and drift

- **State:** WIP / first pane and profile-routing slice implemented
- **Baseline:** `fbf6af76f373bf253488862426f4ea70a9ff25ba`
- Enter only with the exact `Specs/Plans/Done/Terminal_RendererControlAndAccessibility_2026-07-22.md` closeout commit/blob, its predecessor engine/core evidence identities, its exact Renderer `TerminalNative` manifest/companion paths and digests, and its exact renderer Commands archive/digests. Recompute the Done blob and rerun `Tools/Verify-TerminalNativeEvidence.ps1`; a missing/mismatched path, implicit latest-run lookup, skipped native invocation, or repository/engine-lock mismatch blocks implementation.
- Recheck current `FolderWindow*`, `DxUi::TabControl`, `FluentIcons.h`, `Common/PlugInterfaces/{Viewer,Terminal,Factory,Informations}.h`, `ViewerPluginManager.*`, `ManagePluginsDialog.*`, command registry/defaults, the completed pseudo-command-line retirement, Edit/editor resolution, Preview tests, the canonical shared WSL-catalog work, plugin settings/profile helpers, and reentrant callback/module-quiet specs before editing.
- Before implementation, compare those surfaces and plans 1–3 with the baseline. Any material ownership, ABI, lifecycle, settings, test-harness, or UI-contract drift is a STOP: update/re-review this plan and record the new baseline before changing product code.

## Deliverables

- Stable dynamic Folder/Preview/Terminal pane content model, optional generic tab icons, and one generic `ITerminal` plugin object/child per Terminal tab, using the frozen 336-byte Open context with independent source and launch locations.
- Reused `cmd/pane/openCommandShell` (`Ctrl+Alt+T` and `Alt+7`), stable no-default `cmd/pane/openCommandShellWithProfile`, explicit external-shell replacement, directory Edit, contextual/full path insertion, pseudo-command-line retirement, chooser/diagnostic/start/run/exit/restart/close view states.
- Keyboard/UIA contracts for pane tabs, profile chooser, diagnostic, unsafe-paste, exit, and restart surfaces introduced here; reuse the plan-3 base terminal-document provider, with unsafe-paste content/token ownership remaining inside `Terminal.dll`.
- Application-unique `OriginalSourceKey{FolderWindowInstanceId, PaneSourceId}`, separate `PhysicalHostKey{FolderWindowInstanceId, PhysicalPaneId}`, exact swap/zoom/focus/ordering behavior, session-to-view IDs.
- Plugin-owned safe cmd/Windows PowerShell/pwsh/WSL profile catalog and working-directory launch plans, including canonical Windows-shell OS/integration environment deltas consumed by the core's exact bounded UTF-16 block builder, exact Windows-versus-matching-WSL source policy, the WSL five-token/default-user/default-shell terminal-only boundary, and `windowsDefaultProfileId` setting seam.
- Two distinct plugin-only Windows launch protocols: every local, validated-
  plugin-backing, or UNC cmd launch uses one protected-inbound-pipe, exact-root-
  PID, capability/nonce-correlated one-use
  `LaunchCwdBootstrap`, while every local/UNC Windows PowerShell/pwsh launch uses
  the mandatory `PowerShellStartupBootstrap`. The latter restores the requested
  directory after normal profiles, samples the latest integration policy,
  exercises the generic semantic join seam, acknowledges logical cleanup, wins
  one lifecycle-locked Running CAS, and completes an authenticated prompt commit
  before input or profile memory opens. Plan 4 production always resolves the
  semantic branch as `SemanticUnavailable`; only an injected fake may prove the
  `SemanticReady`/two-slot join branch, and no production semantic capability is
  claimed before plan 5.
- Injected `ITerminalPreferenceStore` seam plus an in-memory fake that proves profile-precedence behavior without defining persistence; plan 5 supplies the durable generation-fenced implementation.
- Commands/DxUi/plugin-ABI regression coverage without changing `Alt+6` Preview. `Ctrl+Alt+T` intentionally migrates to embedded Terminal; the former external behavior remains explicit with no default shortcut.
- Pane/view behavior for plugin disable, retained-generation re-enable, irreversible Manage Plugins Refresh, blocked Retry, and application-shutdown supersession without force-closing a user's live shell during normal refresh, including safe/pending-unsafe paste continuity for existing frozen-config views and separate OSC 52 revocation.
- Plugin-owned `maxTabsPerPane` Open admission with a hidden uncommitted host record, exact quota rollback, post-Open host-commit rollback, and no host-side setting/count enforcement.
- One plugin-owned `TerminalInteractionBroker` per committed view, with a single content-free Close/unsafe-paste/OSC 52 presentation slot, exact priority/deadline/focus/UIA/lifecycle rules, plus the host's stable-ID background-close select/reveal/focus/revalidate handoff. No close confirmation or security prompt may be invisible or resurrect after selection/activation loss.

## Implementation and build surfaces

- Consume and complete the plan-3 `Plugins/Terminal/TerminalInteractionBroker.h/.cpp` core; do not create a second pane/host arbiter. Wire it through `TerminalControl.h/.cpp`, `TerminalAccessibility.h/.cpp`, plugin resources in `Terminal.rc`/`resource.h`, and the existing close, paste-authorization, OSC-authorization, posted-payload, Restart, and teardown gates. Add every source/resource to `Terminal.vcxproj` and `.filters`; all presentation identity and sensitive-domain routing remains inside `Terminal.dll`.
- Implement only the generic stable-ID close-glyph handoff in the host's existing `FolderWindow` pane/tab selection, auxiliary-child layout, and focus paths. `DxUi::TabControl` keeps its generic close-button callback contract; the host never owns broker state, domain tokens, paste/clipboard bytes, prompt timers, or authorization decisions and receives no new public terminal ABI.
- Extend `Tests/TerminalTests` and the PaneProfiles native slice with `terminal_interaction_broker_` cases, extend pane/Commands tests with the background close-glyph handoff, and retain focused `DxUiTests` for generic close/reorder geometry. Source-contract tests reject broker/domain implementation in the EXE/Common, multiple broker slots per view, host token/payload storage, raw posted payload pointers, or a hidden/background prompt path.
- Implement both Windows launch protocols in `Plugins/Terminal/TerminalShellIntegration.h/.cpp`, with launch-plan construction in `TerminalProfileCatalog.h/.cpp`, logical-root verification in `TerminalWorkingDirectory.h/.cpp`, and `Starting`/input/profile-memory publication gates in `ConPtySession.h/.cpp` and `TerminalService.h/.cpp`. Compile the exact launch-only `.cmd` script/manifest plus the exact versioned PowerShell startup script/manifest and immutable fail-closed wrapper template into `Terminal.dll` as identity-bound RCDATA through `Terminal.rc`/`resource.h`; transactionally extract and reverify only the cmd and PowerShell launch-script resources through the plugin's owner/access/reparse-safe versioned directory. Invoke cmd only through the exact two-pass-encoded `/S /K call` token and PowerShell only through the exact generated `-Command` wrapper. Add every source/resource/build input to `Terminal.vcxproj` and `.filters`; no host/EXE/Common code, loose package/runtime script, pane-folder script, helper EXE, or public Viewer/Terminal ABI implements either protocol.
- Create/freeze the one canonical fake/production wire-schema source at
  `Plugins/Terminal/ShellAdapters/TerminalSemanticProtocol.v1.json`, deterministic
  generator `Tools/Generate-TerminalSemanticProtocol.ps1`, its checked-in outputs
  `Plugins/Terminal/Generated/TerminalSemanticProtocolV1.generated.h` and
  `Plugins/Terminal/Generated/TerminalSemanticProtocolV1.generated.ps1.txt`,
  closed declarative
  corpus `Tests/TerminalTests/Fixtures/TerminalSemanticProtocolCorpus.v1.json`,
  native test `Tests/TerminalTests/TerminalSemanticProtocolTests.cpp`, and
  generator/identity test `Tools/Test-TerminalSemanticProtocol.ps1`. In Plan 4,
  generated C++/PowerShell constants feed only the injected test seam and the
  generator's `-Check` mode fails on stale output; do not add
  the schema/constants as production RCDATA or package a loose copy. Plan 5 must
  consume those exact unchanged paths and add the identity-bound resource/project
  wiring before production Ready; hand-maintained duplicate constants are
  forbidden.
- Extend `Tests/TerminalTests`, PaneProfiles native evidence, and source/package
  contracts with `terminal_cmd_launch_bootstrap_`,
  `terminal_powershell_startup_bootstrap_`, and their exact resource/transport
  cases. The fake child exposes exact argv/environment/pipe frames,
  deterministic clock/policy/process/profile outcomes, and is the only Plan-4
  participant allowed to report `SemanticReady` or fill either
  `StartupSemanticJoin` slot. Real Windows PowerShell 5.1 and every shipped pwsh
  prove profile-before-script order, post-profile rooting, policy enforcement
  without bypass, deterministic `Prepared(SemanticUnavailable)`, no hook or
  semantic-pipe connection, Running-CAS-before-prompt, and CommitAck-before-
  input ordering. Allowlisted adapter installation, real Ready/handshake, and
  real-shell semantic closeout belong to plan 5.

## Pane model and icons

- Use stable IDs; never vector indexes in callbacks/posts. Folder is permanent first; Preview is singleton second when present; Terminal tabs follow and only Terminals reorder in v1.
- Tab view owns `wil::com_ptr<ITerminal>`, stable `TerminalInstanceId`, and a non-owning ownership-marker-validated plugin child HWND; it never owns/destroys the HWND, service, session, or workers. Explicit `UserTab` close removes only after generation-matching `TerminalCloseApproved`; cancel changes nothing. Forced `WindowClosing`/`ApplicationShutdown` paths emit no approval and follow the noninteractive sequences below. All interactive, forced, plugin, chooser, and natural-exit paths enter the plugin's one per-view retirement-intent arbiter and the host's one `{instanceId,hostViewGeneration}` removal claim. Every path returns without joining; plugin retirement owns quiet.
- Rename/generalize Preview auxiliary host to display one selected auxiliary child; hidden terminals keep parsing but do not paint/blink.
- Extend `DxUi::TabControl` generically with optional glyph/fallback, cached geometry, LTR/RTL logical edges, DPI/theme/high-contrast paint, and decorative UIA. Preserve title-only caller geometry/API.
- Reuse `FluentIcons::kPreview` U+E8FF and `kCommandPrompt` U+E756; Folder text-only. No HICON/bitmap/IconCache path.
- Keep `DxUi::TabControl` close semantics generic. A close glyph on a background Terminal captures the stable tab ID and host-view generation, programmatically selects that exact tab, reveals/layouts its plugin child, and gives the child ordinary terminal focus because this is a user action. After every reentrant selection/show/layout/focus callback, re-resolve the same ID/generation and require it still selected, interactive-visible, foreground, and ownership-marker-valid before `RequestClose(UserTab)`. A failure/removal/selection race before that request emits no close and leaves the race winner untouched. Once selection succeeds there is no rollback to the former tab: `ERROR_BUSY`, `ERROR_RETRY`, or user Cancel leaves the clicked terminal selected. Direct/approved close uses MRU/Folder fallback only after the shared host removal claim wins. Keyboard/context Close already targets the selected visible Terminal and performs the same final revalidation.
- Callbacks are destructively reentrant. Capture stable ID/lifetime token, invoke copied callback, then re-resolve `ITerminal`/control/root/HWND/model before any access. Never clear the callback from inside an `ITerminalCallback`; defer teardown to the host close path.
- `TerminalOpenSiblingRequested(instanceId, requestGeneration, selectionMode,
  cookie)` is the sole plugin-to-host route for `New Terminal With...`. v1 emits
  only nonzero strictly increasing generations with `ForceChooser`. The FIFO
  callback never coalesces; the host ignores stale/duplicate/wrong-instance/
  wrong-mode/detached/retired requests, revalidates the stored source and
  physical-host keys, and dispatches the separate sibling without mutating the
  originating object or substituting terminal cwd/current pane.
- `ITerminalCallback` v1 has exactly six slots in this order:
  `TerminalStateChanged`, `TerminalPendingOpenResolved`,
  `TerminalOpenSiblingRequested`, `TerminalCloseApproved`,
  `TerminalViewRemovalRequested`, `TerminalQuiet`.
  `TerminalPendingOpenResolved(const TerminalInstanceId*, uint64_t
  pendingOpenGeneration, TerminalPendingOpenResult, void*) noexcept` is the
  noncoalescing slot-2 result for the hidden transaction below; its closed result
  enum is `ReadyToCommit=1`, `ProfileChoiceRequired=2`,
  `ProfileNotInsertionCapable=3`, `IntegrationDisabled=4`, `LaunchFailed=5`,
  `PromptUnavailable=6`, and `TimedOut=7`. The fifth slot is
  `TerminalViewRemovalRequested(const TerminalInstanceId*, uint64_t
  viewRemovalGeneration, TerminalViewRemovalReason, void*) noexcept`, with
  the closed enum `ChooserCancelled=1`, `AutomaticCleanExit=2`,
  `AutomaticAnyExit=3`, `PluginCloseAction=4`. A nonzero strictly increasing
  removal generation is allocated by the single retirement-intent arbiter.
  Chooser cancellation is legal only after the required view-slot reservation
  and before any process-lifetime session slot/session/process; accepted deferred
  removal releases that view slot. Automatic-clean requires exit
  code zero, complete-final-snapshot natural-exit quiet, and `closeOnExit=clean`;
  automatic-any requires complete-final-snapshot natural-exit quiet and
  `closeOnExit=always`; plugin Close is legal only in provisional diagnostic/
  Failed/Exited after its applicable no-session/retirement admission and never
  substitutes for `RequestClose(UserTab)` in Starting/Running. Incomplete final
  snapshot, quarantine, `closeOnExit=never`, and interactive-close cancellation
  emit no removal request.
- The plugin arbiter states are exactly `Open`, `PendingUserConfirm`,
  `UserApproved`, `ForcedWindow`, `ForcedApplication`, `ChooserCancelled`,
  `AutomaticCleanExit`, `AutomaticAnyExit`, and `PluginCloseAction`, with one
  nonzero monotonic `retirementIntentGeneration`. A confirming UserTab request
  owns one exact `{instanceId,viewGeneration,retirementIntentGeneration}` token;
  allow alone may emit its one CloseApproved, cancel alone returns to Open, and
  duplicate pending requests return `ERROR_BUSY`. Any legal forced/chooser/
  plugin/automatic takeover of Open or Pending allocates a newer generation,
  dismisses the prompt, and suppresses every late reply/CloseApproved path.
  Approval-first suppresses a later removal callback; cancellation relinquishes
  only the user intent and cannot block a later independently qualifying owner.
- The generic host's atomic removal claim is shared by CloseApproved, every
  removal callback, and forced window/application teardown. The sole winner
  performs hide/removal, MRU/Folder fallback, pending-insertion cancellation,
  callback null/drain, idempotent Close, public-object release, and record
  invalidation; all losers join/no-op. A removal callback performs no host mutation reentrantly. It posts only a
  stable instance ID plus host-view generation. After callback return, the
  deferred task revalidates the still-attached view, applies the generic
  MRU-remaining-tab/else-Folder removal, cancels any pending insertion, drains
  with `SetCallback(nullptr,nullptr)`, calls idempotent `Close`, and releases the
  object. It never calls `RequestClose`, parses Terminal settings, or infers
  removal from lifecycle state. Duplicate/stale/unknown/wrong-instance
  requests, host-view-generation reuse, and user/window/application-shutdown
  races are absorbed by the shared claim and generation check. Chooser/plugin-Close removal follows the last
  retained state event for that action; automatic removal order is exactly
  `TerminalStateChanged(Exited) -> TerminalQuiet ->
  TerminalViewRemovalRequested`.

### Per-view interaction broker, unsafe paste, and OSC 52

- `Terminal.dll` owns exactly one `TerminalInteractionBroker` per committed
  view. It is the sole admission, presentation, focus, UIA, reply, and dismissal
  path for close confirmation, unsafe-paste confirmation, and OSC 52 `ask`.
  Its state is exactly `None|CloseConfirm|UnsafePasteConfirm|Osc52Ask` with one
  nonzero monotonic `interactionGeneration`. A visible token binds
  `{serviceGeneration,sessionId,sessionIncarnation,viewGeneration,
  interactionGeneration,kind,domainToken}`. Every reply revalidates the entire
  identity before the owning close/paste/OSC gate consumes its opaque domain
  token. Wrap, stale/duplicate/wrong-kind/wrong-view, missing posted payload, or
  post failure invalidates and securely scrubs the incoming domain request; no
  generation or superseded interaction is reused.
- The broker retains only one content-free presentation record. Paste and OSC
  authorization gates retain their sensitive buffers and own the final commits.
  The pane host owns only the plugin child container and tab visibility/focus;
  it never reads the clipboard, receives preview/clipboard bytes, stores or
  forwards a domain token, classifies content, tracks broker time, or admits
  input. Every broker post uses `PostMessagePayload`/`TakeMessagePayload`, with
  the child registry initialized at create and drained at teardown.
- `interactive-visible` is plugin-observed on the UI thread: the direct child
  and every ancestor are visible; the exact instance/view ownership marker is
  current; the top-level Folder window is non-minimized, foreground, and not
  closing; and this child is the host-revealed auxiliary child. Show/window-pos,
  focus/activation, and ownership changes advance a presentation generation.
  Before presentation or reply, recheck visibility, foreground root, ownership,
  and that generation. The plugin never queries host model internals.

The one closed priority table is exact; a canceled lower-priority interaction
is never restored:

| Current broker state | Incoming Close | Incoming unsafe paste | Incoming OSC 52 `ask` |
|---|---|---|---|
| `None` | Present only when the close intent/token and interactive-visible state still match | Present if eligible; otherwise cancel/scrub | Present only if interactive-visible and service/session limits reserve; otherwise deny/scrub |
| `CloseConfirm` | Exact duplicate returns `ERROR_BUSY`; other/stale Close is ignored | Reject/scrub new Paste; retain Close | Deny/scrub new OSC; retain Close |
| `UnsafePasteConfirm` | Cancel/scrub Paste, then present Close if still valid | Newer same-view Paste cancels/scrubs old before replacement; replacement failure leaves `None` | Deny/scrub new OSC; retain Paste |
| `Osc52Ask` | Deny/scrub OSC, then present Close if still valid | Deny/scrub OSC, then present Paste if still eligible | Newer same-session OSC denies/scrubs old before replacement under the existing service cap |

- A background, hidden, minimized, nonforeground, inactive, or ownership-stale
  OSC 52 `ask` is denied and completely scrubbed before the per-session or
  eight-service outstanding counter increments. It never selects/reveals a tab,
  moves focus, retains a token, or queues work for later. Only a coalesced
  content-free blocked-in-background status may appear on later activation.
  An admitted visible ask has a fixed monotonic-QPC 30-second strict-before
  deadline; equality denies/scrubs. Selection, visibility, foreground,
  ownership, policy, service, view, session, or incarnation loss denies/scrubs
  immediately and returning never resurrects it.
- Unsafe Paste is user initiated and presentable only while the same view is
  interactive-visible and owns terminal-document input. Its fixed monotonic-QPC
  60-second deadline cancels/scrubs at equality. Selection away, hide,
  deactivation, input-owner change, policy/input-mode revocation, supersession,
  or lifecycle loss cancels and securely scrubs immediately; returning never
  restores the prompt or payload. Reorder alone may preserve it only when the
  same stable view remains selected and interactive-visible.
- Close confirmation is presentable only after the stable-ID host handoff above
  and has no elapsed timeout while interactive-visible. Losing selection,
  visibility, foreground activation, ownership, exact view generation, or the
  broker slot is the retirement arbiter's exact Cancel transition: invalidate
  the token, return the still-matching intent to `Open`, emit no callback, and
  restore no Paste/OSC request. If confirmation is required while the child is
  not interactive-visible, return `ERROR_RETRY` without allocating/changing an
  intent or broker generation; an invisible close prompt is forbidden.
- Opening History or the terminal context menu requires broker `None`. Close
  dismisses either non-sensitive surface before broker admission. OSC 52 `ask`
  while either is open is denied/scrubbed. A Paste routed with
  `FocusOwnerKind=PluginTextInput` acts only on that owned UI and cannot create
  a terminal paste prompt. Diagnostic/profile chooser base states cannot
  receive Running-only Paste or OSC requests.
- Close and Paste confirmations record the exact prior owned child focus,
  disable terminal input, expose localized actions, focus Cancel initially,
  confine Tab/Shift+Tab to the buttons, let Enter/Space invoke only the focused
  button, and map Escape to Cancel. After Cancel or successful paste admission,
  restore terminal-document focus only when the exact broker/view/presentation
  generations remain selected, interactive-visible, and input-capable. Close
  approval, selection loss, teardown, or newer interaction restores nothing.
  OSC 52 is a non-focus-stealing security bar: terminal focus/input remains,
  one polite UIA notification is raised, and only mouse/UIA Invoke or the
  bar-scoped displayed `Alt+A` Allow and `Alt+D` Deny accelerators can decide the
  exact visible token. Hidden/canceled/expired surfaces expose no name/button/
  live-region/stale Invoke target; later provider calls return
  `UIA_E_ELEMENTNOTAVAILABLE`.
- Window/application shutdown, callback detach, forced/plugin/automatic
  retirement, Restart, and view removal first close broker admission, advance
  its generation, dismiss the child/UIA surface, route every live domain token
  through Cancel/Deny plus secure scrub, and drain broker posts before child
  teardown. The host never infers scrub completion from tab state.
- Generic disable/re-enable and accepted normal refresh are not paste-policy
  transitions and do not by themselves revoke safe Paste or a still-eligible
  pending unsafe Paste in an existing live frozen-configuration view. That
  prompt nevertheless remains subject to the 60-second, selection, activation,
  visibility, input-owner, supersession, and lifecycle cancellations above.
  OSC 52 continuity is different: disable/re-enable increments policy generation
  and refresh revokes by service generation, so every ask/allow is denied/
  scrubbed before publication and disabled requests remain effective deny.

### Close reasons and process shutdown

- `UserTab` is the only interactive reason. Call `RequestClose(UserTab)` with the callback registered only after the selected-visible stable-ID revalidation above. A background close glyph must first select/reveal/focus that exact view; once that selection succeeds there is no rollback on `ERROR_BUSY`, `ERROR_RETRY`, or Cancel. If confirmation is needed but the selected child is no longer interactive-visible, the plugin returns `ERROR_RETRY` with no retirement intent/broker allocation. Cancellation leaves the object/session/view and generations unchanged, while acceptance emits exactly one generation-matching `TerminalCloseApproved`. Only that callback admits explicit-tab removal, after which the host hides the view, calls `SetCallback(nullptr,nullptr)` outside callback dispatch, calls idempotent `Close`, releases the object, invalidates the host record, and returns without waiting.
- Natural exit does not implicitly resolve a pending confirmation before complete quiet. At quiet, a complete snapshot qualifying for `always`, or `clean` with exit code zero, takes over and emits only ViewRemovalRequested. A successful but nonqualifying outcome (`never`, nonzero `clean`, or incomplete snapshot) converts the still-exact pending token to UserApproved and emits only CloseApproved; cancel-first retains the exited view. Quiet failure leaves the prompt unresolved. Approval-first suppresses later automatic removal, and a forced window/application path always owns or joins teardown regardless of a prior cancel.
- `WindowClosing` is noninteractive and object-local. For each object owned by that FolderWindow, call `RequestClose(WindowClosing)` while its callback is live; the forced intent first invalidates/dismisses a pending prompt and emits no `CloseApproved`. Then enter the shared host removal claim, whose winner hides, drains, closes, releases, and invalidates the view without awaiting approval. It does not close process-wide service admission and must not affect terminals in another FolderWindow.
- Main `WM_CLOSE` calls generic `TerminalPluginManager::BeginProcessShutdown(validatedMainTarget, nonzeroApplicationLifetimeGeneration)` before any FolderWindow, main message target, or pump is destroyed, then returns without `DestroyWindow` or a wait. The first valid begin captures manager `shutdownT0`; closes terminal/configuration creation, Preferences-deep-link, status/Retry, and page-presentation admission; increments admission generation; and snapshots both the module-scoped `IPluginConfigurationOperations`/coordinator generation when present and all public `ITerminal` objects in stable instance-ID order.
- The terminal-public-object barrier calls `RequestClose(ApplicationShutdown)` on each object while callbacks remain live, then enters the same host removal claim; only its winner hides, callback-drains, closes, releases, and invalidates that record, while a callback/WindowClosing winner is joined. Concurrently, the configuration-public-object barrier marks the retained administration/coordinator generation `Retiring`, suppresses page presentation but keeps its callback/dispatcher/private settlement lane alive, cancels/takes every pre-commit operation, and settles every transaction-bearing state. Before a Settings CAS it aborts the whole participant group and gives each successful Prepare/Arm exactly `Finalize(Failed)`; after CAS start it awaits the actual outcome and gives each successful participant exactly `Finalize(Committed|Failed)`. Commit-won work, typed results, required compensation save, and matching `HostCompensationSave` acknowledgement finish; query results are taken/discarded. Only after no operation, preparation, CAS, compensation, result, or continuation remains may it null-clear/drain the configuration callback, call nonblocking `Close`, and release the administration object. An `ERROR_BUSY` Close after callback clear is a host invariant breach and takes the fatal path.
- Only after both public-object barriers complete does the manager begin the existing module shutdown export through shared `PluginModuleLifecycle`; the zero-terminal/object-free-retiring case begins the service epoch through that export. The generic helper samples the shutdown, `CanUnloadNow`, and process-shutdown-only retention exports and returns only `Busy|Unloaded|RetainedUntilProcessExit`; a false unload vote is `Busy` without an explicit retention vote. It retains the module owner and samples on generation-checked dispatcher turns at intervals of at most 50 ms while the main target/pump remains alive. Unload or exact retention posts exactly one generation-bound continuation; busy reschedules without blocking. The manager outer two-/five-second bounds start at `shutdownT0`; the first accepted terminal request owns the no-later service/session `t0`, and zero-terminal administration remains bounded before export. There is no EXE terminal-service accessor, direct terminal-implementation symbol, or UI-thread join.
- `RedSalamanderPluginRetainModuleUntilProcessExit` is allowed only for the passive external-UIA tail after the export and both public barriers: all providers are disconnected, captured-before-disconnect calls have returned, and callbacks, workers, engines/renderers, process/ConPTY handles, posts, continuations, and every host/service-touching owner are quiet. The only remaining gates may be externally held disconnected provider/range COM objects with self-contained immutable snapshots/passive ledger references; all post-disconnect calls except `IUnknown` return `UIA_E_ELEMENTNOTAVAILABLE` with null/zero output and no host/service access. The helper rechecks `CanUnloadNow` once after a false retention vote to absorb a racing final `Release`; otherwise it retains the module owner through OS teardown and posts the same one continuation. Normal refresh/disable and any active work cannot use this exception; every other five-second busy result is fatal.
- Repeated same-generation shutdown, duplicate `WM_CLOSE`, repeated forced requests, and repeated export begins are idempotent; stale/zero generations, stale targets, stale view/callback generations, and late polls are ignored and cannot remove a live replacement view or post another continuation. Application shutdown supersedes a normal refresh and reuses already-complete barriers/export work without posting a second continuation.

### Normal plugin refresh and retained terminal views

- Generic plugin disable closes new open/Edit/path-insertion/Restart admission but does not begin retirement. Every existing Terminal tab, plugin child, callback, session, and administration generation remains attached and usable; no tab is closed, hidden, disconnected, reordered, or replaced. Safe Paste and a pending unsafe Paste that remains selected, interactive-visible, input-owning, unexpired, and otherwise broker-eligible remain usable without a paste-generation change; disable itself does not revoke them, but no broker cancellation is waived. The plugin-owned OSC 52 gate revokes outstanding ask/allow payloads before disabled publication and treats later requests as deny. Re-enable while retirement is `Idle` reuses that same module/service generation after current persisted configuration reconciliation and never replays an OSC 52 request.
- Only a confirmed Manage Plugins Refresh may call generation-bound `TerminalPluginManager::BeginRefresh`. Cancel before the call changes nothing; `S_OK` is irreversible for the captured old generation. The pane host immediately fences new Terminal creation, sibling-open, directory Edit, insertion, Restart, plugin deep links/status/actions, and any generation-stale pending chooser/insertion/open post with localized refresh status and `HRESULT_FROM_WIN32(ERROR_SHUTDOWN_IN_PROGRESS)` where an HRESULT is returned.
- Refresh does not call `RequestClose`, hide a child, clear an `ITerminalCallback`, call `Close`, release a tab object, or synthesize a view-removal request. Already-open terminals retain keyboard, rendering, accessibility, state callbacks, explicit user close, safe Paste, a still-selected/interactive-visible/unexpired broker-eligible unsafe-Paste confirmation, and natural `closeOnExit` behavior under their frozen effective configuration. Refresh does not change paste-policy generation or itself revoke/reject those user paste paths; selection/activation/visibility/input-owner/expiry/lifecycle cancellation remains independent and final. Its plugin-owned OSC 52 fence is distinct and must revoke/scrub every outstanding ask/allow before refresh publication. The terminal-public-object barrier completes only through those ordinary removals; `RefreshWaitingForTerminalObjects` has no timeout, and a shell may outlive both independent five-second refresh clocks.
- When the last old-generation tab is naturally removed, the normal callback-drain/Close/release path reports the generation-checked barrier completion. Configuration settlement may complete before or after it. Only after both barriers may the generic lifecycle invoke the shutdown export and quiet poll; normal refresh never displays process-fatal UI or uses process-exit module retention. `RefreshBlocked(ConfigurationSettlementTimeout|ModuleQuietTimeout)` keeps creation closed and exposes only the content-free Restart-required status plus host-owned `plugin.refresh.retry` in Manage Plugins.
- Enable/disable churn while refresh is active updates only the latest desired availability for the eventual one replacement; it never restores old pane admission or creates a queued tab/module generation. Successful unload resets the old owning module handle before one verified replacement is mapped and reconciled. Replacement failure leaves existing old views impossible because their natural-release barrier already completed, creates no substitute tab, and exposes the manager's blocked/reinstall guidance.
- Main application shutdown supersedes every refresh state. It force-closes any still-live old-generation tabs only through `RequestClose(ApplicationShutdown)` plus the existing callback-drain/Close/release sequence, then applies the stricter process deadline/fatal/passive-UIA retention policy; normal refresh itself can use none of those outcomes.
- Implement pane-side fencing and natural barrier completion in `RedSalamander/TerminalHost/TerminalPaneHost.h/.cpp`, `TerminalPluginManager.h/.cpp`, the FolderWindow command/view integration files, `ManagePluginsDialog.h/.cpp`, localized resources, and Commands/native lifecycle cases. The pane host owns no module export lookup or Terminal-private runtime state.

## Exact open-terminal command behavior

1. Resolve complete `OriginalSourceKey` and `PhysicalHostKey` independently. Folder or Preview focus selects that owning pane's logical Folder source/current committed location and the opposite physical pane as host; Preview content is never a path.
2. Every Terminal-kind tab—including chooser, starting, diagnostic, running, exited, or failed—selects its stored complete source key/latest committed source location and immutable physical-host key, creating a sibling there. After `SwapPanes`, never recompute a terminal-originated host from the source key's current side.
3. Resolve a known focused pane descendant first. Focus temporarily outside pane content uses the stable active pane plus its selected tab captured at dispatch; never default to Left or a hidden unselected terminal. Resolve the still-live content parent for a hidden/zoomed host, but do not reveal/unzoom, change selection/MRU/chrome, or insert a tab before Open succeeds.
4. Create only one hidden uncommitted host record containing both complete keys,
   the real pane `{sourceGeneration,sourceLocation}` follow baseline, the
   independently named `launchLocation`, admitted
   `TerminalProfileSelectionMode`, stable instance/view generations, and the
   candidate plugin object. The record is not a tab and does not count. Register
   its callback and call `Open` against the hidden parent; Open delivers no
   callback synchronously before return, and the host never reads or
   pre-enforces `maxTabsPerPane`.
5. After all synchronous validation, `Terminal.dll` atomically reserves the
   per-complete-`PhysicalHostKey` view slot before child/async work. Exact
   `HRESULT_FROM_WIN32(ERROR_QUOTA_EXCEEDED)` means null/drain the callback,
   call `Close`, release the object/record, retain prior selection/zoom, and show
   localized `Terminal tab limit reached for this pane` plus `Open Terminal
   settings`; there is no tab/child/process. Any other synchronous failure uses
   its typed path with the same no-half-commit rollback. On Open success, obtain
   and validate the child, then atomically commit the host view: reveal/unzoom,
   insert/select the provisional Terminal tab, attach the child, and begin
   chooser/starting/diagnostic presentation. A failure in that host commit
   null-drains, Closes, releases, restores prior host UI, and releases the slot.
6. Ordinary opens pass value-identical source/launch locations and freeze
   plugin-internal `launchTracksSourceUntilStart=true`; Directory Edit passes
   the pane baseline plus its revalidated child launch and freezes that flag
   false. If resolution requires a chooser, focus it only after host commit.
   Cancel emits only generation-bound
   `TerminalViewRemovalRequested(ChooserCancelled)`; deferred generic removal
   removes the committed provisional tab and no process starts. Resolve/
   revalidate location/profile asynchronously. Launch/location failure
   keeps a closable diagnostic with localized Retry, Choose profile, Close.
   Retry preserves the admitted mode and the applicable admitted launch
   location against a refreshed catalog/latest source generation; Directory
   Edit therefore keeps its child launch fixed while its source follow record
   advances. Choose profile switches the still-unlaunched object to
   `ForceChooser`; neither calls `Open` again. For an ordinary open, a fully
   revalidated newer source update before process creation transactionally
   replaces both pending source and launch locations. Process creation freezes
   launch. Successful launch binds stable session ID, swaps in control, and
   focuses terminal.
7. `cmd/pane/openCommandShell` uses `ResolveDefault` for both `Ctrl+Alt+T` and
   `Alt+7`. Stable no-default `cmd/pane/openCommandShellWithProfile` uses
   `ForceChooser`; it is the command-menu `Open terminal with...` route and the
   target of a validated plugin `TerminalOpenSiblingRequested` context action.
   Repeated invocation creates an independent session. No reuse by path/profile.

Preview retains its no-focus-steal contract. `Alt+6` never closes/moves Terminal. `SwapPanes` swaps logical folder content and its source key; auxiliary tabs/HWNDs remain in physical hosts and followers keep the same key. Zoom/layout never retargets by side.

`TerminalOpenContext` begins with `sizeBytes`. Consumers reject a record shorter
than the current required prefix, accept larger records, and ignore unknown
tails; no concrete compiler size/offset table is frozen. Its canonical semantic
field order is parent window, instance ID, original source key,
physical host key, source generation, source location, launch location,
`TerminalProfileSelectionMode selectionMode`, `TerminalOpenPurpose
openPurpose`, requested profile ID, then `reserved[4]`; the purpose enum is
exactly `InteractiveTab=1|PendingPathInsertion=2`. Source and launch
must each satisfy the complete location cross-fields and resolve to the same
profile family/exact WSL distro. Loaded-DLL tests reject undersized values and
accept current/oversized values.

`InteractiveTab` requires the second `Open` argument to be null and permits
`ResolveDefault|ForceChooser|ExactProfile` with their normal requested-ID
cross-fields. `PendingPathInsertion` requires a nonnull valid copied insertion,
`ResolveDefault`, Windows-compatible source/launch/initiating locations, and a
null requested profile ID. Every WSL, `ForceChooser`, `ExactProfile`, null/non-
null argument mismatch, unknown purpose, or unsupported-family combination is
`E_INVALIDARG` before child, quota, status, or callback effects.

The former external launch implementation moves to `cmd/pane/openExternalCommandShell` and has no default shortcut. It retains its tested Windows Terminal/cmd fallback but is never an automatic fallback for embedded-plugin/profile/location failure.

## Folder Edit and path insertion

- Primary Edit/F4 on an item the provider reports as a directory invokes the open-terminal flow in the opposite physical host. Its Open context keeps the owning pane's real current committed `{sourceGeneration,sourceLocation}` as the follow baseline and puts only the independently revalidated focused child in `launchLocation`; the host never invents a child navigation generation. This route freezes `launchTracksSourceUntilStart=false`, so navigation and Retry cannot replace the child launch. When follow is enabled, the initial baseline does not immediately undo that explicit launch: only a later accepted update with a strictly greater real generation becomes automatic pending follow. Explicit follow off/on may internally requeue the plugin's stored same-generation pane baseline without an ABI update and without lowering/replacing a newer generation. This directory decision happens before editor action resolution and never uses launch failure as a directory probe. File Edit remains editor behavior; Alternate Edit and Edit With never route to Terminal.
- Preserve legacy `cmd/pane/bringFilenameToCommandLine` for `Ctrl+Enter`, add `cmd/pane/bringFullPathToTerminal` for `Ctrl+Shift+Enter`, and migrate `cmd/pane/bringCurrentDirToCommandLine` for Ctrl+Space/its current Shift alias. The distinct full-path ID is mandatory because both Enter chords currently dispatch one ID.
- Each command snapshots its command-originating live Folder source as one atomic `{initiatingSource, initiatingSourceGeneration, initiatingSourceLocation}`. This is the logical Folder source that owns the focused item for the two item commands or owns the current directory for Ctrl+Space; it is independent of the selected target terminal's immutable `originalSource` and follow record. `SwapPanes` changes which physical side contains that Folder source but not its stable source key, and a selected compatible terminal remains a valid destination when its original source key differs. The host never rewrites the initiating triple from the target's followed source, trusted cwd, launch location, or physical side.
- When the opposite selected content is Folder or Preview and the command creates
  a hidden `PendingPathInsertion` object, its complete Open tuple comes from the
  initiating focused Folder, not the destination content: `originalSource` is
  `initiatingSource`; `sourceGeneration`/`sourceLocation` are the captured
  initiating generation/location; `launchLocation` is that same initiating
  location; and that identity alone owns profile-memory resolution, initial cwd,
  follow subscription, and Restart fallback. The opposite selected Folder or
  Preview contributes only its `PhysicalHostKey`, hidden parent, and eligibility
  to host a new auxiliary object; its own location/path/Preview item never enters
  the Open tuple. Immediately before Open and again before fake Commit, revalidate
  the initiating tuple and opposite host/selection independently. Tests use
  deliberately different source/destination paths, selected Preview, navigation
  before each revalidation, and `SwapPanes` before capture/between capture and
  Open/between Ready and Commit to prove no path or identity crosses roles.
- For `Ctrl+Enter` and `Ctrl+Shift+Enter`, the item is exactly the focused item and the selection set is ignored. Destination is the selected content in the opposite physical pane; never search a hidden Terminal. The host calls `InsertPath` only for a selected Terminal whose copied ViewState advertises `InsertCapable`, after its final request/source/item revalidation. Busy, incompatible, provisional, diagnostic, Starting, Exited, Failed, closing, or capability-unavailable states reject with zero bytes and do not create another Terminal. A selected Folder or Preview may start one `PendingPathInsertion` hidden open through the closed transaction below. In Plan 4, a fake `ITerminal` proves both routes and every target/source race; the production plugin advertises `InsertCapable=false`, so direct insertion returns the localized capability-unavailable status and pending insertion resolves `IntegrationDisabled` before process creation with no visible tab.
- `TerminalPathInsertion` begins with `sizeBytes`; its semantic order is initiating source, initiating generation, initiating location, item generation, mode, reserved0, item location, parent location, display leaf, and reserved words. Consumers validate the current required prefix and ignore unknown tails. Immediately before a direct call, the host atomically revalidates the initiating triple against the command-originating Folder source and, for either item mode, revalidates the focused item plus `itemGeneration`; generation-bound pending delivery repeats both checks immediately before the one eventual call.
- Ctrl+Enter marshals `ContextualLeafOrFull`: nonzero item generation, supported item/parent locations, and a nonempty presentation-only display leaf. Ctrl+Shift+Enter marshals `AlwaysFull`: nonzero item generation and supported item location, with Unsupported/empty parent and null/zero display leaf. Ctrl+Space marshals `CurrentDirectoryFull`: zero item generation, Unsupported/empty item and parent, and null/zero display leaf. Every mode includes a supported initiating source location. Plan-4 ABI/fake tests copy the request, keep the initiating triple independent of the target terminal's original/follow source, and prove invalid cross-fields return `E_INVALIDARG` with zero bytes. Authenticated prompt state, leaf/full cwd comparison, Windows/UNC namespace translation, literal quoting, busy/TUI/password/partial-input safety, atomic enqueue, and no-Enter product semantics are plan-5 adapter work and cannot be claimed by Plan-4 production evidence.
- A busy, incompatible, or capability-unavailable existing terminal rejects visibly, never defers to a later prompt, and never causes an implicit new tab. The Plan-4 fake covers unsupported/cross Windows-WSL/distro/provider mapping with zero bytes and `Open Terminal settings` where configuration can help; production rejects earlier at its truthful capability gate.
- A path-insertion command targeting an exact-distribution WSL terminal may insert only the quoted full native Linux path; contextual mode degrades to full and never claims prompt trust. Different/unknown distributions and cross-family paths reject with zero PTY bytes. A selected Folder/Preview source may open a matching WSL terminal for that full-path insertion but must never inject bootstrap or follow commands.

The Plan-4 host/ABI seam nevertheless proves the complete hidden-open
transaction with a deterministic fake `ITerminal`; production takes only the
closed `IntegrationDisabled` path:

- A `PendingPathInsertion` Open synchronously validates/copies the request,
  reserves the normal plugin-owned view slot plus the sole hidden-pending slot
  for its complete `PhysicalHostKey`, creates a nonzero generation, and returns
  without publishing a tab, focus, UIA surface, chooser, diagnostic, or callback
  inside `Open`. A second hidden open for that host returns
  `HRESULT_FROM_WIN32(ERROR_BUSY)` before child/process. The fake deadline is
  `pendingOpenT0 + 15 seconds`, strict-before; equality is `TimedOut`.
- The fake state machine is exactly `AwaitingResolution|Ready|TerminalFailure|
  Committed|Aborted`. It emits one noncoalescing
  `TerminalPendingOpenResolved`. `ReadyToCommit` alone enters `Ready` and holds
  input; each closed failure enters `TerminalFailure` and makes Commit forever
  illegal. Close/forced teardown moves a noncommitted transaction to `Aborted`.
  Stale/duplicate/wrong-generation callbacks or decisions are no-ops and can
  never revive another instance.
- After a Ready callback returns, one UI-lane task revalidates the exact hidden
  instance/host generation, still-selected Folder/Preview source, initiating
  source/location/item generations, window/application state, and unchanged
  focus. `PrepareHiddenTerminalCommit` preallocates every tab/MRU/UIA/child-link/
  title/icon resource into one rollbackable unpublished token. Preparation
  failure calls exact-generation Abort and admits zero bytes.
- With no message pump or callback between final revalidation, successful
  preparation, plugin decision, and finalize, `ResolvePendingOpen(Commit)`
  either atomically admits the complete fake insertion descriptor and enters
  `Committed`, or admits zero and enters `Aborted` with its typed HRESULT.
  Commit `S_OK` is followed contiguously by allocation-free/no-fail
  `FinalizeHiddenTerminalCommit`, which publishes/selects/focuses exactly one
  user-owned tab and performs the deferred `PendingOpenHostFinalized` preference
  eligibility action. An impossible post-Commit invariant breach uses the
  injected fail-fast policy; no ordinary failure may occur after bytes are
  admitted. Abort/TerminalFailure cleanup null-drains, Closes, releases every
  hidden/view/session/module reservation exactly once, and preserves prior
  zoom/selection/focus/UIA.
- In the production Plan-4 plugin, `InsertCapable=false` is immutable and a
  valid hidden request resolves exactly once as `IntegrationDisabled` before
  profile/process/ConPTY creation. It emits zero bytes, writes no preference,
  and follows the same hidden cleanup; `ReadyToCommit`, Commit, and host finalize
  are fake-only until plan 5 supplies authenticated insertion capability.

- After these stable IDs migrate, remove `Create/Destroy/Show/HideCommandLine`, pseudo field/layout/state/debug seams, `LaunchCommandLine`/`cmd.exe /C`, obsolete resources, and old pseudo-input tests. Update `UI_CommandMenuKeyboard.md` as an intentional migration.

`Terminal.dll` solely owns `maxTabsPerPane` and its `1..32` effective value. A successful Open retains one slot keyed by complete `PhysicalHostKey` through chooser, provisional diagnostic, Starting, Running, Failed, Exited, natural-exit teardown timeout, and host removal admission. It releases exactly once only after callback null/drain, valid `Close` child destruction/detach, and public-object release; later internal session retirement/quarantine owns no slot. Folder/Preview, hidden uncommitted records, and content-free service diagnostics after an already-removed view do not count. Live lowering atomically publishes the new limit, evicts nothing, and blocks later reservations for only a key still at/above it; raising resurrects nothing. The service erases zero-count keys and serializes simultaneous/reentrant opens. This view limit remains separate from process-session admission. `confirmCloseRunningProcess` prompts only for explicit close of Starting/Running; Cancel changes no tab/session/view generation. Exited/failed/diagnostic close without confirmation.

The service independently caps process-lifetime sessions at 32 slots retained until complete quiet. Launch/restart atomically reserves before process/engine creation; exhaustion leaves the provisional/exited view diagnostic and starts nothing. Any quarantine closes all new launch/restart admission until every quarantine recovers; existing terminals and file-manager work remain usable. Quiet retained exited views use no slot but remain charged to the fixed 512 MiB process terminal-text/snapshot ledger.

After final output and the core's complete-or-bounded-incomplete final-snapshot outcome publishes `Exited`, defer automatic removal until natural-exit teardown reaches quiet. On quiet success with `finalSnapshotComplete=true`, `closeOnExit=never` retains all; `clean` requests removal only as `TerminalViewRemovalRequested(AutomaticCleanExit)` for exit code 0; `always` requests it only as `AutomaticAnyExit`. With `finalSnapshotComplete=false`, the view retains its last-safe snapshot or empty content-free exit surface, shows exact localized `Exited — final output unavailable (resource limit)`, and permanently suppresses every automatic removal regardless of configured `closeOnExit`. Failed launch and natural-exit quarantine/teardown-timeout remain visible while their view exists. The plugin surface's Close in diagnostic/Failed/Exited state admits the applicable no-session/retirement path and then requests only `PluginCloseAction`; Starting/Running still use `RequestClose(UserTab)`. Explicitly closing an exited-but-not-yet-quiet view removes it immediately without a process prompt, suppresses later automatic removal, and cannot resurrect content/tab; any later quiet failure uses only the service's content-free application diagnostic/status. If quarantined work later reaches quiet, a removed view stays absent and its diagnostic clears; a still-visible natural-exit view enables Restart/Close, retains its snapshot, and never re-runs `closeOnExit`: a complete snapshot becomes localized `Exited — cleanup recovered`, while an incomplete snapshot returns to the exact resource-limit status. Automatic removal and selected-auxiliary close use MRU remaining tab, else Folder and exact chrome/focus restoration, without prompting. Retained exited sessions show status/Restart/Close, but Restart is disabled with a reason until old-session quiet. Restart location precedence is a requested/attached/safely launchable followed source; else the latest explicit cwd only if trusted/valid/filesystem; else original validated launch location only when no explicit cwd was ever reported; else diagnostic. A later invalid/disappeared/non-filesystem report invalidates earlier cwd and launch fallback. Never choose root/home/temp or record a command.

An `OriginalSourceKey`'s two components are opaque, application-lifetime unique, and never reused. A `PhysicalHostKey` combines that unique window ID with its stable physical Left/Right pane ID. Navigation generation is monotonic within the source key. Detach tombstones before canceling work; stale results/posts cannot attach to a later window/source or retarget a host.

## Profile catalog

- Stable profile IDs are exactly `cmd`, `windows-powershell`, `pwsh:<canonical-normalized-DOS-path>`, and `wsl:<normalized-distribution-name>`. The complete variable identity is retained: profile IDs are capped at 32,772 UTF-16 units (the five-unit `pwsh:` prefix plus the 32,767-unit canonical Windows-path maximum), while a WSL ID is at most 260. Catalog lookup/deduplication, chooser rows, `ExactProfile`, Open/callback copies, remembered choice, missing-profile display, history partition, and Restart all use the same complete ID. No layer truncates, hashes, aliases, or replaces canonical material with file identity.
- Obtain the native system directory and Windows directory with
  `GetSystemDirectoryW`/`GetWindowsDirectoryW`, and Program Files locations with
  `SHGetKnownFolderPath`; inherited `COMSPEC`, `SystemRoot`, `windir`,
  Program-Files variables, app/current directory, and PATH never define a
  system or auto-eligible profile. Each launch supplies canonical API-derived
  values for `SystemRoot`, `windir`, `ProgramFiles`, applicable
  `ProgramW6432`/`ProgramFiles(x86)`, `ComSpec`, and every permitted launch/
  integration environment key to the one core environment builder; a parent
  case variant cannot override or survive beside those canonical spellings.
  The PowerShell bootstrap target and semantic nonce/control-pipe name are never
  environment keys or initial-argv values. Initial argv carries one canonical
  fail-closed wrapper token built only from the verified startup-script path and
  bootstrap pipe/key/operation/generation/protocol/mode literals;
  the protected startup exchange sends the target in `RootRequest` and creates/
  sends a semantic nonce/control pipe only in a later `SemanticEnable` decision.
- The builder owns a fresh `GetEnvironmentStringsW` snapshot with
  `wil::unique_environstrings_ptr`, parses checked through the double NUL, and
  deduplicates ordinary nonempty keys with later-parent spelling/value winning
  under ordinal-ignore-case. It separately accepts only
  `=<ASCII drive letter>:=<nonempty absolute same-drive DOS path>`, uppercases
  the drive key, and keeps the last entry per drive. Reject every other
  leading-`=`, malformed key/path/termination, embedded NUL in a profile delta,
  or overflow. After canonical overwrites, total-order sort hidden and ordinary
  `name=value` entries by ordinal-ignore-case then ordinal case-sensitive,
  append one NUL per entry plus the final NUL, and represent empty as exactly
  two NULs.
- Every profile launch passes a mutable command line of at most 32,767 UTF-16
  units including its terminator and an explicit environment block of at most
  524,288 UTF-16 units including both terminal NULs. Exact boundaries succeed;
  +1 or a malformed/over-cap parent or integration overwrite produces a
  localized retryable diagnostic before `CreateProcessW`, with no truncation or
  process. Use creation flags exactly
  `EXTENDED_STARTUPINFO_PRESENT | CREATE_UNICODE_ENVIRONMENT`; neither cap is a
  Preference.
- `cmd` is only the non-reparse system file
  `<GetSystemDirectoryW>\cmd.exe`; Windows PowerShell is only
  `<GetSystemDirectoryW>\WindowsPowerShell\v1.0\powershell.exe`. Use the
  master's one `TerminalExecutableIdentity` validator: explicit absolute path,
  local fixed regular file, component/final reparse rejection for system
  profiles, final DOS path and volume/file identity, exact basename, PE plus
  OS-supported machine, and actual version. Missing/invalid means unavailable,
  never an environment/PATH fallback.
- Enumerate pwsh from API-derived Program Files, then HKLM/HKCU App Paths in
  native and alternate WOW64 views, then one `SearchPathW` over an explicit
  sanitized local list excluding current/empty/relative/app-local/network/
  removable entries. Program Files/HKLM are `autoEligible`; HKCU/PATH are
  visibly `explicitChoiceOnly` and cannot satisfy `auto`/family `pwsh` until an
  exact user choice reaches `Running`. Expand `REG_EXPAND_SZ` only through the
  master's closed OS-derived token map; one balanced quote pair is removable,
  arguments/unknown tokens fail. A local reparse alias may resolve only to a
  final regular local `pwsh.exe` that passes version/PE/machine validation.
- Dedupe volume/file identity then path. Representative source rank is Program
  Files, HKLM, HKCU, PATH, then ordinal-ignore-case path. Persistent ID uses
  that representative canonical DOS path; sort profiles version descending
  then path. Immediately before launch/restart, reopen and compare the current
  target to the catalog snapshot, hold a no-write/no-delete-share WIL handle
  through process creation, pass the reopened canonical final target as exact
  nonnull `CreateProcessW.lpApplicationName`, and begin the mutable command line
  with exactly one unquoted fixed basename—`cmd.exe`, `powershell.exe`, or
  `pwsh.exe`—as child `argv[0]`; never repeat the canonical path. Only the
  matching profile adapter appends reviewed arguments through the canonical
  quoting helper. A swap aborts with zero process creation and refreshes
  the catalog; only a newly valid same-representative in-place upgrade retains
  preference.
- Test malicious relative/quoted/argument/app-local/network COMSPEC, spoofed
  OS-folder/PATH variables, unknown App-Path expansions, reparse components/
  aliases/swaps, catalog-to-launch replacement, wrong basename/PE/machine,
  missing system profiles, explicit-only non-auto-selection, exact
  `lpApplicationName`, fixed basename argv0 under hostile current directory/
  PATH/app-local same-basename fixtures, exact child-token round trip, and held-
  handle zero-process failure. A synthetic terminal-only pwsh record accepts a
  32,767-unit canonical path as the complete 32,772-unit profile ID and reaches
  fake Running with fixed `pwsh.exe` argv0; path/ID +1 fails before publication. A fake child
  recorder also proves the exact creation flags, non-ASCII argv/environment,
  last-parent-wins duplicate casing/spelling, canonical OS/permitted-integration
  overwrite precedence plus semantic-nonce environment rejection, valid hidden-
  drive preservation and malformed hidden rejection,
  deterministic ordering, exact empty double NUL, command-line 32,767-unit and
  environment 524,288-unit inclusive boundaries/+1, malformed parent
  termination, no truncation, and zero process on rejection.
- WSL distributions consume the canonical bounded shared catalog owned by Emberglass R4-ARCH-01 on a cancelable worker; direct use or duplication of either legacy enumerator is forbidden. If the shared catalog is not ready, coordinate/land that one helper and update `Core_SharedHelpers.md` before this slice. Publish non-WSL profiles without waiting for it. The default enumeration deadline is exactly 3 seconds on the injected monotonic clock; a positive override is hard-clamped to 10 seconds. Setting an override `<= 0` returns `E_INVALIDARG`, leaves the prior/default deadline unchanged, and starts no replacement generation. Each enumeration generation owns one serialized atomic state `Pending|Completed|TimedOut|Cancelled`. A result may change `Pending -> Completed` only after rechecking cancellation and while its publication critical section observes monotonic `now < deadline`; at `now >= deadline`, including exact equality, the timeout transition wins and publishes localized `WSL profiles unavailable` plus Retry/settings link. A close/chooser replacement/retry/application-teardown cancellation flag is checked first under the same serialization, so cancellation wins a result/deadline race and publishes no unavailable row for a surface that no longer owns the generation. Once any terminal state commits it is immutable: the timer/result/cancel losers discard, and a canceled/timed-out generation can never publish late results. Distro launchability is independent of integration capability, and non-WSL profiles never wait for this state machine.
- V1 does not discover Git Bash/MSYS2/Cygwin/custom Windows/SSH profiles.

The pane/profile controller accepts a non-owning, lifetime-stable `ITerminalPreferenceStore`. Its cancelable asynchronous read returns one immutable `{captureEnabled, generation, preference?, storeRevision}` result for a normalized Windows launch-folder identity; its asynchronous conditional write carries that read generation, folder identity, resolved profile ID, update UTC, mutation UUID, and originating session/view generation. The controller revalidates its own view/location/profile generation on every completion. Plan 4 uses only an injected in-memory fake with deterministic generation advancement, stale-write rejection, cancellation, and failure injection; it does not read/write Settings or disk and proves Windows precedence plus the exact profile/purpose-specific one-shot eligibility fence below. A WSL open bypasses the store completely: the exact distro comes from the source and every read/write request is unavailable/no-op. No Plan-4 component may infer durable, cross-process, merge, corruption, or shutdown semantics. Plan 5 implements this same seam with `TerminalStateStore` and reruns the unchanged Windows precedence suite plus the WSL no-read/no-write assertions.

The seam calls are exactly `ResolvePreferredProfile(locationKey,generation)` and
`RecordEligibleProfile(locationKey,profileId,sessionGeneration,
CmdLaunchRootedRunning|PowerShellCommitAck|PendingOpenHostFinalized)`. The
`CmdLaunchRootedRunning` discriminator can be formed only after the plugin
accepts the cmd result's endpoint capability, exact root PID, nonce, operation,
generation, target digest, mode/status/drive cross-fields, and (for UNC) exact
reverse mapping, clears the launch transaction, and that same incarnation wins
its lifecycle-locked `Running` CAS.

Classify the source before any applicable preference lookup. Windows
local/UNC/validated-plugin backing permits only cmd/Windows PowerShell/pwsh and
consults `windowsDefaultProfileId=auto|pwsh|windows-powershell|cmd`; `auto` is
newest supported `autoEligible` pwsh, Windows PowerShell, cmd, while family
`pwsh` means only the newest supported `autoEligible` pwsh. A WSL
UNC/`wsl:` source permits only its exact matching `wsl:<distro>` profile and
never consults or mutates folder-profile memory. For a Windows source,
`ResolveDefault` requires an empty requested ID and resolves a valid compatible
remembered exact profile (including a previously explicit HKCU/PATH profile)
when `rememberShellByFolder=true`, then the Windows default policy, otherwise
chooser; a missing/incompatible remembered choice is shown in that chooser. For
a WSL source, `ResolveDefault` deterministically selects the exact source distro
or shows it unavailable, with no memory/default-distro fallback. `ForceChooser`
requires an empty requested ID and always shows
the compatible chooser—even with one member—without auto-launching remembered
or default choice. `ExactProfile` requires one nonempty bounded stable ID and
either selects that compatible record or shows it unavailable; it never
substitutes. Any other mode/span combination is `E_INVALIDARG` before child or
process creation. When memory is false, ignore every retained
preference—including missing-choice UI—and do not mutate it. A missing explicit
Windows default or WSL distro is shown unavailable with settings link and never
substituted.

Profile-memory eligibility is profile- and purpose-specific; process creation or
public `Running` alone is not the general gate:

- An ordinary interactive cmd tab is eligible only after the plugin accepts the
  correlated local, validated-plugin-backing, or UNC launch-root result under
  the exact pipe/PID/target contract below, clears that launch transaction, the
  same incarnation wins the lifecycle-locked `Running` CAS, and startup
  cancellation/retirement has not won. Cmd has no ongoing semantic handshake.
- An ordinary interactive Windows PowerShell/pwsh tab is eligible only after the
  same incarnation wins its `Running` CAS and the server then accepts its
  complete matching post-Running `CommitAck`. Rooting or `Running` without Ack
  writes nothing; terminal-only final dispositions become eligible on that
  accepted Ack too.
- In the injected Plan-4 hidden-open seam (and the Plan-5 production
  implementation), a `PendingPathInsertion` PowerShell/pwsh open must first
  satisfy the same `CommitAck` gate, then retain its eligible write as deferred through
  `ReadyToCommit`, fallible host preparation, and plugin Commit. It publishes
  exactly once only after the contiguous allocation-free/no-fail
  `FinalizeHiddenTerminalCommit` makes the tab user-owned. Abort, timeout,
  hidden cleanup, source/item staleness, host preparation failure, plugin Commit
  failure, or teardown writes nothing. Plan-4 production resolves
  `IntegrationDisabled` before process creation and can never form
  `PendingOpenHostFinalized`.
- WSL never reads or writes profile memory in v1 because its exact profile is
  source-derived and `Running` is not cwd proof.

Every conditional preference write is one-shot and generation-bound. Creation
failure, pre-CAS exit/cancel, a PowerShell post-CAS failure before accepted Ack,
or any ineligible hidden outcome writes nothing. A stale/failed eligible write
is non-success and may offer Retry but cannot change the active terminal. Shell
cwd reports never mutate another folder's choice.

## Working-directory plans

- Generate logical `TerminalLocationIdentity` without collapsing symlinks/junctions/mapped/SUBST/UNC aliases; validation is separate.
- Location identity and display representation preserve every admitted code unit. Component ceilings are Windows path 32,767, Linux path 32,768, distro 256, and plugin short ID 128. Canonical `locationKey` is capped at 33,037 UTF-16 units (the maximum `file:wsl:<decimal-length>:<distro><linux-path>` form), and `displayPath` at 33,040 (the maximum canonical `\\wsl.localhost\<distro><path>` form). Exact maxima succeed independently of process launchability; +1, malformed length prefix, noncanonical spelling, or key/location mismatch rejects the identity. There is no prefix truncation, hash/alias substitution, physical-target replacement, or display spelling used as identity.
- Support local drive/safely bootstrapped UNC/validated plugin local backing only with Windows profiles; matching WSL UNC/`wsl:` only with its exact distro profile. Cross-family/distro and unsupported cloud/archive/MTP/remote locations remain diagnostic; no root/home/temp/default-distro fallback.
- Use explicit executable, mutable Windows command line/argument encoding, the core-built canonical explicit UTF-16 environment block, safe current directory, and exactly `EXTENDED_STARTUPINFO_PRESENT | CREATE_UNICODE_ENVIRONMENT`. No inherited/null environment and no concatenated untrusted shell command.
- Revalidate immediately before process creation. Every cmd launch runs its
  mandatory one-use `LaunchCwdBootstrap` through the verified embedded `.cmd`,
  protected precreated inbound pipe, eight fixed overwritten environment keys,
  and exact bounded ASCII Ack below. Local/plugin-backed applies the fixed
  direct-expanded `cd /d` primitive, while UNC additionally revalidates and
  reverse-maps the script's `pushd` drive. All eight child keys and every secret
  or target-bearing mutable plugin launch value are cleared before
  `Running`/user input and never become ongoing trusted cwd, activity, history,
  follow, insertion, or Restart authority. UNC retains only the bounded
  nonsecret temporary-mapping ownership tuple defined below until root death.

For WSL, resolve/open only the non-reparse
`<GetSystemDirectoryW()>\wsl.exe` and pass that exact path as
`CreateProcessW.lpApplicationName`; PATH, app/current directory,
`ShellExecute`, and command processors are never launch authorities. The exact
logical argv is `[wsl.exe, --distribution, <catalog exact distro name>, --cd,
<absolute profile-native Linux path>]`. Do not add `--exec`, `--user`, `-e`, a
shell/bootstrap command, an extra `wsl.exe`, or home/default-distro fallback.
The configured distribution default user and default shell run unchanged.
`--cd` is a supported-platform invariant under the existing Windows-build-
22000.2600 minimum, not a help/version probe. Reject NUL/CR/LF, an empty or
leading-dash distro, a non-absolute Linux path, and command-line overflow before
creation.
The full Linux path may remain representable as identity while not launchable.
Quote all five tokens through the reviewed shared `CommandLineToArgvW`-inverse
helper and require the mutable command line including NUL to be at most 32,767
UTF-16 units before preflight/ConPTY/process creation. There is no shortening,
hash, environment indirection, response file, wrapper, or second process. With
quote-free one-unit distro `D`, constants consume 31 units including NUL, so
32,736 Linux-path units is the exact boundary and +1 starts no preflight or
process.
V1 performs no pre- or post-launch WSL integration injection: no bootstrap is
concatenated into argv or environment, written into the PTY, installed in a
Linux profile, staged through a guessed `/mnt/<drive>` path, or carried through
`WSLENV`; no wrapper or second process is allowed. Every WSL shell, including
bash, is terminal-only and creates no semantic-integration nonce or epoch.

Before creation, a cancelable off-UI generation-bound
`WslLaunchPathPreflight` freezes exact catalog generation, distro name, Linux
path, and launch generation; opens the exact canonical
`\\wsl.localhost\<catalog exact distro name>\<path>` directory for attributes
with WIL ownership/no delete sharing; revalidates directory and catalog
generation; and retains the handle through `CreateProcessW` plus startup-channel
admission. Its fixed monotonic three-second deadline irrevocably fails the
generation before requesting `CancelSynchronousIo` on its owned worker, so a
late completion can only release its handle. Cancel/timeout/stale/missing/path
failure creates the closable Retry/settings diagnostic and starts no process.
The service owns the `std::jthread` through return under the same quiet/fatal
policy; it never detaches, UI-joins, runs preflight on UI, changes process cwd,
or reuses a cached success as proof.

`wsl.exe --cd` has no shell-independent post-create success acknowledgement.
The terminal may publish `Running` after exact preflight/ConPTY/process
admission only with internal `launchCwdTrust=RequestedUnverified`. A later
`Running` state claims only admitted launcher/preflight/ConPTY/process startup,
not an authenticated Linux cwd. `launchCwdTrust` remains
`RequestedUnverified` for the entire v1 incarnation, there is no WSL trusted-cwd
event, and every attempted transition to `Verified|Diverged` is rejected. App
history, authenticated activity/cwd, pane follow, every path-insertion mode, and
WSL folder-profile reads/writes are disabled. An asynchronous launcher error or
root exit may therefore follow a briefly visible `Running` state after an
external distro/path race, but it never triggers retry, fallback, or memory;
tests must not require every asynchronous `--cd` failure to appear before
`Running`.

A later WSL-integration feature requires a separate approved design before code
changes. Its first gate must prove clean-default Bash plus profile-override/
`exec`/early-exit behavior, bootstrap before user-input admission, zero visible
PTY/native-history input, no Linux-profile edit, configured default-user/
default-shell/startup preservation, no `/mnt/c` assumption, bounded nonce/path-
state erasure, and safe fallback to a fully usable terminal-only session. If it
needs `--exec`, a wrapper, `WSLENV`, or another process, that design must
explicitly replace the five-token invariant and freeze exact argv/environment/
path/ownership/timeout/cleanup semantics. V1 carries no dormant or best-effort
experiment.

### Windows launch-rooting protocols

`shellIntegrationEnabled` controls ongoing prompt/history/follow/insertion
semantics; it never weakens launch rooting. Plan 4 implements two different
plugin-only startup protocols. Every local, validated-plugin-backing, or UNC cmd
launch uses exactly one correlated `LaunchCwdBootstrap`, even while integration
is disabled; the initial `CreateProcessW` cwd is never sufficient. Every Windows
PowerShell/pwsh launch, local or UNC, uses the separate
`PowerShellStartupBootstrap`; calling the cmd transaction or treating
PowerShell as a one-result UNC-only bootstrap is forbidden.

- Cmd starts from the validated local/plugin-backed requested directory, or from
  the API-derived safe local system directory for UNC, and preserves normal
  AutoRun ordering. Plan 4 embeds one versioned reviewed `.cmd` as
  `Terminal.dll` RCDATA, records its resource ID/version/SHA-256, and extracts it
  through the same per-user owner/access/reparse-safe transaction. Immediately
  before every launch the plugin reopens and hashes the extracted bytes. The
  exact child token array is
  `["cmd.exe","/E:ON","/V:OFF","/S","/K",
  "call \"<verified-extracted-script-path>\""]`. The last token is produced by
  one closed two-pass `cmd /S /K`+`call` path encoder and contains no target or
  channel secret. Golden real-cmd tests round-trip that absolute path with
  spaces, Unicode, `%`, `!`, `^`, `&`, `|`, `<`, `>`, and parentheses and prove
  one script invocation; inability to encode it fails before process creation.
- Before `CreateProcessW`, `Terminal.dll` creates exactly one
  `\\.\pipe\RedSalamander.TerminalCmdV1.<32-lowercase-hex-random>` server. It
  is `PIPE_ACCESS_INBOUND | FILE_FLAG_OVERLAPPED |
  FILE_FLAG_FIRST_PIPE_INSTANCE`, message-mode, one-instance, uses
  `PIPE_REJECT_REMOTE_CLIENTS`, and has exactly one overlapped
  `ConnectNamedPipe` pending before child creation. Only after connection
  completion and exact-root-PID validation may the first overlapped `ReadFile`
  arm; claiming a pre-connect pending read is forbidden. Its noninherited
  descriptor is protected, owned by the current-
  user SID, and its DACL contains only LocalSystem full access and that SID
  write/synchronize access. Accept only a client whose
  `GetNamedPipeClientProcessId` is the exact root `cmd.exe` PID.
- For UNC, split the already-normalized requested path into its exact logical
  connection root `\\server\share` with no trailing separator and a remaining
  tail that is empty or begins with one backslash. Snapshot
  `GetLogicalDrives()` before `CreateProcessW`, reject a zero/failing snapshot,
  and require the Ack's reported uppercase drive letter to have been free in
  that bitmap. This is a logical
  UNC proof: it never resolves a symlink, DFS target, provider alias, or physical
  identity.
- The explicit child environment first removes every inherited case-insensitive
  variant, then writes exactly these reserved names:
  `REDSALAMANDER_TERMINAL_CMD_V1_PIPE`,
  `REDSALAMANDER_TERMINAL_CMD_V1_CAPABILITY`,
  `REDSALAMANDER_TERMINAL_CMD_V1_NONCE`,
  `REDSALAMANDER_TERMINAL_CMD_V1_OPERATION_ID`,
  `REDSALAMANDER_TERMINAL_CMD_V1_LAUNCH_GENERATION`,
  `REDSALAMANDER_TERMINAL_CMD_V1_MODE`,
  `REDSALAMANDER_TERMINAL_CMD_V1_TARGET`, and
  `REDSALAMANDER_TERMINAL_CMD_V1_TARGET_UTF16_SHA256`. `PIPE` is the exact full
  random-suffix name above; capability is an independent 256-bit
  `BCryptGenRandom` value encoded as 64 lowercase hex; nonce and operation ID are
  independent 128-bit values encoded as 32 lowercase hex; generation is
  canonical nonzero `uint64_t` decimal; mode is exactly `LOCAL|UNC`; target is
  the copied validated UTF-16 logical path; and digest is 64 lowercase hex for
  SHA-256 over the target's exact UTF-16LE code units without BOM or terminating
  NUL. Target containing NUL, CR, LF, or `"` is invalid; spaces, Unicode, `%`,
  `!`, `^`, `&`, `|`, `<`, `>`, and parentheses are literal required fixtures.
- After normal AutoRun, the script starts exactly with
  `setlocal EnableExtensions DisableDelayedExpansion`, validates/stages only
  those fixed keys, and direct-expands the target once into exactly
  `cd /d "<target>"` for local/plugin-backed or `pushd "<target>"` for UNC.
  It uses no `call set`, delayed/recursive target expansion, child process,
  PowerShell, helper EXE, VT output, DOSKEY, prompt edit, or HMAC. It calls its
  emit subroutine with safe-ASCII fields only, clears every reserved and private
  environment value before the Ack/prompt, executes `endlocal`, and clears any
  restored reserved variants again on every exit path. All eight names are
  permanently absent from the interactive shell after bootstrap; no inherited
  value is restored. The emit subroutine alone writes the result; mutable plugin
  staging is likewise cleared before Running,
  and immutable cmd parser/string buffers are not claimed securely erased.
- The sole logical result is exactly the fixed-case/fixed-order ASCII line
  `RSTCMD1;CAP=<64-lowerhex>;NONCE=<32-lowerhex>;OP=<32-lowerhex>;
  GEN=<canonical-uint64>;MODE=<LOCAL|UNC>;TARGET=<64-lowerhex>;
  STATUS=<0|1|2>;DRIVE=<-|[A-Z]:>\r\n`. `STATUS=0` is primitive success,
  `1` is `cd`/`pushd` primitive failure, and `2` is invalid/missing staging.
  `DRIVE=-` is required except for `STATUS=0;MODE=UNC`, which requires one
  uppercase `[A-Z]:`. The full result is at most 512 bytes. The emit subroutine
  uses one redirection scope, which closes its pipe handle immediately after the
  Ack attempt. Because cmd may issue more than one message-mode write inside
  that scope, the overlapped server accumulates complete messages into one 513-
  byte staging buffer and parses nothing until a clean client EOF/disconnect
  arrives strictly before the deadline. `ERROR_MORE_DATA`, aggregate byte 513,
  empty/partial-at-EOF, first-valid-then-second, second-line, trailing, non-ASCII,
  noncanonical data, or a client that keeps the handle open through the deadline
  fails startup. Only one exact CRLF-terminated line with no extra byte is
  accepted. There is no cmd-side HMAC or Unicode cwd serialization.
- Strictly before the common five-second deadline, the plugin validates the
  exact root PID, capability, nonce, operation, generation, mode, target digest,
  status/drive cross-fields, EOF, and immediate target identity. A local success
  is the reviewed `cd /d` primitive's status bound to that revalidated target;
  the plugin does not pretend to read cmd's Unicode cwd. For UNC, after the
  valid Ack and immediate target revalidation, require
  `GetDriveTypeW(L"<D>:\\") == DRIVE_REMOTE`, then call
  `WNetGetUniversalNameW(L"<D>:\\<tail>", REMOTE_NAME_INFO_LEVEL, ...)` with an
  initial 4,096-byte owned buffer and at most one `ERROR_MORE_DATA` retry to the
  exact reported byte count, which must be greater than 4,096 and at most 262,144
  bytes. Checked arithmetic rejects a second resize/error, zero/non-growing or
  inconsistent returned size, overflow, and any null, misaligned, out-of-buffer,
  or missing-NUL `REMOTE_NAME_INFO` pointer/string. Normalize `lpUniversalName`, `lpConnectionName`, and
  `lpRemainingPath` only with the same logical Windows UNC lexical rules used by
  the request—separator and extended-prefix spelling, root/trailing separator,
  and lexical dot segments—and compare them respectively with the complete
  requested UNC, exact requested connection root, and exact requested tail by
  ordinal-ignore-case. The canonical connection root is exactly
  `\\server\share` with no trailing slash. Accept `lpRemainingPath` only as empty
  or a relative sequence with zero or one leading slash; reject a drive/UNC form,
  two leading slashes, NUL, root escape, or `..` above the share, and canonicalize
  a nonempty remainder to exactly one leading slash while lexically removing `.`
  and legal `..`.
  Require normalized connection plus remainder to equal both normalized
  universal name and requested full path, while root and tail also match their
  respective expected values ordinal-ignore-case. A provider, DFS, share, or
  alias canonicalization change fails even when it reaches physically equivalent
  storage.
- Failure status, missing/late output, connect/read failure, wrong mapping, root
  exit, AutoRun exit/hang, or startup cleanup failure remains `Starting` and
  records no preference. Success closes/drains the pipe and clears capability,
  nonce, operation, digest, target, and all other secret/target-bearing staging
  before the same incarnation may form `CmdLaunchRootedRunning`, admit input, or
  write profile memory. Through root lifetime it retains only
  `{drive, expectedRoot, expectedTail, launchGeneration,
  prelaunchLetterWasFree}`; this one-use correlation result is never an ongoing
  semantic epoch.
- Never unmap the retained UNC drive while the root process lives. After root
  death, one counted off-UI cleanup unit revalidates the same drive and retained
  record with the exact bounded `WNetGetUniversalNameW` proof above and
  `WNetGetConnectionW`. That connection call starts with 4,096 UTF-16 code units,
  permits at most one `ERROR_MORE_DATA` resize to the exact reported count capped
  at 262,144 code units, and rejects malformed/nonterminated data.
  `ERROR_NOT_CONNECTED` is already-success. Only when the same drive is remote,
  both proofs still match, and
  `prelaunchLetterWasFree` is true may the plugin call
  `WNetCancelConnection2W(L"<D>:", 0, FALSE)` exactly once, then it must recheck
  with fresh `WNetGetConnectionW` and require `ERROR_NOT_CONNECTED`. A changed,
  reused, shadowed, in-use, access/error, failed-cancel, still-connected, or
  otherwise unproved mapping is never canceled or forced; leave it untouched,
  emit one localized content-free cleanup failure/counter, and use the existing
  terminal cleanup-failure status without an unbounded wait or retry/broader
  unmap. The cleanup record/worker gate releases only after this attempt. The
  design never assumes process
  exit cleaned a `pushd` mapping: only `popd` is documented to remove that
  temporary assignment, and this protocol never issues `popd`. Same-user reuse
  of the formerly free letter for the exact same logical target before cleanup is
  indistinguishable through WNet and remains inside the declared malicious-same-
  user limitation; the plan does not claim hostile same-user ownership proof.
  Changed-target reuse remains detectable and must be left untouched.
- User-authored cmd AutoRun runs before `/K` and is explicitly trusted, just as
  user-authored PowerShell/pwsh profiles are. The pipe/PID/random fields protect
  integrity, correlation, and replay against PTY bytes, stale launches, and
  unrelated processes; they cannot distinguish reviewed bootstrap execution
  from earlier code already running in that same root process with access to its
  argv/environment. Same-root correct-capability/correct-nonce preemption and a
  malicious same-user process are outside the stated threat model. Tests must
  document that limitation while proving wrong PID/capability/nonce/operation/
  generation/digest, replay, duplicate line, and PTY output fail closed.

PowerShell uses this exact closed exchange:

1. Before process creation, create a fresh 128-bit bootstrap key, 128-bit
   operation ID, nonzero 64-bit launch generation, and exactly one
   `\\.\pipe\RedSalamander.TerminalBootstrap.<32-lowercase-hex-random>` server.
   The pipe is duplex, overlapped, first-instance, message-mode, one-instance,
   with 256-KiB inbound/outbound buffers and `PIPE_REJECT_REMOTE_CLIENTS`. Its
   non-inherited protected security descriptor is owned by the current-user SID
   and grants only that SID read/write/synchronize. Accept only the client whose
   `GetNamedPipeClientProcessId` is the exact launched root PID. This bootstrap
   key is distinct from every cmd launch nonce and later semantic nonce.
2. The exact child token array after fixed basename `argv[0]` is
   `["-NoExit","-Command",bootstrapWrapperCommand]` for both
   `powershell.exe` and `pwsh.exe`; the common Windows quoting helper must
   round-trip those four total tokens exactly. Do not add `-NoProfile`,
   `-NonInteractive`, `-ExecutionPolicy`, `-File`, `-EncodedCommand`, or a token
   after the wrapper. Normal profiles load exactly once before the wrapper and
   policy is never bypassed. Generate `bootstrapWrapperCommand` only from one
   reviewed immutable template plus seven internal values: verified script
   path, launch mode, bootstrap pipe, bootstrap key, operation ID, launch
   generation, and protocol version `1`; it contains no requested target or
   semantic nonce. Encode every value with the one closed PowerShell literal
   primitive: reject NUL/CR/LF or unpaired UTF-16, double every apostrophe, then
   surround with one apostrophe pair. No interpolation, expandable string,
   backtick, expression, environment lookup, or pane/user text is permitted.
   The wrapper is one top-level try with exact canonical-ASCII grammar:
   `try { $__rsResult=@(& <scriptLiteral> -ProtocolVersion <versionLiteral>
   -LaunchMode <modeLiteral> -BootstrapPipe <pipeLiteral>
   -BootstrapKey <keyLiteral> -LaunchOperationId <operationLiteral>
   -LaunchGeneration <generationLiteral> -ErrorAction Stop *>&1);
   if ($__rsResult.Count -ne 1 -or
   $__rsResult[0].PSTypeNames[0] -cne
   'RedSalamander.TerminalBootstrapFinalize.V1' -or
   (@($__rsResult[0].PSObject.Properties.Name) -join ',') -cne
   'Sentinel,Finalize' -or ([string]$__rsResult[0].Sentinel) -cne
   'RSTBOOT-WRAPPER-READY-1' -or -not
   ($__rsResult[0].Finalize -is [scriptblock])) { exit 126 };
   $__rsFinalize=[scriptblock]$__rsResult[0].Finalize; $__rsResult=$null;
   $null = & $__rsFinalize; $__rsFinalize=$null }
   catch [System.Exception] { exit 125 }`; whitespace is the one canonical ASCII
   form frozen with the resource and placeholders are replaced only by those
   literal tokens. Initialization, assignment, validation, and finalizer
   invocation stay inside the catch boundary, so a profile-created read-only/
   AllScope collision exits rather than falling through to a prompt. The fixed
   `*>&1` captures any actually emitted warning/verbose/debug/information stream
   as extra result evidence, but never redirects the plugin's binary pipe
   traffic.
   The `[CmdletBinding(PositionalBinding=$false)]` script declares exactly six
   mandatory named bootstrap parameters plus wrapper-supplied common
   `-ErrorAction Stop`; it validates exact version `1`, `Local|Unc` mode, pipe,
   two exact 32-lowercase-hex values, and canonical nonzero unsigned-decimal
   generation grammar and rejects duplicates/unbound arguments. After
   authenticated `CommitPrompt`, it applies final disposition, prebuilds/HMACs
   CommitAck, zeroes the final key buffer, and returns exactly one typed two-
   property finalize object. Its private one-shot scriptblock owns only the
   bootstrap client and immutable prebuilt Ack, sends that complete frame in one
   synchronous message-mode `Write` with no Flush, retry, later pipe operation,
   or pipeline output, and is the only Ack sender—only after wrapper validation.
   Missing/wrong/duplicate/extra output, invalid object/finalizer,
   invocation/parse/policy failure, finalizer throw/partial write, or named
   runtime exception leaves Ack unaccepted and exits `125|126`; a failed/partial
   write is never repaired/retried. Client disposal is idempotent/output-silent,
   suppresses named disposal errors, and cannot alter a written record. The
   server opens gates only after reading one full valid message, never from
   client Write completion, and `-NoExit` cannot fall through to prompt. A profile exit/hang before
   `Ready` has the same bounded failure; a policy-blocked script must evaluate
   no user `prompt` function.
3. Bootstrap frames are separate from VT effects and the ongoing semantic
   channel. Each is exactly `magic[8]="RSTBST01"`,
   `frameSize@8:uint32`, `version@12:uint32=1`, `kind@16:uint32`,
   `flags@20:uint32=0`, `directionSequence@24:uint64`,
   `launchGeneration@32:uint64`, `operationId@40:uint8[16]`,
   `payloadBytes@56:uint32`, `reserved@60:uint32=0`, payload at offset 64,
   followed by a 32-byte HMAC-SHA256 over all preceding bytes using the
   bootstrap key. `frameSize == 64 + payloadBytes + 32` and is at most 262,144.
   Each direction starts at sequence one, advances exactly once, and cannot
   wrap. Integers are little-endian. Text is a uint32 UTF-16-code-unit count
   followed by exactly that many UTF-16LE units, with no terminator, embedded
   NUL, or unpaired surrogate. Unknown kind, nonzero flag/reserved, trailing
   byte, bad length/tag/generation/operation/sequence, duplicate, replay, or
   oversize is fatal startup failure.
4. The sole exchange is child `Ready=1` sequence 1/zero payload; plugin
   `RootRequest=2` sequence 1/exact validated requested directory; child
   `RootResult=3` sequence 2 with
   `{status:uint32 Success=0|Failure=1,cwdCodeUnits:uint32,cwdUtf16le}`, where
   Failure has zero cwd units; plugin `SemanticDecision=4` sequence 2; child
   `Prepared=5` sequence 3; plugin `ReleasePrompt=6` sequence 3/zero payload;
   child `ReleaseAck=7` sequence 4/zero payload; plugin `CommitPrompt=8`
   sequence 4 with exactly
   `{finalDisposition:uint32 TerminalOnly=0|SemanticActive=1,
   policyGeneration:uint64}`; child `CommitAck=9` sequence 5/zero payload. Local launch initially uses the
   validated requested directory and UNC uses the safe system directory, but in
   both modes profiles run before the script calls only module-qualified
   `Microsoft.PowerShell.Management\Set-Location -LiteralPath <exact
   RootRequest>` and obtains cwd only through module-qualified
   `Microsoft.PowerShell.Management\Get-Location`; profile aliases/functions
   cannot intercept or add pipeline output. It never uses expression evaluation,
   interpolation, provider fallback, or command concatenation. The plugin independently resolves the cwd and requires
   canonical logical identity equal to the immutable request. Rooting enters
   neither PSReadLine nor RedSalamander history.
5. Only after a valid matching `RootResult` establishes the immutable launch-
   rooted barrier does the plugin sample the latest effective
   `shellIntegrationEnabled` and integration-policy generation and send its one
   decision. `SemanticDecision` begins
   `{decision:uint32 TerminalOnly=0|SemanticEnable=1,policyGeneration:uint64}`.
   `TerminalOnly` ends there and creates no semantic nonce/control pipe.
   `SemanticEnable` additionally carries one newly generated 128-bit semantic
   nonce and bounded UTF-16 names for the already-created, distinct
   child-to-plugin `TerminalSemantic` and plugin-to-child `TerminalControl`
   pipes. The target arrives only in `RootRequest`; the semantic nonce and both
   names arrive only in this decision. None is guessed or placed in initial argv
   or the child environment, and no second decision/rekey is allowed.
6. Plan 4 keeps the production adapter seam deliberately closed. For
   `SemanticEnable`, the production script installs no hook, opens no semantic
   pipe, performs no ongoing semantic handshake, and deterministically reports
   `Prepared(SemanticUnavailable)` while user input remains gated. For
   `TerminalOnly`, it likewise installs no hook and opens no semantic pipe.
   `Prepared` is exactly
   `{outcome:uint32 TerminalOnly=0|SemanticReady=1|SemanticUnavailable=2,
   policyGeneration:uint64}`. This production Unavailable result permits a
   truthful terminal-only launch with zero semantic capability. Plan 4 packages
   no allowlisted adapter manifest, observer/control script or resource, install
   path, or production `SemanticReady`; those first belong to plan 5. Only an
   injected fake adapter may return `Prepared(SemanticReady)` and send the sole
   fake authenticated handshake in order to prove the generic join machinery.
   Decision/outcome or immutable decision-generation mismatch is fatal; do not
   compare `Prepared` to a newer live generation. In that fake-only branch,
   Prepared and handshake use independent transports and have no cross-handle
   arrival order. Own one bounded `StartupSemanticJoin` with two optional slots keyed by exact
   launch/session/incarnation/profile/executable/location, semantic nonce/epoch,
   Decision policy generation, and startup generation. Either record may arrive
   first. Independently authenticate and fully bounds/sequence/schema/admission-
   validate each, copy it into its slot, and grant no capability from one slot;
   only their matching conjunction strictly before the common startup deadline
   advances to `ReleasePrompt`. Both handlers use the same session lock only to
   reserve/fill/scrub these fixed slots and post one generation-checked
   continuation; neither waits, parses, writes a pipe, invokes a callback, or
   takes a model/writer/host lock while holding it. Reserve both bounded slots
   before child creation so parser-callback admission is nonblocking; exhaustion
   is deterministic fatal startup failure. The semantic handshake is sequence one and no
   other semantic event is legal before conjunction. Duplicate, conflicting,
   gapped, wrong-identity, oversized, or aggregate-effect queue-admission failure
   is fatal, permanently revokes the staged epoch, and scrubs both slots.
   Production `TerminalOnly` requires both slots absent and a Handshake paired
   with it is fatal. Production Unavailable normally has no Handshake; the fake
   harness additionally proves the sole valid boundary race in which a complete
   authenticated matching sequence-1 Handshake was staged by the server before
   its receive deadline but the child observed its write at/equality-after 50 ms
   and reported Unavailable. Discard/revoke that frame and finish terminal-only;
   malformed, mismatched, or extra data remains fatal. Cancellation, root exit,
   timeout at equality, Restart, or shutdown scrubs both, and no late record can
   complete a later incarnation. Live true-to-false is the sole nonfatal
   exception: mark the join `RevokedTerminalOnly`, permanently close semantic/
   control operation and capability admission, and scrub/ignore staged and late
   old-generation records, but while Starting before the CAS retain both already-
   created fake servers solely as exact-root-PID, no-operation cleanup transports.
   They may finish connect/setup and the fake performs no control writes; disable
   must not manufacture connect/setup failure to signal policy. Stop requiring
   the Handshake and accept otherwise-valid
   `Prepared(SemanticReady|SemanticUnavailable)` only as cleanup proof. It grants
   nothing and absence still times out. Close/cancel both pipes at the Running
   CAS or earlier only for an independently winning retirement/protocol failure.
   `CommitPrompt(TerminalOnly,currentPolicyGeneration)` then requires any
   fake-staged module closure to latch process-lifetime inert/pass-through and
   remove its fake control mapping iff still app-owned before `CommitAck`; the
   Plan-4 production script has neither. False-to-true changes neither the join
   nor an original TerminalOnly decision.
   The injected control fake models the final protected server/client boundary:
   current-user ACE and client rights are exactly
   `ReadData|ReadAttributes|WriteAttributes|Synchronize`, `WriteData` is absent,
   and the exact
   asynchronous/anonymous/noninherited client changes only
   `ReadMode=Message`. Its loaded-fake test proves connect/mode-switch; this
   remains harness proof and adds no Plan-4 production adapter or resource.
7. The session stays `Starting`; ConPTY user input, profile-memory writes,
   trusted activity, and all semantic capabilities remain closed through
   `Prepared`. After accepting it, send `ReleasePrompt`. The script removes the
   target/decision/semantic-nonce variables from script scope, zeroes their
   mutable byte/character buffers, retains only one module-private mutable
   bootstrap-HMAC-key buffer, and sends `ReleaseAck`. It then performs no
   prompt, input, hook dispatch, or other pipe operation except waiting for the
   exact authenticated `CommitPrompt`. EOF, pipe error, cancellation, unknown
   data, or timeout before that frame is fatal: zero the remaining key, close
   the client, and terminate the root without returning to PowerShell's
   interactive loop.
8. Accepting `ReleaseAck` does not commit startup. Under the one session
   lifecycle lock, revalidate the immutable launch/session/view/profile/
   executable/location identities, launch generation, root liveness, no prior
   retirement winner, and strict deadline. Compare current live policy
   generation with the immutable Decision, choose the truthful final semantic
   disposition, then perform the sole CAS from
   `Starting/BootstrapReleaseAcked` to `Running/PromptCommitPending`. That CAS is
   the only Running linearization point and wins or loses atomically against
   Close, WindowClosing, Restart, root exit, timeout, and ApplicationShutdown.
   On loss, send no `CommitPrompt`; startup retirement closes the pipe, the
   script exits fatally, and no prompt/profile write/input/capability appears.
   On success, public lifecycle is Running before any prompt can run, but
   `startupInputGate` remains closed, profile-memory eligibility false, and all
   semantic capabilities unpublished.
9. After the Running CAS, close/cancel both startup-revoked fake
   `TerminalSemantic` and `TerminalControl` servers and send
   authenticated `CommitPrompt` with the CAS-frozen final disposition and
   current policy generation. The exact generic
   Decision/Prepared/generation matrix is below; its `SemanticReady|Active` row
   is reachable in Plan 4 only through the injected fake:
   original TerminalOnly plus Prepared TerminalOnly remains TerminalOnly at any
   equal-or-later generation; SemanticEnable plus SemanticReady with a completed
   join is SemanticActive only at the unchanged generation; SemanticEnable plus
   SemanticUnavailable is TerminalOnly at the unchanged generation; and a later
   latched disable makes either valid Enable outcome TerminalOnly at its current
   strictly greater generation, regardless of a still-later enable. The script
   retains only its immutable Decision generation and Prepared outcome through
   CommitPrompt. Generation decrease/wrap and every other pair are fatal.
   `TerminalOnly` must latch every staged wrapper process-lifetime
   inert/pass-through and remove the control mapping only if still app-owned;
   `SemanticActive` requires the validated completed join. Only after applying
   that disposition may the script prebuild/HMAC `CommitAck`, zero its last
   mutable bootstrap-key buffer, and return the sole typed
   `RSTBOOT-WRAPPER-READY-1` finalize object; it does not send the Ack. The fail-
   closed wrapper validates that exact object, then invokes its private one-shot
   output-silent finalizer, whose single synchronous message-mode Write sends the
   prebuilt Ack with no Flush/retry/later operation. Accept the ack only after
   the server reads the complete valid frame for the same Running incarnation and strictly
   before the deadline. Policy generation is not part of this cleanup-Ack key.
   If the sole intervening policy change is an already-latched true-to-false
   fence, consume the matching ack, close/zero bootstrap state, open ordinary
   user input, and permit the one-shot ordinary Windows PowerShell/pwsh
   preference write whose matching `CommitAck` is the eligibility point, but
   publish zero semantic capability and reject old-generation
   events. An Enable decision disabled before or after the CAS therefore closes/
   cancels both pipes and keeps the wrapper inert/pass-through without failing
   terminal startup. A TerminalOnly decision later enabled likewise consumes
   the ack, remains terminal-only, and reports Restart required. Neither sends a
   second decision. With no downgrade, accepted ack performs the same cleanup/
   input/profile-memory steps and publishes only independently proved semantic
   capabilities. Plan-4 production always took Unavailable and therefore
   publishes zero; a fake Active result proves only the gate implementation and
   cannot satisfy product acceptance. Only an unrelated lifecycle/protocol/root/deadline/wrapper-
   finalization failure after the Running CAS but before accepted
   `CommitAck` retires the already-Running incarnation with input/capabilities/
   profile memory still gated. Whether the child did or did not see
   `CommitPrompt`, first-prompt evaluation can never occur for a never-Running
   launch; disable alone never retires it. A disable winning after the CAS uses
   the ordinary Running fence: immediately close/cancel both fake pipes and
   suppress capability publication even if an already-built fake CommitPrompt
   says `SemanticActive`. `TerminalControl`, not write-only `TerminalSemantic`,
   is the fake child's pending-read revocation sentinel; completed control EOF/
   error latches the interposer inert, while event-pipe closure breaks a pending
    or next fake semantic write. The fake no-Flush OVERLAPPED writer-admission,
    pre-Write fixed-protocol-reserve inert wake placeholder with its final FIFO
    sequence, AwaitingResult/complete-result-identity publication before same-
    descriptor activation/no second admission or priority bypass,
    exactly-one-wake/no-retry, one validated-frame stage where ordinary-wrapper
    `PollControlTransport(NoWait)` may stage/rearm but never dispatch and handler-
    only `PollControlTransport(HandlerWaitOnce)` consumes an occupied stage first
    with zero wait/EndRead on the new read and otherwise uses its sole nonalertable
    `AsyncWaitHandle.WaitOne(250)`, strict post-read deadline/one-shot
    dispatch, mapping-replacement-at-most-one-foreign-wake, closed persistent-
    state/result disposition, ambiguous-input retention, and quiet-drain seams
    match the final Plan-5/Plan-6 production contract but add no Plan-4 product
    adapter/resource.
    The fake persistent states are exactly `Connecting -> Idle -> WritePending ->
    AwaitingResult -> Idle`. Matching authenticated Success and, for Follow/
    Insert only, matching pre-effect RejectedState or BufferChanged release the
    operation/input hold, advance the no-wrap sequence, and return Idle; a
    mechanically proved pre-submit/no-action loss returns Idle without advancing.
    Probe non-Success, Expired, BindingChanged, ShellOperationFailed, timeout,
    malformed/mismatched result, policy/protocol failure, or unresolved ambiguity
    retires control, latches terminal-only, and bars every later operation.
10. PowerShell/.NET immutable strings and the OS process command line cannot be
   securely overwritten. Drop references promptly and zero mutable buffers, but
   keep same-user command-line/managed-heap forensics outside this nonce threat
   boundary. No spec/test may claim physical erasure of immutable strings.
11. One monotonic `launchT0 + 5 seconds` deadline covers connection and every
   message through `CommitAck`; each receipt is strictly before it and equality
   times out. Close, WindowClosing, Restart, root/profile exit, pipe error,
   cancellation, app shutdown, invalid rooting, missing handshake, or no valid
   ack closes startup and staged semantic handles, revokes the staged epoch, and
   remembers no profile. Before the Running CAS, the existing retirement arbiter
   owns the outcome: explicit close/shutdown removes or closes normally, while
   launch/protocol/profile/policy failure leaves the closable Retry/Choose
   Folder/Open Terminal settings diagnostic. After the CAS, follow the matching
   ordinary already-Running retirement path with startup gates still closed. No
   late frame or setting can resurrect either outcome, and execution-policy
   failure is never retried with bypass or a different directory.

### Injected semantic payload schema (fake-only in Plan 4)

Plan 4's injected Ready/join/control fake uses the exact future production wire
ABI below so it proves the shared runtime seam without inventing a second test
protocol. The production Plan-4 plugin still installs/packages no adapter,
semantic/control script, or schema resource and deterministically reports
Unavailable; plan 5 first turns this schema into a generated product resource.

- The schema is closed to exactly
  `Handshake=1|RootReadReady=2|AcceptedInput=3|ControlResult=4`. Decoders perform
  explicit little-endian unaligned scalar loads/copies with checked offset/length
  sums; they never cast a wire buffer to a struct, consume compiler padding, or
  accept a TLV/extension. Every payload starts with `payloadVersion@0:uint32=1`,
  has its exact computed size with no trailing byte, and treats an unknown kind,
  enum, flag bit, reserved value, or cross-field combination as fatal. Transport
  identity already binds profile/session/incarnation/epoch; payloads never
  duplicate or weaken it. C++ reuses the strict conversions in
  `Common/StringConversion.h` before applying protocol-specific control/path
  validation; embedded PowerShell uses exactly
  `System.Text.UTF8Encoding(false, true)`. A Terminal-local UTF converter or
  replacement/fallback encoding is forbidden. No text slice may contain U+FEFF,
  U+0000, or an unpaired surrogate. Tuple IDs and paths reject every control
  character. Command alone permits HT, LF, and CR from the C0/C1 ranges and
  rejects every other C0/C1 code point.
- `Handshake=1` is exactly `payloadVersion@0:uint32=1`,
  `psEdition@4:uint32 Desktop=1|Core=2`,
  `editMode@8:uint32 Windows=1|Emacs=2`, `wakeCodePoint@12:uint32` in exactly
  `{0x1c,0x1d,0x1e,0x1f,0x07}`, nonzero
  `channelGeneration@16:uint64`, `capabilityFlags@24:uint64`,
  `tupleIdBytes@32:uint32`, `psVersionBytes@36:uint32`,
  `psReadLineVersionBytes@40:uint32`, `reserved@44:uint32=0`,
  `adapterDefinitionSha256@48:uint8[32]`, then tuple ID, PowerShell version, and
  PSReadLine version bytes concatenated at offset 80; exact size is checked
  `80 + tupleIdBytes + psVersionBytes + psReadLineVersionBytes`. Tuple ID is
  1..128 ASCII bytes matching only `[A-Za-z0-9._-]`. Each version is 3..43 ASCII
  bytes in canonical `System.Version.ToString()` form: two through four dot-
  separated decimal components, each `0..2147483647`, with no sign or leading
  zero except the single digit `0`; parsing and `ToString()` reproduce the bytes
  exactly. Edition, both versions, tuple ID, and definition SHA-256 equal the one
  exact allowlisted manifest record. `channelGeneration` equals the nonzero
  semantic-channel generation issued in this `SemanticDecision`. The only
  capability bits are `TrustedRootBoundary=0x1`,
  `TrustedFilesystemCwd=0x2`, `ExactAcceptedInput=0x4`, `FollowControl=0x8`, and
  `InsertControl=0x10`; any other bit is fatal. Either control bit requires
  TrustedRootBoundary and FollowControl additionally requires
  TrustedFilesystemCwd. These are provisional tuple claims only: runtime
  publication intersects them with the accepted startup join, current policy/
  state/binding gates, and a matching real authenticated Probe Success.
- `RootReadReady=2` is exactly `payloadVersion@0:uint32=1`, `flags@4:uint32`, nonzero
  `readCycle@8:uint64`, `completedAcceptedSequence@16:uint64`,
  `cwdKind@24:uint32`, `cwdBytes@28:uint32`, and cwd bytes at offset 32. The only
  flags are `CompletedAccepted=0x1`, `PriorStatusAvailable=0x2`, and
  `PriorCommandSucceeded=0x4`; exact size is checked `32 + cwdBytes`.
  `readCycle` starts at one, increments once per owned root-wrapper entry, and
  never wraps. Completed sequence is zero iff CompletedAccepted is absent;
  otherwise it equals the still-pending prior AcceptedInput semantic-frame
  sequence and consumes that candidate exactly once; intervening ControlResult
  frames do not change that relation. PriorStatusAvailable is forbidden without
  CompletedAccepted, and PriorCommandSucceeded requires both other flags.
  `cwdKind` is exactly `UnavailableOrNonFilesystem=0|WindowsFilesystem=1`: kind 0
  requires zero cwd; kind 1 requires 1..131,072 strict-UTF-8 bytes containing the
  byte-identical canonical Windows drive/UNC logical path produced by the
  Location-identity canonicalizer. No other provider/cwd spelling is accepted.
- `AcceptedInput=3` is exactly `payloadVersion@0:uint32=1`, `cwdKind@4:uint32`,
  `readCycle@8:uint64`, `commandBytes@16:uint32`, `cwdBytes@20:uint32`, command
  bytes at offset 24 followed immediately by cwd bytes; exact size is checked
  `24 + commandBytes + cwdBytes`. `readCycle` matches the RootReadReady cycle
  that opened this exact still-active ReadLine call even when intervening
  ControlResult frames exist. Command length is `0..1,048,576` bytes; cwd uses
  the same kind/strict-UTF-8/131,072-byte rules and is the filesystem cwd captured
  at acceptance. The maximum valid AcceptedInput payload is therefore exactly
  1,179,672 bytes and remains below the semantic-frame payload cap. Command bytes
  are strict UTF-8 for the exact returned .NET `String`; empty and multiline CR/
  LF/HT are allowed, while BOM, NUL, unpaired UTF-16, every other C0, and every C1
  control are rejected. The payload never normalizes, trims, or reconstructs
  command text.
- `ControlResult=4` is exactly `payloadVersion@0:uint32=1`,
  `controlKind@4:uint32 Probe=1|Follow=2|Insert=3`,
  `status@8:uint32 Success=0|RejectedState=1|Expired=2|BindingChanged=3|
  BufferChanged=4|ShellOperationFailed=5`, `flags@12:uint32=0`, echoed
  `controlSequence@16:uint64`, echoed `operationId@24:uint8[16]`,
  `cwdBytes@40:uint32`, `reserved@44:uint32=0`,
  `bufferUtf16CodeUnits@48:uint64`, `bufferSha256@56:uint8[32]`, then cwd bytes
  at offset 88; exact size is checked `88 + cwdBytes`. It matches the sole
  awaiting RSTCTL01 kind, nonzero sequence, operation ID, epoch, and embedded
  deadline. All six status values are valid for all three control kinds. Probe
  has zero cwd/buffer result fields for every status. Follow Success has nonempty
  canonical cwd and zero buffer fields; every non-Success Follow has
  `cwdBytes=0`, `bufferUtf16CodeUnits=0`, and an all-zero buffer digest. Insert
  Success has zero cwd, `bufferUtf16CodeUnits` in `1..1,048,576`, and SHA-256 over
  the exact complete post-insert .NET buffer serialized as raw UTF-16LE code
  units with no BOM or terminating NUL; every non-Success Insert has those same
  three result fields zero. Result-data zeroing never applies to the echoed
  sequence/operation ID: both match the current control operation for every
  kind/status. RejectedState, Expired, BindingChanged, and BufferChanged may be
  emitted only before any shell-effect API call and are mechanically no-action;
  ShellOperationFailed means an API call was attempted and is conservatively
  effect-ambiguous. A BeginRead still incomplete after the sole handler
  `WaitOne` emits no result. A ControlResult
  consumes one ordinary semantic sequence only after its bounded semantic write
  succeeds.
- Each epoch creates one nonzero 64-bit `BCryptGenRandom` operation namespace.
  Control `operationId` is exactly namespace little-endian followed by that
  operation's nonzero `controlSequence` little-endian, giving 16 bytes, nonzero
  identity, and no reuse without unbounded state; semantic nonce plus operation
  ID is the cross-epoch identity. Semantic sequence 1 is the sole Handshake;
  every later semantic frame is exact +1 with no wrap. AcceptedInput enters
  RunningRootCommand; only its matching later RootReadReady returns idle and
  publishes applicable cwd. There are no generic Activity or Cwd message kinds.
- The fake consumes exact
  `Plugins/Terminal/ShellAdapters/TerminalSemanticProtocol.v1.json` and the
  generated C++/embedded-PowerShell constants from the two checked-in Generated
  paths above; generator `-Check` and byte-identity source-contract tests prevent
  drift. Plan 5 compiles them as identity-bound `Terminal.dll` resources. The
  closed fake corpus covers every kind's valid min/max cases, astral/multibyte
  text, unknown kind/status/flag, nonzero reserved, bad version/manifest identity,
  illegal capability combination, offset/length overflow/underflow/trailing
  bytes, invalid UTF-8/control/path/version, `1 MiB + 1`, cwd `128 KiB + 1`, full
  frame `2 MiB + 1`, sequence/read-cycle wrap, wrong completed sequence, and wrong
  control ID/kind/data combination.

## Tests

- Existing title-only tabs and Preview behavior before Terminal wiring.
- Icons set/change/clear/title persistence, LTR/RTL geometry, overflow/reorder, hit/close, fallback, DPI/theme/high contrast, UIA name.
- Folder/Terminal/Preview source/host resolution before and after swap, menu/function-bar fallback to active selected content, hidden-host unzoom, cancel/no process, diagnostic/settings actions/tab count, success focus, Preview no steal, and repeated independent sessions through both `Ctrl+Alt+T` and `Alt+7`.
- Stable-command migration: embedded `cmd/pane/openCommandShell`, no-default
  `cmd/pane/openCommandShellWithProfile`, explicit no-default
  `cmd/pane/openExternalCommandShell`, distinct Ctrl+Enter/Ctrl+Shift+Enter IDs,
  Ctrl+Space current-directory insertion, and absence of pseudo command-line
  HWND/layout/launcher.
- Directory Edit versus file/Alternate/Edit With; focused-item-only insertion;
  selected compatible fake Windows terminal versus selected busy/incompatible/
  capability-unavailable no-new-tab rejection versus selected Folder/Preview
  one-shot fake hidden open; hidden-terminal non-targeting; same/different cwd
  leaf/full, always-full, no Enter, partial/TUI/password/follow rejection,
  source/tab generation cancellation, fake Windows/UNC success, every WSL
  target rejecting with truthful status and zero bytes/no new tab, cross-family/
  distro/provider rejection, quoting/metacharacters, and settings link. The
  production Plan-4 plugin separately proves `InsertCapable=false`, direct
  capability-unavailable, hidden `IntegrationDisabled` before process, zero
  bytes/preference, and no visible tab. Assert the exact `416/8` layout/offsets
  and all mode cross-fields; cover initiating source equal to and different
  from the target terminal's immutable original source, `SwapPanes` both before
  capture and between capture/revalidation, and Ctrl+Space proving with the fake
  that only the copied/revalidated initiating source location reaches
  `CurrentDirectoryFull` while target follow location and authenticated cwd are
  deliberately different. Authenticated product translation/quoting/admission
  is deferred to plan 5.
- Fixed Folder/Preview order, Terminal reorder only, close confirmation/cancel identity, MRU/Folder restoration, all `closeOnExit` modes after quiet, incomplete-final-snapshot last-safe/empty status with permanent automatic-removal suppression, disabled-until-quiet restart and exact precedence, complete/incomplete late-quarantine recovery status, natural-exit visible versus explicit-close content-free quarantine, title sanitize, max-tab lowering without eviction, 32 held service slots/denied 33rd with no process, rapid retiring/quarantine admission closure, and late-quiet reopening. Cover the exact callback vtable/removal enum, chooser/plugin-Close legal states, automatic `Exited -> Quiet -> ViewRemovalRequested` order, deferred-not-reentrant host retirement, and no RequestClose/settings/state inference. Deterministic barriers at every arbiter transition prove UserTab allow/cancel, approval-first, cancel-first, forced takeover, natural exit before/at/after quiet, every `closeOnExit`/snapshot/quiet result, prompt generation invalidation, stale/duplicate reply suppression, and exactly one owning callback or forced path. Background close-glyph tests capture stable ID/generation, select/reveal/layout/focus, revalidate after each destructive callback, and issue UserTab only for that still-selected interactive-visible view; reorder/removal/selection races before the request leave their winner untouched, while selection success followed by Busy/Retry/Cancel never rolls back to the former tab and no hidden prompt exists.
- Exact close/shutdown matrix: the single plugin retirement-intent arbiter is the only source of CloseApproved/removal ownership, and the shared host `{instanceId,hostViewGeneration}` claim admits exactly one hide/remove/null-drain/Close/release/MRU mutation when interactive approval, queued removal callback, `WindowClosing`, and `ApplicationShutdown` race. Only `UserTab` can prompt/emit `CloseApproved`; `WindowClosing` forces only its FolderWindow's objects and leaves another window/service admission usable; `ApplicationShutdown` proves stable-ID request-before-hide/drain/close/release ordering for terminal objects and concurrent administration retirement before export. Cover callback losing to forced teardown, host-generation reuse, duplicate/stale/wrong-instance/unknown reason, administration absent/idle and Prepare/Arm/pre-CAS/in-CAS/post-CAS/Finalize/compensation/result states, cancellation-wins/loses, exact Failed/Committed finalization, compensation acknowledgement, retired-page suppression, result ownership/retry, `ERROR_BUSY` Close breach, callback-clear/Close/release, zero-terminal/retiring export start, manager/service no-later two-/five-second deadlines, immediate/deferred quiet, target/pump survival, at-most-50-ms generation-checked polling, exactly one continuation, and no UI wait or premature `DestroyWindow`. Cover duplicate/stale begin, `WM_CLOSE`, target, view/callback generation, export, and late poll. Prove the sole process-exit-retention case is an externally held disconnected UIA provider/range with zero captured calls and passive immutable snapshot/ledger state: it returns `UIA_E_ELEMENTNOTAVAILABLE`, touches no host/service, racing Release may permit unload, and normal refresh/disable cannot use the vote; all active work is fatal at five seconds. Source checks reject an EXE service accessor/direct terminal symbol and every wider mapped-busy escape.
- Normal plugin lifecycle: disable with Starting/Running/Exited tabs leaves every view/object/session usable and starts no retirement; idle re-enable reuses that generation; accepted refresh closes every new host-mediated route while old tabs continue keyboard/render/UIA/state/explicit-close/natural-removal behavior. Cover cancel-before-begin, duplicate/stale/unissued/concurrent refresh generations, old tabs outliving both five-second clocks, last-tab removal before/after configuration settlement, pending chooser/open/insertion cancellation by admission generation, both distinct nonfatal blocked statuses, host Retry, desired-availability churn, replacement failure with no synthetic tab, old-handle-reset-before-one-replacement ordering, and app shutdown superseding every state. Assert no normal-refresh `RequestClose`, hide, callback clear, forced view removal, fatal seam, retention vote, or second module/service generation. With barriers at disable/refresh admission and publication, prove an existing selected interactive-visible live frozen-config view keeps safe Paste and its still-eligible pending unsafe-Paste prompt/token with unchanged paste-policy generation; the transition itself does not cancel it, but the independent 60-second/selection/activation/visibility/input-owner rules still do. Outstanding OSC 52 ask/allow tokens are revoked and scrubbed before the same transitions publish and disabled OSC 52 remains effective deny.
- The broker matrix exercises all 12 table cells plus duplicate/stale kinds, every pairwise and three-way Close/Paste/OSC arrival order, exact interaction-generation and full identity mismatch, missing/failed posted payload, selection-away/back, hide/show, foreground deactivation/reactivation, deadline just-before/equal/after, newer replacement failure, reorder preserving only an unchanged selected view, removal/HWND reuse, and teardown drain. Eight background/inactive OSC asks plus `+1` are each denied/scrubbed before counters, leave zero retained tokens, never select/reveal/focus, and only coalesce content-free status. History and context open only at `None`; Close dismisses them, OSC denies while open, and plugin-text-input Paste stays within its owned surface.
- Unsafe-paste pane/UI tests prove `Terminal.dll` alone owns the frozen preview/classification and opaque token; the host receives no clipboard bytes/preview/token and performs no classification/admission. Exercise the strict 60-second deadline, accessible modal Allow/Cancel, initial Cancel focus, confined Tab order, focused Enter/Space, Escape, conditional focus restoration, selection/hide/deactivation/input-owner cancellation with no resurrection, reorder/zoom/`SwapPanes`, newer-paste supersession, stale UI actions, and exact view/incarnation binding. User close, automatic removal, window close, Restart, and application shutdown must revoke/scrub before plugin child teardown, while generic disable/normal refresh alone preserve a still-visible/eligible prompt and allow it to reach the core's one final current-mode/cap/queue check. No host generation change may cause a clipboard reread, reclassification, token copy, paste replay, or canceled prompt restoration.
- OSC 52 UI tests prove a visible ask has the strict 30-second deadline, never steals focus, raises one polite UIA notification, and accepts only exact-token mouse/UIA Invoke or bar-scoped `Alt+A`/`Alt+D`. Selection/activation/ownership/policy/lifecycle loss, close/paste supersession, expiry, hidden UIA access, duplicate/stale/wrong identity, and teardown deny/scrub exactly once; canceled providers return `UIA_E_ELEMENTNOTAVAILABLE`, and exactly one successful Allow may reach the final clipboard-write gate.
- Profile-launch recorder cases are mandatory in PaneProfiles evidence: exact Unicode creation flags, non-ASCII and duplicate-key round trips, valid/malformed hidden drives, canonical OS/permitted-integration overwrites, rejection of bootstrap target/semantic nonce/control-pipe environment entries, deterministic sort, exact double-NUL empty block, both inclusive caps and +1, and zero process on every rejected environment/command line. Windows rows prove exact held/revalidated nonnull `lpApplicationName`, fixed unquoted basename argv0, exact token round trip under hostile current/PATH/app-local fixtures, complete 32,772-unit pwsh profile ID acceptance, and no truncation/hash. The shared identity corpus accepts exact component, 33,037-unit location-key, and 33,040-unit display maxima and rejects +1/malformed forms without mutation.
- Plugin-owned tab admission: limits 1 and 32 plus exact +1 quota HRESULT; independent complete `PhysicalHostKey`s; simultaneous/reentrant opens; synchronous validation failure before reservation; hidden uncommitted record with no tab/zoom/selection change; callback null-drain/Close/release on quota and other Open failures; post-Open host-commit failure; successful chooser/diagnostic/Failed/Exited/teardown-timeout retention; live lowering without eviction; exactly-once release despite internal quarantine; zero-key erasure; and release/reopen. Assert no host setting read/precheck and no rejected tab, child, process, callback publication, or leaked slot.
- Logical source follows current `SwapPanes` content while physical terminal host stays; two-window key collision, tombstoned detach, destroyed-window late result, late generation/HWND reuse.
- Full source-compatible
  `ResolveDefault|ForceChooser|ExactProfile` cross-field/profile-precedence/
  profile/purpose-specific one-shot eligibility fence, WSL no-read/no-write
  memory boundary, and disabled-memory matrix against the in-memory
  `ITerminalPreferenceStore`, including ForceChooser with one compatible
  profile, mode-preserving Retry, diagnostic Choose-profile transition,
  336-byte x64/ARM64 Open layout with exact offsets including
  `openPurpose@300`, canonical post-header field order, all-zero reserved words,
  and independent source/launch locations; real loaded-DLL plus checked-in
  golden-oracle tests cover `InteractiveTab|PendingPathInsertion`, every
  purpose/second-argument/selection/requested-profile/location cross-field,
  unknown enum values, and ordinary pre-start tracking versus fixed Directory-Edit
  launch, same-generation explicit-follow requeue, generation-bearing
  OpenSiblingRequested stale/duplicate/wrong-mode/order/reentrant cases,
  `windowsDefaultProfileId` auto/pwsh/Windows PowerShell/cmd, deterministic source-derived exact matching WSL with no preference lookup/write,
  incompatible remembered choice, canceled reads, stale generations, write
  failure, and late completions; pwsh discovery
  validation/dedupe/version/stable ID/missing/replacement matrix; and all
  location/quoting/path-disappearance cases. The cmd local/plugin-backed/UNC
  matrix covers the embedded script resource/extraction/reopen hash; the exact
  `["cmd.exe","/E:ON","/V:OFF","/S","/K",
  "call \"<verified-extracted-script-path>\""]` tokens and closed two-pass path
  encoder; one real-cmd invocation for spaces, Unicode, `%`, `!`, `^`, `&`, `|`,
  `<`, `>`, and parentheses; and pre-create rejection when the path cannot
  encode. It covers the exact protected pipe name/security/mode/flags/ready-
  before-create contract; root PID; all eight exact fixed environment names and
  value grammars; removal of mixed-case inherited variants, insertion of only
  canonical names, permanent deletion on every script path with no restoration,
  and duplicate rejection; local `cd /d` versus UNC `pushd`; the prelaunch-free
  drive snapshot; exact logical root/tail parsing; `DRIVE_REMOTE`; bounded
  `REMOTE_NAME_INFO_LEVEL` reverse mapping with initial 4,096-byte buffer, one
  exact-size retry, 262,144-byte cap, pointer/NUL validation, and complete/root/
  tail ordinal-ignore-case comparison;
  direct-expansion metacharacter literalness; and reserved/private cleanup before
  Ack/prompt. Ack cases cover exact fixed-case/fixed-order
  `RSTCMD1;CAP=...;DRIVE=...\r\n` grammar, maximum-valid shape, 512/513-byte
  assembly boundaries, and `ERROR_MORE_DATA`; every field/
  delimiter/status-drive cross-field; chunk assembly through clean EOF; and
  empty, partial-at-EOF, extra, non-ASCII, noncanonical, second-line, over-cap,
  first-valid-then-second/trailing, and first-valid-but-kept-open failures. They
  cover capability, wrong PID/nonce/operation/generation/digest, replay, target-
  identity revalidation, local no-cwd-serialization, and UNC share-root/empty-
  tail/deep-tail plus case/slash/dot lexical normalization. UNC cases include
  zero/failing `GetLogicalDrives`; 4,096 bytes, exact 262,144-byte cap and +1,
  zero/non-growing/inconsistent retry sizes, malformed/null/misaligned/out-of-
  buffer/nonterminated pointers, local/SUBST drives, DFS/
  provider/share aliases, empty/zero-or-one-leading-slash remainder, drive/UNC/
  double-leading/root-escape rejection, and preoccupied/reused drive letters.
  Cleanup cases
  repeat the 4,096/262,144/+1/malformed `WNetGetConnectionW` corpus and prove no
  cancel while the child lives; already-gone success; exact unchanged-mapping
  one-shot cancel and fresh absence recheck; and open-files/API/root-exit/
  remap/remove/changed-target-reuse races that never force or cancel another
  mapping. A limitation probe records indistinguishable same-user same-target
  letter reuse rather than claiming ownership proof.
  Exercise integration enabled/disabled,
  success/failure/root exit, AutoRun cwd change/exit/hang/suppression, strict-
  before/equal/after timeout, unrelated-settings/toggle races, zero PTY/user-byte/
  app/DOSKEY-history interleaving, no failed preference, and complete mutable
  launch-state cleanup without claiming immutable cmd-buffer erasure. One
  explicit limitation probe documents that trusted same-root AutoRun with the
  correct capability/nonce is indistinguishable and outside the threat model;
  wrong PID/capability/nonce/operation/generation/digest and PTY replay still
  fail closed. Cmd proves zero ongoing nonce/history/follow/insertion/activity/
  trusted-cwd state after its launch bootstrap is erased.
- The injected semantic-wire corpus proves canonical JSON/generated C++/embedded-
  PowerShell byte identity and only the four exact payload kinds/offsets/enums/
  flags/reserved/cross-fields, the two checked-in Generated paths, and generator
  `-Check`. It covers every min/max, astral/multibyte text, strict UTF-8/control
  rejection including all-text BOM/NUL/unpaired-surrogate and tuple/path-control
  rejection with HT/LF/CR permitted only for command, tuple/manifest identity,
  exact `payloadVersion`, 3..43-byte canonical two-to-four-component
  System.Version round trips, channel-generation equality, capability dependencies
  and no Follow/Insert publication before real Probe Success, sequence/read-cycle/
  accepted-sequence and control-operation relations, cwd/result zero/nonzero
  rules including all 18 control kind/status combinations with echoed identity,
  intervening ControlResult before the still-pending AcceptedInput completion,
  overflow/trailing/unknown values, command `1 MiB + 1`, cwd `128 KiB + 1`, exact
  1,179,672-byte maximum AcceptedInput payload, and frame `2 MiB + 1`; it
  rejects struct-cast/padding/TLV and generic Activity/Cwd kinds, requires C++
  `Common/StringConversion.h` strict APIs and embedded-PowerShell
  `System.Text.UTF8Encoding(false, true)`, and rejects a local/fallback converter.
  This is fake-
  seam proof only and creates no Plan-4 product resource/capability.
- The separate PowerShell startup matrix covers local and UNC launches for real
  Windows PowerShell 5.1, every shipped pwsh tuple, and the deterministic fake:
  exact four-token fixed-basename/`-NoExit -Command` argv and quoting round
  trip; canonical single-quoted-literal encoder/hostile apostrophe and invalid-
  UTF-16 rejection; exact wrapper text and AST including fixed `*>&1`; six named
  script parameters; `-ErrorAction Stop`; sole typed two-property finalize
  object with exact first PSTypeName/property order/sentinel/scriptblock; profile
  variable collision/read-only/AllScope state; finalize-object property/type/
  scriptblock mismatch and duplicate/throw/partial-Write finalizer; injected
  named disposal error proving idempotent output-silent disposal and no Ack
  repair; exit 125/126
  for named exception, invalid finalize object, or any extra success/
  warning/verbose/debug/information result; one output-silent finalizer and one
  complete synchronous message-mode CommitAck Write with no Flush/retry/later
  pipe operation, with server gates opening only after a full valid read; normal profiles exactly
  once before the wrapper/script; profile
  cwd override restored; slow/hanging/throwing/exiting profiles; execution-
  policy rejection without bypass; parameter rejection; protected pipe DACL/
  noninheritance/client-root-PID match; exact `RSTBST01` HMAC frame layout,
  kind/direction/sequence/generation/operation/length/UTF-16/cap rules; and the
  sole `Ready -> RootRequest -> RootResult -> SemanticDecision -> Prepared ->
  ReleasePrompt -> ReleaseAck -> CommitPrompt -> CommitAck` order. Exercise
  Local/Unc module-qualified literal `Microsoft.PowerShell.Management\Set-Location`
  plus module-qualified `Get-Location` rooting and profile-shadowed aliases/
  functions that try to emit output,
  wrong/failure cwd, every malformed/spoof/replay/trailing/oversize outcome,
  connection and every receipt just before/at/after the common deadline,
  production `TerminalOnly` and deterministic
  `SemanticEnable -> Prepared(SemanticUnavailable)`, plus fake-only
  `SemanticReady`, false-ready, and decision/outcome/generation mismatches. Real
  Windows PowerShell 5.1 and every shipped pwsh prove no adapter hook, semantic-
  pipe connection, handshake, or semantic capability in Plan 4. Settings-change
  races cover Decision construction/delivery, fake semantic-pipe creation and
  fake client connect/handshake enqueue, Prepared Write, Running CAS,
  CommitPrompt, and CommitAck. Only the injected fake exercises
  `StartupSemanticJoin` with Prepared-first,
  handshake-first, exact conjunction, either slot alone through timeout,
  duplicate/conflict/gap/wrong-identity/oversize/queue failure, cancellation/
  root-exit/Restart/shutdown scrub, TerminalOnly/Unavailable with zero slots,
  forbidden handshake with either, pre-child two-slot reservation/exhaustion,
  and handler lock/order probes proving only reserve/fill/scrub plus one posted
  continuation under the session lock with no wait/parse/pipe/callback/model/
  writer/host-lock work. Live-disable `RevokedTerminalOnly` tests the production
  Unavailable outcome with zero semantic connection and the fake Ready cleanup
  outcome with no required handshake, exact-root connection to the retained
  no-operation fake semantic/control cleanup pipes, scrubbed/ignored old records,
  both-pipe close/cancel at CAS, TerminalOnly/current-generation CommitPrompt,
  fake owned-mapping removal/inert latch before ack, and missing-Prepared timeout.
  The fake also stages a valid Handshake before its receive deadline while child
  observation lands at/equality-after the 50-ms write deadline, pairs it with
  Unavailable, and proves nonfatal discard/revocation; malformed/mismatched/extra
  data and Handshake plus TerminalOnly remain fatal.
  Assert no semantic nonce/control pipe before
  `SemanticEnable`; no target/semantic secret in environment or initial argv;
  no second decision; post-decision disable yields Running terminal-only with an
  inert wrapper; post-decision enable yields terminal-only plus Restart required;
  no user input/trusted activity/profile write/capability before `CommitAck`;
  Running CAS precedes `CommitPrompt` and therefore first-prompt evaluation.
  Validate the exact CommitPrompt disposition/generation payload: unchanged
  Enable/Ready/joined Active; same-generation Enable/Unavailable TerminalOnly;
  original TerminalOnly/Prepared-TerminalOnly at equal/greater generation;
  either valid Enable outcome after latched disable TerminalOnly only at its
  current strictly greater generation regardless of later re-enable; minimal
  retained Decision generation/Prepared outcome; and fatal decrease, wrap, or
  every other pair. The unchanged-generation exact
  `SemanticEnable -> Prepared(SemanticUnavailable) ->
  CommitPrompt(TerminalOnly) -> CommitAck` case must reach usable Running with
  ordinary input/profile memory and zero semantic capability. Exercise every
  disable timing for both Ready and Unavailable: before CommitPrompt
  construction, after construction/before send, after send/before child receipt,
  between finalizer Ack Write/server receipt, between wrapper validation/
  finalizer invocation, during Ack acceptance, and immediately after acceptance.
  Also prove a post-CAS disable closes/cancels both pipes, completed control EOF/
  error or semantic-write failure latches the fake interposer inert, and no
  emission crosses the fence even when a built prompt said SemanticActive. The
  injected seam proves exact
  `ReadData|ReadAttributes|WriteAttributes|Synchronize`/no-
  `WriteData` control DACL granted mask and async anonymous noninherited message-
  mode client,
  forbids `FlushFileBuffers` and synchronous pipe Write/Flush, and exercises final
  Plan-5/6 OVERLAPPED writer admission, fixed-protocol-reserve descriptor/one-
  byte inert placeholder and final queue sequence before Write, AwaitingResult/
  complete-result-identity publication before same-descriptor activation with no
  second capacity/admission/sequence or priority bypass,
  exactly one wake/no retry, one validated-frame stage with nonwaiting
  ordinary-wrapper `PollControlTransport(NoWait)` stage/rearm but no dispatch,
  and handler-only owner/binding/
  epoch/lifecycle-revalidated `PollControlTransport(HandlerWaitOnce)` with one
  stage-first consume with zero wait/EndRead on the newly armed read and, only
  when no frame is staged, one nonalertable `AsyncWaitHandle.WaitOne(250)`. Cover
  NoWait completing/staging/rearming during an ordinary wrapper immediately
  before the wake and the handler dispatching that stage exactly once without
  inspecting the new read; also cover read completion
  during that wait and just-before/equal/after the embedded deadline, unsignaled
  timeout with child-side no-action/no-emission/no-revoke, plugin-side operation
  timeout/revoke versus no-operation physical-press liveness and at-most-250-ms
  root-shell block, close/disable EOF/error waking the wait, one-shot rearm-before-
  dispatch, physical-key consumption, and mapping replacement receiving at most
  one raw wake before missing-result retirement. Also cover user-before-reserve
  cancellation, fixed-reserve byte/descriptor boundary and +1, later-user FIFO
  barriers before/during/after submit/completion/state-publication/activation,
  an immediate child result before the completion callback returns, pre-submit zero-byte
  tombstone, submitted-failure inert fence, ambiguous-input hold,
  exact cancel with handle retention through
  one quiet-bounded event drain/terminal observation, normal-success release of
  only per-write storage while persistent pipe/operation/input hold awaits the
  matching result; matching authenticated Success and Follow/Insert pre-effect
  RejectedState or BufferChanged each returning Idle/advancing sequence; pre-
  submit no-action returning Idle without advancing; Probe non-Success, Expired,
  BindingChanged, ShellOperationFailed, and failed/revoked/ambiguous outcomes
  becoming terminal-only with no later control operation; and module-pin behavior
  without packaging a production adapter.
  At every disable timing, matching Ack remains acceptable by incarnation/
  deadline rather than
  live policy generation, bootstrap cleanup/input/integration-independent
  profile memory completes, zero semantic capability/old event survives, and
  disable alone never retires the terminal.
  Race Close/WindowClosing/Restart/root exit/ApplicationShutdown and deadline
  just before/at/after ReleaseAck write, server receipt, validation, Running CAS,
  CommitPrompt, and CommitAck; prove one CAS winner, no never-Running prompt,
  pre-CAS failure with no Running/memory, and post-CAS failure retiring an
  already-Running incarnation with every startup gate still closed. Mutable buffers are zeroed/references dropped;
  and no test claims secure physical erasure of immutable command-line/.NET
  strings.
- Profile-memory tests assert that cmd local, validated-plugin-backing, and UNC
  launches write `CmdLaunchRootedRunning` only after the complete matching
  pipe/PID/correlation/target/mode result, UNC reverse mapping when applicable,
  launch-state cleanup, and the same incarnation's `Running` CAS; ordinary
  Windows PowerShell/pwsh writes only after a matching post-
  Running `CommitAck` and never in the Running-to-Ack gap; the injected hidden
  `PendingPathInsertion` seam retains no durable choice at `Running`, Ack,
  `ReadyToCommit`, host preparation, or plugin Commit and writes exactly once
  only after the allocation-free/no-fail host finalize; every Abort, failure,
  timeout, stale result, and teardown writes zero; production Plan-4 hidden Open
  writes zero through `IntegrationDisabled`; and WSL writes zero both before and
  after `Running`.
- WSL enumeration commits just before 3 seconds, exactly at 3 seconds, and just after it: only the strict-before case may publish results, exact equality/after publish timeout/Retry. Prove cancel wins simultaneous cancel/result/deadline, losing callbacks cannot publish, overrides above 10 seconds clamp to 10, `0`/negative return `E_INVALIDARG` without changing the active/default deadline or starting work, and cmd/Windows PowerShell/pwsh publication/use never waits.
- WSL launch uses the exact reopened System32 executable and five-token `--distribution <exact-distro> --cd <absolute-linux-path>` argv, preserves the configured default user/shell, and adds no `--exec`/`--user`/`-e`/wrapper/bootstrap/extra process/fallback. Fake-argv cases preserve spaces, quotes, backslashes, shell metacharacters, option-like path components, and non-BMP text exactly, while empty/leading-dash distro and other invalid cross-fields start no process. Identity and launch bounds are tested separately: the full 32,768-unit Linux component/33,037-unit key/33,040-unit display form can remain valid identity, while the exact five-token quoted command accepts 32,767 total units including NUL (32,736 quote-free path units for distro `D`) and rejects +1 before preflight/process with no truncation/hash/indirection. Generation-bound preflight tests cover exact `\\wsl.localhost` directory identity, handle retention through process/startup admission, catalog/path staleness, cancel before/during/after open, zero process on failure, and just-before/equal/after the fixed three-second deadline including irrevocable failure before `CancelSynchronousIo` and late handle release. Launch/environment/transport fakes and source contracts prove no WSL-specific argv/environment/`WSLENV` delta, no PTY/Linux-profile/staged-script injection or guessed `/mnt/<drive>` path, and no WSL semantic nonce/epoch. State/preference/UI tests hold `RequestedUnverified` permanently, reject every attempted `Verified|Diverged` transition, keep the raw terminal usable, expose no app-history/follow/activity/authenticated-cwd/profile-memory capability and allow only exact-distribution full-native-path explicit insertion, perform no preference read/write, and truthfully handle a post-Running launcher error without fallback or false memory; they do not require every asynchronous `--cd` failure to precede `Running`.
- Keyboard/UIA coverage for pane tabs, chooser, diagnostic, the one broker surface, plugin-owned unsafe-paste modal actions, non-focus-stealing OSC 52 `Alt+A`/`Alt+D` and UIA Invoke, exit, and restart actions; hidden/canceled broker providers are unavailable, the host never exposes preview/token metadata, and the plan-3 document provider remains unchanged.
- `IID_ITerminal` factory/open/child ownership/state/callback/close paths pass,
  including exact Open layout/enum/profile cross-fields and the plugin context
  request route; `IID_IViewer` is rejected; module quiet remains false through
  each live tab/session/callback. It becomes true after all ordinary and
  external gates release, or the exact passive disconnected-UIA process-exit
  retention vote safely owns the remaining external COM tail. Existing Preview
  cases stay green; former external launch mechanics pass under the renamed
  explicit command.

The hidden-new-session routing matrix fixes the initiating Folder as
`originalSource`, source/launch location, profile-memory key, follow source, and
Restart fallback; the opposite Folder/Preview supplies only physical host and
eligibility. It covers different paths, selected Preview, navigation at every
revalidation boundary, and `SwapPanes` before capture, before Open, and between
Ready and fake Commit.

The loaded-DLL/golden ABI matrix requires the exact six-slot callback order,
exercises noncoalescing slot-2 `TerminalPendingOpenResolved` with every result,
and fails on the former five-slot layout, slot substitution, coalescing,
duplicate/stale generation acceptance, or architecture drift.

## Verification

```powershell
$expectedRepositoryCommit = (& git rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0 -or $expectedRepositoryCommit -cnotmatch '^[0-9a-f]{40}$') { throw 'Cannot resolve exact repository commit.' }
$expectedEngineLockSha256 = (Get-FileHash -LiteralPath '.\External\terminal-engine.lock.json' -Algorithm SHA256).Hash.ToLowerInvariant()
.\build.ps1 -ProjectName DxUiTests -Platform x64 -Configuration Debug
.\build.ps1 -ProjectName TerminalTests -Platform x64 -Configuration Debug
$paneProfilesNativeManifest = .\Tools\Run-TerminalNativeTests.ps1 -Slice PaneProfiles -Platform x64 -Configuration Debug -EvidenceRoot .\Specs\TestRuns -PassThruManifest
.\Tools\Verify-TerminalNativeEvidence.ps1 -Manifest $paneProfilesNativeManifest -ExpectedSlice PaneProfiles -ExpectedRepositoryCommit $expectedRepositoryCommit -ExpectedEngineLockSha256 $expectedEngineLockSha256
.\build.ps1 -ProjectName RedSalamander -Platform x64 -Configuration Debug
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter terminal_ -TimeoutMultiplier 2
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter preview_ -TimeoutMultiplier 2
```

Expected: the exact PaneProfiles `TerminalNative` manifest verifies with nonzero unskipped plugin-ABI/DxUi/profile/working-directory/path-insertion/plugin-owned-tab-admission/normal-refresh/launch-environment cases; all Commands cases pass; no title-only geometry drift; quota rejection leaves the prior zoom/selection with no tab/child/process and release reopens capacity; exact focus/order/icon/swap/zoom/chooser/diagnostic/Edit/insertion contracts hold; every profile uses the canonical bounded explicit UTF-16 environment and exact creation flags; complete profile/location/display identities survive without truncation/hash, Windows launches use fixed basename argv0 plus exact nonnull `lpApplicationName`, and WSL's identity and five-token command boundaries reject +1 independently; cmd remains ongoing terminal-only after its all-Windows-location launch state is erased; the per-view retirement arbiter plus shared host removal claim produces one owning callback/forced path and one host mutation under confirmation/natural-exit/forced-close races; WSL launches remain usable at the exact distro/path with the five frozen tokens and configured default user/shell, permanently report `RequestedUnverified`, create no integration state/input or profile memory, and truthfully disable history/follow/activity/cwd capabilities while allowing only exact-distribution full-native-path explicit insertion; plugin-owned unsafe-paste UI exposes only the frozen request and keeps safe Paste plus an otherwise broker-eligible pending unsafe Paste usable through generic disable/normal refresh, while independent broker cancellation and OSC revoke/scrub remain effective; disable/re-enable preserves the retained generation, accepted refresh never force-closes a live Terminal and blocks only new host-mediated actions, both blocked outcomes and Retry are truthful, application shutdown supersedes refresh, Ctrl+Alt+T/Alt+7 open embedded Terminal; the pseudo command line is absent; and external launch remains available only through its explicit no-default command.

The same expected gate requires the exact one-shot, generation-bound profile-
memory contract: cmd local/validated-plugin-backing/UNC writes only through
`CmdLaunchRootedRunning` after its complete matching pipe/PID/correlation/target/
mode result, applicable prelaunch-free-drive and bounded complete/root/tail UNC
proof with only the nonsecret owner tuple retained, launch cleanup, and the same
incarnation's `Running` CAS; ordinary PowerShell/pwsh writes only after the
accepted matching post-Running `CommitAck`; the injected hidden
`PendingPathInsertion` seam writes only after that Ack plus successful
allocation-free/no-fail host finalize, while Plan-4 production resolves
`IntegrationDisabled` and writes zero; every earlier/failing/stale/aborted/
teardown outcome writes zero; and WSL never reads or writes profile memory.

The same expected gate requires nonzero unskipped `terminal_interaction_broker_` coverage: one content-free per-view slot, exact `None|CloseConfirm|UnsafePasteConfirm|Osc52Ask` state and Close-over-Paste-over-OSC table, stable-ID background close select/reveal/focus/revalidate with no invisible prompt or post-selection rollback, strict 60-second Paste and 30-second OSC equality expiry, immediate background/inactive OSC deny/scrub before counters and without focus, selection/activation cancellation with no resurrection, exact modal/default/restore/Escape/Enter behavior, OSC `Alt+A`/`Alt+D` plus UIA, History/context exclusion, lifecycle/post drain, eight background asks plus `+1`, all overlap orders, reorder/removal races, and stale replies.

The same gate also requires nonzero unskipped
`terminal_powershell_startup_bootstrap_` coverage. Every local/UNC Windows
PowerShell launch must stay `Starting` and input-gated through the exact
client-PID-bound protected HMAC pipe sequence, root after normal profiles, form
one latest-policy `TerminalOnly|SemanticEnable` decision, and accept the matching
production `TerminalOnly|SemanticUnavailable` Prepared with no installed hook,
semantic-pipe connection, handshake, or capability. Only the injected fake may
produce `SemanticReady`; it must prove both orderings and the exact matching
`Prepared`/`StartupSemanticJoin` conjunction (plus the no-handshake
`RevokedTerminalOnly` cleanup outcome),
receive `ReleaseAck`, win the lifecycle-locked Running CAS,
then authenticate `CommitPrompt|CommitAck` before opening input/profile memory/
proved capability. EOF before CommitPrompt is fatal and never releases a
prompt. Cmd retains only its distinct protected inbound-pipe, root-PID-bound,
capability/nonce-correlated one-use `LaunchCwdBootstrap` for every local/plugin-
backed/UNC launch. No
target/semantic nonce/control pipe may enter PowerShell initial argv/environment;
the exact four-token fail-closed wrapper must accept only the sole typed finalize
object before invoking its one-shot output-silent CommitAck writer, and turn
every exception/early/invalid/extra merged-stream result into root exit;
pre-CAS failure publishes no Running/profile memory, while post-CAS failure
retires the already-Running incarnation with all startup gates closed; no policy
race sends a second decision; Plan 4 packages no allowlisted adapter manifest,
observer/control resource or installer and makes no real-shell semantic claim;
and evidence must state logical mutable-buffer cleanup without the
impossible claim that immutable command-line/.NET strings were physically
erased.

## Child completion and move

- **Completion record:** `TERMINAL-HANDOFF-UNRESOLVED` — replace this line with the exact plan-3 closeout-commit/blob pair, evidence repository commits, predecessor renderer native/archive identities, this plan's PaneProfiles native manifest/companion identities, and terminal/Preview Commands archive path/file SHA-256 values required below.

Do not mark this child complete until all of the following are true:

1. Consume the exact predecessor `Specs/Plans/Done/Terminal_RendererControlAndAccessibility_2026-07-22.md`. In this plan's final completion record, name/recompute the accepted plan-3 closeout commit, that commit's Done-plan Git blob object ID, every predecessor evidence `repositoryCommit`, the exact Renderer `TerminalNative` manifest/companion paths/digests, the exact `Specs/TestRuns/<MachineHash>/Commands/<RunId>/` archive, and SHA-256 for canonical `results.json`, `trace.txt`, runner-created `run-all-tests-results.json`, and `perf/perf_metrics.jsonl`. Rerun `Tools/Verify-TerminalNativeEvidence.ps1` with the predecessor's recorded repository/engine-lock identities; any missing file, blob/digest mismatch, skipped case, or implicit newest-run selection blocks entry.
2. Create exactly two x64 Release Commands runs with the canonical wrapper: scenarios `PaneProfilesTerminal` and `PaneProfilesPreview`, each with `-Platform x64 -Configuration Release -RepositoryCommit <exact-commit> -EvidenceRoot .\Specs\TestRuns -PassThruRunPath`. Validate both exact returned paths together with `Tools/Test-TestRunArchive.ps1 -RunPath @(<terminal-path>,<preview-path>)` and retain the exact PaneProfiles `TerminalNative` manifest/companion pair emitted above. Record both run IDs, evidence `repositoryCommit`, every path, SHA-256 for each Commands run's canonical content-free `results.json`, `trace.txt`, and `run-all-tests-results.json`, and the literal no-perf reason where applicable, plus both native-evidence files. Rerun `Tools/Verify-TerminalNativeEvidence.ps1`; all selected plugin-ABI/DxUi/profile/working-directory/Edit/path-insertion/interaction-broker/PowerShell-startup/native invocations and `terminal_pane_profiles_`/`preview_terminal_tabs_` Commands cases must be present, passed, and unskipped.
3. Merge the durable plugin/host boundary, plugin-owned per-host tab admission and hidden host-commit rollback, pane/tab/profile/location/lifecycle, canonical explicit UTF-16 launch environment/flags/caps, every-local/plugin-backed/UNC cmd `LaunchCwdBootstrap` including its verified embedded `.cmd`, two-pass `/S /K call` path encoding, eight-key child environment, protected precreated inbound pipe, root-PID/correlation validation, exact 512-byte-maximum fixed-order ASCII Ack assembled through clean EOF, target-digest/local proof, exact prelaunch-free-drive plus bounded `WNetGetUniversalNameW` logical UNC root/tail proof, retained nonsecret mapping ownership, root-death-only unchanged-mapping `WNetCancelConnection2W` cleanup, and stated same-root trusted-code limitation, the exact all-local/UNC `PowerShellStartupBootstrap` fail-closed wrapper/resource/protected-pipe/HMAC/rooting/decision/two-slot `StartupSemanticJoin`/`RevokedTerminalOnly` paired cleanup transports/Prepared/ReleaseAck/lifecycle-locked Running-CAS/final-disposition CommitPrompt/CommitAck/both-pipe post-CAS fence/control-read-semantic-write inert-latch/policy-race/input-memory-capability-gate/cleanup-failure contract, disable/retained re-enable and normal-refresh view behavior, the one per-view interaction broker and exact background-close/priority/deadline/focus/UIA/history-context/lifecycle contract, plugin-owned unsafe-paste overlay/token plus close/restart revocation and qualified disable/refresh continuity, distinct background/policy/lifecycle OSC 52 deny-and-scrub behavior, command migration, pseudo-input removal, folder Edit, path insertion, Windows/WSL selection, keyboard, and accessibility contracts owned here—including the shared WSL catalog and monotonic deadline CAS/tie/invalid-override outcomes, exact five-token/default-user/default-shell launch, permanent terminal-only/`RequestedUnverified` status, zero WSL memory/integration capability, exact-distribution full-native-path insertion, and different-distro/cross-family rejection—into `Specs/Terminal/Terminal_EmbeddedPane.md`, `Specs/Plugins/Plugins_Terminal.md`, `Specs/UI/UI_ManagePluginsDialog.md`, `Specs/UI/UI_FolderWindow.md`, `Specs/UI/UI_CommandMenuKeyboard.md`, `Specs/UI/UI_DxUiWinUIDesign.md`, and applicable plugin/shared-helper specs. Leave TerminalState persistence/generation/privacy rules to plan 5, but preserve the Windows-only `ITerminalPreferenceStore` boundary and Plan-4 precedence contract.

   The durable startup contract must state that every local/plugin-backed/UNC
   cmd launch uses the protected inbound-pipe, exact-root-PID, capability/nonce-
   correlated one-use bootstrap and that Plan-4
   production PowerShell resolves `SemanticEnable` only as
   `Prepared(SemanticUnavailable)` with zero hook/channel/capability. Preserve
   the generic fake-only Ready/two-slot-join proof, but defer the allowlisted
   adapter manifest, observer/control resources, extraction/install path, real
   semantic handshake, production Ready, and real-shell semantic closeout to
   plan 5.

   The durable insertion contract must separate Plan-4 host/ABI/fake transaction
   proof from product capability: production Plan 4 advertises
   `InsertCapable=false`, rejects direct insertion truthfully, and resolves a
   hidden pending Open as `IntegrationDisabled` before process creation. Plan 5
   first owns authenticated prompt/cwd translation, quoting, atomic insertion,
   no-Enter behavior, and production Ready/Commit/finalize success.

   The durable profile-memory contract must name the one-shot, generation-bound
   eligibility points: `CmdLaunchRootedRunning` only after the complete matching
   pipe/PID/correlation/target/mode result, applicable prelaunch-free-drive and
   bounded complete/root/tail UNC proof, secret/target-bearing cleanup with only
   the nonsecret owner tuple retained, and lifecycle-locked `Running` CAS for cmd;
   that CAS plus accepted matching post-Running `CommitAck` for ordinary Windows
   PowerShell/pwsh; the same Ack plus successful allocation-free/no-fail host
   finalize for the hidden seam/Plan-5 `PendingPathInsertion`, with Plan-4
   production `IntegrationDisabled` writing zero; and no read or write for WSL.
4. Change this document's State from WIP to Done, retain the exact evidence paths/digests in it, and perform exactly:

```powershell
git mv Specs/Plans/WIP/Terminal_PaneTabsProfilesAndLifecycle_2026-07-22.md Specs/Plans/Done/Terminal_PaneTabsProfilesAndLifecycle_2026-07-22.md
```

5. In the same change, update `Specs/Plans/WIP/README.md` N3 to route next to `Terminal_ShellIntegrationHistoryAndPaneFollow_2026-07-22.md`. A move or index update without its paired durable-spec and evidence updates is incomplete.

## STOP conditions

- Dynamic model cannot preserve Folder/Preview behavior or reentrant callback safety.
- A session lifetime must be owned/joined by the UI tab.
- Application shutdown must block the UI, destroy its message target before module quiet, access Terminal-private service/runtime state from the EXE, or retain a still-busy mapped module for process exit.
- Normal disable would retire/close a live Terminal, or normal refresh would force-remove an old-generation tab, time out its natural object barrier, use fatal/process-retention policy, reopen old-generation admission, or allow a replacement before the old owning module handle is reset.
- The host would need clipboard/preview/token ownership or paste classification/admission; a view-generation action could reread/reclassify mutable clipboard data; close/restart/app shutdown could destroy the child before plugin scrub; or generic disable/normal refresh alone would revoke/reject safe or an otherwise broker-eligible pending unsafe Paste on an existing live frozen-config view.
- Disable/re-enable or refresh could publish before the plugin-owned OSC 52 gate revokes/scrubs outstanding ask/allow payloads, or disabled existing terminals could accept a new OSC 52 write.
- More than one broker/presentation slot can exist per committed view; the host must own broker/domain state; the exact Close-over-Paste-over-OSC table cannot be implemented; a canceled lower-priority request can return; or History/context/plugin-text-input routing can bypass broker `None`.
- A background close can call UserTab before stable-ID select/reveal/focus/revalidation, display an invisible prompt, or roll selection back after selection succeeded; or a background/inactive OSC ask can retain/count a token, select/reveal/focus, or wait for later activation.
- Paste/OSC deadlines are not monotonic strict-before 60/30 seconds; selection/visibility/foreground/input-owner loss can preserve or resurrect a prompt; focus default/restore, Escape/Enter, OSC `Alt+A`/`Alt+D`, UIA unavailability, lifecycle scrub, or posted-payload drain cannot be proved.
- `IViewer` must be overloaded, the host must own/destroy/render the plugin child, or terminal-specific profile/launch/insertion/close logic cannot stay inside `Terminal.dll`.
- Logical source identity cannot move with current folder-content swap semantics.
- Profile/location validation would silently search current directory, launch network/removable code, or substitute an unrelated cwd.
- A profile launch would trust inherited OS variables, emit a noncanonical/ANSI/truncated environment, omit either Unicode creation flag, or accept a malformed/over-cap command line/environment.
- Any admitted profile/location/display identity would be truncated, hashed, aliased, or silently disabled in persistence; Windows launch could not use exact nonnull `lpApplicationName` plus fixed basename argv0; or WSL launchability could not reject its five-token command-line +1 independently of valid full logical identity.
- Interactive confirmation, natural-exit removal, chooser/plugin removal, and forced teardown would not share the one plugin retirement-intent arbiter, or any host route could bypass the single `{instanceId,hostViewGeneration}` removal claim and mutate/drain/release twice.
- Any local/plugin-backed cmd could rely on initial `CreateProcessW` cwd without
  the verified embedded `.cmd`, exact two-pass `/K call` path encoder, eight
  fixed overwritten/cleared environment keys, protected ready-before-create
  one-instance inbound pipe, exact root-PID/capability/correlation validation,
  or the exact 512-byte-maximum fixed-order ASCII Ack assembled through clean
  EOF; any HMAC, VT, cmd-side Unicode-cwd
  serialization, alternate inline target/helper, or unqualified authentication
  claim could be introduced; or the trusted same-root AutoRun limitation could
  be hidden. STOP as well if any cmd launch could reach `Running`/input/profile
  memory after AutoRun exit/hang, a missing/late/wrong/duplicate/replayed result,
  wrong target digest, occupied reported drive, nonremote drive, over-cap/
  malformed/inconsistent complete-root-tail UNC reverse map, or before pipe
  drain and complete secret/target-bearing launch-state cleanup. STOP if a live-
  root mapping could be canceled, mapping identity could rely on physical/DFS/
  provider alias collapse, a changed/reused/shadowed mapping could be forced or
  canceled after root death, cleanup could wait unboundedly, or process exit
  could be assumed to perform the unissued `popd`; or if cmd would require an ongoing nonce, trusted
  activity/cwd, history, follow, or insertion adapter afterward.
- Any Windows PowerShell/pwsh local or UNC launch could bypass
  `PowerShellStartupBootstrap`; profiles would not run exactly once before the
  exact four-token fail-closed `-NoExit -Command` wrapper/verified script;
  literal encoding, fixed `*>&1`, typed finalize-object/one-shot complete-frame
  Write, or exit-125/126 behavior would drift; the finalizer could Flush, retry,
  or perform a later pipe operation, or a gate could open from client Write
  completion rather than the server's full valid-frame read; policy
  would be bypassed; target/semantic material would be placed in initial argv/environment; the protected client-PID-bound HMAC pipe,
  exact nine-kind sequence, post-profile literal rooting, decision-time latest
  policy, Prepared outcome, two-slot order-independent `StartupSemanticJoin`,
  or strict five-second deadline could not be proved;
  `ReleaseAck` could itself release a prompt; the lifecycle-locked Running CAS
  could lose nonatomically to retirement; or a prompt could occur before Running
  or a user byte/profile write/semantic capability before accepted `CommitAck`.
- Join slots cannot be reserved before child creation/admitted without blocking;
  a pre-CAS live disable would close/fail rather than retain both exact no-
  operation cleanup transports through either valid cleanup Prepared; the final-disposition/
  policy-generation CommitPrompt cannot force inert-owned-mapping cleanup before
  Ack; or a post-CAS disable cannot close/cancel both and make completed control
  EOF/error or semantic-write failure latch every fake interposer inert before
  emission. STOP if the fake seam cannot prove the exact
  `ReadData|ReadAttributes|WriteAttributes|Synchronize`/no-`WriteData` mode-
  switch boundary, uses Flush/synchronous control I/O, adds a Terminal-local/
  fallback UTF converter instead of the shared C++/exact PowerShell encoders,
  lets `PollControlTransport(NoWait)` wait/dispatch/overwrite an occupied stage,
  lets `PollControlTransport(HandlerWaitOnce)` inspect or wait on the newly armed
  read before consuming an occupied stage, or use anything but its one nonalertable
  250-ms AsyncWaitHandle wait after owner/binding/epoch/lifecycle
  revalidation, acts from an unsignaled read, omits rearm-before-dispatch/final
  strict deadline, submits a control Write without the fixed-reserve descriptor/
  byte/final-sequence placeholder, performs a second wake admission/sequence,
  bypasses later FIFO input, mishandles pre-submit tombstone or submitted-failure
  held-fence disposition, injects a second/retry wake, makes the child revoke merely
  from an unsignaled read, fails to let the plugin timeout/revoke its own
  outstanding operation, allows a replacement binding to receive more than one
  raw wake, releases ambiguous input
  at expiry/disable, or releases an I/O/module
  gate or closes its in-flight handle before exact cancel, quiet-bounded event
  drain, and actual terminal completion; or closes the persistent fake pipe/
  releases operation or input hold merely from successful write rather than the
  matching result; or permits another control operation after Probe non-Success,
  Expired, BindingChanged, ShellOperationFailed, or another failed/revoked/
  ambiguous outcome, or fails to return Idle after matching Success or a Follow/
  Insert pre-effect RejectedState/BufferChanged. A matching cleanup Ack would be rejected solely for the
  policy change, ordinary input/integration-independent profile memory would
  stay closed, or disable alone would retire the otherwise valid terminal.
- CommitPrompt disposition/generation cannot accept exactly all four rows:
  original TerminalOnly plus Prepared TerminalOnly -> TerminalOnly at equal/
  later generation; SemanticEnable plus SemanticReady plus completed join ->
  SemanticActive only at unchanged generation; SemanticEnable plus
  SemanticUnavailable -> TerminalOnly only at unchanged generation; and either
  valid SemanticEnable outcome after latched disable -> TerminalOnly at current
  strictly greater generation regardless of later re-enable. STOP as well if it
  accepts generation decrease/wrap, Active without Enable+Ready+join,
  TerminalOnly outside those rows, or any other pair, or if the script retains
  more than immutable Decision generation plus Prepared outcome through prompt
  commit.
- Plan 4 production PowerShell could install or package an allowlisted adapter,
  observer/control script or resource, open an ongoing semantic channel, report
  `SemanticReady`, publish a semantic capability, or use real-shell semantic
  success instead of fake-only Ready/two-slot-join proof. STOP and leave those
  dependencies to plan 5.
- A PowerShell startup failure could fall back interactively, reroot, remember a
  profile, resurrect from a late frame/settings change, send a second decision,
  publish an unproved staged adapter, or claim secure physical erasure of OS/
  .NET immutable command-line strings.
- Profile memory could form `CmdLaunchRootedRunning` or write for cmd without
  its complete matching pipe/PID/correlation/target/mode result, applicable UNC
  reverse mapping, launch cleanup, and the same-incarnation lifecycle-locked
  `Running` CAS; for ordinary Windows
  PowerShell/pwsh before the accepted matching post-Running `CommitAck`; for a
  hidden `PendingPathInsertion` before that Ack and the successful allocation-
  free/no-fail host finalize; more than once or from a stale generation; or at
  all for WSL.
- The host must read/pre-enforce `maxTabsPerPane`, make a provisional tab visible before Open succeeds, or cannot completely roll back quota/synchronous Open/host-commit failure without changing zoom/selection or leaking a child/process/view slot.
- Plan-4 production could advertise `InsertCapable`, admit direct insertion,
  start a hidden insertion process instead of resolving `IntegrationDisabled`,
  or claim authenticated translation/quoting/atomic/no-Enter success that first
  depends on the plan-5 adapter.
- Command migration cannot distinguish Ctrl+Enter from Ctrl+Shift+Enter, directory Edit cannot use provider metadata, or pseudo-command-line removal would strand a stable command without an explicit migrated contract/test.
- WSL discovery requires another registry reader instead of the canonical bounded shared catalog.
- A v1 WSL launch would add or require `--exec`, `--user`, `-e`, a wrapper, another process, a bootstrap or `WSLENV`/environment/PTY/profile/staged-script injection; alter the configured default user/shell; transition out of permanent `RequestedUnverified`; expose history/follow/insertion/activity/authenticated-cwd/profile-memory capability; or make terminal usability depend on integration. STOP and require a separately approved WSL-integration design rather than weakening the frozen five-token contract.
- Any material baseline drift is unresolved, the preference seam would need Plan-4 disk semantics, or WSL discovery lacks the exact monotonic deadline/cancel CAS priority, accepts a non-positive override, blocks the UI/other profiles, or fails silently.
