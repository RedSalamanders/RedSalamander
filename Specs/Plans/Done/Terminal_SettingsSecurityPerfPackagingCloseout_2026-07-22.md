# Terminal Settings, Security, Performance, Packaging, And Closeout

> **Status — Completed 2026-08-03.** Terminal settings, Preferences integration, security/resource limits,
> performance instrumentation, centralized package topology, exact private-runtime identity, complete notices, and
> dependency check/upgrade/rollback procedures are implemented. Durable authority is
> `../../Terminal/Terminal_EmbeddedPlugin.md`; later WIP/checkpoint language below is historical.

## Current checkpoint (2026-08-01)

Terminal-typed Preferences, validated first-slice settings, exact dependency
rebuild identity, centralized nested staging, focused lifecycle testing, and
initial terminal performance metrics now exist. This is not closeout: trusted
shell state/history, advanced rendering/accessibility, loader hash/signature
binding, native ARM64 product builds, portable/package validation, security
stress, and reviewed performance budgets remain open.

> **Authority:** Execution plan 6 of 6 for `Terminal_EmbeddedConPtyLibGhosttyVt_IntegrationPlan_2026-07-13.md`. This is the integration/closeout slice; it cannot waive failures in plans 1–5.

## Status and entry gate

- **State:** WIP / first settings, dependency, staging, and perf seams implemented
- **Drift baseline:** `fbf6af76f373bf253488862426f4ea70a9ff25ba`
- Begin final closeout only when the selected engine is locked, core/render/pane/shell tests pass, and every master capability row has a truthful disposition.
- Consume the exact predecessor `Specs/Plans/Done/Terminal_ShellIntegrationHistoryAndPaneFollow_2026-07-22.md` and every exact predecessor handoff path/digest it records for the engine lock/finalization, core, renderer, pane, shell-integration/history, state-store, and Commands evidence. Recompute each recorded digest and reject missing, changed, non-passing, unverified, or latest-discovered evidence.
- Recheck Settings Store v16 opaque per-plugin configuration, `IInformations`/generic Preferences plugin configuration, `IID_ITerminal`/module quiet, themes, localization, UIA, `RuntimeDependencies.props`, installer/portable specs, perf-budget syntax, and WIP closeout policy.
- Compare the master, all five predecessor Done plans and durable contracts, the paths named above, packaging/release scripts, schemas, themes, localization resources, installer contracts, testing guidance, and release policy against the baseline before implementation and again before closeout. Any material drift in behavior, evidence schema, resource/security/privacy rules, package topology, artifact naming/versioning, native architecture, signing/install policy, or closeout ownership is a STOP: reconcile every affected authoritative spec/plan, review the change, and advance the baseline deliberately before continuing.

## Final integration surfaces

- Revalidate that all PowerShell observer implementation remains inside `Plugins/Terminal/TerminalShellIntegration.h/.cpp` and plugin-owned support. `Plugins/Terminal/ShellAdapters/PowerShellReadLineAdapters.json`, its closed schema/corpus, and the exact versioned wrapper/control scripts are source inputs compiled as identity-bound `Terminal.dll` RCDATA through `Terminal.rc`/`resource.h`; `Terminal.vcxproj` and `.filters` list every input. Neither the EXE, host, Common, another plugin, nor a loose runtime file may implement or configure the observer.
- Revalidate the one wire-schema source at
  `Plugins/Terminal/ShellAdapters/TerminalSemanticProtocol.v1.json`.
  `Tools/Generate-TerminalSemanticProtocol.ps1` deterministically emits the
  checked-in `Plugins/Terminal/Generated/TerminalSemanticProtocolV1.generated.h`
  and
  `Plugins/Terminal/Generated/TerminalSemanticProtocolV1.generated.ps1.txt`;
  the build runs the generator in `-Check` mode and fails on stale output.
  `Terminal.vcxproj`, `.filters`, `Terminal.rc`, and `resource.h` compile the
  canonical JSON plus generated PowerShell constants as identity-bound RCDATA
  and include the generated C++ header. The closed declarative
  corpus is exactly
  `Tests/TerminalTests/Fixtures/TerminalSemanticProtocolCorpus.v1.json`; native
  byte/parser cases live in
  `Tests/TerminalTests/TerminalSemanticProtocolTests.cpp`; and generator/resource/
  PowerShell identity checks live in `Tools/Test-TerminalSemanticProtocol.ps1`.
  No generated or source schema/constant file is a loose runtime payload, and no
  hand-maintained duplicate constant table is permitted.
- Revalidate the distinct startup implementation in the same plugin boundary: every local, validated-plugin-backing, or UNC cmd launch uses one protected-inbound-pipe, exact-root-PID, capability/nonce-correlated one-use `LaunchCwdBootstrap`, and every local/UNC Windows PowerShell/pwsh launch uses `PowerShellStartupBootstrap`; both live in `TerminalShellIntegration.h/.cpp`. `TerminalProfileCatalog.h/.cpp` builds the exact shell argv; `TerminalWorkingDirectory.h/.cpp` verifies logical rooting; and `ConPtySession.h/.cpp` plus `TerminalService.h/.cpp` own `StartupSemanticJoin`, the lifecycle-locked Running CAS, and startup input/profile-memory/capability gates. The exact launch-only `.cmd` script/manifest, exact versioned PowerShell startup script/manifest, and immutable fail-closed wrapper template are identity-bound `Terminal.dll` RCDATA. Only the two verified launch scripts are extracted transactionally through the owner/access/reparse-safe plugin directory; cmd uses only the exact two-pass-encoded `/S /K call` token and PowerShell uses only the generated `-Command` wrapper. Project/filter/resource maps name every input. Reject a host/EXE/Common implementation, public ABI dependency, helper EXE, pane-folder script, or loose runtime/package payload.
- The final build/test inventory includes `Tools/Test-TerminalPowerShellAdapters.ps1`, its read-only `-AuditInstalled` mode, `Tools/Test-TerminalPowerShellControl.ps1`, deterministic `Tests/TerminalTests` observer/control plus `terminal_cmd_launch_bootstrap_` and `terminal_powershell_startup_bootstrap_` cases, the PaneProfiles and ShellState native slices, `terminal_pane_profiles_`/`terminal_shell_state_` Commands cases, source-contract tests, and extracted-package resource tests. They verify exact cmd/PowerShell startup and adapter resource ID/version/SHA-256 bindings and reject a loose manifest/script, a second parser, host semantic interpretation, or startup/observer implementation outside `Terminal.dll`.
- Every adapter tuple remains exact: PowerShell edition and host interval, PSReadLine version, module path/file SHA-256, exact UTF-8 effective `FunctionInfo.Definition` SHA-256, reflected static `Microsoft.PowerShell.PSConsoleReadLine.ReadLine` overload, strong assembly identity, one audited two-/three-argument recipe, wrapper/manifest identities, clean-default real-shell result, and upstream-license record. Duplicate/overlapping/floating tuples, nearest/range fallback, unknown overload/recipe, function-text patching, or unhashed input are release failures.
- Revalidate `Plugins/Terminal/TerminalInteractionBroker.h/.cpp` as the sole per-committed-view Close/unsafe-paste/OSC 52 presentation arbiter, wired through `TerminalControl.h/.cpp`, `TerminalAccessibility.h/.cpp`, `Terminal.rc`/`resource.h`, the domain authorization gates, posted-payload registry, and lifecycle teardown. `Terminal.vcxproj`/`.filters` include every file/resource. The host changes only its stable-ID background close-glyph select/reveal/layout/focus/revalidate path; it never owns broker state, deadlines, UI tokens, or sensitive buffers, and no public ABI is added for host model inspection.
- Keep hyperlink validation and opening inside `Terminal.dll`, adjacent to the terminal-document effect gate in `TerminalControl.h/.cpp`. The plugin-local helper is intentional: the shared-helper catalog has no URL opener with this Terminal-specific immutable-document generation, scheme, length/control/syntax, failure-localization, and one-side-effect contract. Recheck `Specs/Core/Core_SharedHelpers.md` before implementation; add a shared helper only if its reviewed semantics truly become cross-consumer, and then update that catalog and the authoritative owner. The host owns no URL policy or open call. Add the implementation and its localized failure/status resources to `Terminal.vcxproj`, `.filters`, `Terminal.rc`, and `resource.h` as applicable.
- Final test/source/package inventory includes nonzero `terminal_interaction_broker_` TerminalTests in Core and PaneProfiles native evidence, pane/Commands background-close cases, Preferences/UIA/security/lifecycle cases, generic DxUi close/reorder regression, and source-contract checks. Reject multiple broker slots, broker/domain implementation in EXE/Common, host token/payload storage, raw posted payloads, a loose broker resource, an invisible/background prompt, or a package where this implementation is outside exact `Plugins\Terminal.dll`.

## Settings contract

Do **not** add a top-level `terminal` object. Persist the complete canonical plugin value under existing `plugins.configurationByPluginId["builtin/terminal"]` without a schema bump. `Terminal.dll` exposes schema/defaults/validation/canonical output through `IInformations`; the host preserves the JSON opaquely and uses the generic plugin configuration transaction. Use yyjson copy APIs and positional resource placeholders.

`Preferences > Plugins > Terminal` exposes every master setting in groups, including `windowsDefaultProfileId=auto|pwsh|windows-powershell|cmd`, exact matching-WSL launch/status, `followPathWhenIdle`, debounce, close/tab/profile/history/appearance/input/security/image bounds, privacy actions, and collapsed Advanced values. Advanced is never JSON-only. Capability status is exact: cmd is usable but ongoing terminal-only after its mandatory pipe/PID/capability/nonce-correlated one-use launch-root bootstrap on every local, validated-plugin-backing, or UNC launch; this is integrity/correlation/replay protection under trusted AutoRun, not authentication against earlier same-root startup code. Every local/UNC Windows PowerShell/pwsh launch first completes the mandatory post-profile startup exchange, and may expose only independently proved features from the one semantic epoch created and delivered by that exchange's post-rooting `SemanticEnable` decision plus one exact allowlisted ReadLine adapter tuple. `TerminalOnly` creates no epoch. Input/profile memory/capabilities stay gated until the Running CAS is followed by authenticated `CommitPrompt|CommitAck`. Unknown/custom/missing/Vi/replaced adapters and observer failure are terminal-only/conservatively busy with no history/follow/insertion until Restart. A revoked epoch is permanent until explicit Restart/new session and nested integration is unavailable. The WSL status is explicit and truthful: exact distro plus `--cd` launch is available, the configured default user/shell is preserved, v1 cwd trust is permanently `RequestedUnverified`, and app history/follow/activity/authenticated-cwd/profile-memory capabilities are unavailable, while explicit insertion is limited to the quoted full native Linux path in the exact launch distribution. No setting or hidden Advanced field enables cmd/WSL ongoing semantics, a second/live PowerShell handshake, an unlisted adapter, prompt/key/cell inference, or a WSL bootstrap/adapter. Diagnostics link to the exact group/field. Generic plugin enable/disable is the only availability switch: disable blocks new open/Edit/insertion/Restart requests but leaves existing sessions alive and usable, retains the administration/module/service generation, and never implicitly starts refresh.

The Terminal schema-v2 navigation identity is closed. Its seven case-sensitive
section IDs, ascending `order`, first-entry collapse policy, and editable
elements in UI order are exactly:

| Section ID | Order | Collapsed by default | Editable element IDs in UI order |
|---|---:|---:|---|
| `shells` | 10 | false | `windowsDefaultProfileId`, `shellIntegrationEnabled` |
| `pane-behavior` | 20 | false | `followPathWhenIdle`, `maxTabsPerPane`, `confirmCloseRunningProcess`, `closeOnExit`, `scrollbackMaxLines` |
| `history-privacy` | 30 | false | `rememberShellByFolder`, `rememberCommandHistory`, `historyIgnoreLeadingSpace`, `historyDeduplicateConsecutive`, `historyMaxCommandsPerFolder`, `historyMaxTotalCommands`, `historyMaxFolders`, `historyMaxAgeDays` |
| `appearance` | 40 | false | `fontFamily`, `fontSizePoints`, `fontWeight`, `ligatures`, `cellWidthScale`, `lineHeightScale`, `cursorStyle`, `cursorBlink`, `cursorBlinkIntervalMs`, `backgroundOpacity`, `boldIsBright` |
| `interaction` | 50 | false | `copyOnSelect`, `trimTrailingWhitespaceOnCopy`, `scrollToBottomOnInput`, `showScrollbar`, `warnOnUnsafePaste`, `pasteWarningLineThreshold`, `bellStyle` |
| `security-graphics` | 60 | false | `kittyGraphicsEnabled`, `kittyImageStorageMiB`, `kittyApcMaxMiB`, `osc52Policy`, `hyperlinkPolicy`, `allowApplicationTitle` |
| `advanced` | 70 | true | `followPaneDebounceMs`, `scrollbackMaxMemoryMiB`, `inputQueueMaxMiB`, `historyMaxCommandBytes`, `historyMaxStateFileMiB`, `pasteMaxBytes`, `kittyGlobalCpuStorageMiB`, `kittyGlobalGpuStorageMiB`, `kittyMaxSingleImagePixels` |

Each stable case-sensitive `elementId` equals its exact configuration key and
its element order is the one-based list position multiplied by 10. Advanced
membership is explicit schema metadata from this table, never inferred from
type, range, comment, or control kind. The visible editable ID set equals those
48 unique keys exactly. `configurationVersion` is the one hidden metadata field,
making 49 total fields, and is forbidden from sections, status/action records, edit
controls, and deep-link targets. Exact-set schema/renderer tests reject any
missing, extra, duplicate, renamed, case-changed, multiply placed, or hidden-
version element. A generated row for every localized Terminal diagnostic/status/
action link must resolve exact
`{pluginId="builtin/terminal",sectionId,elementId}` against the active schema;
unknown or empty pairs fail source-contract tests rather than opening an
approximate page.

Every such link uses
`IHostPreferences2::RequestOpenPluginConfiguration(const
HostPluginPreferencesRequest*)`; the frozen public request type is exactly
`HostPluginPreferencesRequest` (`80/8` on x64 and ARM64); no shortened alias is
valid. Require exact size/version, zero reserved words,
copied/bounded NUL-free spans, canonical case-insensitive plugin lookup, case-
sensitive section/element lookup, posted-payload ownership, same-plugin newest-
complete coalescing with cross-plugin FIFO, UI-generation/schema/owner
revalidation, and the master's exact HRESULT/stale/disabled-element outcomes.
Loaded-DLL and checked-in golden-oracle tests freeze its name, IID, layout, QI/
controlling-IUnknown/agility, bounds, result codes, and x64/ARM64 parity.
Implement only the generic public service in `Common/PlugInterfaces/Host.h`, the
schema-v2 parser/validation in `Common::PluginConfiguration`, and the generic
`ShowPreferencesDialogPluginElement(...)` UI route in `Preferences.h/.cpp` plus
its normal project/filter/localization/test surfaces. `Terminal.dll` queries
and retains `IHostPreferences2` from factory `IHost` and supplies its schema,
mapping, resources, and exact link IDs; it never calls an EXE symbol directly.
The host contains no Terminal setting, default, section map, diagnostic choice,
or fallback target. Source/package tests prove all Terminal-specific runtime
implementation remains in `Plugins\Terminal.dll` while these narrowly required
generic host/ABI integration surfaces remain reusable by any plugin.

Important fixed/default limits:

- `maxTabsPerPane` is a plugin-owned effective `1..32` live limit, enforced only by `ITerminal::Open` through an atomic per-complete-`PhysicalHostKey` view reservation after synchronous validation and before child/async work. The host never reads/pre-enforces it. Live lowering evicts nothing and blocks later reservations for a key until its count is below the new value; raising resurrects nothing. Quota returns exactly `HRESULT_FROM_WIN32(ERROR_QUOTA_EXCEEDED)` with no visible tab/zoom/child/process, and a successful diagnostic/Exited view remains charged until callback null-drain, valid Close, and public release.
- Exactly 32 process-lifetime service session slots, held through Starting/Running/Failed/Draining/Closing/retiring/quarantine until complete quiet; any quarantine closes launch/restart admission service-wide until all recover.
- 64 MiB terminal/scrollback simultaneous peak/session, including grids, scrollback/metadata, immutable snapshots, UIA/copy/provider buffers, and reflow/compression staging; 10,000 default lines, lower bound wins.
- 512 MiB terminal/scrollback/snapshot CPU/process, including hidden sessions, UIA/RPC-held old referenced publications, and quiet retained exited snapshots; deterministic activation-only cross-session LRU with stable session ID and one-waiter/release-epoch retry never invalidates a visible grid or in-use publication.
- 8 MiB default/hard-maximum normal input bytes, live-lowerable to 1 MiB; 4096 total descriptors and fixed 256 KiB/256 descriptor protocol/control reserve used only by required engine responses plus the one in-flight foreground-control wake placeholder. The placeholder reserves its descriptor, one byte, and final FIFO sequence before the control-pipe Write and is activated in place after exact completion.
- 4 MiB default paste, configurable only up to 8 MiB and still all-or-nothing. Immutable normalized payload plus preview awaiting authorization has fixed non-configurable caps of 8 MiB per terminal and 32 MiB service-wide, charged before materialization and securely scrubbed; only one unsafe-paste confirmation may be pending per terminal, it must own that view's one broker slot, and its fixed monotonic 60-second strict-before deadline is not a setting.
- Kitty: 64 MiB decoded CPU/session, 256 MiB CPU/process, 256 MiB GPU/process, 32 MiB APC default, 16M pixels/image, one conversion generation per session total.
- Named non-Kitty hard caps replace the 64 KiB unclassified body cap: OSC 8 is 16 KiB record/8 KiB URI/1 KiB parameters; title/status 4 KiB before 256-code-point/80-grapheme sanitization; cwd/path 128 KiB; OSC 52 1 MiB encoded/768 KiB decoded; authenticated integration 2 MiB frame/1 MiB command/128 KiB cwd; pending effects 256 descriptors/2 MiB encoded per session. OSC 52 defaults deny; a foreground interactive-visible `ask` permits exactly one outstanding request per session and eight service-wide with a fixed monotonic 30-second strict-before deadline. Background/hidden/inactive asks deny/scrub before either counter increments and never focus/select/reveal; clipboard reads are ignored. Hyperlinks require Ctrl+click; application title is allowed only after sanitization. No notification setting exists until the runtime exposes an enforceable effect.
- History defaults and bounds exactly as master; capture disable performs durable purge/fence. Shell-choice disable independently purges/fences choices.
- Fixed identity/launch ceilings are not Preferences: profile ID 32,772 UTF-16 units, canonical location key 33,037, display path 33,040, mutable process command line 32,767 including NUL, and explicit environment block 524,288 including both terminal NULs. Exact identity maxima remain complete across chooser/Open/preference/history/Restart and +1 rejects the whole identity or mutation; no truncation/hash/alias is permitted. Windows shells use exact held/revalidated nonnull `lpApplicationName` plus one fixed unquoted basename argv0. WSL keeps valid logical identity distinct from its exact five-token quoted-command bound; for quote-free distro `D`, a 32,736-unit absolute Linux path is the exact launch boundary and +1 creates no preflight/process.

`TerminalState.schema.json` uses JSON Schema Draft 2020-12, whose `maxLength`
counts Unicode code points only. Treat it as a coarse structural backstop, never
as proof of a UTF-16-code-unit or UTF-8-byte ceiling. Every applicable property
has the exact nonassertive annotation
`x-redsalamander-maxUtf16CodeUnits` or
`x-redsalamander-maxUtf8Bytes`; the runtime validator authoritatively enforces
32,772/33,037/33,040 UTF-16 code units and the 1-MiB strict-UTF-8 command bound,
and the serializer emits only values that pass both runtime and schema. Every
applicable ceiling has closed `bmp-boundary-both-pass` and
`bmp-plus-one-both-reject` classes; each UTF-16 annotation also has
`astral-utf16-plus-one-schema-pass-runtime-reject`, and each UTF-8 annotation
has `multibyte-utf8-plus-one-schema-pass-runtime-reject`. A
`schema-invalid-runtime-valid` case is forbidden; no test or document
may claim blanket schema/runtime parity or that `maxLength` rejects every
UTF-16/UTF-8 +1 value.

Appearance applies live. Lowering the input ceiling atomically blocks new normal admission until accepted bytes drain under it and never changes the protocol reserve. Lowering a memory cap blocks reservations, prunes/evicts deterministically off UI, and commits only once usage fits; on failure retain the old setting. Visible grid/current snapshot/paint is never evicted, no allocation is grandfathered, and raising a limit resurrects nothing. Profile/launch defaults otherwise apply to new/restarted sessions. `followPathWhenIdle` is likewise sampled once at Open or explicit Restart; existing incarnations and session overrides are unchanged by a persisted-default commit, and Restart discards the prior override before resampling. Every PowerShell launch performs startup rooting regardless of `shellIntegrationEnabled`; only the one post-rooting `SemanticDecision` samples the latest effective value/generation. A false-to-true commit before that decision may therefore produce `SemanticEnable`; after a `TerminalOnly` decision it sends no second decision and that incarnation remains terminal-only with Restart required. A true-to-false commit before the decision produces `TerminalOnly`; after `SemanticEnable` it permanently revokes the staged/current epoch under the normal generation fence and marks any incomplete `StartupSemanticJoin` `RevokedTerminalOnly`. That sole nonfatal exception closes semantic/control operation and capability admission and scrubs/ignores both slots/late records, but a Starting pre-CAS session retains both protected servers as exact-root-PID, no-operation cleanup transports through Prepared and the Running CAS. It requires otherwise-valid Prepared Ready or Unavailable only as cleanup proof; missing Prepared times out. At the Running CAS both servers close/cancel before final TerminalOnly CommitPrompt makes the script latch inert/remove an app-owned mapping before ack. A disable after the CAS closes/cancels both immediately: `TerminalControl` is the child's pending-read revocation sentinel, while `TerminalSemantic` closure breaks a pending or next semantic WriteAsync. Matching CommitAck acceptance remains keyed to bootstrap/session/incarnation/deadline, not live policy generation: a disable before or after CommitPrompt/Ack never fails or retires startup by itself, and accepted Ack still completes bootstrap cleanup, opens ordinary input, and permits integration-independent Windows profile memory while publishing zero semantic capability/rejecting old events. For an already Running incarnation, true-to-false atomically and permanently invalidates its sole epoch/trust/capabilities; cancels and purges pending follow, path insertion, and any pending accepted-command candidate; stops app-history capture; leaves an owned wrapper as inert same-recipe pass-through; and removes the control mapping only if still exactly app-owned. False-to-true never retrofits/rekeys an incarnation whose one decision already occurred. Cmd's separate startup-owned all-Windows-location `LaunchCwdBootstrap` remains admitted/unchanged because it proves launch cwd; cmd and WSL create/revoke/restart no ongoing epoch and both directions inject zero semantic bytes. Generic `enabled=false` prevents new sessions but does not kill existing sessions. A setting can lower but never exceed hard caps. Deterministic races cross both integration-toggle directions with both `StartupSemanticJoin` arrival orders/slot failures, every PowerShell bootstrap frame, ReleaseAck write/receipt/validation, Running CAS, just-before/after CommitPrompt construction/send, one-shot finalizer CommitAck Write, and server receipt, including same-generation Enable/Unavailable TerminalOnly and both valid Enable outcomes after a latched disable, plus nested ownership, pending follow/insertion/candidate, `AcceptedInput`/matching `RootReadReady`, and close/restart/root-exit/shutdown. They prove the single CAS winner, exact CommitPrompt disposition/generation validation, no never-Running prompt, no input/profile memory/capability before CommitAck, correct already-Running gated retirement for unrelated post-CAS failure, no retirement solely for disable, cleanup-Ack acceptance and ordinary input/memory with zero semantic capability after disable, no second decision/rekey, stale semantic completions restore no capability/history, and rooting remains independent. The same races against cmd/WSL prove no ongoing epoch/input/capability change.

For every live field, configuration apply constructs one immutable
`TerminalLivePolicySnapshot{configurationGeneration,...}` off UI. One service-
owned serialized configuration coordinator is the sole linearization point for
Apply, Open registration, and registry removal; it is a queue/coordinator, not
a lock held while session work runs. `BeginApply` reserves the next nonzero
generation, snapshots active session identities in stable ascending order, and
queues later Open registrations and registry removals behind that transaction.
A root exit may mark its own session retiring locally before it reaches the
coordinator, but it cannot remove or replace captured membership.

Prepare visits the captured identities one at a time. Under only that session's
serial gate it validates and stages the candidate policy and fences the affected
actions, then releases the gate before advancing; it never holds the service or
another session lock and never calls/waits into the coordinator. A session that
was already retiring when Prepare enters acknowledges a no-op. Root exit, close,
or retirement arriving after that session's Prepare queues behind commit/abort.
Any validation/staging failure runs Abort over every prepared session in the
same stable order, restores its old staged state, retains the old immutable
snapshot, publishes no generation, and then releases queued coordinator work.

After every Prepare succeeds, one atomic coordinator commit token publishes the
complete immutable snapshot and generation. Commit/release then visits each
prepared session gate one at a time before Apply reports success. Prepare has
already reserved and validated every needed resource, so every post-publication
Commit/release operation is allocation-free, nonblocking, `noexcept`, and
infallible; no ordinary rollback or partial-failure branch exists after the
commit token. An impossible token, captured-identity, or staged-state mismatch
closes service admission and invokes injected
`FatalConfigurationInvariantPolicy`: production calls `RaiseFailFastException`,
while tests record the fatal boundary and must never continue in a mixed-
generation state. Queued Opens bind wholly to the new snapshot; queued removals
then retire normally. No lock is held across a wait, callback, UI/COM call,
process/pipe I/O, or coordinator call, and at most one session gate is held.
Each action reads one complete snapshot at the gate below and never mixes fields
or rereads a newer generation halfway through a side effect. Deterministic and
source-contract tests cover Apply against Open, root exit, explicit close/
retire, a captured session becoming retiring, Prepare failure plus full
rollback, atomic commit-token publication, allocation-free/nonblocking/noexcept
Commit/release, fatal token/identity/staged mismatch without continued service,
actions on both sides of commit, stable visitation, and deadlock/mixed-
generation absence. The closed timing table is:

| Setting | Applies to | Exact snapshot/commit boundary and race result |
|---|---|---|
| `confirmCloseRunningProcess` | Existing and new sessions, on the next explicit `UserTab` close | The retirement-intent arbiter snapshots the Boolean and configuration generation in the same lock/CAS that changes `Open` to the close intent. A configuration commit that linearizes first governs that intent; a later change neither creates, dismisses, nor answers an already-issued prompt. A canceled intent returns to `Open`, so a later independent close snapshots the then-current policy. Forced/window/application/automatic paths never consult it. |
| `closeOnExit` | Existing and new sessions whose root exit has not yet been observed | Root-exit admission snapshots the enum and configuration generation in the same lifecycle transition that records exit code and `DrainingAfterRootExit t0`. That frozen value alone decides automatic removal after final snapshot/quiet. A later settings change never closes or resurrects that exited incarnation; a Restart creates a new incarnation using then-current policy. |
| `followPathWhenIdle` | Each new session incarnation at Open or explicit Restart | Open registration and Restart admission serialize through the configuration coordinator and snapshot the persisted Boolean plus configuration generation exactly once. A successful Apply commit that linearizes first governs the new incarnation; otherwise it receives the prior committed value. Existing incarnations and their session-only overrides never change from this persisted-default commit. Restart ends the old override and the replacement resamples the committed default. Cancel publishes no generation and changes nothing; toggling the session override never persists. |
| `followPaneDebounceMs` | Existing and new follow-capable sessions | A commit increments the follow-policy generation, cancels every not-yet-dispatched old-generation timer, retains only each session's newest pending source update, and schedules it from the commit instant using the new delay (`0` queues the normal serialized admission immediately). A timer and commit race through the same generation gate: dispatch admitted before commit keeps its already-frozen operation; otherwise the stale timer is a no-op. All authenticated-idle checks still occur at final dispatch. |
| `copyOnSelect` | Selection completions after commit | The final selection-complete UI action snapshots `copyOnSelect`, `trimTrailingWhitespaceOnCopy`, document/selection generation, and policy generation together immediately before clipboard transaction admission. A settings commit before that gate governs; after it, the one admitted action finishes under its frozen snapshot. Drag/intermediate selection never copies. |
| `trimTrailingWhitespaceOnCopy` | Explicit Copy and admitted copy-on-select actions | Copy admission freezes the Boolean with the exact immutable selected-text generation before transformation. Trimming applies per logical copied line and never mutates terminal cells/selection. A later setting change cannot reinterpret bytes already admitted to the clipboard transaction. |
| `scrollToBottomOnInput` | Each later normal user-input descriptor; protocol responses and integration operations never use it | Writer admission freezes the Boolean, input descriptor identity, and model/view generation atomically. If true, that descriptor schedules one bottom-follow effect after its bytes win admission; if false it preserves viewport. A settings race is resolved solely by which publication/admission linearizes first; queued earlier descriptors are not rewritten. |
| `showScrollbar` | Existing and new views | Commit publishes the new snapshot then posts one generation-coalesced UI update per live view. Layout, paint, hit testing, UIA, and scroll input each consume one view-policy generation; the old scrollbar remains wholly valid until the UI thread installs the new generation, then the new geometry is complete. No mixed old hit-test/new paint frame is allowed, and hiding it never changes scroll position. |
| `bellStyle` | Existing views and bell effects parsed after commit | Bell admission is per view on the injected monotonic clock. The first effect is eligible; after an accepted effect, another is accepted only at or after `lastAccepted + 250 ms`. An earlier effect is consumed, increments one content-free dropped counter, and is never queued/replayed. Admission freezes view, view generation, policy generation, and `none|visual|system`; parser/model callbacks only enqueue bounded content-free work. `none` has no side effect. `visual` paints exactly one 2-DIP inner border for 150 ms, bound to that view generation: opacity decreases linearly from 1 to 0 when client animation is enabled, otherwise stays full until clearing; it uses `cursorArgb` normally and host-resolved high-contrast `foregroundArgb` when `TerminalTheme.highContrast != 0`; it changes no layout/focus/UIA/cell/snapshot. Committed `none`, view-generation change, or teardown cancels its timer/coalesced repaint. `system` increments the module/service quiet fence, obtains a pin through `AcquireModuleReferenceFromAddress(...)`, then calls `TrySubmitThreadpoolCallback` from the non-UI admission continuation. The callback first executes `FreeLibraryWhenCallbackReturns(instance, pin.release())`, calls `MessageBeep(MB_OK)` once, and decrements quiet only afterward. Submitted work is not canceled by a later setting/view teardown; quiet waits. Pin/submission failure consumes the effect, releases any pin/quiet ownership, emits no beep/retry/fallback, and increments one content-free failure counter; `MessageBeep` failure increments only that counter. OS mute is honored. |
| `hyperlinkPolicy` | Existing and new hyperlinks at each activation | `ctrl-click` activation first snapshots Ctrl state, policy, immutable hyperlink/document/view generations, and configuration generation under one gate; `disabled` rejects. The immutable OSC 8 URI must already fit the 8-KiB encoded-field cap. Before parsing, reject NUL, every ASCII C0 code point `U+0000..U+001F`, ASCII space, or absence of one explicit leading ASCII scheme plus `:`. `Terminal.dll` calls `CreateUri` once with reserved `0` and exactly `Uri_CREATE_CANONICALIZE | Uri_CREATE_CANONICALIZE_ABSOLUTE | Uri_CREATE_NO_DECODE_EXTRA_INFO | Uri_CREATE_NO_CRACK_UNKNOWN_SCHEMES | Uri_CREATE_NO_PRE_PROCESS_HTML_URI | Uri_CREATE_NO_IE_SETTINGS`; parse/canonicalization failure denies. Compare the canonical scheme ordinal-ASCII-case-insensitively against exactly `http|https|mailto`; `http|https` require nonempty canonical host and empty user-info/password, while `mailto` requires a nonempty opaque address/path. Credentials, every other scheme, malformed authority, relative text, or empty required component deny. Immediately recheck live policy and the exact view/document generations before the sole plugin-local `ShellExecuteW(pluginChildHwnd, L"open", validatedImmutableUri.c_str(), nullptr, nullptr, SW_SHOWNORMAL)` side effect; a disabling commit or generation change that wins this final gate revokes. Only return `>32` is success. Parameters/cwd are null, no process handle exists, and command processor/`CreateProcess`/concatenation/host URL policy are forbidden. Failure posts one localized status plus content-free diagnostic and performs no retry/fallback/second side effect. Hover styling is generation-coalesced and grants no authority. |
| `allowApplicationTitle` | Existing and new sessions | Each OSC title effect and a settings commit serialize through the session title gate. While false, title effects are rejected before publication. A `true -> false` commit clears the last accepted application title, restores the deterministic profile/default title, increments title generation, and queues one coalesced tab/UIA update before settings success is published. `false -> true` restores no old title; only a later valid OSC effect may set one. If an OSC wins first, disable clears it; if disable wins first, the OSC is rejected. |

Focused deterministic tests place a configuration commit on both sides of and
exactly at each gate: explicit-close intent/prompt, root-exit observation,
Open/Restart persisted-follow-default snapshot, follow-timer dispatch,
copy/input admission, scrollbar-generation installation, bell admission,
hyperlink initial and final activation gates, and OSC-title acceptance/reset.
Apply-false/true versus concurrent Open/Restart proves commit-first gets the new
default and admission-first gets the old, existing incarnation/override
stability, Restart resampling, Cancel no-effect, and session-toggle
nonpersistence. Other barriers prove a single linearization winner, no mixed
snapshot, no duplicate clipboard/input/external-open/bell effect, no stale
follow dispatch, no retroactive close-on-exit, complete title reset before
commit success, no title resurrection on re-enable, and only the explicitly
frozen bell/hyperlink revocations after first admission.

Bell tests use an injected monotonic clock/effect sink at first/249/250 ms
boundaries independently for two views; prove earlier effects increment one
content-free dropped counter rather than queue; exercise visual 149/150 ms,
exact 2-DIP/color/high-contrast geometry, enabled linear-fade versus disabled
static animation, no model/UIA/layout/focus mutation, stale timer, and
generation/`none`/teardown cancellation. Replace `MessageBeep` with a recorder
proving one `MB_OK` call on a non-UI callback,
`AcquireModuleReferenceFromAddress(...)` before
`TrySubmitThreadpoolCallback`, callback-entry
`FreeLibraryWhenCallbackReturns(instance, pin.release())`, OS mute/failure with
no fallback, pin/submit failure consumed with any pin/quiet ownership released,
no beep/retry/fallback, and one content-free failure counter; `MessageBeep`
failure increments only that counter. Cover already-submitted setting/teardown
races, callback/unload pressure, and quiet remaining false until `MessageBeep`
returns and the callback/module-pin transfer is complete. A
source-contract test rejects destroying the last `wil::unique_hmodule` in the
callback or omitting that callback-entry pin transfer. Hyperlink
tests cover the 8-KiB boundary/+1, NUL, every C0 class, ASCII space, missing/
malformed/non-ASCII scheme, and the exact one-call `CreateUri` flags/reserved-
zero contract. Cover ASCII case variants of `http|https|mailto`, `http|https`
empty host and user-info/password rejection, `mailto` empty opaque address/path,
every other scheme including `file` and shell/custom, relative/malformed
authority, Ctrl pressed/not pressed, document/view/link generation replacement,
and policy commit before/exactly-at/after both activation gates. The external-
open recorder
must observe at most one exact `ShellExecuteW` call with the plugin child HWND,
`L"open"`, immutable validated URI, null parameters/null working directory, and
`SW_SHOWNORMAL`; exercise `<=32` and `>32` returns, prove one localized failure/
diagnostic with no retry/fallback/second side effect, and source-contract reject
`CreateProcess`, command construction, a retained process handle, or host-owned
URL validation/opening.

Changing `warnOnUnsafePaste`, `pasteWarningLineThreshold`, `pasteMaxBytes`, or `inputQueueMaxMiB` increments the service paste-policy generation and, under the paste-authorization gate, revokes every pending confirmation and securely scrubs its frozen payload/preview before the effective change is published. A more-permissive result never approves or replays a request classified under an older generation. Generic plugin disable/re-enable and normal refresh are deliberately not paste-policy transitions: they neither increment this generation nor by themselves revoke/reject safe or already-pending unsafe paste on an existing selected interactive-visible frozen-configuration view. That continuity is still bounded by the broker's strict 60-second deadline and immediate selection/hide/deactivation/input-owner/supersession/lifecycle cancellation; returning to the view never resurrects a canceled prompt. Tests cover both the lifecycle continuity and these independent cancellation rules.

Every effective `osc52Policy` change and generic plugin disable/re-enable is an OSC 52 sensitive-effect policy transition; each increments `osc52PolicyGeneration` under the dedicated authorization gate. Accepted normal-refresh `BeginRefresh` is a revoking lifecycle transition bound by service generation. In every case, revoke/invalidate every ask or prompt-free-allow token and securely scrub its payload before publishing policy, availability, refresh, or teardown state. A more-permissive value never upgrades/replays an earlier request. While disabled, existing terminals remain usable but later OSC 52 writes are effective deny; re-enable applies configured policy only to later requests. The generic host receives no clipboard bytes and cannot publish success before the plugin-owned fence completes.

Preferences exposes product settings, privacy actions, and truthful capability/gated status. Apply/cancel uses a working configuration copy and hot reload calls the plugin transactionally. Startup, hot reload, and Preferences use the master's one pure closed parser: wrong-type/range/enum/Boolean/font fields receive its exact field repair plus degraded/compensation flow; whole-root/version/migration invalidity uses its complete safe policy; v1 has no cross-field-invalid setting combination because related valid limits compose through the named runtime lower caps. A failed live cap-lowering transition retains the prior committed configuration. `fontFamily` uses the master's exact 1..128 UTF-16-unit and empty/NUL/control/surrogate policy before missing-font fallback. Host code contains no terminal default/clamp/enum/cross-field table. Run the master's all-three-entrypoint corpus, including every Boolean/enum wrong type, every font invalid class, 128/129 units, and representative multi-limit combinations.

A later WSL-integration feature is not unlocked by configuration/package drift. It requires a separate approved design whose first gate proves clean-default Bash plus profile override/`exec`/early-exit behavior, bootstrap before input, zero visible PTY/native-history input, no Linux-profile edit, configured default-user/default-shell/startup preservation, no `/mnt/c` assumption, bounded state erasure, and safe terminal-only fallback. If it needs `--exec`, a wrapper, `WSLENV`, or another process, that design must explicitly replace the five-token invariant and freeze argv/environment/path/ownership/timeout/cleanup. V1 has no dormant/best-effort toggle, resource, or payload for that experiment.

For a Terminal Apply, one privacy field changing true-to-false selects its single-domain disable journal kind and both privacy fields changing true-to-false select the one `disable-history-and-shell-choice` kind, which scrubs both domains atomically in one family transaction. A candidate that changes one privacy field false-to-true while the other changes true-to-false is rejected before any journal, state, or Settings mutation with `HRESULT_FROM_WIN32(ERROR_INVALID_STATE)` and localized field guidance to Apply the disable first and then enable the other field. V1 never sequences two journals or creates a circular pending-enable.

### Pending-enable privacy fence

- A false-to-true privacy transition is the only operation that creates the
  master schema's `privacyTransition=kind="pending-enable"` record. Prepare
  snapshots exactly `Present` plus both generation UUIDs or an explicit
  UUID-free `Absent` sentinel retained through Finalize; only Arm may replace
  Absent once with generated identities. Arm acquires
  the family mutex and then the validated current-target mutex under the shared
  two-second deadline, enumerates/freezes the complete bounded family, and
  proves/rechecks journal absence plus the exact Present generations or still-
  absent sentinel. For Absent, every effective target-configuration true domain
  must be listed or Arm returns revision mismatch. It generates two distinct
  baseline UUIDs and computes the exact serialized empty/pending-state size with
  checked arithmetic. Absent Arm admits creation only when the resulting Release
  family has at most four active targets (Debug exactly one) and the complete
  family bytes are at most 512 MiB Release or 256 MiB Debug. It never retires an
  older or recovery-bearing target. A four-old-plus-absent-current Release family
  or byte-cap breach closes Arm before mutation with
  `HRESULT_FROM_WIN32(ERROR_QUOTA_EXCEEDED)`, content-free `family-cap`
  Retry/Reset status, no preparation consumption, and no Settings save. After
  that preflight it creates one empty v1 false/false state, gives every listed
  domain a mutually/baseline-distinct next UUID, preserves each unlisted baseline
  UUID/false, atomically includes pending in that first publication, and binds
  the generated identities into the armed preparation. A crash exposes either
  continued absence or one complete within-cap pending state, never a fifth/
  over-cap partial target. It writes/flushes/
  verifies pending false with the
  coordinator's target Settings digest. It releases target and then family
  before Arm success or any Settings CAS. The preparation retains only immutable
  transition/generation/digest identities; no mutex, thread, module pin held
  solely for that mutex, or other live transaction owner spans persistence.
- Finalize-Committed, Finalize-Failed, startup reconciliation, Retry, and
  `HostCompensationSave` continuation each independently reacquire the family
  mutex and then current-target mutex under the shared cancelable two-second
  rule, enumerate/freeze the complete family, reload and validate the exact
  durable state, and recheck family-journal absence. Before any pending
  completion/cleanup publication they compute the exact serialized replacement
  delta with checked arithmetic, revalidate every family byte/count cap, publish
  atomically, and release current then family. They require the exact transition ID, domains, expected/next
  generations, target digest, and current/representable target configuration
  true for every listed domain. Exact Committed rotates/sets true and clears
  pending. If an observer already completed exactly that transition, Committed
  recognizes no pending, every prepared next generation/listed capture true,
  unchanged unlisted identities/captures, journal absence, and exact persisted
  target/configuration as idempotent `Enable` bit-0 success. Failed clears only
  the matching pending record while retaining false.
- Any other missing/mismatched transition or unexpected true state must not
  request false compensation directly. Finalize releases both current and family
  mutexes, then restarts from family-first acquisition for the canonical ordinal
  target union required to establish the durable journaled false fence for every
  affected domain; it never expands the held lock set or inverts order. Only then
  may it return canonical false with bit 1; exact persisted false
  observed later may settle that journal. `Finalize(Failed)` carries zero
  persisted identity and its Prepare-captured baseline may be stale: when pending
  already became true or mismatched, Failed retains the family false journal in
  `awaiting-settings-false` with bit 1 and never removes it or infers current
  Settings false. Only clearing the still-exact pending record whose captures
  never became true may use Failed bit 0. If that fence cannot be established,
  return an ordinary noncompensating failure, close
  admission/require reload, and never claim rollback. `HostCompensationSave`
  consumes false only when the matching journal/pending false fence or an exact
  completed false state derived from it exists, never while capture is true with
  no durable family fence.
- Every observer of pending enable keeps capture false and waits only for the
  family mutex and then target mutex, never an in-memory owner. After the same
  complete-family freeze and exact replacement-delta/cap checks, an exact latest
  matching Settings digest may complete through the same atomic publication.
  An exact latest nonmatching observation atomically clears
  pending while retaining false through that family-gated publication and either accepts already-false policy or
  requests canonical-false compensation for persisted true. A concurrent
  in-flight CAS then conflicts, or Finalize observes the exact false abort,
  releases both locks, restarts the family/ordinal-union false-fence transaction,
  and only then compensates false. Conversely,
  an observer-completed exact target is idempotent `Enable` bit-0 success to the
  initiator. Stale/foreign observations return reload-required and
  mutate nothing; crash recovery requires no surviving in-memory lock or owner.

### Preferences privacy/status/action contract

Implement the master plan's family-wide privacy deletion and incompatible-state reset as one Terminal-plugin-owned contract. The generic Preferences host renders only the validated schema/status/action metadata and never enumerates state files, interprets Terminal state, chooses deletion targets, or performs a partial fallback.

- The exact Preferences v1 action IDs in the History & privacy section are
  `history.clear-all`, `shell-choice.forget-all`, and
  `state.reset-incompatible`. Every invocation carries the nonzero result
  generation of the latest QueryStatus taken/rendered for the same
  administration object/page, plus `contextVersion=0`, a null/zero context span,
  and an all-zero context digest. The plugin binds the action to that status
  publication's compact configuration/Settings/family token and rejects zero/
  foreign/not-latest input synchronously or a post-admission identity change
  asynchronously without mutation. Folder-scoped clear/forget commands remain
  in the plugin-owned history UI and are not additional Preferences action IDs.
- The generic schema-v2 action shape adds optional
  `visibleWhenStatusKey`, which must name a declared Boolean status. An action is
  hidden while status is loading/failed/false and visible only for a
  non-`unavailable` literal true; the plugin revalidates this gate at invocation.
  The schema declares `state.reset-incompatible` with
  `visibleWhenStatusKey="state.reset-available"`, `refreshStatus=true`, and
  destructive confirmation required. It is visible/invocable only for a
  repairable current-family condition whose persisted Terminal configuration is
  current/representable so both privacy fields can be compensated false. Future-
  configuration quarantine keeps Reset hidden and shows its upgrade/reinstall
  diagnostic; stale/wrong availability closes
  asynchronously with revision mismatch and no mutation.
- The localized confirmation states that Reset permanently deletes all Terminal command history, remembered shell choices, incompatible active state, and recovery copies for the current build flavor; Release means the complete retained Release family and Debug means Debug only. It states that the operation cannot be undone and leaves both memory domains disabled. Buttons are exactly localized `Reset` and `Cancel`, with Cancel the default. No path, command, profile, state/recovery filename, mutex digest, or malformed JSON is rendered.
- `BeginQueryStatus` returns exact IDs
  `state.recovery-file-count` (integer), `state.recovery-byte-count` (integer),
  `state.family-reason` (text), and `state.reset-available` (Boolean). The reason
  is exactly `normal|incompatible|deletion-incomplete|recovery-cap|family-cap|
  manual-cleanup-required`. Overlapping conditions select one reason with strict
  highest-first precedence `manual-cleanup-required > deletion-incomplete >
  family-cap > recovery-cap > incompatible > normal`; unsafe/repair-ceiling
  overflow never appears repairable, while a valid resumable journal dominates
  an ordinary cap/incompatibility diagnosis. Reset availability is true only for one of the four
  repairable nonnormal reasons plus representable current configuration, and is
  false for normal/manual cleanup/future configuration. Count/bytes
  include only exact owned recovery files and are bounded by the repair
  ceilings. These are the only recovery details exposed to Preferences,
  diagnostics, logs, metrics, or archived test output; status never includes
  filenames, paths, data excerpts, command/profile content, or raw failure text.
  Schema `valueKind` and bounds must agree with the operation ABI's exact JCS
  status grammar.

An absent Arm rejected only by its exact projected active-count/family-byte
equation retains one content-free, in-memory
`ProspectiveFamilyCapObservation` scoped to that module-administration
generation. It stores only the immutable failed-preflight family identity/
generation, build flavor, projected active delta, exact projected replacement-
byte delta, and one/two-domain bitmask—never a path, command, profile,
configuration bytes, or Settings digest. The failed preparation still receives
normal Finalize-Failed. A subsequent QueryStatus for the same administration
generation reacquires family then current under the shared deadline, revalidates
the family identity, and recomputes the checked stored equation. While it still
fails, the QueryStatus token binds `state.family-reason="family-cap"` and
repair-ceiling/current-configuration-qualified `state.reset-available=true`
even when the disk family alone remains at its legal boundary. Reset revalidates
the token and prospective equation before mutation. Any family, configuration,
Settings, availability, or administration-generation change, successful Arm/
Reset, different failed-Arm shape, or `Close` clears/supersedes the observation.
It is never persisted or restored after restart; restart reports the disk-only
reason, and retrying the same enable recreates the observation if the exact
projection still cannot fit.

The family transaction covers every retained active target and every exact owned recovery sibling of the current flavor, including a recovery whose derivable target active file is absent. Normal Release has at most four canonical active targets and 512 MiB across active files plus recovery siblings; Debug has one active target and a 256 MiB data cap. The journal is separately reserved overhead. Per canonical target leaf derived from the recovery filename, allow at most four exact owned `.bad.<UTC>.<lowercase-UUID>` siblings, each at most 64 MiB and at most 256 MiB aggregate, regardless of active-file presence. Count/size admission uses checked arithmetic; a safe fifth orphan sibling or safe orphan-derived aggregate above 256 MiB is `recovery-cap` and remains Reset-repairable within the separate repair ceilings, while an unsafe alias or 257th complete-family recovery entry is `manual-cleanup-required`. Active-file absence never relaxes admission. Family cap or safely enumerable incompatible state remains capture-off/read-only and exposes Reset rather than silently evicting opaque bytes. Unsafe/malformed identity or scope beyond the Reset repair ceilings remains capture-off/read-only with `manual-cleanup-required`; Reset cannot claim it fixed.

Add and corpus-test `Specs/Terminal/TerminalStatePrivacyTransaction.schema.json`. Its sole runtime journal is exact `<AppId>-release.terminal-privacy-transaction.v1.json` or `<AppId>-debug.terminal-privacy-transaction.v1.json`, JCS UTF-8, content-free, non-reparse, single-link, fail-if-unexpected, and at most 256 KiB. Freeze the exact root and enums from the master, including `kind=clear-folder-history|clear-all-history|disable-history|forget-folder-shell-choice|forget-all-shell-choices|disable-shell-choice|disable-history-and-shell-choice|reset-incompatible`. The three disable kinds target respectively `{command-history}`, `{shell-choice}`, and the ordinal set `{command-history,shell-choice}`; combined disable is one journal and one atomic family transaction, never nested single-domain journals. Also freeze null versus folder-only `scopeKeySha256`, ordinal active-state target records with expected/next domain generations and final capture Booleans, and ordinal unique `recoveryLeaves`. Each recovery leaf uniquely derives its canonical target state leaf by stripping its final exact `.bad.<UTC>.<UUID>` suffix; that active target may be absent, no active record is invented, and the leaf itself contains no path/content. `recoveryLeaves` records the complete post-lock actual-case owned-recovery set (zero through 256) and is the only recovery set deletion/resume may consume. Normal actions admit four Release/one Debug active targets; Reset may repair at most 64 active targets and 256 exact owned recovery siblings, and when the canonical current output is absent may add exactly one synthetic current-output target record with null expected generations. The 65th/257th, a second synthetic record, an unsafe/case alias, incomplete enumeration/deadline, or oversized journal is manual cleanup with no mutation.

The frozen journal also requires `phase=deleting|awaiting-settings-false`.
Clear/forget remains `deleting`. A disable starts `deleting`, performs and
verifies the complete family scrub, and removes the journal immediately only
when its bound persisted plugin configuration already has every Boolean in the
  kind's targeted set false; otherwise it atomically changes to
  `awaiting-settings-false` until the source-specific persistence/restart rule
  proves every target false. Reset starts
  `awaiting-settings-false` iff either bound persisted privacy Boolean is true
  and otherwise starts `deleting`. A live continuation requires an exact newer
  manager-local observation matching its retained compensation expectation and
  every targeted field false (false/false for Reset). After process loss, current
  canonical all-targeted-false `StartupLoad`—or a fresh Retry bound to latest current status
  when no live expectation remains—may settle only after full lock/journal/
  family revalidation and never compares manager-local sequences across
  processes. The journal gates every version in either phase.

Every active-state publication, recovery creation, import/retention, Reset, pending-enable transition, and privacy action acquires the protected flavor-family mutex first and enumerates/identity-checks the complete bounded family under one shared two-second deadline. A multi-target family transaction then acquires the unique union of recorded active/synthetic target mutexes and target mutexes derived from every recovery leaf in canonical ordinal path order; a one-target ordinary publication/recovery/pending transition acquires only that target mutex. An orphan recovery is thereby locked through its derivable absent-active target without a fabricated active record. Under both locks every one-target publisher reloads/merges its target, computes the exact serialized replacement delta with checked arithmetic, and revalidates journal absence plus every active/recovery byte/count cap. It publishes only when the post-replacement family remains within cap; otherwise it returns `HRESULT_FROM_WIN32(ERROR_QUOTA_EXCEEDED)` with no bytes changed. Active-active, active-recovery, import, Reset, and absent-Arm growth serialize on one family owner; exactly one boundary grower may commit and no conforming writer creates a transient over-cap family. Reads may remain target-local, but publication never acquires target first. Lock/enumeration/revalidation failure releases in reverse order and mutates nothing. The action gates the affected domain in its process, cancels pending mutations and chooser data, and must durably publish the journal before touching state bytes; every other process gates before its next read/use/write when watcher/preflight sees the journal. Journal-publication failure returns failure and retains the affected capture-off fail-safe. Ordinary writers reject or reschedule publication while that journal exists. After a scrub/journal commit that needs Settings persistence, release every target mutex in reverse order and then the family mutex before returning the staged result; no named lock is held across the generic coordinator's Settings CAS. On the qualifying exact newer persisted observation, reacquire the family mutex and then the ordinal unique union of journal-target and recovery-leaf-derived target mutexes under the same shared deadline, re-enumerate, and revalidate the unchanged `transactionId`, `kind`, `phase`, target records, recovery leaves, and safe identity set before phase advance, deletion, or journal removal. Only journal-authorized already-completed deletion absences are admissible. Any lock contention, mismatch, unrecorded/new owned entry, unsafe identity, or repair-cap/manual-cleanup condition mutates nothing, retains the cross-version journal fence, and leaves the family degraded/capture-off.

A one-time compatible older-Release import acquires family then the canonical
ordered current/source/retirement target union, selects the exact newest source
and exact oldest zero-recovery retirement candidates, freezes/revalidates the
family, and computes the exact serialized imported bytes with checked
arithmetic. It admits the atomic current-target publication only when the
complete new target plus every still-present old active/recovery byte already
fits the four-target/512-MiB caps before any retirement. It never deletes old
state to make room. Admission failure returns
`HRESULT_FROM_WIN32(ERROR_QUOTA_EXCEEDED)`, writes/deletes nothing, preserves
every source byte, and exposes content-free `family-cap` Reset/Retry. Only after
durable new-target publication may it make one non-gating best-effort pass over
prevalidated oldest targets. Crash before/within publication yields either all
sources with current absent or all sources plus a complete current target, both
within cap. Crash/delete failure during retirement yields the complete current
target plus a prefix-subset retired and remains an ordinary capture-enabled
within-cap family. It creates no durable cleanup marker or Retry/Reset debt, and
startup never infers a dead process's intended set; any later pass freshly
enumerates and selects its own safe candidates. A later absent-target import
that cannot fit fails before mutation with `family-cap`. No branch removes
opaque/unsafe state, publishes a partial target, or creates a transient
fifth/over-byte family.

- All-history clear/disable and all-choice forget/disable rewrite every valid retained active file, rotate the targeted deletion-fence generation, and preserve the opposite domain. Combined disable removes both domains, rotates both generations, and leaves both captures false in the same active-file rewrite and one `disable-history-and-shell-choice` journal. They delete every owned recovery sibling because opaque recovery bytes cannot be selectively scrubbed. Clear/forget verifies, removes the `deleting` journal last while locked, may restore its prior true capture, releases, and only then succeeds. A disable leaves every targeted state-capture Boolean false. If its bound persisted plugin value already has every targeted field false, it verifies the scrub/recovery deletion, removes the journal while locked, releases, and succeeds. Otherwise it atomically changes the journal from `deleting` to `awaiting-settings-false`, releases in the common order, and returns `DisableFenceCommitted` plus complete canonical all-targeted-false compensation; this is staged, not action success. Exact newer all-targeted-false persistence reacquires/revalidates through the common protocol, verifies the already-scrubbed state and absent recorded recovery leaves, removes the journal last, releases, and only then completes the action.
- Any incompatible active member blocks a normal domain-specific action.
  `state.reset-incompatible` is the sole v1 action allowed to delete future,
  wrong-format/flavor, corrupt, oversized, recovery-cap, family-cap, or
  deletion-incomplete state within its repair ceilings. If the bound persisted
  privacy configuration is not already false/false, it stops after publishing
  the `awaiting-settings-false` journal and deletes nothing. Exact
  `HostCompensationSave` acknowledgement advances to `deleting`; only then does
  it delete every recorded active/recovery file and create the current empty v1
  false/false state. It never touches the other build flavor, main Settings
  directly, or unrelated/unsafe files. If a valid deletion-incomplete journal
  exists, confirmed Reset first resumes/removes that exact journal and only then
  starts a separately journaled Reset; failed resume returns no Reset success.
  Malformed/unsafe journal or repair-ceiling overflow disables Reset and shows
  manual cleanup.
- Success is reported only after every resulting/deleted identity is reopened and verified, the directory is flushed through the shared transaction seam where supported, the journal is removed last, and locks are released. The schema-triggered QueryStatus refresh then must observe that final state. A stale current/older Release writer must reject changed/absent generations and can never recreate a removed target or resurrect data after downgrade.
- Reset also leaves both effective privacy configuration Booleans false. If
  persisted plugin JSON had either true, the staged InvokeAction returns the
  master's configuration-only compensation bit, complete canonical false/false
  plugin value, and outstanding expectation only after all file locks and then
  the family lock have been released. The generic coordinator performs
  one complete-snapshot revision CAS preserving availability/unrelated settings
  and acknowledges it through `HostCompensationSave`, which advances/deletes/
  verifies before full action success/status refresh, after the common
  reacquisition and exact journal/family-identity revalidation. Conflict gets
  the one bounded reload/reconcile retry. Write failure leaves the awaiting
  journal and data untouched; acknowledgement/reacquisition/deletion failure
  leaves false Settings plus the journal and hidden data. Both report the exact partial
  `Reset pending; settings/data remain safely disabled` and expose Retry; neither
  lets stale Settings true re-enable capture or claims deletion prematurely.
- Any crash or failure after journal publication is `deletion-incomplete`, never success. Keep every transaction-affected capture false (both domains for combined disable and Reset), hide cached data, retain the journal and content-free diagnostic, and permit one bounded off-UI startup resume plus Retry Reset. A compatible older binary observes the same journal fence and cannot read, capture, write, or recreate state around it. A live awaiting disable/Reset continuation requires its exact newer local observation plus retained expectation. After process loss, canonical all-targeted-false `StartupLoad` or a fresh latest-status Retry with no live expectation may settle only through the common family/file reacquisition and complete journal/identity revalidation, never by comparing `hostLoadSequence` with the dead process. Restored A/B/A false safely settles; any restored/current targeted true retains the journal and canonical false compensation. Resume treats a missing recorded target/recovery leaf as an already-completed deletion only where authorized by the journal and treats any unrecorded owned active/recovery leaf as an unsafe post-fence identity change. Settings CAS/access failure, plugin generation/quarantine failure, lock contention, identity mismatch/manual cap, or any delete, verification, recreation, flush, or journal-removal failure returns action failure and preserves the capture-off/read-only journal fence; unsafe/malformed journal or file identity requires manual removal and is never claimed fixed.

Schema/corpus/native tests must cover all three action IDs and zero-context rules; zero/foreign/stale/latest QueryStatus-generation binding and post-status identity races; generic Boolean-status action visibility, Reset visibility and stale/wrong-availability rejection; exact destructive confirmation/default button; exact status IDs/value kinds/bounds/ordering, all six reasons, every pair plus representative three-way overlap proving strict precedence `manual-cleanup-required > deletion-incomplete > family-cap > recovery-cap > incompatible > normal`, resulting Reset availability, and redaction; valid/invalid privacy journals and every kind/phase/scope/target/recovery-leaf/generation cross-field rule, including the one combined-disable journal and its ordinal two-domain target set; exact recovery suffix stripping plus malformed/ambiguous rejection; orphan-recovery-only families; orphan per-derived-target count/byte `N` and `N+1` with exact `recovery-cap`, 257th-family/unsafe exact manual-cleanup outcomes, and Reset recovery; the sole synthetic null-expected current-output record when current active is absent; Release four-target and Debug isolation; every recovery/family count/byte boundary and +1; separately reserved journal; concurrent cross-target recovery admission; Reset at active/recovery repair ceilings and manual cleanup at +1; future/corrupt/oversized/unsafe state; normal opposite-domain preservation; all-recovery deletion including orphans; partial-family failure; crash/resume after partial orphan deletion; watcher miss/poll recovery; two-process stale-writer rejection; and upgrade-purge-downgrade non-resurrection. Prove family-first ordinal acquisition and reacquisition of the unique active/synthetic/recovery-derived target-mutex union, including orphan-derived absent-target locks, plus recovery create/rename/delete races and identity/suffix changes before resume. At exact family boundaries, race active growth on two targets, recovery creation against active growth, and absent Arm against old-target growth; every permutation proves one family owner, exactly one grow commit, `ERROR_QUOTA_EXCEEDED`/no-write for the loser, and no transient over-cap state. Release-import cases prove canonical family/current/source/retirement locking, exact source/zero-recovery retirement selection, all-old-sources-counted imported-byte admission at count/byte boundary and +1 before mutation, quota failure with current absent/byte-identical sources/no deletion plus `family-cap` Reset/Retry, atomic absent-or-complete current crash states, and post-publication prefix retirement crash/delete failure that leaves the complete current target capture-enabled, creates no cleanup marker or Retry/Reset debt, and requires any later hygiene pass to re-enumerate and reselect fresh under family/target locks. Pending-enable tests cover Prepare Present/two-UUID versus UUID-free Absent, current target initially present/absent, family-then-current shared-deadline acquisition, complete-family freeze, Arm-first/family-action-first orderings, absent-to-present race rejection, target-config true-domain completeness, two distinct false baselines, mutually/baseline-distinct listed next UUIDs, unlisted baseline false, binding generated IDs into the armed preparation, exact serialized first-publication size with checked arithmetic, post-create active-count and 512/256-MiB family-byte boundary admission, four-old-plus-absent-current and byte-cap +1 rejection before mutation with `ERROR_QUOTA_EXCEEDED`, `family-cap` Retry/Reset status, unconsumed preparation and no Settings save, recovery-bearing would-be retirement with no eviction, atomic crash outcome of absence or one complete within-cap pending state, exact prospective-observation creation, same-admin QueryStatus ordered-lock/equation revalidation and token-bound Reset, family/configuration/Settings/availability/admin-generation/success/new-shape/Close clearing, restart nonpersistence and retry recreation, atomic first pending-false create/write/flush, reverse target/family release before Arm success/Settings CAS, no retained live owner, independent Finalize-Committed/Failed/startup/Retry/observer/HCS family-then-current acquisition, complete-family freeze, exact replacement-delta/cap revalidation, atomic publication, current-then-family release, and release-both/restart-family-then-ordinal-union on mismatch, exact transition/generation/digest/journal revalidation, matching completion, Failed cleanup, crash before/after CAS, timeout/access/abandonment, journal conflict, and unsafe/missing/changed state. Race observer, Finalize, and HCS pending completion/cleanup against boundary growth and family privacy work; prove one-way lock order, no transient cap excess, no partial pending publication, and deterministic retry/fence result. Assert exact result bits: Arm and successful enable bit 0; exact observer completion before initiator Finalize is idempotent `Enable` bit 0; only clearing a still-exact never-true pending record may use Failed bit 0; `Finalize(Failed)` zero persisted identity/stale Prepare baseline cannot remove a family false journal after pending became true or mismatched, so that branch retains `awaiting-settings-false` with bit 1 until a later exact persisted observation; mismatched persisted true may yield bit 1 only after a durable family false journal/fence. Inject fence-creation failure and require ordinary noncompensating failure, closed admission/reload, and no false expectation consumable by `HostCompensationSave`; capture true without a durable fence can never acknowledge false. Race exact matching and nonmatching observers against in-flight CAS and require observer-wins success or conservative pending-false abort plus CAS conflict/fenced Finalize compensation respectively. For single-domain history/shell-choice disable, combined history-and-shell-choice disable, and Reset, distinguish live newer-expectation settlement from post-crash `StartupLoad`/fresh latest-status Retry settlement; the combined path uses one atomic scrub/journal and settles only after both persisted fields are false through restart/HCS/Retry. Inject every combined scrub/crash/compensation boundary and prove no partial success or nested journal. Reject opposing false-to-true/true-to-false Apply before any mutation with `ERROR_INVALID_STATE` and localized two-step guidance. A/B/A and restored false settle after full revalidation without cross-process sequence comparison, while any restored/current targeted true retains the journal and false compensation. Also test the direct already-persisted-all-targets-false path and every scrub/journal/CAS/generation/crash/older-binary/lock/reacquisition/identity/manual-cap/phase/delete/recreate/verify/journal-removal boundary; full success is impossible before the source-specific false rule and journal removal. Every failure retains the applicable journal and hidden capture-off/degraded state without false success or cached-data exposure.

### Normal plugin refresh and disable/re-enable lifecycle

- Implement the availability states `LoadingConfiguration|EnabledReady|EnabledDegraded|EnabledConfigurationBlocked|DisabledRetained` and the separate retirement states `Idle|RefreshWaitingForTerminalObjects|RefreshSettlingConfiguration|RefreshWaitingForQuiet|RefreshBlocked|Unloaded`. Generic plugin disable is not retirement: it persists the one availability value, closes new terminal admission, and retains running terminals plus the module-scoped administration/service generation. Existing live views retain safe Paste and a still-selected/interactive-visible/input-owning/unexpired broker-eligible unsafe Paste with unchanged paste-policy generation under frozen configuration; disable itself does not revoke it, but every broker cancellation remains effective. Before `DisabledRetained` publishes, the plugin-owned OSC 52 gate revokes/scrubs outstanding ask/allow payloads; later OSC 52 requests are effective deny. Metadata, schema, current configuration, status, and re-enable UI remain reachable. With retirement `Idle`, re-enable reuses that generation, reconciles the latest persisted snapshot, increments admission, and exposes enabled status only after success; there is no second Terminal Boolean and no OSC 52 replay.
- Explicit confirmed Manage Plugins Refresh is the sole normal-runtime retirement route and calls generic `TerminalPluginManager::BeginRefresh(nonzeroManagerIssuedApplicationLifetimeGeneration)`. `S_OK` atomically changes `Idle` to `RefreshSettlingConfiguration`; the same generation returns `S_FALSE`; zero/stale/unissued returns `E_INVALIDARG`; another active generation returns `HRESULT_FROM_WIN32(ERROR_BUSY)`. Cancel before Begin has no effect. After `S_OK`, old-generation retirement is irreversible: no cancel, queued refresh, old-generation reopen, or second module/service generation is allowed.
- Begin captures the old module generation, increments/closes terminal/configuration/deep-link/status/action/page admission, and retains only host-copied schema/metadata plus content-free refresh status. New public calls reject with `HRESULT_FROM_WIN32(ERROR_SHUTDOWN_IN_PROGRESS)`. Existing terminals receive no `RequestClose`, hide, disconnect, callback clear, Close, or release and remain usable until ordinary user/lifecycle removal. Their safe Paste and still-selected/interactive-visible/input-owning/unexpired broker-eligible pending unsafe Paste remain usable with unchanged paste-policy generation; refresh itself cannot dismiss/revoke/reject that interaction, while selection/visibility/activation/input-owner/expiry/lifecycle cancellation remains final. Separately, the OSC 52 authorization fence revokes/scrubs outstanding ask/allow payloads before refresh publication. Their unbounded `RefreshWaitingForTerminalObjects` barrier is independent of both five-second clocks.
- In parallel, the retained `PluginConfigurationCoordinator` enters `Retiring`, suppresses page presentation, and settles only already-accepted work. Cancel/take every operation that has not crossed `CommitStarted`; before Settings CAS, abort the complete participant group and give each successful Prepare/Arm exactly one `Finalize(Failed)`; after CAS begins, await the actual CAS result and give every successful participant exactly one matching `Finalize(Committed|Failed)`. Complete commit-won reconcile/action/finalize work, every typed result, required compensation save, and matching `HostCompensationSave` acknowledgement before callback clear/drain, nonblocking `Close`, and release. Missing/idle administration completes without invented work. Neither callback nor private settlement admission is cleared while any preparation, CAS, result, compensation, or continuation remains.
- The configuration-settlement deadline is five seconds from accepted Begin and does not include terminal lifetime. Expiry is nonfatal `RefreshBlocked(ConfigurationSettlementTimeout)`: retain the old module/coordinator, leave all new admission closed, and show exact host-owned localized `Restart RedSalamander to finish plugin refresh`. Mandatory private continuations may complete safely, but completion only enables explicit Retry and never silently unloads or loads new code.
- After configuration settlement and natural release of every old-generation public terminal object, call the standard shutdown export once and enter `RefreshWaitingForQuiet`. Poll `CanUnloadNow` on generation-checked UI turns no more often than 50 ms. This second five-second clock begins at export. True unregisters the resource owner and resets the old owning `wil::unique_hmodule` before any replacement path is opened/mapped. False at expiry is nonfatal `RefreshBlocked(ModuleQuietTimeout)`, retains the old mapping, and stops automatic polling. Normal refresh never invokes the application fatal policy or queries/honors process-exit retention.
- Expose host-owned stable `plugin.refresh.retry` only in `RefreshBlocked`; it is not a Terminal schema action and never depends on the retiring plugin page. For `ConfigurationSettlementTimeout`, return `HRESULT_FROM_WIN32(ERROR_BUSY)` without a plugin call until mandatory settlement and both public-object barriers complete. Otherwise one explicit Retry starts one bounded five-second quiet window; false returns to the same blocked outcome. There is no background retry, repeated automatic polling, or unload/load race.
- Settings enable/disable accepted during refresh persists normally and updates one latest-wins `desiredAvailabilityAfterRefresh`. It never invokes the retiring administration object, cancels refresh, or reopens/queues the old generation; other Terminal configuration edits remain unavailable. After old-handle reset, re-resolve/verify the current path and map at most one replacement, create one administration object, and reconcile the latest complete Settings snapshot before exposing its page or terminal admission. Desired disabled reaches `DisabledRetained`; desired enabled reaches `EnabledReady|EnabledDegraded` only after success. Missing/wrong-architecture/corrupt/ABI-incompatible/startup-failed replacement leaves no old executable generation and reaches visible `EnabledConfigurationBlocked` with reinstall/Retry guidance.
- `BeginProcessShutdown` supersedes every normal-refresh state. It joins already-complete barrier/export steps without duplication, force-closes remaining terminal objects through the application-shutdown contract, and uses its own outer deadlines. Only process shutdown may choose the exact passive-UIA retention outcome or fatal policy; normal `RefreshBlocked` never weakens that rule.
- Implement this contract across `RedSalamander/TerminalHost/TerminalPluginManager.h/.cpp`, `RedSalamander/PluginModuleLifecycle.h/.cpp`, `PluginConfigurationCoordinator.h/.cpp`, `ManagePluginsDialog.h/.cpp`, `Preferences.Plugins.cpp`, `Preferences.Plugin.Configuration.h/.cpp`, project/filter files, localized resources, and Commands/native lifecycle tests. Reuse the one loader/coordinator; do not duplicate export lookup, Terminal configuration parsing, or module state.
- Deterministic tests cover disable with running tabs and in-flight Preferences work without retirement; idle re-enable reuse; every BeginRefresh return/precedence case; terminal objects outliving both clocks; configuration pause before/after `CommitStarted` and CAS; exact Finalize/compensation/result settlement and late safe completion; last-object release before/after settlement; immediate/deferred quiet; both distinct blocked reasons; Retry busy/one-window outcomes; desired-availability churn in every state; replacement load/reconcile failure; old-handle reset before one replacement; no overlapping mapped module/service generations; no post-fence public request; and application shutdown superseding every state. Assert normal refresh never calls retention/fatal seams, abandons a callback/private transaction, or silently resumes after timeout. With deterministic transition barriers, prove generic disable/normal refresh alone leave a selected interactive-visible eligible view's safe Paste and pending unsafe prompt/token usable and generation-stable, while selection/deactivation/expiry still cancels without resurrection, OSC 52 ask/allow is revoked/scrubbed before transition publication, disabled later requests are denied, and re-enable/replacement never replays an old request.

### Configuration administration during process shutdown

- Main `WM_CLOSE` calls the generic
  `TerminalPluginManager::BeginProcessShutdown(validatedMainTarget,
  nonzeroApplicationLifetimeGeneration)` before destroying a FolderWindow,
  main posted-message target, or pump. Its first valid call captures manager
  `shutdownT0`; closes terminal/configuration creation, Preferences-deep-link,
  status/Retry, and page-presentation admission; increments admission
  generation; and snapshots both the retained module-scoped
  `IPluginConfigurationOperations`/coordinator generation when present and all
  public `ITerminal` objects. The terminal-object and configuration-object
  barriers proceed concurrently and never block the UI thread.
- The coordinator marks its administration generation `Retiring`, suppresses
  page output while keeping the manager callback/dispatcher and private
  settlement lane live, calls `CancelOperation` for each accepted pre-commit
  operation, and awaits/takes/releases every result. If no global Settings CAS
  started, it aborts the complete participant group and consumes every
  successful Prepare/Arm with exactly one `Finalize(Failed)`. Once CAS starts it
  is noncancelable: await its actual outcome and issue exactly
  `Finalize(Committed)` or `Finalize(Failed)` to every successful participant.
  Commit-won reconcile/action/finalize work, each transaction-bearing result,
  and every required fail-safe/compensation save plus matching
  `HostCompensationSave` acknowledgement complete; QueryStatus results are
  validated, taken, and discarded. No preparation, result, compensation, or
  continuation is abandoned or inferred from current Settings.
- Only after no accepted operation, unconsumed preparation, host CAS,
  compensation, retained result, or coordinator continuation remains may the
  manager call `SetCallback(nullptr,nullptr)`, synchronously drain, call
  nonblocking idempotent `Close`, and release the administration object.
  `Close` returns `HRESULT_FROM_WIN32(ERROR_BUSY)` without retirement mutation
  while transaction-bearing state remains; encountering that result after the
  callback was irreversibly cleared is a host invariant breach and invokes the
  application fatal policy. Missing/idle administration completes without
  manufacturing operations. The standard module shutdown export runs only
  after both public-object barriers; the outer two-/five-second bounds start at
  `shutdownT0` and cannot be reset by polling or settlement.
- Deterministic shutdown tests cover administration absent/idle and every
  accepted Prepare/Arm/pre-CAS/in-CAS/post-CAS/Finalize/compensation/result
  phase, cancellation wins/loses, exact Failed/Committed finalization,
  compensation acknowledgement, retired-page suppression, result-buffer
  retry/release, callback-clear/Close/release ordering, `ERROR_BUSY` invariant
  failure, duplicate/stale shutdown, zero-terminal export start, export only
  after both barriers, one continuation, and the no-UI-wait/two-/five-second
  deadline contract.

## Security completion

- Treat PTY bytes, images, titles, URLs, clipboard effects, and integration messages as untrusted.
- Keep file/temp/shared-memory Kitty, clipboard reads, notifications, automatic URL/file/process side effects, unsupported media/protocols disabled.
- Unsafe paste is one immutable plugin-owned authorization transaction: exactly one UI-thread `CF_UNICODETEXT` read, bounded UTF-16 validation and strict UTF-8/CR normalization, frozen payload/classification/preview, complete service/session/incarnation/view/paste-policy/input-mode identity, opaque one-shot token, one pending confirmation/session, fixed 8 MiB/session and 32 MiB/service pending-storage caps, newer-request supersession, strict 60-second broker deadline, and secure full-capacity scrub. Approval never rereads the clipboard; it snapshots current selected-engine paste mode and, under the gate, revalidates all identity/generations/liveness plus current paste/input/descriptor/queue caps before one atomic complete event or zero bytes. Setting/input-mode/lifecycle and broker selection/visibility/activation/input-owner/expiry transitions revoke before publication/teardown. Generic disable and normal refresh alone preserve safe and a still-selected/visible/eligible pending unsafe paste for an existing frozen-config view; they do not waive broker cancellation.
- OSC 52 ask/allow remains bounded/asynchronous and never blocks callbacks. Every supported write has immutable service/session/incarnation/view/policy/effect identity, opaque token, plugin-owned securely scrubbed payload, one outstanding ask/session, and eight/service. A background/inactive ask denies/scrubs before counters and makes no focus/selection change; an admitted ask has the strict 30-second broker deadline. Newer same-session request supersedes; ninth retained distinct-session request denies. Policy changes, disable/re-enable, normal refresh, newer request, selection/activation loss, close/restart/retirement, and shutdown revoke/scrub before publication/teardown; disabled means effective deny. Approved ask revalidates the full broker/domain identity, presentation generation, strict deadline, interactive-visible ownership, current policy, enabled nonrefreshing live Running state, and then performs the sole clipboard replacement while holding the same gate; prompt-free explicit allow follows its authorization gate without becoming an ask. Both consume once on success/failure, and stale/duplicate/revoked replies make no clipboard call. V1 hyperlinks allow only ASCII-case-insensitive `http`, `https`, and `mailto`; every other scheme has zero external side effect. `Terminal.dll` performs the only open as `ShellExecuteW(pluginChildHwnd, L"open", validatedImmutableUri.c_str(), nullptr, nullptr, SW_SHOWNORMAL)` and treats only `>32` as success; failure produces one localized diagnostic with no retry/fallback/second side effect. Parameters/cwd remain null, no process handle exists, and host URL policy, `CreateProcess`, or command concatenation is forbidden. Ordinary oversized optional effects are consumed then ignored/denied with one content-free counter/coalesced status and parsing continues; authenticated integration overflow permanently invalidates that PowerShell/pwsh incarnation's sole epoch until explicit Restart/new session and rejects a second handshake; required-response exhaustion in the fixed protocol/control reserve remains session-fatal after any admitted control-placeholder charge. Never truncate/reclassify an oversized record.
- Enforce every exact non-Kitty boundary before buffer growth, sanitize/bound titles/status, use checked arithmetic everywhere, reserve allocations before work, and run the selected-engine fuzz/allocator/decompression/teardown corpus.
- No command, screen, nonce, clipboard, path, profile executable, distro, URL, image, or user name in logs/metrics/crash annotations/archives.

### Final Windows startup-bootstrap contract

- Cmd starts from the validated local/plugin-backed requested directory, or from
  the API-derived safe local system directory for UNC, and preserves normal
  AutoRun ordering. One versioned reviewed `.cmd` is embedded as `Terminal.dll`
  RCDATA with resource ID/version/SHA-256 and extracted through the per-user
  owner/access/reparse-safe transaction; immediately before every launch the
  plugin reopens and hashes it. The exact child token array is
  `["cmd.exe","/E:ON","/V:OFF","/S","/K",
  "call \"<verified-extracted-script-path>\""]`. The last token comes from one
  closed two-pass `cmd /S /K`+`call` path encoder and contains no target or
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
  variant, then writes exactly `REDSALAMANDER_TERMINAL_CMD_V1_PIPE`,
  `REDSALAMANDER_TERMINAL_CMD_V1_CAPABILITY`,
  `REDSALAMANDER_TERMINAL_CMD_V1_NONCE`,
  `REDSALAMANDER_TERMINAL_CMD_V1_OPERATION_ID`,
  `REDSALAMANDER_TERMINAL_CMD_V1_LAUNCH_GENERATION`,
  `REDSALAMANDER_TERMINAL_CMD_V1_MODE`,
  `REDSALAMANDER_TERMINAL_CMD_V1_TARGET`, and
  `REDSALAMANDER_TERMINAL_CMD_V1_TARGET_UTF16_SHA256`. `PIPE` is the exact full
  random-suffix name; capability is an independent 256-bit `BCryptGenRandom`
  value encoded as 64 lowercase hex; nonce and operation ID are independent
  128-bit values encoded as 32 lowercase hex; generation is canonical nonzero
  `uint64_t` decimal; mode is exactly `LOCAL|UNC`; target is the copied validated
  UTF-16 logical path; and digest is 64 lowercase hex for SHA-256 over the
  target's exact UTF-16LE code units without BOM or terminating NUL. Target with
  NUL, CR, LF, or `"` is invalid; spaces, Unicode, `%`, `!`, `^`, `&`, `|`, `<`,
  `>`, and parentheses are literal required fixtures.
- After normal AutoRun, the script starts exactly with
  `setlocal EnableExtensions DisableDelayedExpansion`, validates/stages only
  those fixed keys, and direct-expands the target once into exactly
  `cd /d "<target>"` for local/plugin-backed or `pushd "<target>"` for UNC. It
  uses no `call set`, delayed/recursive target expansion, child process,
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
  exit, AutoRun exit/hang, or startup cleanup failure remains Starting and
  records no preference. Success closes/drains the pipe and clears capability,
  nonce, operation, digest, target, and all other secret/target-bearing staging
  before the same incarnation may form `CmdLaunchRootedRunning`, admit input, or
  write profile memory. Through root lifetime it retains only
  `{drive, expectedRoot, expectedTail, launchGeneration,
  prelaunchLetterWasFree}`; this one-use correlation result is never an ongoing
  semantic epoch. Every local or UNC
  Windows PowerShell/pwsh instead uses `PowerShellStartupBootstrap`; cmd's
  protocol is never reused for PowerShell and both exchanges run even when
  semantic integration is disabled.
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
  malicious same-user process are outside the stated threat model. Tests document
  that limitation while proving wrong PID/capability/nonce/operation/generation/
  digest, replay, duplicate line, and PTY output fail closed.
- Before a PowerShell process, create a fresh 128-bit bootstrap key, 128-bit
  operation ID, nonzero 64-bit launch generation, and one duplex/overlapped/
  first-instance/message-mode/one-instance 256-KiB-in/out
  `\\.\pipe\RedSalamander.TerminalBootstrap.<32-lowercase-hex-random>` with
  `PIPE_REJECT_REMOTE_CLIENTS`, noninherited current-user-owned DACL granting
  only that SID read/write/synchronize, and exact root-PID verification through
  `GetNamedPipeClientProcessId`. Keep this key distinct from cmd and semantic
  nonces.
- Both shell families use exactly four total child tokens:
  `[fixedBasename,"-NoExit","-Command",bootstrapWrapperCommand]`.
  `-NoProfile`, `-NonInteractive`, `-ExecutionPolicy`, `-File`,
  `-EncodedCommand`, and trailing tokens are forbidden. Generate the wrapper
  only from the reviewed immutable template plus verified script path, mode,
  bootstrap pipe/key, operation ID, generation, and protocol `1`; never include
  target or semantic nonce. The sole literal encoder rejects NUL/CR/LF/unpaired
  UTF-16, doubles apostrophes, and wraps one apostrophe pair; interpolation,
  expandable strings, backticks, expressions, environment lookup, and pane/user
  text are forbidden. Its canonical ASCII grammar is one top-level try:
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
  catch [System.Exception] { exit 125 }`, with placeholders replaced only by
  those literals. Initialization, assignment, validation, and finalizer
  invocation all remain inside the catch boundary, so a profile-created read-
  only/AllScope collision exits without a prompt. Fixed `*>&1` captures every
  actually emitted warning/verbose/debug/information stream as extra result
  evidence and never redirects binary pipe traffic. Whitespace is the one
  canonical ASCII form frozen with the resource, placeholders denote only the
  literal tokens above, and exact generated-token/AST tests cover both shell
  families. Normal profiles run exactly once before the wrapper. The
  `[CmdletBinding(PositionalBinding=$false)]` script accepts exactly six named
  bootstrap parameters plus common `-ErrorAction Stop`, validates the same
  closed version `1`, `Local|Unc` mode, pipe, two exact 32-lowercase-hex values,
  and canonical nonzero unsigned-decimal generation grammar, and rejects
  duplicate or unbound arguments. After authenticated CommitPrompt it applies the final
  disposition, prebuilds/HMACs CommitAck, zeroes the last key buffer, and returns
  exactly one typed two-property finalize object whose first PSTypeName is
  `RedSalamander.TerminalBootstrapFinalize.V1`, whose exact ordered properties
  are `Sentinel,Finalize`, whose Sentinel is
  `RSTBOOT-WRAPPER-READY-1`, and whose Finalize is a one-shot scriptblock. That
  private closure owns only the bootstrap client and immutable prebuilt Ack and
  performs exactly one synchronous message-mode Write of the complete frame,
  with no Flush, retry, later pipe operation, or pipeline output. It alone sends
  CommitAck, only after wrapper validation. Missing/wrong/duplicate/extra
  output, invalid object/finalizer, parse/invocation/policy failure, finalizer
  throw/partial Write, or named exception leaves Ack unaccepted and exits
  125/126 without an interactive prompt. Client disposal is idempotent and
  output-silent, suppresses named disposal errors, and cannot change a written
  record; a failed/partial Write is never repaired. Server gates open only after
  reading one complete valid frame, never from client Write completion.
- Every bootstrap frame is exactly `magic[8]="RSTBST01"`,
  `frameSize@8:uint32`, `version@12:uint32=1`, `kind@16:uint32`,
  `flags@20:uint32=0`, `directionSequence@24:uint64`,
  `launchGeneration@32:uint64`, `operationId@40:uint8[16]`,
  `payloadBytes@56:uint32`, `reserved@60:uint32=0`, payload at 64, then the
  32-byte HMAC-SHA256 over all prior bytes. Size is exactly
  `64 + payloadBytes + 32`, at most 262,144; each direction starts at one,
  advances once, and never wraps. Integers are little-endian; text is an exact
  uint32 UTF-16-unit count plus UTF-16LE with no terminator/NUL/unpaired
  surrogate. Bad/unknown/trailing/duplicate/replay/oversize/tag/identity/
  sequence data is fatal.
- The only nine-kind exchange is child `Ready=1` sequence 1/empty; plugin
  `RootRequest=2` sequence 1/exact request; child `RootResult=3` sequence 2 with
  `{status:uint32 Success=0|Failure=1,cwdCodeUnits:uint32,cwdUtf16le}` and zero
  cwd on failure; plugin `SemanticDecision=4` sequence 2; child `Prepared=5`
  sequence 3; plugin `ReleasePrompt=6` sequence 3/empty; child `ReleaseAck=7`
  sequence 4/empty; plugin `CommitPrompt=8` sequence 4 with exactly
  `{finalDisposition:uint32 TerminalOnly=0|SemanticActive=1,
  policyGeneration:uint64}`; child
  `CommitAck=9` sequence 5/empty. After normal profiles the script calls only
  module-qualified `Microsoft.PowerShell.Management\Set-Location -LiteralPath
  <exact RootRequest>` and obtains cwd only through module-qualified
  `Microsoft.PowerShell.Management\Get-Location`, so profile aliases/functions
  cannot intercept or add pipeline output; the plugin independently requires
  canonical logical equality. No command
  concatenation/evaluation/interpolation/provider fallback or history entry is
  allowed.
- After matching RootResult, sample latest effective integration value and
  generation once. `SemanticDecision` begins
  `{decision:uint32 TerminalOnly=0|SemanticEnable=1,
  policyGeneration:uint64}`. TerminalOnly has no nonce/pipe. SemanticEnable
  additionally sends the one fresh semantic nonce and already-created bounded
  semantic-pipe name; target arrives only in RootRequest, and semantic material
  only here, never in environment/initial argv. `Prepared` is exactly
  `{outcome:uint32 TerminalOnly=0|SemanticReady=1|SemanticUnavailable=2,
  policyGeneration:uint64}`. Unsupported/replaced hook may report Unavailable
  and leave an inert pass-through wrapper for truthful terminal-only launch.
- `Prepared(SemanticReady)` and semantic handshake have no cross-transport
  order. One bounded `StartupSemanticJoin` owns two optional slots keyed by exact
  launch/session/incarnation/profile/executable/location, semantic nonce/epoch,
  Decision generation, and startup generation. Either may arrive first; each is
  independently authenticated and fully bounds/sequence/schema/admission-
  validated, and neither grants capability alone. Only their matching
  conjunction strictly before the common deadline advances. The handshake is
  semantic sequence one and no other semantic event precedes conjunction. Both
  handlers use the same session lock only to reserve/fill/scrub fixed slots and
  post one generation-checked continuation; neither waits, parses, writes a
  pipe, invokes a callback, or takes a model/writer/host lock while holding it.
  Reserve both slots before child creation so parser-callback admission never
  blocks; exhaustion is deterministic fatal startup failure.
  Duplicate/conflict/gap/wrong identity/oversize/queue failure is fatal, revokes
  the staged epoch, and scrubs both. TerminalOnly requires both slots absent and
  any Handshake paired with it is fatal. Unavailable normally requires no
  Handshake; its sole exception is one structurally/authentically valid matching
  sequence-1 Handshake staged before the server receive deadline and paired with
  child Unavailable under the exact server-read-before/client-observation-at-or-
  after-write-deadline race. That frame is scrubbed/revoked and completes
  terminal-only nonfatally; every malformed/mismatched/extra frame remains
  fatal. Cancellation/root exit/timeout/Restart/shutdown scrubs both. Live true-
  to-false alone becomes nonfatal `RevokedTerminalOnly`: close semantic/control
  operation and capability admission and scrub/ignore staged/late semantic
  records, but retain both Starting pre-CAS servers solely as exact-root-PID, no-
  operation cleanup transports. Permit exact-PID connect/setup, perform no
  control writes, stop requiring a handshake, and accept otherwise-valid
  Prepared Ready or Unavailable only as cleanup proof; missing Prepared still
  times out. Close/cancel both pipes at Running CAS or earlier only for an
  independent retirement/protocol failure. Final
  `CommitPrompt(TerminalOnly,currentPolicyGeneration)` requires process-lifetime
  inert/pass-through latch and removal of the control mapping iff still app-owned
  before Ack. False-to-true changes no prior TerminalOnly decision.
- Through Prepared the session is Starting and input/profile memory/trusted
  activity/capabilities are closed. ReleasePrompt makes the script remove
  target/decision/semantic variables, zero their mutable buffers, retain only
  one module-private mutable bootstrap-HMAC key, and send ReleaseAck. It then
  does nothing except await authenticated CommitPrompt; EOF/error/cancel/
  unknown/timeout first zeroes the key and fatally exits without returning to
  the interactive loop. ReleaseAck alone commits nothing. Under the session
  lifecycle lock, revalidate all immutable identities/generation/root/deadline/
  no-retirement-winner and final policy, then perform the sole
  `Starting/BootstrapReleaseAcked -> Running/PromptCommitPending` CAS against
  Close/WindowClosing/Restart/root-exit/timeout/ApplicationShutdown. CAS loss
  sends no CommitPrompt and yields no prompt/input/memory/capability. CAS success
  makes public Running while all gates remain closed.
- After CAS, close/cancel both startup-revoked `TerminalControl` and
  `TerminalSemantic` servers before sending authenticated
  CommitPrompt with CAS-frozen disposition/generation. The exact matrix is:
  original TerminalOnly plus Prepared TerminalOnly remains TerminalOnly at any
  equal-or-later generation; Enable plus Ready with a completed join is Active
  only at unchanged generation; Enable plus Unavailable is TerminalOnly at
  unchanged generation; and either valid Enable outcome after a latched disable
  is TerminalOnly at its current strictly greater generation regardless of a
  later re-enable. The script retains only immutable Decision generation plus
  Prepared outcome through CommitPrompt; decrease/wrap/every other pair is
  fatal. TerminalOnly latches wrappers inert/removes an app-owned mapping;
  SemanticActive requires the completed join. Only then may the script prebuild/
  HMAC CommitAck, zero the last key, and return the sole typed finalize object;
  it does not send Ack. The wrapper validates the exact object and invokes its
  private one-shot output-silent finalizer, whose one synchronous message-mode
  Write sends the complete prebuilt Ack with no Flush/retry/later operation.
  Accept Ack only after the server reads the full valid frame, by
  same Running incarnation/deadline, not live policy generation. A latched
  disable before or after CAS/CommitPrompt/Ack still consumes matching Ack,
  cleans bootstrap, opens ordinary input, and permits integration-independent
  Windows profile memory, but publishes zero capability/rejects old events and
  never retires solely for disable. Later enable after TerminalOnly likewise
  consumes Ack, remains terminal-only, and reports Restart required. Without
  downgrade, publish only proved capability. Only an unrelated lifecycle/
  protocol/root/deadline/wrapper-finalization failure after the Running CAS but
  before accepted Ack retires the already-Running
  gated incarnation. A post-CAS disable closes/cancels both pipes immediately;
  completed `TerminalControl` EOF/error latches every interposer inert and a
  pending or next `TerminalSemantic` WriteAsync fails, even if a built
  CommitPrompt said SemanticActive. The exact nonblocking poll, pre-fence writer,
  queued-input ambiguity, and quiet-drain contracts below govern the race. Thus first-
  prompt evaluation never occurs for a never-Running launch.
- One monotonic `launchT0 + 5 seconds` deadline covers connection through
  CommitAck, strict-before with equality timeout. Pre-CAS protocol/profile/root/
  retirement failure other than the defined live-disable exception uses the
  existing arbiter and never publishes Running/memory; unrelated post-CAS
  failure uses ordinary already-Running retirement with startup gates closed.
  Live disable alone follows the nonfatal cleanup path above. No late record/
  setting resurrects either. Never bypass
  execution policy. Drop immutable-string references and zero mutable buffers,
  but explicitly exclude same-user OS-command-line/.NET-heap forensics and never
  claim physical erasure of immutable strings.

### Final PowerShell semantic transport write contract

- The one semantic epoch owns one
  `\\.\pipe\RedSalamander.TerminalSemantic.<32-lowercase-hex-random>` inbound
  server. It is `PIPE_ACCESS_INBOUND | FILE_FLAG_OVERLAPPED |
  FILE_FLAG_FIRST_PIPE_INSTANCE`, message-mode, one-instance, has a 64 KiB
  inbound buffer, uses `PIPE_REJECT_REMOTE_CLIENTS`, and is in overlapped
  Connect before `SemanticDecision`. Its noninherited protected descriptor
  is owned by the current-user SID and grants only LocalSystem full access and
  that SID write/synchronize access. The module connects write-only and the
  server accepts only when `GetNamedPipeClientProcessId` equals the exact root
  PowerShell/pwsh PID. Before `SemanticEnable`, the plugin charges the session/
  process memory ledgers and reserves exactly one fixed 2 MiB semantic-frame
  staging buffer plus fixed Connect/Read OVERLAPPED state. One
  `ConnectNamedPipe` is pending before the decision; only after it completes and
  exact PID validation succeeds does the plugin arm the first `ReadFile`.
  Claiming a pre-connect pending read is forbidden. The same one client/server
  connection persists for the entire epoch and is never reconnected. One read is
  pending whenever the connected epoch is idle. A completed message is validated
  from the reservation, staging is reset, and the next read is armed before any
  resulting semantic action is dispatched. Reserve/connect/arm/rearm failure
  closes the server and permanently revokes the epoch; parser admission never
  allocates or waits. The distinct outbound `TerminalControl` pipe is only app-
  to-root request transport, never an alternate marker path.
- One semantic frame is one pipe message emitted by exactly one
  `NamedPipeClientStream.WriteAsync(byte[], 0, frameLength,
  CancellationToken)` call. Construct the client with local server `.`, only the
  decision-delivered canonical `RedSalamander.TerminalSemantic.<32-lowercase-
  hex-random>` leaf, `PipeDirection.Out`, `PipeOptions.Asynchronous`, and
  `TokenImpersonationLevel.Anonymous`; no arbitrary full path is accepted. The
  server accumulates only that one message in the fixed staging buffer. Each
  `ERROR_MORE_DATA` completion continues that same message and the final
  successful `ReadFile` completion delimits it; EOF is never a delimiter. EOF
  with nonempty staging is truncation, clean unexpected EOF permanently revokes
  the epoch, and EOF after an installed lifecycle/policy fence is normal cleanup.
  Filling 2 MiB while `ERROR_MORE_DATA` reports another byte is the 2-MiB+1
  fatal case; no prefix is parsed. Layout is exactly `magic[8]="RSTSEM01"`,
  `frameSize@8:uint32`, `version@12:uint32=1`, `kind@16:uint32`,
  `flags@20:uint32=0`, `sequence@24:uint64`,
  `semanticNonce@32:uint8[16]`, `payloadBytes@48:uint32`,
  `reserved@52:uint32=0`, payload at offset 56, then a 32-byte HMAC-SHA256 over
  every preceding byte using the semantic nonce. `frameSize` includes the tag,
  equals `56 + payloadBytes + 32`, and is at most 2 MiB. Unknown kind/flag/
  reserved/trailing byte, bad length/tag/nonce, partial message, or over-cap byte
  permanently invalidates the epoch. Sequence starts at 1 with the sole
  Handshake and advances exactly once per successful later ordinary message.
  Ordinary frames 2 through N on the same connection are valid and required; a
  second Handshake is fatal and sequence never wraps. Before startup join
  completion only the one Handshake is legal; any extra semantic frame is fatal.
  No selected-engine tagged effect, PTY/OSC/APC output, second
  pipe, or second VT parser authorizes an update.
- The semantic payload schema is closed to exactly
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
  frames do not change that relation.
  PriorStatusAvailable is forbidden without CompletedAccepted, and
  PriorCommandSucceeded requires both other flags. `cwdKind` is exactly
  `UnavailableOrNonFilesystem=0|WindowsFilesystem=1`: kind 0 requires zero cwd;
  kind 1 requires 1..131,072 strict-UTF-8 bytes containing the byte-identical
  canonical Windows drive/UNC logical path produced by the Location-identity
  canonicalizer. No other provider/cwd spelling is accepted.
- `AcceptedInput=3` is exactly `payloadVersion@0:uint32=1`, `cwdKind@4:uint32`,
  `readCycle@8:uint64`, `commandBytes@16:uint32`, `cwdBytes@20:uint32`, command
  bytes at offset 24 followed immediately by cwd bytes; exact size is checked
  `24 + commandBytes + cwdBytes`. `readCycle` matches the RootReadReady cycle
  that opened this exact still-active ReadLine call even when intervening
  ControlResult frames exist. Command length is `0..1,048,576` bytes; cwd uses
  the same kind/strict-UTF-8/131,072-byte rules and is the filesystem cwd captured
  at acceptance. The maximum valid AcceptedInput payload is therefore exactly
  1,179,672 bytes and remains below the semantic-frame payload cap. Command bytes
  are strict UTF-8 for the exact returned .NET
  `String`; empty and multiline CR/LF/HT are allowed, while BOM, NUL, unpaired
  UTF-16, every other C0, and every C1 control are rejected. The payload never
  normalizes, trims, or reconstructs command text.
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
- The one canonical JSON wire-schema source is exactly
  `Plugins/Terminal/ShellAdapters/TerminalSemanticProtocol.v1.json`; it is
  compiled into an identity-bound `Terminal.dll` resource and generates both C++
  and embedded-PowerShell constants at the two exact checked-in Generated paths
  above. The build's generator `-Check` plus exact corpus/test paths prove byte-
  identical enums, offsets, masks, bounds, and cross-field tables. The closed
  corpus covers every kind's valid min/max cases, astral/multibyte text, unknown
  kind/status/flag, nonzero reserved, bad version/manifest identity, illegal
  capability combination, offset/length overflow/underflow/trailing bytes,
  invalid UTF-8/control/path/version, `1 MiB + 1`, cwd `128 KiB + 1`, full frame
  `2 MiB + 1`, sequence/read-cycle wrap, wrong completed sequence, and wrong
  control ID/kind/data combination.
- Before production Ready, every supported real Windows PowerShell 5.1/pwsh
  tuple proves from its loaded runtime API and real message-mode pipes that this
  exact four-argument `WriteAsync` overload exists, returns without a synchronous
  large-write stall, and preserves one message boundary through the 2 MiB cap.
  A failing tuple is terminal-only. Synchronous Write, helper process, background
  runspace, alternate runtime leaf, and second transport are forbidden.
- Every received message has a strict assembly deadline. `receiveT0` is the QPC
  sample captured with that message's first completed read chunk and
  `receiveDeadline = checked(receiveT0 + ceil(qpcFrequency/20))`. The final
  successful `ReadFile` must complete strictly before that 50-ms deadline;
  equality or overflow closes/cancels the connection and revokes the epoch. No-
  first-chunk Handshake remains bounded by the existing strict
  `launchT0 + 5 seconds` startup deadline: connection alone does not start the
  50-ms assembly clock because adapter installation follows connect. Later
  messages use the same first-chunk-plus-50-ms rule. A never-read/stalled message
  cannot retain staging forever.
- Every emission owns one fresh cancellation-token source, one frame byte array,
  and the returned Task. A module-injected monotonic Stopwatch starts before
  serialization/HMAC and defines `writeT0 + 50 ms`. Success requires the sole
  Task to be `RanToCompletion`, synchronously observed, with the post-observation
  sample strictly before the deadline; equality is timeout. The foreground hook
  may wait only the nonnegative whole-millisecond remainder rounded down and
  never schedules a wait beyond the deadline. On Windows PowerShell 5.1 that
  wait is only `((IAsyncResult)task).AsyncWaitHandle.WaitOne(floorRemainingMs)`;
  `Task.Wait`, `.Result`, a waiting `GetAwaiter`, spin/poll loops, continuations,
  and callbacks are forbidden. Only successful observation advances sequence
  and permits buffer zero/release.
- Synchronous throw, faulted/canceled Task, timeout, disable, close, Restart,
  root exit, or teardown cancels the token, immediately closes the client stream
  so the server observes lifecycle-fenced EOF, and latches every installed
  adapter hook process-lifetime inert. There is no retry or replacement frame
  for that failed emission, fallback transport, Task continuation/callback,
  helper thread, or background runspace.
  `AcceptedInput` still returns the exact user line and `RootReadReady` still
  proceeds to the underlying shell read; other failed marker emissions return
  ordinary raw shell control. The module privately retains the Task, token
  source, and byte array until a later owned foreground entry or teardown sees
  completion, then observes any named exception, zeroes mutable bytes, disposes,
  and releases them. No overlapping semantic write is admitted while that record
  exists. Teardown applies the same cancel/close rule and never waits in a prompt
  hook; bounded root-process shutdown is the final cleanup boundary if the Task
  does not settle. Failed handshake emission before `Prepared` makes the partial
  hook inert and reports `Prepared(SemanticUnavailable)`; it is neither a retry
  nor terminal-startup failure. Until Prepared, child outcome is authoritative:
  a structurally/authentically valid matching sequence-1 Handshake staged by the
  server before its receive deadline but paired with
  `Prepared(SemanticUnavailable)`—including server-read-before/client-observation-
  at-or-after-write-deadline—is scrubbed/discarded, closes/revokes the staged
  epoch, and completes terminal-only nonfatally. `Prepared(SemanticReady)` still
  requires that matching staged Handshake. Malformed/wrong identity/nonce/HMAC/
  sequence, a pre-Prepared extra semantic frame, or Handshake plus
  `Prepared(TerminalOnly)` remains fatal.

### Final semantic/control revocation and foreground-control contract

- `TerminalSemantic` is inbound/write-only from the root module; it has no child
  read and cannot be the pre-emission EOF sentinel. `TerminalControl` is the
  distinct app-to-root channel whose root-side `BeginRead` supplies that sentinel.
  A pre-CAS true-to-false transition immediately marks
  `StartupSemanticJoin=RevokedTerminalOnly`, permanently closes semantic/control
  operation and capability admission, and scrubs/ignores every staged or late
  old-generation semantic record. It retains the already-created
  `TerminalSemantic` and `TerminalControl` servers solely as exact-root-PID, no-
  operation cleanup transports through Prepared and the Running CAS; they may
  finish exact-PID connect/setup and the plugin performs no control write, so a
  policy change never manufactures a client setup failure. Either otherwise-
  valid Ready or Unavailable Prepared proves cleanup; missing Prepared times out.
  At the winning Running CAS, before `CommitPrompt(TerminalOnly)`, close both
  server admissions and cancel their outstanding operations. A disable that wins
  at/after that CAS, Close, Restart, root exit,
  application shutdown, or independent protocol retirement installs the session
  generation/admission fence and cancels both immediately. A handle with no
  outstanding I/O closes immediately; an in-flight handle follows the exact
  cancel/drain/close ownership rule below rather than closing under its
  OVERLAPPED storage.
- After closing admission, cancel and drain every accepted Connect/Read/Write
  OVERLAPPED completion, fixed frame-staging/parser unit, semantic/control
  dispatch unit, reserved/armed control frame, and posted result. Each owns the
  session operational gate and plugin module pin through scope-exit release;
  RAII handles, OVERLAPPED storage, fixed buffers, and parser/dispatch gates stay
  counted through actual final release. `CanUnloadNow` remains false until both
  pipes and root/process quiet are complete. The existing off-UI five-second
  quiet deadline governs the drain; a miss uses the existing per-session
  quarantine or application-shutdown fatal policy and never unloads reachable
  code.
- The root adapter calls output-silent `PollControlTransport(NoWait)` at every
  ordinary owned read-wrapper entry. The integration-owned key-handler entry alone
  calls `PollControlTransport(HandlerWaitOnce)` and may perform the sole bounded
  wait frozen below. `NoWait` touches an already-complete read only when the sole
  staged-record slot is empty, validates/stages and rearms, and never dispatches.
  `HandlerWaitOnce` first consumes an occupied stage with zero wait/`EndRead` on
  the new read; only an empty stage permits it to inspect/wait on that current
  read. Either path invokes `EndRead` exactly once only when the current
  `TerminalControl` `BeginRead` was already complete or the handler wait signaled.
  Zero bytes/EOF or a named read
  error latches every installed hook process-lifetime inert before it emits or
  dispatches new work. A complete valid frame is retained as the one staged
  control record and the next `BeginRead` is armed before any action. Closing
  `TerminalSemantic` breaks an in-flight child
  `WriteAsync`, or the next attempted emission when none was pending; that
  failure follows the same inert/no-retry rule. Pipe closure does not pretend to
  recall a control frame whose plugin writer admission already won, but every
  result/semantic frame received after the plugin fence is rejected and grants
  no state, history, or capability.
- The one `TerminalControl` server is outbound, overlapped, first-instance,
  message-mode, one-instance, has 256-KiB inbound/outbound buffers, uses
  `PIPE_REJECT_REMOTE_CLIENTS`, and has a protected noninherited descriptor owned
  by the current-user SID. Its only current-user ACE is exactly
  `ReadData|ReadAttributes|WriteAttributes|Synchronize`; `WriteData` is absent.
  The verified
  post-profile module opens the local server through the exact noninheriting
  `NamedPipeClientStream` constructor with
  `PipeAccessRights.ReadData|PipeAccessRights.ReadAttributes|
  PipeAccessRights.WriteAttributes|PipeAccessRights.Synchronize`,
  `PipeOptions.Asynchronous`,
  `TokenImpersonationLevel.Anonymous`, and `HandleInheritability.None`, then sets
  `ReadMode=Message`. `WriteAttributes` exists solely for that mode change and
  permits no pipe-data write. Every supported Windows PowerShell 5.1/pwsh tuple
  must prove this exact constructor/API path against the production DACL or
  report control unavailable.
- The root module owns one fixed 262,144-byte `TerminalControl` buffer, exactly
  one bounded validated-control-frame staging slot, and always keeps one
  `NamedPipeClientStream.BeginRead` pending without a PowerShell callback, helper
  thread, or background runspace. Ordinary owned wrapper entries use
  `PollControlTransport(NoWait)`: only with an empty slot, they may `EndRead` an
  already-complete read, validate it into the slot, and rearm, but never wait or
  dispatch; an occupied slot leaves the current read untouched. Only the
  integration-owned PSReadLine wake handler may use
  `PollControlTransport(HandlerWaitOnce)`, perform the one bounded wait frozen
  below, or execute a staged control frame. After owner/binding/epoch/lifecycle
  revalidation it first atomically takes an occupied slot, with zero wait and
  zero `EndRead` on the newly armed read. Only when the slot is empty may it call
  `EndRead` exactly once for an already-complete or signaled read, validate into
  that slot, arm the next BeginRead, and take the slot before any action.
  The child never branches on unobservable plugin-side `Armed|InFlight` state.
  `FlushFileBuffers`
  and every synchronous pipe Write/Flush are forbidden. The plugin reserves the
  fixed frame buffer and freezes its immutable payload before `operationT0`; it
  does not claim a complete frame exists before that sample. Its off-UI control
  worker holds only the generation/input-operation gate, performs final
  validation, atomically reserves one inert fixed-protocol-reserve descriptor
  plus its eventual one-byte wake, assigns that placeholder its final input-
  queue sequence, samples `operationT0`, serializes/HMACs into the reservation, then
  submits exactly one message-mode OVERLAPPED `WriteFile` and releases that gate
  immediately after submission. Failure to reserve either fixed unit is a
  mechanically proved pre-submit/no-action outcome. Close/cancel calls `CancelIoEx` for that exact
  OVERLAPPED but keeps the handle open. It performs one off-UI manual-reset-event
  drain wait bounded only by the existing session quiet deadline, then calls
  `GetOverlappedResult(..., FALSE)` once to observe terminal normal/aborted
  completion. Only then may it close/release the handle, event, OVERLAPPED, frame,
  and pin. A quiet miss retains every owner and handle under quarantine/fatal
  policy. This drain admits no wake/effect/result, and no detached worker exists.
- A control frame is exactly `magic[8]="RSTCTL01"`, little-endian
  `frameSize@8:uint32`, `version@12:uint32=1`, `kind@16:uint32`,
  `flags@20:uint32=0`, `controlSequence@24:uint64`,
  `deadlineQpc@32:uint64`, `qpcFrequency@40:uint64`,
  `operationId@48:uint8[16]`, `payloadBytes@64:uint32`,
  `reserved@68:uint32=0`, payload at offset 72, then a 32-byte HMAC-SHA256.
  `frameSize` includes the tag, equals `72 + payloadBytes + 32`, and is at most
  262,144. Kind is `Probe=1|FollowSetLocation=2|InsertWithoutAccept=3`; Probe has
  zero payload and the other kinds have nonzero UTF-8 without NUL/CR/LF bounded
  by the frame and owning path/input cap. Sequence starts at one for Probe,
  increments exactly once, and never wraps. HMAC uses the sole 128-bit semantic
  nonce over every byte before the tag and is fixed-time compared. After final
  state validation and immediately before serialization, sample `operationT0`
  and set `deadlineQpc = checked(operationT0 + ceil(qpcFrequency/4))`; reject a
  nonpositive frequency or overflow before any Write. OVERLAPPED completion,
  same-descriptor wake activation, and the handler's final pre-effect sample must all be
  strictly before that 250-ms deadline; equality expires without action, and the
  root requires its `Stopwatch.Frequency` to equal `qpcFrequency`.
- Exactly one control operation is `Armed|InFlight`. Under the generation/input-
  operation gate, user input that linearizes before placeholder reservation
  cancels the undispatched operation. Reservation winning first assigns the
  inert placeholder the earlier immutable FIFO sequence; every later user
  descriptor receives a later sequence and the writer stops at that head.
  `writer-admitted` is exactly successful OVERLAPPED submission
  (`WriteFile` returns `TRUE` or `FALSE/ERROR_IO_PENDING`), not completion, client
  read, placeholder reservation, or wake activation. A reserve/capacity failure,
  user-input fence, or disable fence winning before submit converts any reserved
  placeholder to a consumed zero-byte tombstone before later input is writable,
  issues no Write, has zero shell effect, and does not advance `controlSequence`.
  Submission winning first is the sole
  pre-fence winner and may cause at most one shell action before embedded expiry,
  even if disable wins before completion or wake. Each operation owns one manual-
  reset OVERLAPPED event. A `TRUE` `WriteFile` is the immediate-completion path;
  `FALSE/ERROR_IO_PENDING` is followed by exactly one off-UI
  `GetOverlappedResultEx` with the nonnegative whole-millisecond remainder to the
  250-ms deadline rounded down and `bAlertable=FALSE`. `INFINITE`, polling,
  spinning, a second wait, and UI-thread wait are forbidden. Completion is
  accepted only for the exact frame byte count and a post-return QPC sample
  strictly before the deadline. Completion plus an epoch-live revalidation must
  both win strictly before the deadline before the already-sequenced placeholder
  is atomically activated in place with exactly one wake byte under the normal
  input-writer gate. Activation performs no second capacity check, admission,
  sequence assignment, or priority bypass; there is no retry, timer, or second
  wake. Write success does not prove the client's
  `BeginRead` completed. Normal successful write completion releases only that
  per-write event/OVERLAPPED/frame/write pin; the epoch-persistent server handle
  remains open, and neither the control-operation record nor later-user-input
  hold releases before a matching authenticated `ControlResult`. Timeout/
  equality/short write/other error revokes and
  activates no wake and retains the placeholder as an inert held fence; call
  `CancelIoEx` on the exact OVERLAPPED and keep the handle
  open. One off-UI event drain wait, bounded by the session quiet deadline rather
  than the admission deadline, precedes one nonwaiting
  `GetOverlappedResult(..., FALSE)` terminal-completion observation. Only then
  close/release. Retain the handle/event/OVERLAPPED/frame/module pin plus any
  ambiguous input gate throughout; a quiet miss quarantines/fails fatally as
  applicable and never frees or unloads reachable work. “No second wait” forbids
  a second admission/deadline wait, not this one ownership drain, which authorizes
  no wake/effect/result. Disable after submission admits no later wake, while
  later user input remains held until the closed result/no-action disposition
  below. Persistent channel states are exactly `Connecting -> Idle ->
  WritePending -> AwaitingResult -> Idle`; only Idle admits an operation. On exact
  successful write completion, one input-operation-gate transaction publishes
  AwaitingResult plus the complete expected result identity before its final
  release step activates/makes runnable the placeholder. The writer cannot emit
  from WritePending, so even a result arriving before the completion callback
  returns observes AwaitingResult. A publication-fence loss activates nothing and
  retires the submitted ambiguity. A submitted failure/disable/timeout leaves an inert fence;
  explicit Discard/confirmed Close retires it without emitting the wake. A fully validated matching
  Success, or for Follow/Insert only a matching pre-effect
  `RejectedState|BufferChanged`, releases the operation/input hold, advances the
  no-wrap control sequence, and returns Idle; the latter two report safe no-
  action failure. A mechanically proved pre-submit/no-action loss returns Idle
  without advancing sequence. Probe non-Success, Expired, BindingChanged,
  ShellOperationFailed, timeout, malformed/mismatched result, protocol/policy
  failure, or unresolved ambiguity enters Retiring, closes through the exact
  cancel/drain contract, latches terminal-only, and admits no later control
  operation; quiet success reaches Retired and a quiet miss reaches Quarantined/
  fatal shutdown. Strict expiry proves only
  that no future action can begin; it does not prove whether a pre-expiry Follow/
  Insert effect already occurred. An ambiguous pre-expiry effect or lost result
  therefore keeps queued input held, leaves the incarnation terminal-only/
  untrusted, and offers only `Discard queued input` or confirmed Close. Disable
  never auto-releases that ambiguous input, and I/O/frame/module gates remain
  held through actual completion.
  `Discard queued input` is a plugin-owned one-shot action identified by
  `{sessionId,incarnation,inputHoldGeneration,operationId}` and is enabled only
  after the strict expiry/admission fence enters Retiring and makes wake
  activation irrevocably impossible. Under the generation/input-operation gate
  it revalidates identity/state, captures the queue tail, converts any remaining
  placeholder to a permanent zero-byte no-wake tombstone at its original
  sequence, tombstones and securely scrubs every held normal-user/paste/
  integration descriptor through that tail, preserves required engine responses
  in FIFO, releases the post-wake hold, and increments a no-wrap input-recovery
  generation. The epoch remains terminal-only/untrusted and only newly admitted
  raw input may flow. Late completion/result, stale/duplicate action, and the
  Close race are inert/serialized by the existing fences; physical I/O/module
  owners still drain through ordinary quiet and are not released by recovery.
- Ordinary owned wrapper entries use `PollControlTransport(NoWait)` and may only
  stage/rearm an already-complete frame when the sole slot is empty; they never
  dispatch. On key-handler entry only, `PollControlTransport(HandlerWaitOnce)`,
  after exact owner-token, key-binding, epoch, and lifecycle revalidation, first
  take and dispatch/scrub an occupied slot once with no wait or `EndRead` on the
  newly armed read. Only with no staged frame does an incomplete current
  BeginRead perform exactly one nonalertable
  `((IAsyncResult)read).AsyncWaitHandle.WaitOne(250)` and no other wait, callback,
  `Task.Wait`, `.Result`, blocking `GetAwaiter`, poll, or spin. A signaled/already-
  complete read is consumed by exactly one `EndRead`; zero-byte EOF or a named
  read error latches every child hook inert, while one complete frame is validated
  into the sole slot and the next BeginRead is armed before the handler takes and
  dispatches it. An unsignaled timeout performs
  zero action, emits zero semantic/result frame, and does not itself revoke child
  integration. If a plugin operation exists, its missing-result/strict embedded
  250-ms deadline retires and closes both pipes, waking the read wait; if none
  exists, a physical press may block only this root shell for at most 250 ms and
  integration remains live. After a complete read, the handler's final QPC sample
  and one-shot consume must be strictly before the embedded deadline; equality or
  later scrubs without action. A completed pre-fence frame may be dispatched once by
  the admitted wake or physical press only while exact operation ID, sequence,
  one-shot ownership, binding, empty buffer/cursor-zero state, and strict deadline
  still match. Exactly one injected wake bounds a mapping-replacement race:
  replacement after child/plugin validation may invoke the later foreign handler
  at most once, after which missing result retires the epoch; no retry invokes it
  again. At/equality-after expiry scrub without action before control-
  mapping removal/pass-through. Malformed/partial/oversized frame, wrong HMAC/
  sequence/operation/deadline/client, result mismatch, or aggregate admission
  failure after a completed read permanently revokes the epoch. A post-fence result or semantic frame grants no
  plugin state/history/capability. If action completion cannot be proved after a
  lost result, retain later user input and offer only the already-defined
  `Discard queued input` or confirmed Close path.

### Final per-view interaction-broker contract

- Exactly one plugin-owned broker exists per committed view. Its closed state is
  `None|CloseConfirm|UnsafePasteConfirm|Osc52Ask`, with one nonzero monotonic
  `interactionGeneration`; its content-free visible token binds
  `{serviceGeneration,sessionId,sessionIncarnation,viewGeneration,
  interactionGeneration,kind,domainToken}`. Reply validation precedes domain-
  token consumption. Wrap, stale/duplicate/wrong-kind/wrong-view, missing or
  failed posted payload invalidates and scrubs the incoming domain request and
  never reuses a generation. Sensitive bytes remain exclusively in the owning
  paste/OSC gate.
- `interactive-visible` requires a visible direct child and ancestors, current
  exact instance/view ownership marker, non-minimized foreground non-closing
  Folder root, and the host-revealed auxiliary child. Show/window-position,
  focus/activation, and ownership changes advance the presentation generation;
  visibility, foreground root, ownership, and generation are rechecked before
  presentation and reply without querying host model internals.

The exact closed table is mandatory; no canceled lower-priority interaction is
restored:

| Current broker state | Incoming Close | Incoming unsafe paste | Incoming OSC 52 `ask` |
|---|---|---|---|
| `None` | Present only for a matching close intent/token and interactive-visible view | Present only if eligible; else cancel/scrub | Present only if interactive-visible and counters reserve; else deny/scrub |
| `CloseConfirm` | Exact duplicate returns `ERROR_BUSY`; stale/other Close is ignored | Reject/scrub new Paste; retain Close | Deny/scrub new OSC; retain Close |
| `UnsafePasteConfirm` | Cancel/scrub Paste, then present still-valid Close | Newer same-view Paste cancels/scrubs old before replacement; replacement failure leaves `None` | Deny/scrub new OSC; retain Paste |
| `Osc52Ask` | Deny/scrub OSC, then present still-valid Close | Deny/scrub OSC, then present eligible Paste | Newer same-session OSC denies/scrubs old before replacement under the service cap |

- For a background Terminal close glyph, preserve generic `DxUi::TabControl`
  semantics: capture stable tab ID and host-view generation; select that exact
  tab; reveal/layout and ordinarily focus its child; after every reentrant step
  re-resolve and require the same selected, foreground interactive-visible,
  ownership-valid view; only then issue UserTab. Failure before the request
  leaves the race winner unchanged. After successful selection there is no
  rollback: Busy, Retry, or Cancel leaves the clicked Terminal selected. A
  required confirmation on a noninteractive child returns `ERROR_RETRY` without
  an intent/broker allocation; invisible close is forbidden. An admitted Close
  has no elapsed timeout while interactive-visible, but any selection/
  visibility/foreground/ownership/generation loss performs exact Cancel.
  Approved/direct
  close uses normal MRU/Folder fallback only after the shared removal claim.
- Background/hidden/inactive OSC asks deny and scrub before per-session/eight-
  service counters, never select/reveal/focus or queue for activation, and may
  retain only coalesced content-free status. A visible ask expires on the strict
  monotonic-QPC `now < deadline` rule at 30 seconds; equality denies. Paste uses
  the same strict rule at 60 seconds. Selection/visibility/foreground/
  ownership/presentation/input-owner/policy/lifecycle loss cancels immediately;
  returning never resurrects a prompt, domain token, payload, or UIA element.
  Reorder preserves only an unchanged stable selected interactive-visible view.
- History and the terminal context menu open only with broker `None`; Close
  dismisses them before admission, OSC ask while either is open denies/scrubs,
  and `FocusOwnerKind=PluginTextInput` Paste acts only on that owned UI. Chooser/
  diagnostic base states receive no Running-only Paste/OSC request.
- Close/Paste modal UI records prior owned child focus, disables terminal input,
  initially focuses Cancel, confines Tab/Shift+Tab, invokes only the focused
  button on Enter/Space, and maps Escape to Cancel. Restore terminal-document
  focus after Cancel or successful Paste only for the same selected,
  interactive-visible, input-capable broker/view/presentation generations;
  Close approval, selection loss, teardown, or newer interaction restores
  nothing. OSC ask never steals focus, raises one polite UIA notification, and
  accepts only exact-token mouse/UIA Invoke or displayed bar-scoped `Alt+A`
  Allow/`Alt+D` Deny. Hidden/canceled/expired provider calls return
  `UIA_E_ELEMENTNOTAVAILABLE`.
- Window/application shutdown, callback detach, forced/plugin/automatic
  retirement, Restart, and view removal close broker admission first, advance
  its generation, dismiss child/UIA, route live tokens through Cancel/Deny and
  secure scrub, then drain `PostMessagePayload` registry entries before child
  teardown. Generic disable/normal refresh preserves a still-eligible Paste
  only because it is not itself a paste-policy transition; every independent
  broker cancellation remains effective, while OSC is revoked before those
  transitions publish.

## Theme, localization, and accessibility

- Add all master terminal semantic colors/ANSI 0–15 to `AppTheme`, defaults, schemas, overrides, and every shipped theme; preserve session OSC overrides. High contrast remains readable and may force opacity.
- Add localized command/chooser/diagnostic/follow/history/privacy/paste/OSC-52/exit/restart/settings/status strings and all satellite translations, including broker Close/Paste actions, Cancel-default text, content-free deadline/background/rejection/cap statuses, OSC 52 Allow/Deny with displayed `Alt+A`/`Alt+D`, ask/disabled/revoked states, the exact Reset-incompatible destructive confirmation, `Reset`/`Cancel` buttons, deletion-incomplete/recovery-cap/family-cap states, Retry Reset, and manual-removal fallback. Resource placeholders are positional and token-identical; no clipboard payload/token appears in a resource argument, log, or archive.
- Complete keyboard-only/UIA coverage for the single broker surface, including plugin-owned modal Close/unsafe-paste actions with Cancel initially focused, confined Tab order, focused Enter/Space, Escape, conditional focus restoration, and non-focus-stealing OSC 52 exact-token mouse/UIA Invoke plus bar-scoped `Alt+A`/`Alt+D`. Hidden/canceled/expired surfaces disappear and return `UIA_E_ELEMENTNOTAVAILABLE`; one polite visible OSC notification is allowed. The host/UIA tree exposes only bounded user-visible metadata, never hidden clipboard bytes, identity generations, or tokens. The terminal document provider never exposes hidden/closed history. History/context exist only at broker `None`; an open history chooser exposes exactly visible command labels for selection/Copy, no offscreen records/nonce/hidden metadata, and nothing after dismissal.
- About/diagnostics reports exact selected engine pin/version/fingerprint and capability states only.

## Performance contract

Implement every master metric before optimization and deterministic cases:

- Reused open-command (`Ctrl+Alt+T`/`Alt+7`) tab-visible, directory Edit, contextual/full path insertion, local first prompt, pane-follow confirmation, input-visible, background close-glyph select/reveal/focus/broker-visible, one-slot Close/Paste/OSC arbitration, immutable one-read unsafe-paste prompt/final admission, paste-storage boundary/supersession/revocation and disable/refresh continuity, OSC 52 visible/background one/eight ask admission/immediate denial/revocation/final clipboard gate, output storm, resize/reflow, a 100,000-line deterministic feed bounded by the lower of line/64 MiB peak caps with retained count reported, exact 512 MiB process boundary/+1 with UIA/RPC-held charging and activation-only stable prune order, release-wakeup single-retry, bounded incomplete final-exit snapshot, 32 slots/denied 33rd/rapid retiring/quarantine recovery, retained-exit charging, hidden multitab, tab-icon overflow, deterministic Kitty recency/tie-break burst, every non-Kitty boundary/+1 disposition, rapid quiet close/module unload, and exact broker/state-lock/shutdown-flush deadlines. Emit `terminal.history.family_lock_wait_us` for every active-state publication and exercise ordinary saves plus active-active, recovery-growth, absent-Arm/old-growth, observer/Finalize/HCS, and family-privacy contenders in `terminal_perf_history_bounded_flush`; all broker/paste/OSC metrics and tags remain content-free and exclude clipboard length/payload/token/identity values.
- Hard invariants gate immediately: no UI blocking calls, no hidden unchanged paint loop, exact 32-slot/512 MiB/queue/Kitty/non-Kitty caps, no final-snapshot retry loop, no quarantine admission leak, no teardown timeout, and no sensitive tags.
- Establish same-machine test-enabled Release baselines with adequate samples, then write `Specs/Testing/TerminalPerfBudgets.json5`.
- Gate using `-PerfBudgetPath Specs\Testing\TerminalPerfBudgets.json5 -RequirePerfBudgets`; never claim Release from Debug/ad hoc/single sample.
- Archive terminal Commands result/trace/aggregate/PerfJsonl evidence only through the master's content-free `TerminalCommandsEvidence.psm1`/`Run-TerminalCommandsEvidence.ps1` path under `Specs/TestRuns/<MachineHash>/Commands/<RunId>/`, validate each exact returned path, and analyze only an explicit `Tools/Show-PerfRuns.ps1 -Run <path>`; finish and run `TerminalCommandsEvidence.Tests.ps1` plus the focused `TestRunArchive.Tests.ps1` cases. The generic selftest repo archive is not terminal evidence, and Gate-0 engine evidence remains under `TerminalEngine`.

## Packaging and supply chain

- Verify root `RuntimeDependencies.props` and every consumer use the plan-2 implementation of `SourceRoot=VcpkgTriplet|RepoRoot|TerminalEngineOutput` and package-relative `PackagePath` identically. Legacy omissions mean exactly `VcpkgTriplet` and `Plugins\<OutputName>`; new entries are explicit. Confirm `VcpkgTriplet` resolves through canonical `$(VcpkgInstalledDir)$(RSVcpkgTriplet)\` and canonical entrypoints pass those exact resolved values to non-MSBuild consumers without ambient-environment inference. Confirm every applicable path-bearing field on dependency/removal rows—`Source` where present, `OutputName`, and `PackagePath`—is literal/token/wildcard-free and valid, and OutputName equals the package leaf. Confirm the former `Flavor=Any` AWS rows are split into literal Debug/Release sources and the real zlib duplicate pairs are normalized into unioned-project rows before strict case-insensitive destination uniqueness. Every package contains exact `Plugins\Terminal.dll`; a dynamic selected-engine primary/closure exists only at its lock-declared `Plugins\TerminalRuntime\<leaf>` paths; state schema/notices are app-root. No duplicate engine basename or terminal implementation is staged app-root. Re-run focused `Tools\Tests\BuildReproducibility.Tests.ps1` plus x64 Debug/`ASan Debug`/Release and ARM64 Debug/Release source/destination tests, with ARM64 ASan only when supported. A source/static winner is linked into Terminal.dll and declares no runtime file. Required missing inputs fail centrally; obsolete outputs are removed declaratively.
- Source-contract and extracted-package tests reject every WSL-specific bootstrap, wrapper, Linux-profile payload, staged script, `WSLENV` setup, or loose integration resource/dependency. They also reject a loose PowerShell adapter manifest/schema/script, a package-side override, observer code outside `Terminal.dll`, a loose interaction-broker resource, broker/domain implementation outside the plugin, or host-owned prompt/token logic; the reviewed exact resources and broker are compiled into that DLL. Runtime launch recorders must still prove the reopened System32 `wsl.exe`, exact five-token argv, unchanged configured default user/shell, and zero pre/post-launch integration bytes or extra process. Packaging cannot turn a terminal-only WSL row into an integration-capable one.
- Normal MSBuild is offline. Release CI rebuilds/verifies locked x64/ARM64 `Terminal.dll` and selected-engine artifacts. Never add per-project PostBuildEvent/xcopy.
- ZIP portable packaging consumes the same manifest and passes clean-extraction smoke per `Specs/Installer/Installer_PortableZip.md`; MSI/MSIX match architecture and notices. Extend `Tools/ReleaseArtifactPolicy.ps1` with MSI naming and a shared exact selector. Unsigned artifacts live only under `.build\AppPackages\<ReleaseVersion>\<Platform>\<Kind>\`; signed MSIX lives under `.build\SignedAppPackages\<ReleaseVersion>\<Platform>\Msix\`. Canonical filenames are exactly `RedSalamander-<version>-<Platform>-Portable.zip`, `RedSalamander-<version>-<Platform>.msi`, and `RedSalamander-<version>-<Platform>.msix`. A build cleans only its exact leaf and writes one file. Selection requires `-ReleaseVersion`, rejects extra/duplicate/mismatched files in that leaf, validates content version/architecture, ignores older sibling version leaves, and release collection consumes only the current-version tree. Add focused Pester coverage.
- Before any official build, check in `Specs/Terminal/TerminalPackageToolchain.lock.json` and add `Tools/New-TerminalPackageProvenance.ps1 -RepositoryCommit <40-lowercase-hex> -ReleaseVersion <three-part> -Mode Build|VerifyBuildInputs|TransferredSmoke|Review -ToolchainLock <path> -PassThruManifest`. Derive the platform-specific official Release context first and only with pure `Get-RSVersionContext -BuildNumber <approved> -OfficialRelease`; do not read `.build/version/current-version.json`. Require its PackagingVersion, BuildNumber, Platform, Configuration, and OfficialRelease fields to equal the approved lane identity before provenance or package output. `Build` then requires `HEAD` to equal the approved commit and `git status --porcelain=v1 --untracked-files=all` to be empty before any build/test/archive output exists. It creates, under `.build/TerminalPackageProvenance/<RepositoryCommit>/`, one JCS object exactly `{formatId="red-salamander-terminal-package-source-inputs",schemaVersion=1,repositoryCommit,entries,submodules}`. `entries` is the sorted array of every non-submodule Git-tree entry as `{mode,repositoryRelativePath,blobSha256,sizeBytes}`; `submodules` is sorted `{repositoryRelativePath,commit}`. It hashes exact blob bytes, uses strict UTF-8, `/` repository-relative NFC paths and ordinal ordering, rejects case aliases/duplicate paths, and returns only the manifest path. `VerifyBuildInputs -ExpectedManifest <exact>` is read-only/no-output: immediately before the official Release build and separately before every ZIP/MSI/MSIX build it requires the same HEAD/version, no staged or unstaged tracked/submodule drift, a recomputed identical source-input identity, and the exact unchanged manifest/digest. After every build, `Read-RSVersionContext` must equal the already-approved pure context before tests or package consumption. `TransferredSmoke` and `Review` still require no staged or tracked modification and use mode-specific untracked allowlists: signed smoke admits only the exact signed/unsigned package leaves, the unchanged two-file unsigned-evidence run directory, and provenance scratch; review admits only the exact seven two-file evidence run directories and provenance scratch. Neither accepts `.build/TerminalEngine`; every other untracked path fails. The path-bearing manifest is never archived; only its SHA-256 is evidence. `Mode VerifyCommitIdentity -RepositoryCommit <commit> -ExpectedSourceInputManifestSha256 <digest>` is the scratch-independent closeout path: it forbids all path/release/tool/pass-through parameters, sets `GIT_NO_LAZY_FETCH=1` and `GIT_NO_REPLACE_OBJECTS=1`, and reconstructs the exact same JCS object in memory from NUL-safe `git ls-tree`/`git cat-file` reads of the named local commit/tree/blob objects without reading or mutating HEAD, index, worktree, submodule checkout, or `.build`. Gitlinks contribute their exact path/commit only. A missing commit/tree/blob (including a shallow/promisor object that would require lazy fetch), invalid/non-NFC path, alias, duplicate, digest mismatch, output byte, or any filesystem mutation fails offline. Success has exit zero and zero success output. A clean working tree makes the build-time bytes the build inputs, while this object-only reconstruction makes the recorded identity independently reproducible after all provenance scratch is deleted or from a second checkout root; engine/runtime locks bind declared non-Git inputs. Central packagers consume only declared inputs, never ignored/untracked wildcard discoveries. Tests cover a clean checkout with absent saved context, stale/mismatched saved context, wrong approved version with no output, tracked drift between package invocations, scratch deletion, second-root equality, missing/shallow/promisor objects, digest substitution, and strict no-output/no-mutation behavior.
- Check in `Specs/Terminal/TerminalPackageToolchain.lock.json`,
  `Specs/Terminal/TerminalPackageToolchainLock.schema.json`, and the exact
  valid/invalid `TerminalPackageToolchainLockCorpus/`. The lock is exactly the
  closed JCS object `{formatId="red-salamander-terminal-package-toolchain-lock",
  schemaVersion=1,tools,updateSources,signingPolicy}` with no additional
  properties. `signingPolicy` is exactly
  `{certificateRawSha256,publisher,codeSigningEkuOid,timestampUri,
  fileDigestAlgorithm,timestampDigestAlgorithm,requireCertificateCurrentlyValid,
  requireTrustedChain,requireRfc3161Timestamp,signToolVerifyPolicy,
  requireMakeAppxValidate}`: the two algorithms are `SHA256`, the EKU is exact
  code-signing OID `1.3.6.1.5.5.7.3.3`, `timestampUri` is credential-free
  HTTPS, all four `require*` values are `true`, and
  `signToolVerifyPolicy="pa-all"` means exact `/pa /all`. `publisher` is the
  exact AppxManifest Publisher and certificate subject string.
  `tools` contains exact
  `{toolId,platform,kind,mode,version,sha256}` tuples for `msbuild`, `cl`,
  `link`, `powershell-host`, `system-io-compression`, `windows-sdk`,
  `makeappx`, `signtool`, `wix`, `signature-verifier`,
  `appx-deployment-api`, and `terminaltests-activation-helper`; the selected
  `makeappx.exe`/`signtool.exe` records also bind raw-file SHA-256 and native PE
  machine for their exact purpose/platform. The closed signing policy freezes
  the leaf raw-certificate SHA-256, exact Publisher/subject, code-signing EKU,
  validity/chain verification policy, HTTPS RFC-3161 timestamp URI, SHA-256
  file/timestamp digests, and signed-container verification policy. It contains
  no certificate/PFX bytes, password, private-key locator, secret name, store
  handle, or other secret material.
- `updateSources` contains one ordinal-by-`toolId` exact record
  `{toolId,sourceKind,canonicalUri,trackedChannel,allowPrerelease,
  advisoryUriOrNull}` for every distinct tool ID and no orphan/duplicate; it is
  discovery metadata only and is never copied into a package-run `tools` tuple.
  For a single file, `sha256` hashes exact bytes; for `windows-sdk` or another
  resolved set, it hashes the RFC-8785 sorted `{name,version,sha256}` member
  array, including the canonical resolver result. ZIP requires the common
  compiler/build records plus PowerShell and System.IO.Compression; MSI adds
  WiX; unsigned MSIX adds Windows SDK/MakeAppx; signed production/smoke adds the
  exact SignTool/signature-verifier policy, AppX deployment/API OS-build
  identity, and activation helper. Inapplicable tuples are absent, never
  synthesized as another architecture. Every run records the exact applicable
  sorted tuple set and semantic verification requires it to match the lock for
  that platform/kind/mode; x64 and ARM64 may differ only where the lock declares
  platform-specific tuples.
- Add `Tools/Publish-TerminalSignedMsix.ps1` with mandatory exact
  `-UnsignedPackageRoot`, `-SignedPackageRoot`, three-part `-ReleaseVersion`,
  `-Platform x64|ARM64`, `-ToolchainLock`, `-CertificatePath`, caller-created
  `SecureString -CertificatePassword`, and `-PassThruPackage`. Before any
  mutation it validates the lock/schema, canonical contained roots and exact
  version/platform/name/architecture, exactly one unsigned input, absent signed
  output leaf, pinned MakeAppx/SignTool paths plus raw digests/PE machines, and
  an ephemeral read-only PFX load. The PFX leaf must own a private key and match
  the locked raw-certificate SHA-256, Publisher/subject, code-signing EKU,
  validity, and chain policy. A pre-existing matching certificate in
  `Cert:\CurrentUser\My` is rejected so cleanup can never remove user state.
- Signing is fail-closed and ordered: hash the unsigned MSIX; create one unique
  owned staging leaf under the resolved signed version/platform parent; copy
  the unsigned bytes and rehash both; temporarily import only the inspected
  certificate; invoke the pinned SignTool on only the staged copy with exact
  `/fd SHA256`, the store-selected certificate, the locked HTTPS RFC-3161 URI, and
  `/td SHA256`—never `/f`, `/p`, PATH/`Get-Command`, latest-SDK discovery,
  wildcard input, or a password in a child command line. Remove that imported
  certificate in `finally`; prove the unsigned source is still byte-identical;
  run pinned `signtool verify /pa /all` and pinned `makeappx validate`; and
  verify signed leaf certificate, Publisher, trusted RFC-3161 timestamp,
  architecture, and version. Only then atomically rename the one staged file/
  leaf into canonical
  `.build\SignedAppPackages\<ReleaseVersion>\<Platform>\Msix\RedSalamander-<ReleaseVersion>-<Platform>.msix`.
  Success writes only that
  signed copy and emits only its path. Every failure/interruption publishes no
  signed artifact, leaves the unsigned bytes immutable, and removes only the
  helper's identity-verified staging path and certificate import; cleanup never
  guesses at or deletes pre-existing state.
- Add `Tools/Tests/TerminalMsixSigning.Tests.ps1` with TestSandbox fake selector,
  SignTool, MakeAppx, clock/timestamp, certificate-store, and atomic-publish
  seams. Cover x64/ARM64 source/destination identity, unsigned-before/after and
  signed hashes, exact operation order/sole output, PFX/password/certificate
  non-disclosure, wrong certificate/Publisher/EKU/validity/chain/timestamp/tool
  path/digest/machine, pre-existing certificate, extra leaf, escape/reparse,
  copy/sign/verify/publish failure, interrupt/cleanup ownership, and clean retry.
  Source-contract tests reject signing under `.build\AppPackages`, in-place or
  recursive-wildcard signing, PATH/`Get-Command` SignTool, latest-SDK scanning,
  `/f`/`/p`, unsigned-as-release substitution, or missing platform-distinct
  uploads.
- Update `.github/workflows/release.yml` and `Tools/ReleaseArtifactPolicy.ps1`.
  Each x64 and ARM64 lane first normalizes/validates the one unsigned MSIX under
  `.build\AppPackages`, then calls `Publish-TerminalSignedMsix.ps1` exactly once
  for that platform. Upload unsigned and signed roots as distinctly named
  workflow artifacts; release attachment selects and verifies only the signed
  current-version MSIX and rejects the same-name unsigned file as a substitute.
  Signing-disabled/unavailable lanes may retain an unsigned inspection artifact
  but never label or attach it as installable. Remove the existing in-place
  AppPackages signing, PATH/latest-SDK lookup, and recursive MSIX wildcard.
- Extend plan 1's single `Tools/Get-TerminalDependencyStatus.ps1` view with mandatory `-PackageToolchainLock .\Specs\Terminal\TerminalPackageToolchain.lock.json` support after this lock exists. Offline mode validates the lock/schema/corpus identity, every applicable tuple including pinned MakeAppx/SignTool bytes and PE machine, the complete non-secret certificate/Publisher/EKU/timestamp verification policy, absence of secret-bearing lock properties, each terminal runtime-manifest row, notice/license binding, and every engine/package cross-reference without writing; online mode reports tool/SDK/WiX/signing/security update observations from canonical credential-free lock metadata without applying them. Add these cases to `TerminalDependencyLifecycle.Tests.ps1`. A toolchain update is one reviewed lock/provenance/policy change followed by the complete native/package/signed-smoke matrix; it never rides implicitly with an engine update. The authoritative runbook must show both the fast offline check and optional online check.
- The same lifecycle runbook has a separate read-only adapter command, `Tools/Test-TerminalPowerShellAdapters.ps1 -AuditInstalled`. It validates the source schema/corpus and embedded-manifest/resource identities, inspects installed Windows PowerShell/PSReadLine/pwsh tuples without editing profiles/modules/manifests or repository state, and reports exactly `supported|terminal-only-unknown-version|custom-hook|missing`. A newly observed version never becomes compatible by range or nearest match: inspect the upstream effective function plus assembly/API diff, generate comparison only in disposable `.build`, deliberately review one exact tuple/resource/license change, rebuild, then pass real Windows PowerShell 5.1 and every shipped pwsh tuple plus the full native/package matrix before `autoEligible`. Adapter, engine, and package-toolchain upgrades are three independently reviewed changes and never ride implicitly together.
- Implement the master's exact `Tools/Test-TerminalPackage.ps1` interface, including mandatory `-RepositoryCommit`, `-BuildNumber`, `-ReleaseVersion`, `-ProvenanceManifest`, `-ToolchainLock`, `-PassThruManifest`, and mode-required `-UnsignedPackageRoot` plus `-UnsignedEvidenceManifest` for `SignedInstallSmoke`. `-EngineArtifactsRoot` is mandatory only for `ExtractAndSmoke` and forbidden for `SignedInstallSmoke`. Fail before filesystem mutation unless repository/provenance/tool identities pass; signed mode must be exactly MSIX with both unsigned inputs and no engine root, extract mode must have the engine root and forbid both unsigned inputs, and ARM64 MSI is invalid. Add corpus tests for every missing or unexpected mode-specific argument. `ExtractAndSmoke` selects one artifact through `ReleaseArtifactPolicy`, uses TestSandbox scratch, reuses portable smoke for ZIP, `msiexec /a` for x64 MSI, and pinned-SDK `makeappx validate/unpack` for MSIX, then verifies manifest/PE architecture, schema/notices, exact `Plugins\Terminal.dll`, and lock-derived private runtime closure. The exact mutation matrix and recipes are frozen below. `TerminalTests --package-root-smoke` calls the production plugin manager/factory, requires `RedSalamanderCreate(IID_ITerminal,...,L"builtin/terminal")`, rejects `IID_IViewer`, and proves `available` or the exact unavailable category with zero callbacks/handles; the shipped app diagnostic accepts those same expected categories and proves the app composition root still starts while only Terminal is unavailable. ZIP/MSI run it for every mutation. Source/static proves Terminal.dll's compiled engine identity and absence of separate engine/closure DLLs; plugin-DLL mutations always remain applicable.
- Unsigned MSIX is inspection-only. `SignedInstallSmoke` selects and re-inspects the exact same-version/platform unsigned candidate through mandatory `-UnsignedPackageRoot`, verifies its artifact/evidence digest, directly validates/unpacks the exact signed artifact, and validates the lock-derived dynamic closure or source/static compiled identity in that package without a separate engine-artifact root. It repeats manifest/PE/schema/notices/no-unapproved-artifact and unpacked production-loader checks. Canonical ordinal path/size/raw-SHA-256 manifests cover every extracted file. Cross-package equality excludes exactly `[Content_Types].xml`, `AppxBlockMap.xml`, `AppxMetadata/CodeIntegrity.cat`, and `AppxSignature.p7x`; every other path, size, and hash must be identical. Those four records remain in evidence with signed/unsigned presence, size, digest, and validated role. No fifth exclusion or wildcard is allowed, and corpus tests mutate an ordinary payload file to prove rejection. A release-claimed MSIX additionally requires the install lane on clean disposable native x64 and ARM64 machines: trusted signature/Publisher match; refuse any pre-existing package family; `Add-AppxPackage`; derive AUMID; activate via the `TerminalTests` `IApplicationActivationManager` helper with `--terminal-package-smoke=available`; wait/exit 0; remove exactly that installed package in `finally`. Missing signing/native/payload/safe-clean-VM evidence blocks MSIX. Never install unsigned or alter an existing user package.
- Package evidence under `Specs/TestRuns/<MachineHash>/TerminalPackage/<RunId>/` contains only digests, result categories, and call ordering—no package/extraction/certificate/path/content data.
- Normal packages contain no source, tool cache, unapproved PDB, test history, or stale engine artifact.

### TerminalPackage evidence manifest v1

Before package testing, check in `Specs/Terminal/TerminalPackageEvidence.schema.json`, exact valid/invalid fixtures under `Specs/Terminal/TerminalPackageEvidenceCorpus/`, the toolchain lock above, and reuse `Tools/TerminalEvidence.psm1`; a second serializer/digest/archive writer is forbidden. Each `TerminalPackage/<RunId>` directory contains exactly `terminal-package-evidence.v1.json` plus `terminal-package-evidence.v1.sha256`; transferring evidence means transferring that entire directory in its unchanged `Specs/TestRuns/<MachineHash>/TerminalPackage/<RunId>/` layout. The schema is draft 2020-12 with `additionalProperties: false` at every object, explicit bounds/enums, and top-level `oneOf` for `package-run` versus `evidence-set`. It uses the TerminalEngine v1 RFC-8785 JCS/UTF-8/no-trailing-newline, exact-byte lowercase SHA-256 companion, unsigned-decimal-string, ASCII-ID uniqueness/ordinal-sort, exact shared `Get-RSTestMachineHash` 12-lowercase-hex algorithm, and correct `IsWow64Process2` native-attestation rules. Its `runId` and directory leaf use the same collision-safe form as TerminalNative: exact UTC `yyyy-MM-dd_HHmmss.fffffffZ_<32-lowercase-hex-UUID-without-hyphens>`, create-new/fail-if-exists, regenerating a UUID collision before any evidence write; no second-resolution collision, overwrite, or suffix probing is allowed.

Every record requires exactly `formatId="red-salamander-terminal-package-evidence"`, `schemaVersion=1`, `recordKind`, `runId`, RFC-3339 UTC `createdUtc`, 40-lowercase-hex `repositoryCommit`, `machineHash`, `schemaSha256`, and `harness{version,sourceSha256}`. A `package-run` additionally requires `releaseVersion`, unsigned-decimal-string `buildNumber`, `platform=x64|ARM64`, `kind=Zip|Msi|Msix`, `mode=ExtractAndSmoke|SignedInstallSmoke`, `nativeAttestation`, `frozen`, `tools`, `artifact`, `evidenceObjects`, `steps`, `cases`, mode-specific `signedComparison`, and `overallResult=pass|fail`. `releaseVersion` is exactly `<major>.<minor>.<buildNumber>` with all three components canonical unsigned decimals and the third equal to `buildNumber`; `artifact.contentVersion` must encode that same release version under the kind's existing content-version rule. `Msi+ARM64`, signed non-MSIX, extract mode with signed inputs, and signed mode without both unsigned inputs are schema-invalid.

`frozen` is exactly `{engineLockSha256,runtimeDependenciesSha256,releaseArtifactPolicySha256,terminalStateSchemaSha256,noticesManifestSha256,terminalEvidenceHelperSha256,sourceInputManifestSha256,packageToolchainLockSha256,packageProvenanceHelperSha256}`. Every digest is 64 lowercase hex. `repositoryCommit` must equal the explicit approved commit and the commit embedded in the build/smoke provenance manifest; at evidence creation the source-input digest is recomputed from that manifest. At later cross-closeout verification, where scratch is forbidden and absent, the same digest is recomputed from the named commit's local Git objects through `VerifyCommitIdentity`; neither route permits a path or digest claim without byte-level reconstruction. `tools` is the applicable tuple array from the reviewed toolchain lock, sorted ordinally by `{toolId,platform,kind,mode}`, with each record exactly `{toolId,platform,kind,mode,version,sha256}`. `artifact` always means the artifact selected from this run's `PackageRoot`—the unsigned artifact for `ExtractAndSmoke`, the signed artifact for `SignedInstallSmoke`—and is exactly `{sha256,sizeBytes,contentVersion,architecture,payloadManifestSha256}`. `architecture` must equal the requested native platform.

`evidenceObjects` is a sorted unique array of `{evidenceId,sha256,sizeBytes,resultCategory}`; `resultCategory` is the union of the exact step and case result categories below. It contains content-free digests only, while raw tool output remains in TestSandbox and is never archived. A step is exactly `{sequence,stepId,status,resultCategory,evidenceRef}`. Sequence starts at 1 with no gap. `status` is `pass|fail|not-applicable|not-run`; pass uses only `passed`, not-applicable uses only `not-applicable-msix-inspection|not-applicable-empty-transitive-closure|not-applicable-source-static`, not-run uses only `blocked-by-prior-failure`, and fail uses one of `timeout|timeout-containment-failed|loader-matrix-failed|app-matrix-failed|payload-mismatch|signature-invalid|publisher-mismatch|preexisting-package-family|install-failed|activation-failed|diagnostic-exit-failed|remove-failed|post-remove-family-present|cleanup-failed`. Pass/fail has a resolving evidence reference whose category matches; not-applicable/not-run has `evidenceRef=null`. After the first failure every later side-effecting step is present as `not-run`, never omitted.

The harness reaches an identity-ready point only after provenance/native/tool checks, exact artifact selection, container validation/extraction, and payload inspection have succeeded and every required top-level identity/digest exists. A failure before that point—including provenance, native/tool identity, artifact selection/validation, extraction, unsigned-evidence validation, or payload inspection—emits no `package-run`; it uses the invalid/incomplete nonzero exit and leaves only non-archived scratch diagnostics. An emitted `package-run` therefore has every pre-identity step `pass`. `overallResult=fail` is reserved for a complete, schema-valid post-identity behavioral result such as a loader/app case, payload comparison, signature, install, activation, removal, cleanup, or containment failure.

For `ExtractAndSmoke`, `steps` is exactly this order: `provenance`, `select-artifact`, `validate-container`, `extract-container`, `inspect-payload`, `loader-matrix`, `app-matrix`. `app-matrix` is `not-applicable-msix-inspection` only for unsigned MSIX and otherwise must pass. For `SignedInstallSmoke`, it is exactly: `provenance`, `validate-unsigned-evidence`, `select-unsigned-artifact`, `select-signed-artifact`, `validate-signed-container`, `extract-signed-container`, `inspect-signed-payload`, `compare-payloads`, `verify-signature-publisher`, `refuse-preexisting-family`, `install-package`, `derive-aumid`, `activate-diagnostic`, `await-diagnostic-exit`, `remove-package`, `verify-family-absent`.

An extract case is exactly `{sequence,caseId,targetId,status,expectedCategory,actualCategory,resultCategory,appCompositionStarted,terminalCallbacksCreated,terminalHandlesCreated,evidenceRef}`. `caseId` is `available|missing-plugin|corrupt-plugin|wrong-architecture-plugin|missing-engine-primary|corrupt-engine-primary|wrong-architecture-engine-primary|missing-transitive|corrupt-transitive|source-static-no-engine-dll-identity`; expected categories are `available|unavailable:plugin-missing|unavailable:plugin-corrupt|unavailable:plugin-wrong-architecture|unavailable:engine-primary-missing|unavailable:engine-primary-corrupt|unavailable:engine-primary-wrong-architecture|unavailable:transitive-missing|unavailable:transitive-corrupt|not-applicable-empty-transitive-closure|not-applicable-source-static`, and `actualCategory` additionally permits `unexpected-category|no-result`. `status` is `pass|fail|not-applicable|not-run`. Case result categories are `passed|category-mismatch|app-composition-start-failed|terminal-resource-leak|diagnostic-timeout|case-containment-failed|not-applicable-empty-transitive-closure|not-applicable-source-static|blocked-by-prior-failure`. Counts are unsigned-decimal strings for executed cases and null only for not-applicable/not-run cases. `appCompositionStarted` is Boolean for ZIP/MSI and null for inspection-only unsigned MSIX/not-applicable/not-run. A passing unavailable ZIP/MSI case requires app start true plus both counts zero; a passing unavailable MSIX case requires both counts zero. Pass requires expected and actual categories to match byte-for-byte, `resultCategory=passed`, and a resolving evidence reference. Fail uses the precise failure result and a resolving reference; its actual category is the observed closed category or `unexpected-category|no-result`. Not-applicable requires matching expected/actual/result not-applicable values and no reference; not-run requires `actualCategory=no-result`, `resultCategory=blocked-by-prior-failure`, and no reference. The derived full case array is always present. Independent cases continue after an ordinary assertion failure; only failed containment after identity-ready makes every unsafe remaining case `not-run`. `SignedInstallSmoke` has `cases=[]` exactly.

Every matrix begins `available`, then `missing-plugin`, `corrupt-plugin`, and `wrong-architecture-plugin` with `targetId=terminal-plugin`; the target is exact `Plugins\Terminal.dll`, and the wrong-architecture replacement is the reviewed opposite-platform Release plugin after its PE machine is proved. For a dynamic winner append `missing-engine-primary`, `corrupt-engine-primary`, and `wrong-architecture-engine-primary` with `targetId=engine-primary`; the target is the exact locked `Plugins\TerminalRuntime\` primary. Then, for every locked non-system private transitive dependency in ascending ordinal `dependencyId`, append `missing-transitive` then `corrupt-transitive` with that non-path ID as `targetId`; testing one representative is forbidden. An empty dynamic closure appends exactly two not-applicable records with `targetId=empty-closure`. A source/static winner appends `source-static-no-engine-dll-identity`, then the three engine-primary mutation IDs and both transitive IDs once each as `not-applicable-source-static` with `targetId=source-static`; it still executes all three plugin mutations and proves `Terminal.dll`'s compiled identity plus absence of every separate engine/closure DLL. Every executed mutation starts from a new verified extracted-root clone and changes only its target: missing omits the target; corrupt replaces it with exactly 64 bytes of `0xA5`; wrong architecture replaces the plugin/engine primary with the verified locked opposite-platform Release artifact from the same repository/pin and first proves its PE machine is the opposite architecture. The harness rejects a mutation whose before/after tree differs anywhere else.

A canonical payload manifest is an RFC-8785 array of `{path,sha256,sizeBytes}` over every regular unpacked file. A path is relative, NFC, `/`-separated, has no empty/`.`/`..` component, drive/UNC prefix, trailing slash, alternate stream, or reparse point; aliases equal under ordinal-ignore-case are rejected. Preserve normalized case and sort by ordinal normalized path. Hash exact raw file bytes and use unsigned-decimal size strings. ZIP/MSI include every file. MSIX excludes exactly, by ordinal path before enumeration comparison, `[Content_Types].xml`, `AppxBlockMap.xml`, `AppxMetadata/CodeIntegrity.cat`, and `AppxSignature.p7x`; no wildcard or case-folded exclusion exists. The path-bearing manifest remains in scratch, never the archive; `payloadManifestSha256` is its exact JCS-byte digest. Signed/unsigned equality means the two canonical arrays are byte-identical after those exclusions.

`signedComparison` is forbidden for extract runs and required for signed runs. It is exactly `{unsignedEvidenceManifestSha256,unsignedArtifactSha256,signedArtifactSha256,unsignedPayloadManifestSha256,signedPayloadManifestSha256,payloadsEqual,exclusions,signatureResult,publisherResult}`. `signatureResult` is `trusted-valid|invalid|not-run` and `publisherResult` is `match|mismatch|not-run`; a passing signed run requires `trusted-valid` and `match`, while an earlier payload failure records both as `not-run`. The exclusion array is exactly the four paths above in that order; `payloadsEqual` must be true for pass; the base `artifact` digest must equal `signedArtifactSha256`. The unsigned record must match repository/provenance/frozen/release/build/platform/kind/harness identities, its artifact digest must equal `unsignedArtifactSha256`, and its payload digest must equal `unsignedPayloadManifestSha256`. The signed harness validates the supplied unsigned JSON and companion, preserved parent MachineHash/RunId layout, and exact unsigned artifact before any install.

An `evidence-set` has the common fields plus exactly `releaseVersion`, `buildNumber`, `sourceInputManifestSha256`, `frozen`, `inputHarness`, `inputs`, `bindings`, and `overallResult`. `inputs` has exactly seven records `{sequence,platform,kind,mode,machineHash,runId,manifestSha256}` in this order: x64 Zip extract, x64 Msi extract, x64 Msix extract, ARM64 Zip extract, ARM64 Msix extract, x64 Msix signed, ARM64 Msix signed. `bindings` has exactly x64 then ARM64 records `{platform,unsignedManifestSha256,signedManifestSha256,unsignedArtifactSha256,signedArtifactSha256,payloadManifestSha256}`. The set's `harness` identifies the verifier; `inputHarness` is the byte-identical package-run harness identity. The complete shared `frozen` object and source-input digest must be identical across all inputs; tool arrays must each match the platform/kind/mode tuples in the one shared toolchain lock rather than being falsely byte-identical across architectures.

Add `Tools/Verify-TerminalPackageEvidence.ps1 -Mode Create|VerifyExpectedOnly -EvidenceManifest <path[]> -EvidenceRoot <path> -RepositoryCommit <commit> -BuildNumber <n> -ReleaseVersion <version> -ToolchainLock <path> [-ProvenanceManifest <scratch-json>] [-ExpectedSourceInputManifestSha256 <digest>] [-PriorEvidenceSetManifest <canonical-json>] [-ExpectedEvidenceSetManifest <canonical-json>] [-PassThruManifest]`. Every evidence-manifest argument is already in the master's canonical repository-relative `Specs/TestRuns/.../TerminalPackage/...` form. The verifier rejects absolute, wrong-case, backslash, dot-segment, outside, reparse, alias, and ordinal-ignore-case duplicate spellings, resolves only through `Resolve-RSCanonicalTestRunPath`, and requires each path to be the JSON member of an exact two-file run directory under its declared MachineHash/RunId; it performs no latest discovery or caller normalization. It validates all schemas/companions/references, recomputes frozen/tool identities, version/build/architecture relations, step/case order and status/category mapping, exact dynamic/source-static case completeness, payload equality, native attestation, and signed bindings. Plan 6 uses `Create -ProvenanceManifest <live-scratch> -PassThruManifest` without prior/expected-source arguments and receives one resolved local path for exactly one new pair, then immediately converts it to canonical form before recording or transfer. Phase-9 Stage A uses `Create` with canonical `-PriorEvidenceSetManifest` and exact `-ExpectedSourceInputManifestSha256`, forbids `-ProvenanceManifest`, invokes `VerifyCommitIdentity` before any evidence mutation, and requires the prior pair to equal the recomputed matrix before one fresh pair is emitted. Phase-9 Stage B uses `VerifyExpectedOnly` with canonical prior and expected sets plus that same expected digest; it likewise forbids provenance scratch, invokes offline reconstruction, forbids `-PassThruManifest`, has zero success output and no filesystem mutation, and proves both sets and the currently recomputed inputs/frozen files are substantively equal. Invalid/incomplete evidence emits no evidence-set and uses one nonzero exit; Create with a complete matrix containing a real failed run emits `overallResult=fail` and a distinct nonzero exit; only a semantically recomputed complete pass succeeds. Tests cover every illegal mode-specific parameter combination, substitution, reorder, missing/changed prior or expected set/companion, every noncanonical or duplicate path, byte-identical records transferred under two different absolute repository roots with identical cross-machine identity digests, deletion of all provenance scratch, missing/shallow/promisor objects with lazy fetch disabled, source-input digest disagreement, repeated VerifyExpectedOnly no-output/no-mutation, and stale/different/unrecorded Phase-9 identity.

Hard timeouts are not skips. Launch `msiexec`, `makeappx`, TerminalTests, and app diagnostics suspended through the existing `Start-RSContainedProcess`/`Close-RSContainedProcess` kill-on-close JobObject seam; never use an uncontained `Start-Process`. The deadline is 120 seconds for each extraction/validation and each add/remove operation, and 30 seconds for each package-root or installed activation diagnostic. On process timeout, close the containment job, allow at most five additional seconds to observe root/job quiet, and classify `timeout`; failure to prove quiet is `timeout-containment-failed`. A post-identity timeout is recorded in the run; a pre-identity timeout follows the no-manifest invalid/incomplete rule above. Delete scratch only after quiet; deletion failure records `cleanup-failed`, leaves the path only inside TestSandbox for its reaper, and archives no path/content. Because Windows Installer may continue service-owned administrative extraction after the client exits, every MSI run has a unique scratch root; timeout quarantines that root from deletion/reuse for the rest of the run and blocks closeout even if it later settles. Run each `Add-AppxPackage`/remove cmdlet in its own contained no-profile PowerShell child, but treat AppX deployment as service-owned: an add/remove timeout or uncertain registration state immediately taints the disposable VM, forbids further evidence from it, attempts only bounded exact-family cleanup when state is observable, and requires VM discard/recreation even if cleanup later appears successful. Missing signing trust, clean disposable VM, matching native architecture, or containment fails preflight and blocks the signed deliverable; none is an environmental success/skip.

## Final verification

Preferences-navigation acceptance is mandatory in Preferences/UIA, ABI,
source-contract, schema, package, and final Commands lanes. It proves the exact
seven section IDs/orders/collapse states and exact 48-element mapping above;
element order `10,20,...` within each section; Advanced membership only from
schema metadata; `configurationVersion` hidden from every visible/action/status/
link set; and no duplicate, missing, extra, renamed, case-changed, or multiply
placed element. The generated localized-link corpus resolves every exact plugin/
section/element tuple and rejects unknown/empty/approximate targets. Loaded-DLL
and golden-oracle tests prove `HostPluginPreferencesRequest` `80/8` on x64/
ARM64 plus exact validation, HRESULT, QI/agility, posted ownership, coalescing,
stale-schema, disabled-element, destroyed-owner, shutdown, and drain behavior;
any shortened request-type spelling is a source-contract failure.

Live-settings acceptance is mandatory in Core, Preferences/UIA, lifecycle,
source-contract, and final Commands lanes. It proves the one service-owned
serialized coordinator as the sole Apply/Open-registration/registry-removal
linearization point; stable ascending captured membership; later Open/removal
queueing; one-session-at-a-time Prepare/Abort/Commit; retiring-at-Prepare no-op;
post-Prepare root-exit/close/retirement queueing; complete rollback with the old
snapshot and no published generation; one atomic commit-token publication; and
allocation-free/nonblocking/noexcept/infallible Commit/release before Apply
success. Apply races against Open, root exit, explicit close/retire,
disappearing/captured-retiring sessions, failing Prepare, actions on both sides
of commit, rollback, and success must prove no lock held across wait/callback/UI/
COM/process/pipe/coordinator work, no two session gates held together, no
ordinary post-publication failure, no deadlock, and no partial/mixed generation.
Injected token/identity/staged mismatches must close admission and reach the test
fatal-policy recorder; source contract requires production
`RaiseFailFastException` and forbids continued execution.

Bell/hyperlink acceptance is mandatory in Core, Renderer, Preferences/UIA,
lifecycle, source-contract, package, and final Commands lanes. Bell proves the
per-view first/249/250-ms no-queue/dropped-counter gate, exact 2-DIP 149/150-ms
generation-bound visual flash with animation/high-contrast rules, `none`/view-
generation/teardown cancellation, and one non-UI
`TrySubmitThreadpoolCallback` path pinned by
`AcquireModuleReferenceFromAddress(...)` and transferred at callback entry by
`FreeLibraryWhenCallbackReturns(instance, pin.release())`. It proves one
`MessageBeep(MB_OK)`, OS mute/failure no-fallback, pin/submit failure consumed
with released ownership/one content-free failure counter/no beep/retry/fallback,
submitted-work teardown quiet, and quiet decrement only after `MessageBeep`
returns. Hyperlink activation proves the 8-KiB cap, C0/space/scheme precheck,
one exact reserved-zero/flagged `CreateUri`, canonical component/credential
rules, only case-insensitive `http|https|mailto`, all other schemes/relative/
malformed/control/over-limit inputs with zero side effect, and both Ctrl/policy/
document/view-generation gates. Every winner has at most one exact plugin-owned
`ShellExecuteW` call; `>32` alone is success, `<=32` posts one localized status
plus content-free diagnostic with no retry/fallback/second side effect, and the
exact null-parameter/null-cwd/no-process-handle boundary is proven.

Refresh/availability acceptance is mandatory in the Core, PaneProfiles, and final Commands lanes: the implementation files named above must be present in the project/filter maps; `terminal_plugin_refresh_` and `terminal_plugin_availability_` cases must be nonzero, passed, and unskipped; disable must retain a live generation; accepted refresh must preserve live terminal usability while closing new admission; the configuration and post-export quiet clocks must produce their distinct nonfatal blocked reasons; host Retry and latest desired availability must obey the frozen precedence; the old owning module handle must reset before the single replacement; and process shutdown must supersede every state. Content-free evidence records only state/reason/generation counters and never configuration, path, terminal, or history content.

Tab-limit acceptance is likewise mandatory: settings/schema/Preferences accept only `1..32`; the host contains no setting read/count precheck; Open reserves atomically per complete physical-host key and returns the exact quota HRESULT; the hidden uncommitted record and every synchronous/host-commit failure roll back without tab/zoom/selection/child/process; diagnostic and Exited objects retain their slot; callback-drain/Close/public-release releases it exactly once while later internal quarantine does not retain it; live lowering never evicts; and boundary/+1, reentrant, independent-host, rollback, retention, release, and reopen cases are nonzero, passed, and unskipped in native/Commands evidence.

Unsafe-paste acceptance is mandatory in Core, PaneProfiles, Preferences/UIA, lifecycle, and final Commands lanes: exactly one UI-thread clipboard read; strict frozen normalization/classification/preview; complete immutable identity plus paste-policy/input-mode generations and opaque one-shot token; one pending/session and fixed 8 MiB/session plus 32 MiB/service storage; one broker slot; newer-request supersession; strict 60-second equality expiry; selection/hide/deactivation/input-owner/policy/lifecycle revoke-and-scrub with no resurrection; no more-permissive upgrade; final current-mode/liveness/`pasteMaxBytes`/normal-byte/3840-descriptor/queue-capacity atomic admission; zero partial/stale/new-incarnation bytes; and full source/preview/encoded-buffer scrub. Generic disable and normal refresh alone must leave safe and a still-selected/visible/eligible pending unsafe paste usable for an existing frozen-config view with unchanged paste-policy generation; tests must separately prove that broker cancellation still wins.

OSC 52 acceptance is mandatory in Core, PaneProfiles, Preferences/UIA, lifecycle, security, and final Commands lanes: immutable service/session/incarnation/view/policy/effect identity; opaque tokens and plugin-owned scrubbed bytes; exactly one retained visible ask/session and eight/service; background/hidden/inactive immediate deny-and-scrub before counting and without selection/focus; strict 30-second equality expiry; ninth/newer-request outcomes; selection/activation/ownership/policy/disable/re-enable/refresh/close/restart/retirement/shutdown revocation before publication/teardown; effective deny while disabled; stale/duplicate/wrong-generation no-op; non-focus-stealing bar with exact-token UIA Invoke and scoped `Alt+A`/`Alt+D`; final exact policy/enabled/non-refreshing/interactive-visible live-Running revalidation; one clipboard replacement linearized under the same gate; token consumption and scrub on success/failure; no replay/retry/partial write; and ignored clipboard reads.

Interaction-broker acceptance is mandatory in Core, PaneProfiles, Preferences/UIA, lifecycle, source-contract, package, and final Commands lanes. Require one content-free broker per committed view; exact `None|CloseConfirm|UnsafePasteConfirm|Osc52Ask` state, full identity, and all 12 cells of the Close-over-Paste-over-OSC table; no restoration of a canceled lower-priority request; stable-ID background close select/reveal/layout/focus and revalidation after every reentrant callback before UserTab; no invisible prompt or rollback after successful selection; exact interactive-visible/presentation-generation checks; History/context/plugin-text-input exclusion; modal Cancel default, confined Tab, focused Enter/Space, Escape, conditional focus restore; OSC no-focus UIA behavior; and lifecycle/post-registry drain before child teardown. Deterministic tests cover every pairwise/three-way order, eight background asks plus `+1` retaining/counting zero, deadline just-before/equal/after, selection away/back, deactivate/reactivate, reorder/removal/HWND reuse, replacement failure, missing/failed posts, and stale/duplicate/wrong-kind/wrong-view replies.

Identity/launch acceptance is mandatory in ABI, Core, PaneProfiles, ShellState/schema, Preferences, Restart, and final Commands lanes. One shared runtime/serializer corpus accepts the complete 32,772-unit pwsh profile ID, 33,037-unit canonical location key, 33,040-unit display path, and every exact component ceiling; preserves each byte-for-byte through chooser/`ExactProfile`, Open, preference/history save/load/merge, missing profile, and Restart; and rejects every runtime +1/malformed form before mutation with no truncation/hash/alias. Windows launch recorders require held/revalidated nonnull `lpApplicationName`, fixed unquoted `cmd.exe|powershell.exe|pwsh.exe` argv0, and hostile current/PATH/app-local immunity. WSL identity remains independent of exact five-token command-line representability, including the 32,736-unit quote-free path boundary for distro `D` and zero-preflight/process +1 rejection.

State-string acceptance is mandatory in ShellState/schema, source-contract, package, and final Commands lanes. Draft 2020-12 `maxLength` is tested only as a code-point backstop; the exact nonassertive `x-redsalamander-maxUtf16CodeUnits` and `x-redsalamander-maxUtf8Bytes` annotations are present and frozen; runtime validation is authoritative; and serializers pass both layers. Every applicable ceiling proves `bmp-boundary-both-pass` and `bmp-plus-one-both-reject`; each UTF-16 annotation also proves `astral-utf16-plus-one-schema-pass-runtime-reject`, and each UTF-8 annotation proves `multibyte-utf8-plus-one-schema-pass-runtime-reject`, with no schema-invalid/runtime-valid case. Any blanket parity assertion is a test defect.

Windows-startup acceptance is mandatory in Core, PaneProfiles, ShellState,
Preferences/security/lifecycle, source-contract, package, and final Commands
lanes. Cmd proves the correlated one-use bootstrap/cleanup for every local,
validated-plugin-backing, and UNC launch: embedded script identity/reopen hash;
exact six-token `/E:ON /V:OFF /S /K call` argv and two-pass path encoder; exact
protected ready-before-create inbound pipe and root PID; all eight fixed child-
environment names/value grammars, mixed-case inherited-key removal, canonical
  insertion, and permanent cleanup/no restoration; exact direct-expanded `cd /d`
  or `pushd`; prelaunch-free drive snapshot; exact logical root/tail parsing;
  `DRIVE_REMOTE`; bounded `REMOTE_NAME_INFO_LEVEL` reverse mapping with initial
  4,096-byte buffer, one exact-size retry, 262,144-byte cap, pointer/NUL
  validation, and complete/root/tail ordinal-ignore-case comparison; and the
  fixed-order `RSTCMD1` ASCII Ack with maximum-
valid shape, 512/513-byte assembly boundary, `ERROR_MORE_DATA`, and clean EOF.
Cases cover every field/cross-field, status, chunk boundary,
empty/partial/non-ASCII/noncanonical/extra/second/over-cap line,
first-valid-plus-trailing/second data, first-valid-but-kept-open, strict-before/
equal/after deadline, target revalidation, AutoRun cwd-change correction, and
AutoRun exit/hang/suppression/missing/late/wrong-result failure with zero
Running/input/memory. `CmdLaunchRootedRunning` forms only after exact PID/
capability/nonce/operation/generation/mode/target-digest/status/drive/EOF proof,
  applicable UNC share-root/empty-tail/deep-tail proof and safe root-death
  cleanup, and the same-incarnation Running CAS;
cmd has no HMAC/cwd serialization or ongoing epoch. One explicit limitation
probe records trusted AutoRun and same-root correct-capability/correct-nonce
preemption plus malicious same-user code as out of scope, while wrong PID/
capability/nonce/operation/generation/digest, replay, duplicate line, and PTY
  output fail closed. UNC cases include case/slash/dot lexical normalization;
  zero/failing `GetLogicalDrives`; 4,096 bytes, exact 262,144-byte cap and +1;
  zero/non-growing/inconsistent retry sizes; malformed/null/misaligned/out-of-
  buffer/nonterminated pointers; local/SUBST drives; DFS/
  provider/share aliases; empty/zero-or-one-leading-slash remainder, drive/UNC/
  double-leading/root-escape rejection; and preoccupied/reused drive letters.
  Cleanup cases
  repeat the 4,096/262,144/+1/malformed `WNetGetConnectionW` corpus and prove no
  cancel while the child lives; already-gone success; exact unchanged-mapping
  one-shot cancel and fresh absence recheck; and open-files/API/root-exit/
  remap/remove/changed-target-reuse races that never force or cancel another
  mapping. A limitation probe records indistinguishable same-user same-target
  letter reuse rather than claiming ownership proof.
Every local/UNC real Windows PowerShell 5.1 and shipped pwsh tuple plus the fake
must prove the exact four-token fail-closed `-NoExit -Command` wrapper, closed
literal encoder/exact AST with fixed `*>&1`, six script parameters, sole exact
typed finalize object with first PSTypeName
`RedSalamander.TerminalBootstrapFinalize.V1`, ordered properties exactly
`Sentinel,Finalize`, Sentinel exactly `RSTBOOT-WRAPPER-READY-1`, and a
scriptblock Finalize, plus exact property/type/scriptblock mismatch rejection,
profile variable collision/read-only/AllScope failure, wrapper-validation-before-finalizer, one output-silent
synchronous complete-frame CommitAck Write with no Flush/retry/later pipe
operation, duplicate/throw/partial-Write finalizer failure, and injected named
disposal error proving idempotent output-silent disposal with no Ack repair,
server-full-read-before-gates, and 125/126 exit behavior for invalid
objects/finalizers plus every extra success/warning/verbose/debug/information
result; normal-profile-before-script ordering, no policy bypass or
policy-failure prompt, protected client-PID-bound bootstrap pipe, exact HMAC
layout and nine-kind sequence, module-qualified post-profile literal rooting
immune to profile-shadowed Set/Get-Location output, one latest-policy
decision, and no target/semantic material in environment or initial argv.
Toggle races cover Decision construction/delivery, semantic-pipe creation,
before/during/after client connect, adapter install, handshake enqueue, either
join slot, Prepared Write, Running CAS, CommitPrompt, and CommitAck.
Require both orderings and every failure of the two-slot `StartupSemanticJoin`,
pre-child reservation/nonblocking handler lock discipline, the exact nonfatal
  `RevokedTerminalOnly` retained no-operation semantic/control cleanup-pipes/
  both-Prepared path, the valid staged-Handshake plus Unavailable 50-ms boundary
  disposition, ReleaseAck as
noncommit, the lifecycle-locked Running CAS against every retirement race,
authenticated CommitPrompt/CommitAck with the exact Active/TerminalOnly
generation matrix, including the exact unchanged-generation
`SemanticEnable -> Prepared(SemanticUnavailable) ->
CommitPrompt(TerminalOnly) -> CommitAck` usable terminal-only path,
minimal retained Decision-generation/Prepared-outcome state, both valid Enable
outcomes after a latched disable, and no input/profile memory/capability before
CommitAck. Repeat every disable timing for Ready and Unavailable: before
CommitPrompt construction, after construction/before send, after send/before
child receipt, between finalizer Ack Write/server receipt, between wrapper
validation/finalizer invocation, during Ack acceptance, and immediately after
acceptance.
CommitAck acceptance ignores live policy generation for the latched-disable
cleanup case; before/after Prompt/Ack disable races must still open ordinary
  input/integration-independent memory, publish zero semantic capability,
  cancel/close both pipes, make completed control EOF/error or semantic-write
  failure latch interposers inert, and never retire solely for disable. Every receipt is strict-before the common
five-second deadline. Pre-CAS failure has no Running/memory; post-CAS failure
unrelated to disable retires the already-Running incarnation with startup gates closed; a
never-Running launch cannot evaluate prompt. Mutable cleanup is proved without
claiming physical erasure of immutable OS/.NET strings. A fake-only, skipped, or
ReleaseAck/EOF-based prompt result cannot close the plan.

Shell-semantic acceptance is mandatory in ShellState, Preferences/UIA, security, lifecycle, source-contract, dependency-audit, package, and final Commands lanes. Cmd proves launch plus complete one-use local/plugin-backed/UNC launch-bootstrap erasure and zero ongoing nonce/history/follow/insertion/activity/authenticated-cwd capability. Real Windows PowerShell 5.1 and every shipped `autoEligible` pwsh tuple independently prove exactly one semantic epoch per incarnation, nonce/control transport only in the authenticated post-rooting `SemanticEnable` decision, copy into the root module closure, mutable cleanup before prompt commit, no initial-argv/environment/nested-child inheritance, `StartupSemanticJoin` conjunction before readiness, permanent invalidation for malformed/replay/gap/overflow/queue failure, rejection of every later marker/second handshake, and restart/new-session-only rekey. A fake-only result or environmental skip cannot close the plan. Nested integration is unavailable; top-level capabilities merely pause while nested input owns the terminal and resume only at the next valid in-sequence root event in the same epoch.

Semantic-write acceptance is mandatory in ShellState, lifecycle, security,
source-contract, dependency-audit, package, and final Commands lanes. Every real
supported shell tuple proves the exact protected root-PID-bound inbound
`TerminalSemantic` pipe, precharged fixed 2-MiB assembly reservation against its
64-KiB pipe buffer, Connect-before-first-Read order, one persistent epoch
connection with no reconnect, one overlapped read/rearm-before-dispatch path, and
exact `RSTSEM01` frame/HMAC/sequence/bounds. It proves exactly one four-argument
asynchronous `NamedPipeClientStream.WriteAsync` call per frame, no synchronous
large-write stall, one preserved message through the 2-MiB cap, fresh per-write
Task/token/buffer ownership, strict-before 50-ms synchronous observation with
whole-millisecond floor and the exact PowerShell-5.1 `IAsyncResult` wait, no
overlapping owned write, and cancel/stream-close/inert fallback with no retry or
replacement for a failed emission/continuation/helper/background runspace. The
server proves `ERROR_MORE_DATA` continues only the current message, the final
successful Read delimits it, EOF never delimits, partial EOF truncates,
unexpected clean EOF revokes, and lifecycle-fenced EOF is normal. Cover 64-KiB,
2-MiB, and +1 boundaries; ordinary frames 2..N, second/extra Handshake failure;
first-chunk-plus-50-ms receive assembly just-before/equal/after; no-first-chunk
Handshake under the strict five-second startup deadline; sync throw, fault/
cancel/write-timeout/equality; disable/close/Restart/root-exit/teardown; late Task
settlement/exception observation/scrub; and failure-before-Prepared terminal-only
downgrade. A valid matching Handshake staged before the receive deadline but
paired with Unavailable after the child's write observation reaches/eclipses its
50-ms deadline must be scrubbed/revoked terminal-only; malformed/mismatched/extra
data and Handshake plus TerminalOnly remain fatal. A tuple that cannot prove all
of this is terminal-only, never skipped or allowed another transport.

Semantic-payload acceptance is mandatory in the same lanes. Exact canonical-
JSON/generated-C++/embedded-PowerShell byte identity and loaded resource identity
must pass, including the two checked-in Generated paths and generator `-Check`.
The parser accepts only the four frozen kinds and their exact offsets,
little-endian unaligned loads, enum/flag/reserved/cross-field rules, tuple/version/
manifest digest identity, sequence/read-cycle/accepted-sequence relations,
canonical cwd, exact .NET-string UTF-8, control result/operation ID, and result-
specific zero/nonzero fields; it rejects struct casts/padding/TLV/trailing data,
unknown values/bits, overflow, malformed UTF-8/control text, provider cwd,
sequence/operation mismatch, generic Activity/Cwd kinds, command `1 MiB + 1`, cwd
`128 KiB + 1`, AcceptedInput payload maximum mismatch from exactly 1,179,672, and
frame `2 MiB + 1`. Handshake cases prove exact `payloadVersion`, 3..43-byte
canonical two-to-four-component System.Version round trips, channel-generation
equality, capability dependencies, and no Follow/Insert publication before real
Probe Success. RootReadReady cases prove a still-pending AcceptedInput may remain
the completion candidate across intervening ControlResult frames. Every text
slice rejects BOM/NUL/unpaired-surrogate; tuple IDs/paths reject every control and
only command admits HT/LF/CR. Every kind's min/max plus astral/multibyte corpus is nonzero,
passed, and unskipped; all 18 ControlResult kind/status combinations prove the
exact success/failure data-zeroing table while retaining echoed identity. Source
contracts require C++ reuse of
`Common/StringConversion.h` strict APIs and exact embedded-PowerShell
`System.Text.UTF8Encoding(false, true)`, with no Terminal-local/fallback
 converter. Only AcceptedInput enters RunningRootCommand and only its
matching RootReadReady returns idle/cwd.

Semantic/control revocation acceptance is mandatory in the same lanes. Every
pre-CAS disable retains both exact-PID no-operation cleanup servers only through
Prepared and Running CAS, where both close/cancel before TerminalOnly
CommitPrompt; every post-CAS disable and lifecycle/retirement path fences then
closes/cancels both immediately. Tests prove `TerminalControl`, not write-only
`TerminalSemantic`, is the child read sentinel; semantic closure breaks pending/
next WriteAsync; ordinary wrapper `PollControlTransport(NoWait)` is nonwaiting,
stages/rearms only into the sole empty slot and never dispatches; the key
handler's `PollControlTransport(HandlerWaitOnce)` first consumes an occupied
stage with no wait/`EndRead` on the new read and otherwise alone performs the sole
bounded wait; `EndRead` is called once only on an
already-complete or handler-wait-signaled BeginRead, and the next read is armed
before action; and all Connect/Read/Write/parser/dispatch/
posted-result buffers and module pins drain before quiet/unload, with five-second
quarantine/fatal outcomes rather than unsafe unload.

Foreground-control acceptance additionally proves the fixed 262,144-byte read,
production DACL/client interoperability for exact current-user
`ReadData|ReadAttributes|WriteAttributes|Synchronize` with `WriteData` absent,
an exact granted-mask assertion, and the exact
asynchronous/anonymous/noninherited client constructor and successful
`ReadMode=Message` switch on real Windows PowerShell 5.1 and every shipped pwsh,
exact `RSTCTL01` frame/HMAC/sequence/deadline, fixed buffer plus immutable payload
reservation before `operationT0`, fixed-protocol-reserve descriptor/one-byte
inert wake placeholder with its final input-queue sequence before pipe Write,
serialization/HMAC after that sample, and
exactly one OVERLAPPED WriteFile with a manual-reset event. Pending completion
uses exactly one off-UI `GetOverlappedResultEx` with floor-of-remaining timeout
and `bAlertable=FALSE`; accept only exact byte count and strict-before post-return
QPC. `FlushFileBuffers` and every synchronous pipe
Write/Flush are source-contract failures. Writer admission is successful submit
under the generation/input-operation gate and is distinct from earlier queue-
placeholder reservation; cover user-before-reservation cancellation, fixed-
reserve descriptor/byte boundary and +1, later-user FIFO barriers before/after
reserve, submit, completion, AwaitingResult identity publication, and activation;
prove an immediate child result before the completion callback returns sees the
awaiting record; cover pre-submit zero-byte tombstone,
post-submit inert held fence, fence before/at/after submit, completion, epoch-live
revalidation, same-descriptor wake activation with no second capacity/admission/
sequence or priority bypass, exactly one wake with no retry/timer/second wake,
and final pre-effect sample. Ordinary wrapper
`PollControlTransport(NoWait)` calls are nonwaiting, may stage/rearm only into the
sole empty slot, and never dispatch. Cover a frame completing during an ordinary
wrapper entry so `NoWait` stages it and rearms before the injected wake; the
handler must take that staged frame exactly once with zero wait/`EndRead` on the
new read. Only the owner/binding/epoch/
lifecycle-revalidated key handler's `PollControlTransport(HandlerWaitOnce)` may use
one nonalertable `((IAsyncResult)read).AsyncWaitHandle.WaitOne(250)`; cover
BeginRead completing during the wait and just-before/equal/after embedded
deadline, an unsignaled timeout's zero child action/emission/revocation, a no-
operation physical press blocking only its root shell for at most 250 ms, close/
disable EOF/error waking the wait, exact one `EndRead`, rearm before dispatch,
physical press consuming a complete frame, and mapping replacement receiving at
most one raw wake before missing-result retirement. Cover malformed/replay/
expiry/equality and close/cancel using
exact-OVERLAPPED `CancelIoEx`, handle retention, one quiet-bounded off-UI event
drain wait, one nonwaiting `GetOverlappedResult`, then close/release. A read still
unsignaled after the sole handler wait performs child-side zero action/emission/
revocation; an existing plugin operation times out/revokes on the plugin side.
A matching authenticated safe result—Success, or for Follow/Insert only pre-effect
RejectedState or BufferChanged—or mechanically proved pre-submit/no-action
releases queued input; strict
expiry alone never proves an earlier Follow/Insert did not execute, so every
ambiguous lost-result case retains queued input and exposes only Discard or
confirmed Close, including after disable. Post-fence results/semantic frames
grant no state/history/capability, and no I/O/frame/module gate releases before
actual completion. Matching authenticated Success and, for Follow/Insert only,
pre-effect RejectedState or BufferChanged return Idle and advance the sequence;
mechanically proved pre-submit/no-action returns Idle without advancing. Probe
non-Success, Expired, BindingChanged, ShellOperationFailed, failed/revoked/
ambiguous outcomes are terminal-only with no later control operation. Exercise
immediate and pending completion, deadline-minus-one,
equality, sub-millisecond floor-to-zero, short write, cancel/completion race,
success versus `ERROR_OPERATION_ABORTED`, and five-second drain miss with no
free/unload. Successful write proves that only per-write storage/pin releases,
while the persistent pipe plus operation/activated-placeholder input hold survive to matching
ControlResult; cancel/error/disable/retirement drains before closing that pipe.
The drain never admits wake/effect/result. Submitted failure/ambiguity retains
the inert fence until Discard/confirmed Close; safe completion proves the wake
was written at its original earlier queue sequence before any later user byte.

The same mandatory real-shell matrix proves the exact embedded tuple and allowlisted unshadowed PSReadLine `PSConsoleHostReadLine` interposer. Its first executable statement is `$lastRunStatus = $?`; it uses only the tuple's exact audited `Set-StrictMode -Off` recipe: `StrictModeOffThenReadLine2` calls `ReadLine($host.Runspace,$ExecutionContext)`, while `CaptureStatusStrictModeOffThenReadLine3` calls `ReadLine($host.Runspace,$ExecutionContext,$lastRunStatus)`. It calls that static method once, passes the captured Boolean unchanged where supported, and returns the exact same non-null `System.String` object once. `RootReadReady` occurs at wrapper entry after the untouched user prompt and before the block; exact `AcceptedInput` occurs only after a real ReadLine return and before host return; only its matching next-root `RootReadReady` completes/persists the candidate. Parse failure completes at that next root read, while native/nested execution stays busy and null/EOF/exit/process loss/replacement/restart first discards. The separate app control wake never returns the line and cannot count as acceptance. Prompt, `$LASTEXITCODE`, `$Error`, provider, key bindings, PSReadLine options/history, and native history remain unchanged.

Unknown/custom/missing/Vi/shadowed/replaced adapters and any observer/event/UTF-8/queue/sequence/epoch failure stay usable but conservatively busy with no history/follow/insertion until Restart; no prompt/key/cell/control-wake inference is allowed. Live disable/revocation leaves an installed owned wrapper as process-lifetime inert pass-through using the same recipe and emitting/dispatching nothing, never fabricates a `FunctionInfo` restoration, and removes only a still-owned control mapping. The installed-adapter audit is read-only and reports the exact closed classifications; a dependency upgrade follows the reviewed exact-tuple procedure before it can ship.

Retirement acceptance is mandatory in Core, PaneProfiles, lifecycle, and final Commands lanes. Exactly one plugin per-view retirement-intent arbiter linearizes UserTab allow/cancel, chooser/plugin actions, every natural-exit/quiet/`closeOnExit` result, WindowClosing, and ApplicationShutdown; its Close prompt can exist only through the one broker after the stable-ID interactive-visible handoff. Broker/selection/visibility/activation loss performs exact Cancel, and a required nonvisible confirmation returns Retry without intent/token allocation. Takeover invalidates/dismisses a pending prompt generation and stale/duplicate replies are inert. The generic host's one atomic `{instanceId,hostViewGeneration}` removal claim is shared by CloseApproved, removal callbacks, and both forced paths. Deterministic barriers prove pending-confirmation/natural-exit/forced-close races, callback losing to forced teardown, cancel-first/approval-first outcomes, host-generation reuse, and exactly one deferred hide/remove/null-drain/Close/release/MRU mutation.

WSL terminal-only acceptance is mandatory in Core, PaneProfiles, ShellState, Preferences, source-contract, and extracted-package lanes: exact System32 executable plus `[wsl.exe, --distribution, <catalog exact distro name>, --cd, <absolute profile-native Linux path>]`; configured default user/shell unchanged; no `--exec`/`--user`/`-e`/wrapper/extra process/fallback; no argv/environment/`WSLENV`/PTY/Linux-profile/staged-script injection; no semantic nonce/epoch; permanent `RequestedUnverified` with `Verified|Diverged` transitions rejected; no app history/follow/activity/authenticated-cwd/profile-memory read/write; exact-distribution full-native-path-only explicit insertion; truthful capability UI; exact launch usability; and permitted post-Running launcher failure/root exit without fallback or false memory. The tests must not require every asynchronous `--cd` failure to precede `Running`.

The required configuration matrix is x64 Debug, x64 Release, x64 `ASan Debug`, ARM64 Debug, and ARM64 Release. Build/run ARM64 `ASan Debug` only when the selected MSVC/toolchain proves support; otherwise archive the exact unsupported-toolchain evidence and never create, map, stage, or label a Debug artifact as ASan. Runtime-dependency matrix tests cover the five required configurations plus ARM64 ASan only when supported. The following is a two-lane matrix, not one shell transcript: use the same approved repository commit, release version, and build number in both lanes, run the x64 block on native x64, and run the ARM64 block plus ARM64 package smoke on native ARM64. Transferred artifacts retain exact hashes; a commit/version mismatch blocks evidence.

Direct native test proof uses only `Tools/Run-TerminalNativeTests.ps1` and its exact manifest/companion; raw `TerminalTests.exe`, `DxUiTests.exe`, or `SettingsSchemaTests.exe` console success is not closeout evidence. The required v1 matrix contains exactly 14 runs:

| Slice | x64 Debug | x64 Release | x64 ASan Debug | ARM64 Debug | ARM64 Release |
|---|---:|---:|---:|---:|---:|
| Core | required | required | required | required | required |
| Renderer | required | required | not applicable | not applicable | required |
| PaneProfiles | required | required | not applicable | not applicable | required |
| ShellState | required | required | not applicable | not applicable | required |

The later slices' x64 Debug runs reprove the child handoffs at the final repository commit; their x64/ARM64 Release runs are the full native closeout lanes. Core owns the sanitizer boundary. If ARM64 ASan support is positively proved, add exactly one `Core/ARM64/ASan Debug` run and pair; the other three ARM64 ASan slice tuples remain v1 not-applicable. If unsupported, retain the exact toolchain proof instead. Every applicable runner result must be semantically verified immediately, and its entire unchanged `Specs/TestRuns/<MachineHash>/TerminalNative/<RunId>/` two-file directory (`terminal-native-evidence.v1.json` plus `.sha256`) must be retained with both file digests. Missing/skipped/raw-only evidence blocks closeout.

Every slice runs against the centrally staged matching `Plugins\Terminal.dll`, records a content-free plugin build/PE/IID identity, creates through `IID_ITerminal`, and rejects `IID_IViewer`; direct static test linkage cannot satisfy the lane. Full solution builds and package tests assert no terminal implementation import/library is linked into `RedSalamander.exe`, no product Terminal source exists outside `Plugins/Terminal` except the explicit ABI/generic host allowlist, callback/module quiet gates unload, and a source/static engine is contained in `Terminal.dll`. Add and run exact `Tools/Tests/TerminalPluginBoundarySourceContracts.Tests.ps1` in the canonical Full and CI suites rather than relying on directory convention; its reviewed allowlist is the master file-map contract.

Native x64 lane:

```powershell
if ($env:RS_APPROVED_REPOSITORY_COMMIT -notmatch '^[0-9a-f]{40}$') { throw 'Set RS_APPROVED_REPOSITORY_COMMIT to the reviewed terminal implementation commit.' }
$approvedRepositoryCommit = [string]$env:RS_APPROVED_REPOSITORY_COMMIT
$actualRepositoryCommit = [string](git rev-parse HEAD)
if ($LASTEXITCODE -ne 0 -or $actualRepositoryCommit.Trim() -cne $approvedRepositoryCommit) { throw 'HEAD does not equal RS_APPROVED_REPOSITORY_COMMIT.' }
$worktreeStatus = @(git status --porcelain=v1 --untracked-files=all)
if ($LASTEXITCODE -ne 0 -or $worktreeStatus.Count -ne 0) { throw "Official package evidence requires a clean build worktree: $($worktreeStatus -join '; ')" }
if ($env:RS_APPROVED_BUILD_NUMBER -notmatch '^[1-9]\d*$') { throw 'Set RS_APPROVED_BUILD_NUMBER to the shared positive integer build number.' }
[int]$approvedBuildNumber = $env:RS_APPROVED_BUILD_NUMBER
$approvedReleaseVersion = [string]$env:RS_APPROVED_RELEASE_VERSION
if ($approvedReleaseVersion -notmatch '^\d+\.\d+\.\d+$') { throw 'Set RS_APPROVED_RELEASE_VERSION to the matching three-part release version.' }
. .\Tools\Versioning.ps1
$approvedVersionContext = Get-RSVersionContext -RepoRoot (Get-Location).Path -Configuration Release -Platform x64 -BuildNumber $approvedBuildNumber -OfficialRelease
$releaseVersion = [string]$approvedVersionContext.PackagingVersion
if ($releaseVersion -notmatch '^\d+\.\d+\.\d+$' -or $releaseVersion -cne $approvedReleaseVersion -or [int]$approvedVersionContext.BuildNumber -ne $approvedBuildNumber -or [string]$approvedVersionContext.Platform -cne 'x64' -or [string]$approvedVersionContext.Configuration -cne 'Release' -or -not [bool]$approvedVersionContext.OfficialRelease) { throw 'Pure x64 package version context does not match the approved identity; no provenance/package output is allowed.' }
$provenanceManifest = .\Tools\New-TerminalPackageProvenance.ps1 -RepositoryCommit $approvedRepositoryCommit -ReleaseVersion $releaseVersion -Mode Build -ToolchainLock .\Specs\Terminal\TerminalPackageToolchain.lock.json -PassThruManifest
if (-not (Test-Path -LiteralPath $provenanceManifest -PathType Leaf)) { throw 'Provenance preflight did not return its exact manifest path.' }
$provenanceSha256 = (Get-FileHash -Algorithm SHA256 -LiteralPath $provenanceManifest).Hash.ToLowerInvariant()
function Assert-RSTerminalOfficialBuildInputs {
    param([Parameter(Mandatory)][string]$Purpose)
    $verifiedManifest = .\Tools\New-TerminalPackageProvenance.ps1 -RepositoryCommit $approvedRepositoryCommit -ReleaseVersion $releaseVersion -Mode VerifyBuildInputs -ExpectedManifest $provenanceManifest -ToolchainLock .\Specs\Terminal\TerminalPackageToolchain.lock.json -PassThruManifest
    if (-not $? -or [IO.Path]::GetFullPath($verifiedManifest) -cne [IO.Path]::GetFullPath($provenanceManifest) -or (Get-FileHash -Algorithm SHA256 -LiteralPath $provenanceManifest).Hash.ToLowerInvariant() -cne $provenanceSha256) { throw "Official build-input recheck failed before $Purpose." }
}
function Assert-RSSavedTerminalVersionContext {
    param([Parameter(Mandatory)][string]$Purpose)
    $saved = Read-RSVersionContext -RepoRoot (Get-Location).Path
    if ($null -eq $saved -or [string]$saved.PackagingVersion -cne $releaseVersion -or [int]$saved.BuildNumber -ne $approvedBuildNumber -or [string]$saved.Platform -cne 'x64' -or [string]$saved.Configuration -cne 'Release' -or -not [bool]$saved.OfficialRelease) { throw "Saved x64 version context differs from the pure approved context after $Purpose." }
}
Import-Module .\Tools\TerminalEvidence.psm1 -Force
. .\Tools\TestRunPlan.ps1
foreach ($testFile in @(
    '.\Tools\Tests\BuildReproducibility.Tests.ps1',
    '.\Tools\Tests\ReleaseArtifactPolicy.Tests.ps1',
    '.\Tools\Tests\TestRunArchive.Tests.ps1',
    '.\Tools\Tests\TerminalCommandsEvidence.Tests.ps1',
    '.\Tools\Tests\TerminalDependencyLifecycle.Tests.ps1',
    '.\Tools\Tests\TerminalMsixSigning.Tests.ps1',
    '.\Tools\Tests\TerminalPluginBoundarySourceContracts.Tests.ps1'
)) {
    $pesterParameters = New-RSPesterInvokeParameters -Path (Resolve-Path $testFile).Path
    $pesterResult = Invoke-Pester @pesterParameters
    $failedProperty = $pesterResult.PSObject.Properties['FailedCount']
    $failed = if ($failedProperty) { $failedProperty.Value } else { $pesterResult.PSObject.Properties['Failed'].Value }
    if ([int]$failed -ne 0) { exit 1 }
}
$dependencyStatusBefore = @(git status --porcelain=v1 --untracked-files=all)
if ($LASTEXITCODE -ne 0) { throw 'Cannot capture repository status before the terminal dependency check.' }
.\Tools\Get-TerminalDependencyStatus.ps1 -Mode Locked -EngineLock .\External\terminal-engine.lock.json -PackageToolchainLock .\Specs\Terminal\TerminalPackageToolchain.lock.json
if (-not $?) { throw 'Locked terminal dependency/toolchain status failed.' }
$dependencyStatusAfter = @(git status --porcelain=v1 --untracked-files=all)
if ($LASTEXITCODE -ne 0 -or ($dependencyStatusAfter -join "`n") -cne ($dependencyStatusBefore -join "`n")) { throw 'Locked terminal dependency/toolchain status changed repository state.' }
.\build.ps1 -Platform x64 -Configuration Debug
$adapterAuditStatusBefore = @(git status --porcelain=v1 --untracked-files=all)
if ($LASTEXITCODE -ne 0) { throw 'Cannot capture repository status before the PowerShell adapter audit.' }
.\Tools\Test-TerminalPowerShellAdapters.ps1 -AuditInstalled
if (-not $?) { throw 'Installed PowerShell adapter audit failed.' }
$adapterAuditStatusAfter = @(git status --porcelain=v1 --untracked-files=all)
if ($LASTEXITCODE -ne 0 -or ($adapterAuditStatusAfter -join "`n") -cne ($adapterAuditStatusBefore -join "`n")) { throw 'PowerShell adapter audit changed repository state.' }
.\Tools\Test-TerminalPowerShellControl.ps1 -PluginPath .\.build\x64\Debug\Plugins\Terminal.dll -RequireWindowsPowerShell51 -RequirePwsh
if (-not $?) { throw 'Real Windows PowerShell/pwsh observer matrix failed.' }
$expectedEngineLockSha256 = (Get-FileHash -LiteralPath '.\External\terminal-engine.lock.json' -Algorithm SHA256).Hash.ToLowerInvariant()
$x64NativeManifests = [System.Collections.Generic.List[string]]::new()
foreach ($slice in @('Core', 'Renderer', 'PaneProfiles', 'ShellState')) {
    $nativeManifest = .\Tools\Run-TerminalNativeTests.ps1 -Slice $slice -Platform x64 -Configuration Debug -EvidenceRoot .\Specs\TestRuns -PassThruManifest
    .\Tools\Verify-TerminalNativeEvidence.ps1 -Manifest $nativeManifest -ExpectedSlice $slice -ExpectedRepositoryCommit $approvedRepositoryCommit -ExpectedEngineLockSha256 $expectedEngineLockSha256
    $x64NativeManifests.Add($nativeManifest)
}
.\build.ps1 -ProjectName TerminalTests -Platform x64 -Configuration "ASan Debug"
$nativeManifest = .\Tools\Run-TerminalNativeTests.ps1 -Slice Core -Platform x64 -Configuration "ASan Debug" -EvidenceRoot .\Specs\TestRuns -PassThruManifest
.\Tools\Verify-TerminalNativeEvidence.ps1 -Manifest $nativeManifest -ExpectedSlice Core -ExpectedRepositoryCommit $approvedRepositoryCommit -ExpectedEngineLockSha256 $expectedEngineLockSha256
$x64NativeManifests.Add($nativeManifest)
Assert-RSTerminalOfficialBuildInputs -Purpose 'x64 Release build'
try {
    $env:RSBuildEnableTests = 'true'
    .\build.ps1 -Platform x64 -Configuration Release -BuildNumber $approvedBuildNumber -OfficialRelease
} finally {
    Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue
}
Assert-RSSavedTerminalVersionContext -Purpose 'x64 Release build'
foreach ($slice in @('Core', 'Renderer', 'PaneProfiles', 'ShellState')) {
    $nativeManifest = .\Tools\Run-TerminalNativeTests.ps1 -Slice $slice -Platform x64 -Configuration Release -EvidenceRoot .\Specs\TestRuns -PassThruManifest
    .\Tools\Verify-TerminalNativeEvidence.ps1 -Manifest $nativeManifest -ExpectedSlice $slice -ExpectedRepositoryCommit $approvedRepositoryCommit -ExpectedEngineLockSha256 $expectedEngineLockSha256
    $x64NativeManifests.Add($nativeManifest)
}
if ($x64NativeManifests.Count -ne 9) { throw 'The required x64 TerminalNative matrix did not produce exactly nine manifests.' }
if (@($x64NativeManifests | Select-Object -Unique).Count -ne 9) { throw 'The required x64 TerminalNative manifests are not nine unique exact paths.' }
$x64NativeCloseoutIdentities = @($x64NativeManifests | ForEach-Object {
    $manifestPath = (Resolve-Path -LiteralPath $_).Path
    $canonicalManifestPath = ConvertTo-RSCanonicalTestRunPath -RepoRoot (Get-Location).Path -ResolvedPath $manifestPath -Area TerminalNative -ExpectedLeaf 'terminal-native-evidence.v1.json'
    $recordDirectory = Split-Path -Parent $manifestPath
    $recordFiles = @(Get-ChildItem -LiteralPath $recordDirectory -File | Sort-Object Name | Select-Object -ExpandProperty Name)
    if (($recordFiles -join '|') -cne 'terminal-native-evidence.v1.json|terminal-native-evidence.v1.sha256') { throw "TerminalNative directory is not an exact two-file record: $recordDirectory" }
    $companionPath = Join-Path $recordDirectory 'terminal-native-evidence.v1.sha256'
    [pscustomobject][ordered]@{
        manifestPath = $canonicalManifestPath
        manifestSha256 = (Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash.ToLowerInvariant()
        companionSha256 = (Get-FileHash -LiteralPath $companionPath -Algorithm SHA256).Hash.ToLowerInvariant()
    }
})
if ($x64NativeCloseoutIdentities.Count -ne 9 -or @($x64NativeCloseoutIdentities.manifestPath | Select-Object -Unique).Count -ne 9) { throw 'The x64 TerminalNative canonical closeout identity set is incomplete or duplicated.' }
$x64PerfCommandsRun = .\Tools\Run-TerminalCommandsEvidence.ps1 -Scenario FinalPerf -Platform x64 -Configuration Release -RepositoryCommit $approvedRepositoryCommit -EvidenceRoot .\Specs\TestRuns -PassThruRunPath
$x64FullCommandsRun = .\Tools\Run-TerminalCommandsEvidence.ps1 -Scenario FinalFull -Platform x64 -Configuration Release -RepositoryCommit $approvedRepositoryCommit -EvidenceRoot .\Specs\TestRuns -PassThruRunPath
if (@(@($x64PerfCommandsRun, $x64FullCommandsRun) | Select-Object -Unique).Count -ne 2) { throw 'The x64 final Commands evidence paths are missing or duplicated.' }
.\Tools\Test-TestRunArchive.ps1 -RunPath @($x64PerfCommandsRun, $x64FullCommandsRun)
if ($LASTEXITCODE -ne 0) { throw 'The exact x64 Commands evidence failed validation.' }
Assert-RSTerminalOfficialBuildInputs -Purpose 'x64 ZIP package build'
.\build.ps1 -Configuration Release -Platform x64 -BuildNumber $approvedBuildNumber -OfficialRelease -Zip
Assert-RSSavedTerminalVersionContext -Purpose 'x64 ZIP package build'
Assert-RSTerminalOfficialBuildInputs -Purpose 'x64 MSI package build'
.\build.ps1 -Configuration Release -Platform x64 -BuildNumber $approvedBuildNumber -OfficialRelease -Msi
Assert-RSSavedTerminalVersionContext -Purpose 'x64 MSI package build'
Assert-RSTerminalOfficialBuildInputs -Purpose 'x64 MSIX package build'
.\build.ps1 -Configuration Release -Platform x64 -BuildNumber $approvedBuildNumber -OfficialRelease -Msix
Assert-RSSavedTerminalVersionContext -Purpose 'x64 MSIX package build'
$x64ZipManifest = .\Tools\Test-TerminalPackage.ps1 -PackageRoot .\.build\AppPackages -RepositoryCommit $approvedRepositoryCommit -BuildNumber $approvedBuildNumber -ReleaseVersion $releaseVersion -ProvenanceManifest $provenanceManifest -ToolchainLock .\Specs\Terminal\TerminalPackageToolchain.lock.json -Kind Zip -Platform x64 -Mode ExtractAndSmoke -EngineLock .\External\terminal-engine.lock.json -EngineArtifactsRoot .\.build\TerminalEngine -EvidenceRoot .\Specs\TestRuns -PassThruManifest
$x64MsiManifest = .\Tools\Test-TerminalPackage.ps1 -PackageRoot .\.build\AppPackages -RepositoryCommit $approvedRepositoryCommit -BuildNumber $approvedBuildNumber -ReleaseVersion $releaseVersion -ProvenanceManifest $provenanceManifest -ToolchainLock .\Specs\Terminal\TerminalPackageToolchain.lock.json -Kind Msi -Platform x64 -Mode ExtractAndSmoke -EngineLock .\External\terminal-engine.lock.json -EngineArtifactsRoot .\.build\TerminalEngine -EvidenceRoot .\Specs\TestRuns -PassThruManifest
$x64UnsignedMsixManifest = .\Tools\Test-TerminalPackage.ps1 -PackageRoot .\.build\AppPackages -RepositoryCommit $approvedRepositoryCommit -BuildNumber $approvedBuildNumber -ReleaseVersion $releaseVersion -ProvenanceManifest $provenanceManifest -ToolchainLock .\Specs\Terminal\TerminalPackageToolchain.lock.json -Kind Msix -Platform x64 -Mode ExtractAndSmoke -EngineLock .\External\terminal-engine.lock.json -EngineArtifactsRoot .\.build\TerminalEngine -EvidenceRoot .\Specs\TestRuns -PassThruManifest
pwsh .\Tools\Show-PerfRuns.ps1 -Area Commands -Run $x64PerfCommandsRun -Metric terminal.render.paint_us -FailOnQuality -ShowBuildFlavor
```

Native ARM64 lane:

```powershell
if ($env:RS_APPROVED_REPOSITORY_COMMIT -notmatch '^[0-9a-f]{40}$') { throw 'Set RS_APPROVED_REPOSITORY_COMMIT to the same reviewed commit used by the x64 lane.' }
$approvedRepositoryCommit = [string]$env:RS_APPROVED_REPOSITORY_COMMIT
$actualRepositoryCommit = [string](git rev-parse HEAD)
if ($LASTEXITCODE -ne 0 -or $actualRepositoryCommit.Trim() -cne $approvedRepositoryCommit) { throw 'HEAD does not equal RS_APPROVED_REPOSITORY_COMMIT.' }
$worktreeStatus = @(git status --porcelain=v1 --untracked-files=all)
if ($LASTEXITCODE -ne 0 -or $worktreeStatus.Count -ne 0) { throw "Official package evidence requires a clean build worktree: $($worktreeStatus -join '; ')" }
if ($env:RS_APPROVED_BUILD_NUMBER -notmatch '^[1-9]\d*$') { throw 'Set RS_APPROVED_BUILD_NUMBER to the same value used by the x64 lane.' }
[int]$approvedBuildNumber = $env:RS_APPROVED_BUILD_NUMBER
$approvedReleaseVersion = [string]$env:RS_APPROVED_RELEASE_VERSION
if ($approvedReleaseVersion -notmatch '^\d+\.\d+\.\d+$') { throw 'Set RS_APPROVED_RELEASE_VERSION to the same value used by the x64 lane.' }
. .\Tools\Versioning.ps1
$approvedVersionContext = Get-RSVersionContext -RepoRoot (Get-Location).Path -Configuration Release -Platform ARM64 -BuildNumber $approvedBuildNumber -OfficialRelease
$releaseVersion = [string]$approvedVersionContext.PackagingVersion
if ($releaseVersion -notmatch '^\d+\.\d+\.\d+$' -or $releaseVersion -cne $approvedReleaseVersion -or [int]$approvedVersionContext.BuildNumber -ne $approvedBuildNumber -or [string]$approvedVersionContext.Platform -cne 'ARM64' -or [string]$approvedVersionContext.Configuration -cne 'Release' -or -not [bool]$approvedVersionContext.OfficialRelease) { throw 'Pure ARM64 package version context does not match the approved identity; no provenance/package output is allowed.' }
$provenanceManifest = .\Tools\New-TerminalPackageProvenance.ps1 -RepositoryCommit $approvedRepositoryCommit -ReleaseVersion $releaseVersion -Mode Build -ToolchainLock .\Specs\Terminal\TerminalPackageToolchain.lock.json -PassThruManifest
if (-not (Test-Path -LiteralPath $provenanceManifest -PathType Leaf)) { throw 'Provenance preflight did not return its exact manifest path.' }
$provenanceSha256 = (Get-FileHash -Algorithm SHA256 -LiteralPath $provenanceManifest).Hash.ToLowerInvariant()
function Assert-RSTerminalOfficialBuildInputs {
    param([Parameter(Mandatory)][string]$Purpose)
    $verifiedManifest = .\Tools\New-TerminalPackageProvenance.ps1 -RepositoryCommit $approvedRepositoryCommit -ReleaseVersion $releaseVersion -Mode VerifyBuildInputs -ExpectedManifest $provenanceManifest -ToolchainLock .\Specs\Terminal\TerminalPackageToolchain.lock.json -PassThruManifest
    if (-not $? -or [IO.Path]::GetFullPath($verifiedManifest) -cne [IO.Path]::GetFullPath($provenanceManifest) -or (Get-FileHash -Algorithm SHA256 -LiteralPath $provenanceManifest).Hash.ToLowerInvariant() -cne $provenanceSha256) { throw "Official build-input recheck failed before $Purpose." }
}
function Assert-RSSavedTerminalVersionContext {
    param([Parameter(Mandatory)][string]$Purpose)
    $saved = Read-RSVersionContext -RepoRoot (Get-Location).Path
    if ($null -eq $saved -or [string]$saved.PackagingVersion -cne $releaseVersion -or [int]$saved.BuildNumber -ne $approvedBuildNumber -or [string]$saved.Platform -cne 'ARM64' -or [string]$saved.Configuration -cne 'Release' -or -not [bool]$saved.OfficialRelease) { throw "Saved ARM64 version context differs from the pure approved context after $Purpose." }
}
Import-Module .\Tools\TerminalEvidence.psm1 -Force
$expectedEngineLockSha256 = (Get-FileHash -LiteralPath '.\External\terminal-engine.lock.json' -Algorithm SHA256).Hash.ToLowerInvariant()
$arm64NativeManifests = [System.Collections.Generic.List[string]]::new()
.\build.ps1 -ProjectName TerminalTests -Platform ARM64 -Configuration Debug
$nativeManifest = .\Tools\Run-TerminalNativeTests.ps1 -Slice Core -Platform ARM64 -Configuration Debug -EvidenceRoot .\Specs\TestRuns -PassThruManifest
.\Tools\Verify-TerminalNativeEvidence.ps1 -Manifest $nativeManifest -ExpectedSlice Core -ExpectedRepositoryCommit $approvedRepositoryCommit -ExpectedEngineLockSha256 $expectedEngineLockSha256
$arm64NativeManifests.Add($nativeManifest)
$arm64AsanSupport = '<supported|unsupported copied from the approved TerminalEngineDecision record>'
if ($arm64AsanSupport -ceq 'supported') {
    .\build.ps1 -ProjectName TerminalTests -Platform ARM64 -Configuration "ASan Debug"
    $nativeManifest = .\Tools\Run-TerminalNativeTests.ps1 -Slice Core -Platform ARM64 -Configuration "ASan Debug" -EvidenceRoot .\Specs\TestRuns -PassThruManifest
    .\Tools\Verify-TerminalNativeEvidence.ps1 -Manifest $nativeManifest -ExpectedSlice Core -ExpectedRepositoryCommit $approvedRepositoryCommit -ExpectedEngineLockSha256 $expectedEngineLockSha256
    $arm64NativeManifests.Add($nativeManifest)
} elseif ($arm64AsanSupport -cne 'unsupported') {
    throw 'Replace ARM64 ASan support with the exact approved supported or unsupported disposition before running this lane.'
}
Assert-RSTerminalOfficialBuildInputs -Purpose 'ARM64 Release build'
try {
    $env:RSBuildEnableTests = 'true'
    .\build.ps1 -Platform ARM64 -Configuration Release -BuildNumber $approvedBuildNumber -OfficialRelease
} finally {
    Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue
}
Assert-RSSavedTerminalVersionContext -Purpose 'ARM64 Release build'
foreach ($slice in @('Core', 'Renderer', 'PaneProfiles', 'ShellState')) {
    $nativeManifest = .\Tools\Run-TerminalNativeTests.ps1 -Slice $slice -Platform ARM64 -Configuration Release -EvidenceRoot .\Specs\TestRuns -PassThruManifest
    .\Tools\Verify-TerminalNativeEvidence.ps1 -Manifest $nativeManifest -ExpectedSlice $slice -ExpectedRepositoryCommit $approvedRepositoryCommit -ExpectedEngineLockSha256 $expectedEngineLockSha256
    $arm64NativeManifests.Add($nativeManifest)
}
$expectedArm64NativeCount = if ($arm64AsanSupport -ceq 'supported') { 6 } else { 5 }
if ($arm64NativeManifests.Count -ne $expectedArm64NativeCount) { throw "The applicable ARM64 TerminalNative matrix did not produce exactly $expectedArm64NativeCount manifests." }
if (@($arm64NativeManifests | Select-Object -Unique).Count -ne $expectedArm64NativeCount) { throw 'The applicable ARM64 TerminalNative manifests are not unique exact paths.' }
$arm64NativeCloseoutIdentities = @($arm64NativeManifests | ForEach-Object {
    $manifestPath = (Resolve-Path -LiteralPath $_).Path
    $canonicalManifestPath = ConvertTo-RSCanonicalTestRunPath -RepoRoot (Get-Location).Path -ResolvedPath $manifestPath -Area TerminalNative -ExpectedLeaf 'terminal-native-evidence.v1.json'
    $recordDirectory = Split-Path -Parent $manifestPath
    $recordFiles = @(Get-ChildItem -LiteralPath $recordDirectory -File | Sort-Object Name | Select-Object -ExpandProperty Name)
    if (($recordFiles -join '|') -cne 'terminal-native-evidence.v1.json|terminal-native-evidence.v1.sha256') { throw "TerminalNative directory is not an exact two-file record: $recordDirectory" }
    $companionPath = Join-Path $recordDirectory 'terminal-native-evidence.v1.sha256'
    [pscustomobject][ordered]@{
        manifestPath = $canonicalManifestPath
        manifestSha256 = (Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash.ToLowerInvariant()
        companionSha256 = (Get-FileHash -LiteralPath $companionPath -Algorithm SHA256).Hash.ToLowerInvariant()
    }
})
if ($arm64NativeCloseoutIdentities.Count -ne $expectedArm64NativeCount -or @($arm64NativeCloseoutIdentities.manifestPath | Select-Object -Unique).Count -ne $expectedArm64NativeCount) { throw 'The ARM64 TerminalNative canonical closeout identity set is incomplete or duplicated.' }
$arm64FullCommandsRun = .\Tools\Run-TerminalCommandsEvidence.ps1 -Scenario FinalFull -Platform ARM64 -Configuration Release -RepositoryCommit $approvedRepositoryCommit -EvidenceRoot .\Specs\TestRuns -PassThruRunPath
.\Tools\Test-TestRunArchive.ps1 -RunPath $arm64FullCommandsRun
if ($LASTEXITCODE -ne 0) { throw 'The exact ARM64 Commands evidence failed validation.' }
Assert-RSTerminalOfficialBuildInputs -Purpose 'ARM64 ZIP package build'
.\build.ps1 -Configuration Release -Platform ARM64 -BuildNumber $approvedBuildNumber -OfficialRelease -Zip
Assert-RSSavedTerminalVersionContext -Purpose 'ARM64 ZIP package build'
Assert-RSTerminalOfficialBuildInputs -Purpose 'ARM64 MSIX package build'
.\build.ps1 -Configuration Release -Platform ARM64 -BuildNumber $approvedBuildNumber -OfficialRelease -Msix
Assert-RSSavedTerminalVersionContext -Purpose 'ARM64 MSIX package build'
$arm64ZipManifest = .\Tools\Test-TerminalPackage.ps1 -PackageRoot .\.build\AppPackages -RepositoryCommit $approvedRepositoryCommit -BuildNumber $approvedBuildNumber -ReleaseVersion $releaseVersion -ProvenanceManifest $provenanceManifest -ToolchainLock .\Specs\Terminal\TerminalPackageToolchain.lock.json -Kind Zip -Platform ARM64 -Mode ExtractAndSmoke -EngineLock .\External\terminal-engine.lock.json -EngineArtifactsRoot .\.build\TerminalEngine -EvidenceRoot .\Specs\TestRuns -PassThruManifest
$arm64UnsignedMsixManifest = .\Tools\Test-TerminalPackage.ps1 -PackageRoot .\.build\AppPackages -RepositoryCommit $approvedRepositoryCommit -BuildNumber $approvedBuildNumber -ReleaseVersion $releaseVersion -ProvenanceManifest $provenanceManifest -ToolchainLock .\Specs\Terminal\TerminalPackageToolchain.lock.json -Kind Msix -Platform ARM64 -Mode ExtractAndSmoke -EngineLock .\External\terminal-engine.lock.json -EngineArtifactsRoot .\.build\TerminalEngine -EvidenceRoot .\Specs\TestRuns -PassThruManifest
```

Then, on a clean disposable native x64 release-smoke machine, transfer the matching unsigned artifact, signed artifact, and the entire unchanged x64 unsigned-MSIX evidence run directory (JSON plus companion in its MachineHash/RunId layout). Do not transfer or accept a separate engine-artifact root; signed-package closure is checked in place against the production lock and bound unsigned evidence. Then run:

```powershell
if ($env:RS_APPROVED_REPOSITORY_COMMIT -notmatch '^[0-9a-f]{40}$') { throw 'Set RS_APPROVED_REPOSITORY_COMMIT to the reviewed package commit.' }
$approvedRepositoryCommit = [string]$env:RS_APPROVED_REPOSITORY_COMMIT
$actualRepositoryCommit = [string](git rev-parse HEAD)
if ($LASTEXITCODE -ne 0 -or $actualRepositoryCommit.Trim() -cne $approvedRepositoryCommit) { throw 'HEAD does not equal RS_APPROVED_REPOSITORY_COMMIT.' }
if ($env:RS_APPROVED_BUILD_NUMBER -notmatch '^[1-9]\d*$') { throw 'Set RS_APPROVED_BUILD_NUMBER to the approved package build number.' }
[int]$approvedBuildNumber = $env:RS_APPROVED_BUILD_NUMBER
$releaseVersion = [string]$env:RS_APPROVED_RELEASE_VERSION
if ($releaseVersion -notmatch '^\d+\.\d+\.\d+$') { throw 'Set RS_APPROVED_RELEASE_VERSION to the signed artifact three-part version.' }
$unsignedEvidenceManifest = [string]$env:RS_UNSIGNED_MSIX_EVIDENCE_MANIFEST
if (-not (Test-Path -LiteralPath $unsignedEvidenceManifest -PathType Leaf)) { throw 'Set RS_UNSIGNED_MSIX_EVIDENCE_MANIFEST to the transferred x64 unsigned MSIX evidence manifest.' }
$unsignedEvidenceDirectory = Split-Path -Parent $unsignedEvidenceManifest
$unsignedEvidenceFiles = @(Get-ChildItem -LiteralPath $unsignedEvidenceDirectory -File | Sort-Object Name | Select-Object -ExpandProperty Name)
if (($unsignedEvidenceFiles -join '|') -cne 'terminal-package-evidence.v1.json|terminal-package-evidence.v1.sha256') { throw 'The transferred unsigned evidence run directory must contain exactly its JSON and SHA-256 companion.' }
$provenanceManifest = .\Tools\New-TerminalPackageProvenance.ps1 -RepositoryCommit $approvedRepositoryCommit -ReleaseVersion $releaseVersion -Mode TransferredSmoke -ToolchainLock .\Specs\Terminal\TerminalPackageToolchain.lock.json -PassThruManifest
if (-not (Test-Path -LiteralPath $provenanceManifest -PathType Leaf)) { throw 'Transferred-smoke provenance preflight did not return its exact manifest path.' }
$x64SignedMsixManifest = .\Tools\Test-TerminalPackage.ps1 -PackageRoot .\.build\SignedAppPackages -UnsignedPackageRoot .\.build\AppPackages -UnsignedEvidenceManifest $unsignedEvidenceManifest -RepositoryCommit $approvedRepositoryCommit -BuildNumber $approvedBuildNumber -ReleaseVersion $releaseVersion -ProvenanceManifest $provenanceManifest -ToolchainLock .\Specs\Terminal\TerminalPackageToolchain.lock.json -Kind Msix -Platform x64 -Mode SignedInstallSmoke -EngineLock .\External\terminal-engine.lock.json -EvidenceRoot .\Specs\TestRuns -PassThruManifest
```

On a separate clean disposable native ARM64 release-smoke machine, transfer the matching artifacts and the entire unchanged ARM64 unsigned-MSIX evidence run directory. Do not transfer or accept a separate engine-artifact root. Then run:

```powershell
if ($env:RS_APPROVED_REPOSITORY_COMMIT -notmatch '^[0-9a-f]{40}$') { throw 'Set RS_APPROVED_REPOSITORY_COMMIT to the reviewed package commit.' }
$approvedRepositoryCommit = [string]$env:RS_APPROVED_REPOSITORY_COMMIT
$actualRepositoryCommit = [string](git rev-parse HEAD)
if ($LASTEXITCODE -ne 0 -or $actualRepositoryCommit.Trim() -cne $approvedRepositoryCommit) { throw 'HEAD does not equal RS_APPROVED_REPOSITORY_COMMIT.' }
if ($env:RS_APPROVED_BUILD_NUMBER -notmatch '^[1-9]\d*$') { throw 'Set RS_APPROVED_BUILD_NUMBER to the approved package build number.' }
[int]$approvedBuildNumber = $env:RS_APPROVED_BUILD_NUMBER
$releaseVersion = [string]$env:RS_APPROVED_RELEASE_VERSION
if ($releaseVersion -notmatch '^\d+\.\d+\.\d+$') { throw 'Set RS_APPROVED_RELEASE_VERSION to the signed artifact three-part version.' }
$unsignedEvidenceManifest = [string]$env:RS_UNSIGNED_MSIX_EVIDENCE_MANIFEST
if (-not (Test-Path -LiteralPath $unsignedEvidenceManifest -PathType Leaf)) { throw 'Set RS_UNSIGNED_MSIX_EVIDENCE_MANIFEST to the transferred ARM64 unsigned MSIX evidence manifest.' }
$unsignedEvidenceDirectory = Split-Path -Parent $unsignedEvidenceManifest
$unsignedEvidenceFiles = @(Get-ChildItem -LiteralPath $unsignedEvidenceDirectory -File | Sort-Object Name | Select-Object -ExpandProperty Name)
if (($unsignedEvidenceFiles -join '|') -cne 'terminal-package-evidence.v1.json|terminal-package-evidence.v1.sha256') { throw 'The transferred unsigned evidence run directory must contain exactly its JSON and SHA-256 companion.' }
$provenanceManifest = .\Tools\New-TerminalPackageProvenance.ps1 -RepositoryCommit $approvedRepositoryCommit -ReleaseVersion $releaseVersion -Mode TransferredSmoke -ToolchainLock .\Specs\Terminal\TerminalPackageToolchain.lock.json -PassThruManifest
if (-not (Test-Path -LiteralPath $provenanceManifest -PathType Leaf)) { throw 'Transferred-smoke provenance preflight did not return its exact manifest path.' }
$arm64SignedMsixManifest = .\Tools\Test-TerminalPackage.ps1 -PackageRoot .\.build\SignedAppPackages -UnsignedPackageRoot .\.build\AppPackages -UnsignedEvidenceManifest $unsignedEvidenceManifest -RepositoryCommit $approvedRepositoryCommit -BuildNumber $approvedBuildNumber -ReleaseVersion $releaseVersion -ProvenanceManifest $provenanceManifest -ToolchainLock .\Specs\Terminal\TerminalPackageToolchain.lock.json -Kind Msix -Platform ARM64 -Mode SignedInstallSmoke -EngineLock .\External\terminal-engine.lock.json -EvidenceRoot .\Specs\TestRuns -PassThruManifest
```

On the review machine, transfer the exact seven entire run directories (each JSON plus companion) under their unchanged MachineHash/RunId layouts, name the seven JSON paths explicitly, and finalize them:

```powershell
if ($env:RS_APPROVED_REPOSITORY_COMMIT -notmatch '^[0-9a-f]{40}$') { throw 'Set RS_APPROVED_REPOSITORY_COMMIT to the reviewed package commit.' }
$approvedRepositoryCommit = [string]$env:RS_APPROVED_REPOSITORY_COMMIT
$actualRepositoryCommit = [string](git rev-parse HEAD)
if ($LASTEXITCODE -ne 0 -or $actualRepositoryCommit.Trim() -cne $approvedRepositoryCommit) { throw 'HEAD does not equal RS_APPROVED_REPOSITORY_COMMIT.' }
if ($env:RS_APPROVED_BUILD_NUMBER -notmatch '^[1-9]\d*$') { throw 'Set RS_APPROVED_BUILD_NUMBER to the approved package build number.' }
[int]$approvedBuildNumber = $env:RS_APPROVED_BUILD_NUMBER
$releaseVersion = [string]$env:RS_APPROVED_RELEASE_VERSION
if ($releaseVersion -notmatch '^\d+\.\d+\.\d+$') { throw 'Set RS_APPROVED_RELEASE_VERSION to the approved three-part package version.' }
$provenanceManifest = .\Tools\New-TerminalPackageProvenance.ps1 -RepositoryCommit $approvedRepositoryCommit -ReleaseVersion $releaseVersion -Mode Review -ToolchainLock .\Specs\Terminal\TerminalPackageToolchain.lock.json -PassThruManifest
if (-not (Test-Path -LiteralPath $provenanceManifest -PathType Leaf)) { throw 'Review provenance preflight did not return its exact manifest path.' }
Import-Module .\Tools\TerminalEvidence.psm1 -Force
$noticesManifest = '<exact lock-declared notices manifest path>'
if (-not (Test-Path -LiteralPath $noticesManifest -PathType Leaf)) { throw 'Replace noticesManifest with the exact lock-declared reviewed file.' }
$packageRecordIds = @(
    'x64-zip-extract', 'x64-msi-extract', 'x64-msix-extract',
    'arm64-zip-extract', 'arm64-msix-extract',
    'x64-msix-signed', 'arm64-msix-signed'
)
$packageManifests = @(
    '<canonical x64 ZIP manifest>', '<canonical x64 MSI manifest>', '<canonical x64 unsigned MSIX manifest>',
    '<canonical ARM64 ZIP manifest>', '<canonical ARM64 unsigned MSIX manifest>',
    '<canonical x64 signed MSIX manifest>', '<canonical ARM64 signed MSIX manifest>'
)
$packageRecordIdentities = [System.Collections.Generic.List[object]]::new()
$canonicalPackagePaths = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
for ($index = 0; $index -lt $packageManifests.Count; $index++) {
    $manifest = $packageManifests[$index]
    if (-not $canonicalPackagePaths.Add([string]$manifest)) { throw "Duplicate or case-alias package evidence path: $manifest" }
    $manifestPath = Resolve-RSCanonicalTestRunPath -RepoRoot (Get-Location).Path -RepositoryRelativePath $manifest -Area TerminalPackage -ExpectedLeaf 'terminal-package-evidence.v1.json'
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { throw "Missing exact package evidence manifest: $manifest" }
    if ((ConvertTo-RSCanonicalTestRunPath -RepoRoot (Get-Location).Path -ResolvedPath $manifestPath -Area TerminalPackage -ExpectedLeaf 'terminal-package-evidence.v1.json') -cne $manifest) { throw "Package evidence path is not canonical: $manifest" }
    $recordDirectory = Split-Path -Parent $manifestPath
    $recordFiles = @(Get-ChildItem -LiteralPath $recordDirectory -File | Sort-Object Name | Select-Object -ExpandProperty Name)
    if (($recordFiles -join '|') -cne 'terminal-package-evidence.v1.json|terminal-package-evidence.v1.sha256') { throw "Evidence directory is not an exact two-file record: $recordDirectory" }
    $companion = Join-Path $recordDirectory 'terminal-package-evidence.v1.sha256'
    $packageRecordIdentities.Add([pscustomobject][ordered]@{
        recordId = $packageRecordIds[$index]
        manifestPath = $manifest
        manifestSha256 = (Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash.ToLowerInvariant()
        companionSha256 = (Get-FileHash -LiteralPath $companion -Algorithm SHA256).Hash.ToLowerInvariant()
    })
}
$packageEvidenceSetOutput = @(& .\Tools\Verify-TerminalPackageEvidence.ps1 -Mode Create -EvidenceManifest $packageManifests -EvidenceRoot .\Specs\TestRuns -RepositoryCommit $approvedRepositoryCommit -BuildNumber $approvedBuildNumber -ReleaseVersion $releaseVersion -ProvenanceManifest $provenanceManifest -ToolchainLock .\Specs\Terminal\TerminalPackageToolchain.lock.json -PassThruManifest)
if ($LASTEXITCODE -ne 0 -or $packageEvidenceSetOutput.Count -ne 1) { throw 'Package evidence finalization did not return exactly one manifest path.' }
$packageEvidenceSetPath = (Resolve-Path -LiteralPath ([string]$packageEvidenceSetOutput[0])).Path
$packageEvidenceSet = ConvertTo-RSCanonicalTestRunPath -RepoRoot (Get-Location).Path -ResolvedPath $packageEvidenceSetPath -Area TerminalPackage -ExpectedLeaf 'terminal-package-evidence.v1.json'
$packageEvidenceSetDirectory = Split-Path -Parent $packageEvidenceSetPath
$packageEvidenceSetFiles = @(Get-ChildItem -LiteralPath $packageEvidenceSetDirectory -File | Sort-Object Name | Select-Object -ExpandProperty Name)
if (($packageEvidenceSetFiles -join '|') -cne 'terminal-package-evidence.v1.json|terminal-package-evidence.v1.sha256') { throw 'Package evidence set is not an exact two-file record.' }
$packageEvidenceSetCompanion = Join-Path $packageEvidenceSetDirectory 'terminal-package-evidence.v1.sha256'
$packageEvidenceSetRecord = Get-Content -LiteralPath $packageEvidenceSetPath -Raw | ConvertFrom-Json
if ([string]$packageEvidenceSetRecord.overallResult -cne 'pass') { throw 'Package evidence set did not record pass.' }
[pscustomobject][ordered]@{
    engineDecisionSha256 = (Get-FileHash -LiteralPath '.\Specs\Terminal\TerminalEngineDecision.md' -Algorithm SHA256).Hash.ToLowerInvariant()
    noticesManifestPath = $noticesManifest
    frozen = $packageEvidenceSetRecord.frozen
    inputHarness = $packageEvidenceSetRecord.inputHarness
    verifierHarness = $packageEvidenceSetRecord.harness
    packageRecords = $packageRecordIdentities.ToArray()
    packageEvidenceSetManifest = $packageEvidenceSet
    packageEvidenceSetSha256 = (Get-FileHash -LiteralPath $packageEvidenceSetPath -Algorithm SHA256).Hash.ToLowerInvariant()
    packageEvidenceSetCompanionSha256 = (Get-FileHash -LiteralPath $packageEvidenceSetCompanion -Algorithm SHA256).Hash.ToLowerInvariant()
}
```

Expected: native x64/ARM64 builds and required runtime suites, budgets, and extracted-package smokes pass; MSI remains x64-only. Signed MSIX install/activate/uninstall passes before it is claimed for release. Unsafe paste passes immutable one-read/frozen-identity/storage/supersession/revocation/final-admission/scrub tests; generic disable/normal refresh alone preserve a still-selected/visible/input-owning/unexpired broker-eligible prompt, while independent broker cancellation remains final. OSC 52 passes immediate background denial before one/eight counters, strict deadline, non-focus UIA/accelerator behavior, revoke-before-publication/effective-disable-deny, exact final clipboard linearization, and scrub tests. Every package contains exact `Plugins\Terminal.dll`, schema, and notices, contains no loose PowerShell startup/adapter/broker manifest/schema/script/wrapper/resource or WSL bootstrap/wrapper/profile/staged-script payload, and preserves the exact embedded startup/observer/broker implementation plus the usable five-token WSL launch/permanent truthful terminal-only state. Plugin missing/corrupt/wrong-architecture mutations and every dynamic private engine/closure mutation degrade only Terminal while app composition starts; source/static packages contain no separate engine DLL and pass the compiled identity seam. Available cases create only `IID_ITerminal`, reject `IID_IViewer`, and prove callback/module quiet. Archived evidence is Release-quality and content-free. Closeout additionally requires the exact returned `Specs/TestRuns/<MachineHash>/TerminalPackage/<RunId>/terminal-package-evidence.v1.json` evidence-set path and its verified companion digest; that evidence set must contain exactly the seven named run-manifest digests and prove equality across x64 and ARM64 for the approved `repositoryCommit`, source-input manifest, `buildNumber`, `releaseVersion`, the complete `frozen` identity object, and evidence-helper/harness identities, with every run's resolved tool identities matching the shared reviewed toolchain lock.

The same expected result includes complete untruncated profile/location/display identities and fixed-basename Windows argv0 plus the independent WSL five-token boundary; every frozen named schema/runtime boundary disposition; cmd's exact embedded-script/six-token-argv/eight-key/protected-pipe/root-PID/fixed-order-ASCII-through-EOF bootstrap and truthful ongoing terminal-only status, including its trusted-AutoRun/same-root limitation probe; exact all-local/UNC PowerShell startup with the four-token fail-closed wrapper, protected nine-kind exchange, order-independent `StartupSemanticJoin`, lifecycle-locked Running CAS, authenticated prompt commit, and CommitAck-gated input/memory/capability; one decision-delivered noninherited PowerShell/pwsh semantic epoch with permanent revocation and restart-only rekey; exact manifest-gated first-statement-`$?` ReadLine observation with `AcceptedInput`/matching `RootReadReady`, no prompt/control-wake inference, conservative failure, inert-pass-through disable, real Windows PowerShell 5.1/every-shipped-pwsh proof, and a read-only upgrade audit; no nested integration; and exactly one plugin retirement owner plus one host removal claimant through every pending-confirmation/natural-exit/forced-close race. These cases are nonzero, passed, and unskipped in their required native, Commands, Preferences/UIA, lifecycle, source-contract, dependency-audit, schema, and package lanes.

Broker closeout additionally proves one plugin-owned content-free slot per view, the exact four-state Close-over-Paste-over-OSC table, stable-ID background close selection/reveal/focus/revalidation with no invisible prompt or post-selection rollback, strict 60-/30-second equality expiry, background/inactive OSC denial and scrub before counters with no focus, selection/activation cancellation with no resurrection, modal focus/default/Escape/Enter restoration rules, OSC `Alt+A`/`Alt+D` plus UIA behavior, History/context exclusion, and broker/domain/post cleanup before teardown. Eight background asks plus `+1`, every pairwise/three-way overlap, reorder/removal/HWND reuse, missing/failed posts, and stale replies are nonzero, passed, unskipped, and content-free.

## Authoritative documentation and WIP closeout

Before completion, merge durable requirements into:

- `Specs/Terminal/Terminal_EmbeddedPane.md` for the exact untruncated identity/launch boundaries, cmd terminal-only launch bootstrap including prelaunch-free-drive capture, bounded exact logical UNC root/tail reverse proof, retained nonsecret mapping ownership and root-death-only unchanged-mapping cancellation, all-local/UNC PowerShell fail-closed wrapper/protected nine-kind startup/post-profile rooting/decision/pre-reserved join/`RevokedTerminalOnly` paired cleanup servers/split Running-prompt commit/two-pipe policy-race contract, the one decision-delivered noninherited semantic epoch/restart-only-rekey and persistent-message/receive-deadline contract, exact no-Flush OVERLAPPED control/wake/deadline/ambiguous-input rules, adapter manifest/interposer/ReadLine/`AcceptedInput`/`RootReadReady`/replacement/inert-disable/upgrade-audit rules, the single retirement-intent arbiter/shared host removal claim, one per-view interaction broker with stable-ID background close and the full state/priority/deadline/focus/UIA/history-context/lifecycle contract, immutable one-read unsafe-paste and generation-bound OSC 52 authorization/revocation transactions plus qualified disable/refresh paste continuity, and `TerminalState.schema.json`, `TerminalStatePrivacyTransaction.schema.json`, their Draft 2020-12 code-point-backstop/nonassertive UTF-16/UTF-8 annotation/runtime-authority contract, and exact named valid/invalid corpus;
- new authoritative `Specs/Plugins/Plugins_Terminal.md` for the terminal factory/IID/configuration/child-HWND/callback/quiet/unload contract, plugin-only Windows startup/semantic implementation boundary, and the per-view interaction-broker/domain-token ownership boundary, and `Specs/Plugins/Plugins_ViewerPlugins.md` only for the explicit not-`IViewer`/not-association/not-Preview boundary;
- `Specs/Terminal/TerminalPackageEvidence.schema.json`,
  `Specs/Terminal/TerminalPackageEvidenceCorpus/`,
  `Specs/Terminal/TerminalPackageToolchain.lock.json`,
  `Specs/Terminal/TerminalPackageToolchainLock.schema.json`, and its
  valid/invalid corpus, including the non-secret signing-certificate/timestamp
  policy;
- `Specs/UI/UI_CommandMenuKeyboard.md`, `UI_DxUiWinUIDesign.md`, `UI_FolderWindow.md`, `UI_PreferencesDialog.md`, and `UI_ManagePluginsDialog.md` for generic conditional action rendering, destructive confirmation, refresh/Retry status, background-close selection/focus/no-rollback behavior, the one broker surface, modal default/focus restoration/Escape/Enter, OSC `Alt+A`/`Alt+D` and UIA, History/context exclusion, accessible plugin-owned unsafe-paste and OSC 52 actions with no host payload/token exposure, qualified paste continuity versus OSC revocation status, and truthful WSL launchable-but-terminal-only capability/settings status;
- `Specs/Core/Core_SettingsStore.md` and `Specs/SettingsStore.schema.json` for opaque `plugins.configurationByPluginId["builtin/terminal"]` with no top-level terminal domain, plus `Specs/Core/Core_SharedHelpers.md` for the canonical WSL catalog and any actual helper additions;
- theme specs/files, installer/portable specs, testing coverage/performance specs, and the `Specs/TestRuns/README.md` definition for the content-free `TerminalPackage` archive area (purpose, exact two-file record, required manifest/digest/provenance/tool fields, retention/verifier entrypoint, and prohibited source/package/extraction/path/secret/terminal content), including source/package rejection of loose PowerShell startup/adapter resources and every WSL bootstrap/wrapper/profile/staged-script payload plus the exact v1 WSL acceptance matrix; plan 1 owns the existing `TerminalEngine` definition, which plan 6 revalidates but does not redefine;
- `Specs/Plugins/Plugins_PluginAPI.md` for the shared loader's separate normal-refresh and process-shutdown outcome policies, plus `LICENSE.txt`, capability/engine decision records, and the dependency lifecycle runbook with the exact offline/online status commands, lock ownership and periodic/manual cadence, engine exact-pin Collect/Finalize/apply matrix, separate package-toolchain update rule, security-advisory escalation, one-change consistency rule, and previous-lock rollback/rebuild procedure.

Link final engine/terminal perf/platform/package evidence in the master. Confirm no normative requirement remains only in WIP/child files. Each child moves to `Specs/Plans/Done/` after its own implementation, evidence, and durable-spec updates are complete. Move the master last, only after all six children are Done and every master Done checkbox passes; update the WIP index at each move.

Those durable Terminal, Preferences, lifecycle, UIA, security, and testing owners
must also carry the exact seven-section/48-editable-element navigation map,
hidden `configurationVersion`, diagnostic-link corpus, and frozen
`HostPluginPreferencesRequest` `80/8` ABI; the one serialized Apply/Open-
registration/registry-removal
coordinator, stable captured membership, Prepare/Abort/atomic-commit-token/
allocation-free-noexcept Commit-release algorithm, fatal invariant policy,
queueing/retiring outcomes, no-lock-across-external-work rule, and immutable
live-policy snapshot table; the exact 250-ms/no-queue/drop-counter and 2-DIP/
150-ms/animation/high-contrast visual bell contract, the non-UI
`TrySubmitThreadpoolCallback`/`AcquireModuleReferenceFromAddress(...)`/
callback-entry `FreeLibraryWhenCallbackReturns(instance, pin.release())`
one-shot `MessageBeep(MB_OK)` rule, pin/submission/OS-failure content-free-
counter behavior, quiet-after-call/submitted-work rule, the `none`/view-
generation/teardown visual cancellation rule, and the 8-KiB/C0/space/scheme
precheck, exact `CreateUri` flags, canonical component rules, closed case-
insensitive `http|https|mailto` hyperlink allowlist, atomic Ctrl/policy/view/
document-generation activation races, and the sole plugin-owned
`ShellExecuteW(pluginChildHwnd, L"open", validatedImmutableUri.c_str(), nullptr,
nullptr, SW_SHOWNORMAL)`/`>32` success/no-fallback contract. The durable helper
catalog records why this remains plugin-local unless a genuinely matching
cross-consumer helper is introduced. These contracts may not remain only in
this closeout plan.

The durable Terminal, shell-integration, lifecycle, security, dependency, and
testing owners must likewise carry the exact 64-KiB-buffer/2-MiB-reserved
`TerminalSemantic` server, Connect-before-Read and one-persistent-connection
rules, successful-Read/message-boundary plus EOF dispositions, `RSTSEM01`
framing/ordinary-sequence rules, closed four-kind byte-exact payload ABI and
canonical-JSON/generated-constant identity, strict first-chunk receive deadline, single
four-argument asynchronous `WriteAsync` message, real-tuple proof, strict 50-ms
foreground observation/PowerShell-5.1 wait, Task/token/buffer ownership, no-
overlap and handshake-boundary-race disposition; the paired `TerminalControl`
read sentinel, both-pipe disable/teardown close, all-I/O/parser/module quiet gate;
and exact read/read-attributes/write-attributes/synchronize/no-write-data DACL/
client mode-switch proof,
no-Flush OVERLAPPED `RSTCTL01` submit/writer-admission/wake/deadline/physical-key/
ambiguous-input-hold contract. It may not remain only in the master
or this child.

## Child completion and move

- **Completion record:** `TERMINAL-HANDOFF-UNRESOLVED` — replace this line with the exact plan-5 closeout-commit/blob pair; all predecessor evidence repository commits/digests; engine-decision digest; the lock-declared notices-manifest path and complete nine-field package `frozen` object; package-run input and evidence-set verifier harness identities; final TerminalNative and Commands/perf archive identities; all seven ordered TerminalPackage manifest/companion paths and digests; and the returned package evidence-set path plus manifest/companion digests required below.

This child is complete only when all of the following are true:

1. Revalidate `Specs/Plans/Done/Terminal_ShellIntegrationHistoryAndPaneFollow_2026-07-22.md` and every predecessor evidence path/digest from its handoff. Record/recompute the accepted plan-5 closeout commit, that commit's Done-plan Git blob object ID, and every predecessor evidence `repositoryCommit`. All five prior child basenames must exist under `Specs/Plans/Done/`, each must contain exactly `- **State:** Done`, resolve to its recorded blob at its recorded closeout commit, and rely on no implicit latest run.
2. Finish the master-defined content-free runner and explicit validator, including the generic-archive suppression/receipt tests and `Tools/Test-TestRunArchive.ps1 -RunPath <string[]>`; directory scanning or implicit latest selection is forbidden in terminal closeout. Retain exactly three final Commands runs returned by `Run-TerminalCommandsEvidence.ps1`: x64 Release `FinalPerf`, x64 Release `FinalFull`, and native ARM64 Release `FinalFull`. Each has canonical `results.json`, `trace.txt`, and `run-all-tests-results.json`; only `FinalPerf` requires `perf/perf_metrics.jsonl`, while each `FinalFull` records exact `not-applicable:scenario-has-no-perf` unless that frozen scenario emits reviewed metrics. Validate the three exact paths together with `Test-TestRunArchive.ps1 -RunPath`, run `Show-PerfRuns.ps1 -Area Commands -Run <exact-x64-FinalPerf-path> -Metric terminal.render.paint_us -FailOnQuality -ShowBuildFlavor`, and record every canonical file digest plus the recomputed master-defined `CommandsEvidenceSetSha256`. Reverify and record the exact 14 required TerminalNative JSON/companion pairs—plus only the applicable ARM64 Core ASan pair—and compute `TerminalNativeManifestSetSha256` through plan 1's `Get-RSJcsIdentitySetSha256` over the master's exact record shape. All hard invariants, plugin-only source contract, IID/child/callback/module-quiet tests, plugin-owned per-host view-admission/rollback/lowering/release tests, complete normal disable/re-enable/refresh/Retry/replacement/app-shutdown lifecycle tests, the exact profile/location/display identity-cap and Windows-fixed-argv0/WSL-command-boundary corpus, cmd's complete correlated local/plugin-backed/UNC bootstrap and terminal-only limitation matrix plus all-local/UNC PowerShell startup and one decision-delivered noninherited epoch/permanent-revocation/restart-only-rekey tests, the single retirement-intent arbiter/shared-host-removal-claim race matrix, the immutable one-read unsafe-paste identity/storage/supersession/revocation/final-admission/scrub matrix including generic disable/refresh continuity, the OSC 52 identity/one-eight-cap/revoke-before-publication/effective-deny/final-write/scrub matrix, Preferences schema/all-fields/link tests, persisted `followPathWhenIdle` Apply-false/true versus concurrent Open/Restart snapshot races, existing-override stability, Restart resampling, Cancel no-effect, and session-toggle nonpersistence, exact privacy action/context/status/confirmation tests, the privacy-transaction schema corpus, family-wide recovery deletion/reset and stale-writer/downgrade non-resurrection tests, command migration/Edit/path insertion tests, the complete WSL exact-launch/no-injection/permanent-`RequestedUnverified`/zero-capability/no-memory/truthful-UI/source-package rejection matrix, full regressions, required native x64/ARM64 lanes, reviewed budgets, and sample-quality gates must pass.
   Startup completion within item 2 additionally requires the exact plugin files/project/resource maps and real Windows PowerShell 5.1/every shipped pwsh tuple: four-token fail-closed wrapper, closed literal encoder/exact AST with fixed `*>&1`, six parameters, exact typed finalize object, wrapper-validation-before-one-shot output-silent complete-frame CommitAck Write with no Flush/retry/later operation and server-full-read-before-gates, exit 125/126 for invalid/finalizer/merged-stream output, profile-before-script/policy-blocked-no-prompt behavior, protected root-PID-bound pipe, exact HMAC frame/nine-kind order/module-qualified post-profile rooting, decision-time policy, both pre-reserved `StartupSemanticJoin` orders/every slot failure/nonblocking lock discipline, `RevokedTerminalOnly` retained no-operation `TerminalSemantic`/`TerminalControl` cleanup servers plus Ready/Unavailable Prepared outcomes and the valid-staged-Handshake/Unavailable boundary race, ReleaseAck noncommit, one lifecycle-locked Running CAS winner, exact CommitPrompt disposition/generation matrix including same-generation Enable/Unavailable TerminalOnly and either valid Enable outcome after latched disable, CommitAck keyed independently of a latched live disable, before/after Running-CAS/Prompt/finalizer-Write/server-receipt disable races with cleanup/input/integration-independent memory but zero capability/no retirement, both-pipe close/cancel with completed-control-read or semantic-write-failure inert latching, strict deadline, unrelated pre/post-CAS retirement outcomes, no input/memory/capability before CommitAck, no target/semantic secret in argv/environment, and no false immutable-string erasure claim. No loose startup resource, fake-only result, skip, policy bypass, ReleaseAck/EOF prompt release, or unproved handshake is admissible.
   Semantic-transport completion within item 2 additionally requires every real
   supported shell tuple to prove the protected exact-root-PID `TerminalSemantic`
   pipe, precharged 2-MiB assembly reservation against the 64-KiB pipe buffer,
   Connect-before-Read, persistent no-reconnect epoch, one overlapped read/rearm-
   before-dispatch path, successful-Read delimiter/EOF dispositions, strict first-
   chunk receive clock, exact `RSTSEM01` frame/ordinary sequences/second-
   Handshake rejection, and one four-argument asynchronous `WriteAsync`
   preserving a 2-MiB message without synchronous large-write stall. The fresh
   Task/token/buffer, strict-before 50-ms observation, exact PowerShell-5.1 wait,
   successful-only sequence, no-overlap, late observation/scrub, failed-emission
   no-replacement, and cancel/stream-close/inert/no-callback/no-background
   behavior plus valid-staged-Handshake/Unavailable boundary race must pass every
   bound/failure/lifecycle race; an unproved tuple is `SemanticUnavailable`,
   never skipped or given a fallback. The same gate proves the canonical JSON,
   generated C++/PowerShell byte identity and complete four-kind payload schema/
   min-max-invalid corpus. It also proves `TerminalControl` is
   the read sentinel; exact ReadData/ReadAttributes/WriteAttributes/Synchronize/
   no-WriteData
   production DACL plus async anonymous noninherited client mode-switch proof;
   both-pipe fence/close/drain and module quiet; exact
   `RSTCTL01`, no `FlushFileBuffers`/synchronous Write, one manual-reset-event
   OVERLAPPED submit and writer-admission point, one bounded
   `GetOverlappedResultEx`, exactly-one-wake/no-retry,
   one staged-record slot, `PollControlTransport(NoWait|HandlerWaitOnce)`, NoWait
   stage/rearm without dispatch, handler stage-first zero-wait/EndRead path,
   handler-only nonalertable `WaitOne(250)` when empty, already-complete/wait-
   signaled EndRead, rearm/final-QPC/one-shot
   dispatch, physical/no-operation/mapping-replacement dispositions, the closed
   Success and Follow/Insert safe-pre-effect return-to-Idle table, and ambiguous-
   effect input retention through actual completion/five-second quiet.
   In that startup completion gate, "exact typed finalize object" means first PSTypeName `RedSalamander.TerminalBootstrapFinalize.V1`, ordered properties exactly `Sentinel,Finalize`, Sentinel exactly `RSTBOOT-WRAPPER-READY-1`, and a scriptblock Finalize. Repeat each Ready and Unavailable disable case before Prompt construction, after construction/before send, after send/before child receipt, between wrapper validation/finalizer invocation, between finalizer Ack Write/server receipt, during Ack acceptance, and immediately after acceptance.
   Observer/schema completion within item 2 additionally requires real Windows PowerShell 5.1 and every shipped pwsh tuple, the exact embedded manifest/interposer/first-statement-`$?`/two-or-three-argument ReadLine recipes, same-object exact returned line, `AcceptedInput` then matching next-root `RootReadReady`, parse-failure completion, zero control-wake acceptance, unchanged prompt/native state, conservative replacement/failure behavior, process-lifetime inert-pass-through disable, and the byte-preserving read-only installed-adapter audit. It also requires the four closed named BMP/astral/multibyte schema/runtime dispositions, no schema-invalid/runtime-valid case, and serializer acceptance by both layers. No fake-only shell result, skip, blanket parity assertion, or loose adapter resource is admissible.
   Interaction completion within item 2 additionally requires the exact broker files/project maps, one slot/four states/full token identity, all 12 priority-table cells, stable-ID background-close handoff, no invisible prompt or selection rollback, 60-/30-second strict deadlines, immediate background OSC deny before counters, no focus stealing, selection/deactivation cancellation with no resurrection, exact modal/accelerator/UIA/History/context behavior, and lifecycle/post scrub. The eight-background-asks-plus-`+1`, overlap, reorder/removal/HWND-reuse, failed-post, and stale-reply cases must pass in every named lane; no host broker/token/payload implementation or loose package resource is admissible.
3. Finalize the package matrix only from the five exact unsigned and two exact signed two-file run directories. Record, in canonical order, every run JSON path plus its manifest and companion file SHA-256. Record the returned exact `Specs/TestRuns/<MachineHash>/TerminalPackage/<RunId>/terminal-package-evidence.v1.json` evidence-set path plus both its manifest and companion file SHA-256. Completion requires `overallResult=pass`, exactly seven unique ordered input-manifest digests, correct signed-to-unsigned bindings, the approved `repositoryCommit`, engine-decision digest, and cross-architecture equality of the source-input manifest, approved `buildNumber`, `releaseVersion`, and complete `frozen` object (`engineLockSha256`, `runtimeDependenciesSha256`, `releaseArtifactPolicySha256`, `terminalStateSchemaSha256`, `noticesManifestSha256`, `terminalEvidenceHelperSha256`, `sourceInputManifestSha256`, `packageToolchainLockSha256`, and `packageProvenanceHelperSha256`), plus input-harness and evidence-set-verifier identities. Every per-run tool tuple must match the shared lock. These exact identities are the only admissible Phase-9 handoff; a missing companion, substitution, reorder, mismatch, implicit latest discovery, or unverifiable identity blocks completion.
4. Pass `TerminalDependencyLifecycle.Tests.ps1`, then run the fast locked status command with both `External/terminal-engine.lock.json` and `TerminalPackageToolchain.lock.json`; it must validate the complete engine/tool/runtime/package/license/ABI inventory offline and leave `git status --porcelain=v1 --untracked-files=all` byte-for-byte unchanged. Online discovery is exercised deterministically against local fake sources; the live `-Mode Online -RequireOnline` command remains an explicit operator/security-maintenance action and never silently blocks an otherwise reproducible build because a metadata service is down.
5. Merge all remaining durable requirements into the exact owners listed above:
   `Specs/Terminal/Terminal_EmbeddedPane.md`, new
   `Specs/Plugins/Plugins_Terminal.md`, `Specs/Plugins/Plugins_PluginAPI.md`, the
   not-Terminal cross-reference in `Specs/Plugins/Plugins_ViewerPlugins.md`,
   `Specs/Terminal/TerminalState.schema.json`,
   `Specs/Terminal/TerminalStatePrivacyTransaction.schema.json` and its corpus,
   `Specs/Terminal/TerminalPackageEvidence.schema.json`,
   `Specs/Terminal/TerminalPackageEvidenceCorpus/`,
   `Specs/Terminal/TerminalPackageToolchain.lock.json`,
   `Specs/Terminal/TerminalPackageToolchainLock.schema.json` and its corpus,
   `Specs/Core/Core_SettingsStore.md`, `Specs/SettingsStore.schema.json`,
   `Specs/Core/Core_Localization.md`, `Specs/Core/Core_SharedHelpers.md`,
   `Specs/UI/UI_CommandMenuKeyboard.md`, `Specs/UI/UI_DxUiWinUIDesign.md`,
   `Specs/UI/UI_FolderWindow.md`, `Specs/UI/UI_PreferencesDialog.md`,
   `Specs/UI/UI_ManagePluginsDialog.md`, shipped theme/installer/testing/perf
   owners, the plan-6 `Specs/TestRuns/README.md` `TerminalPackage` section,
   `RuntimeDependencies.props`, `LICENSE.txt`, and every engine/update/adapter/
   capability/signing/release/provenance/toolchain contract actually changed.

   The durable Terminal/Preferences/Manage Plugins contracts must own plugin-
   local `maxTabsPerPane`; disable/re-enable/refresh/admission/settlement/Retry/
   replacement/shutdown lifecycle; and cmd's exact embedded-script, argv,
   environment, protected pipe, PID/correlation, clean-EOF Ack, prelaunch-free
   drive, bounded `WNetGetUniversalNameW` complete/root/tail proof, retained
   nonsecret owner tuple, root-death-only unchanged-mapping
   `WNetCancelConnection2W` cleanup, and AutoRun/same-root limitations. They must
   also own the exact all-local/UNC PowerShell wrapper/HMAC startup, pre-reserved
   join including valid-staged-Handshake/Unavailable boundary race,
   `RevokedTerminalOnly` paired cleanup servers, ReleaseAck/Running CAS/
   CommitPrompt/CommitAck, both-pipe policy fence, control-read/semantic-write
   inert latch, persistent semantic-message transport and four-kind generated
   payload ABI, plus no-Flush OVERLAPPED control/wake/deadline/ambiguous-input/
   quiet contract, and all input/memory/
   capability/failure gates.

   The same owners retain the exact observer, broker, unsafe-paste, OSC 52,
   action/context, privacy recovery, and WSL terminal-only contracts listed in
   this plan. Revalidate plan 1's `TerminalEngine` archive without taking
   ownership or silently rewriting it. Schemas, runtime, package contents, and
   documentation must agree with the frozen Draft-2020-12 code-point backstop
   versus authoritative runtime-unit relationship; no normative rule may remain
   only in this child or the master.
6. Link the exact Commands paths/digests and TerminalPackage evidence-set path/digest in the master, but leave the master WIP. In the same reviewed closeout change, replace this file's state line with exactly `- **State:** Done`, run the exact move below, and update the N3 next-step cell in `Specs/Plans/WIP/README.md` to exactly: `All six terminal child plans are complete under Specs/Plans/Done/. Continue only with Phase 9 of Terminal_EmbeddedConPtyLibGhosttyVt_IntegrationPlan_2026-07-13.md; the master remains WIP until every Phase 9 check passes.`

```powershell
$source = 'Specs/Plans/WIP/Terminal_SettingsSecurityPerfPackagingCloseout_2026-07-22.md'
$destination = 'Specs/Plans/Done/Terminal_SettingsSecurityPerfPackagingCloseout_2026-07-22.md'
if (-not (Select-String -LiteralPath $source -SimpleMatch '- **State:** Done' -Quiet)) { throw 'Set the closeout child state to Done before moving it.' }
if (Test-Path -LiteralPath $destination) { throw "Destination already exists: $destination" }
git mv -- $source $destination
if (-not (Select-String -LiteralPath 'Specs/Plans/WIP/README.md' -SimpleMatch 'All six terminal child plans are complete under Specs/Plans/Done/. Continue only with Phase 9 of Terminal_EmbeddedConPtyLibGhosttyVt_IntegrationPlan_2026-07-13.md; the master remains WIP until every Phase 9 check passes.' -Quiet)) { throw 'Update WIP README N3 to master Phase 9.' }
```

The master remains under `Specs/Plans/WIP/` until Phase 9 independently revalidates the six Done children, their durable-spec ownership, the exact evidence digests, and every master Done checkbox.

## STOP conditions

- Preferences could omit/add/duplicate/rename/case-change/multiply place any of
  the exact 48 editable element IDs; infer Advanced membership; expose
  `configurationVersion`; alter a section ID/order/collapse state or per-section
  UI order; allow a diagnostic to open an approximate/unknown target; use a
  shortened request type; or drift the frozen
  `HostPluginPreferencesRequest` `80/8` ABI/validation/posted-lifetime contract.
- Live Apply could publish without the single coordinator/commit token; admit an
  Open or registry removal into captured membership after `BeginApply`; hold the
  coordinator, service, two session gates, or any session gate across a wait,
  callback, UI/COM call, process/pipe I/O, or coordinator call; omit stable-order
  Abort/full rollback; let a retiring/disappearing session invalidate membership;
  let post-Prepare root exit/close bypass the outcome; report success before
  allocation-free/nonblocking/noexcept/infallible Commit/release; introduce an
  ordinary post-publication rollback/failure; continue after a token/identity/
  staged mismatch instead of closing admission and invoking production
  `RaiseFailFastException` through `FatalConfigurationInvariantPolicy`; or expose
  a partial/mixed policy generation.
- Bell handling could queue/replay rate-limited effects, mix per-view policy
  generations, omit dropped/failure counters, render a visual bell other than
  the exact 2-DIP 150-ms animation/high-contrast contract, mutate model/UIA/
  layout/focus, clear a newer/reused-view flash from a stale timer, leave a flash
  after `none`/view-generation/teardown, call `MessageBeep` on parser/UI, omit
  `TrySubmitThreadpoolCallback` or the pre-submit
  `AcquireModuleReferenceFromAddress(...)` pin, fail to transfer the pin at
  callback entry with
  `FreeLibraryWhenCallbackReturns(instance, pin.release())`, destroy the last
  `wil::unique_hmodule` inside the callback, cancel already-submitted work at
  teardown, leak pin/quiet ownership or beep/retry/fallback after pin/submit
  failure, add an OS-mute fallback, or decrement quiet before `MessageBeep`
  returns.
- Hyperlink activation could permit any v1 scheme outside case-insensitive
  `http|https|mailto`; omit the 8-KiB/C0/space/explicit-scheme precheck, exact
  one-call `CreateUri` flags/reserved zero, canonical host/user-info/password/
  mailto component validation, or final policy/view/document-generation gate;
  issue an external open from relative/malformed/control/over-limit/stale text;
  use host-owned policy, `CreateProcess`, command concatenation, non-null
  parameters/cwd, or a retained process handle; omit the plugin child HWND or
  immutable URI from the exact `ShellExecuteW` call; treat `<=32` as success; or
  retry/fallback/perform a second side effect after failure.

- Settings would require a v16 bump without a separately approved safe migration.
- Settings require a top-level terminal object, host-side terminal validation/defaults, a JSON-only required control, or a second enable switch that can disagree with plugin state.
- Unsafe paste would reread/reclassify mutable clipboard data, omit identity or policy/input-mode generations, exceed one pending/session or fixed 8 MiB/session and 32 MiB/service storage, authorize without the final current-mode/liveness/cap/descriptor/queue atomic check, fail to scrub, expose payload/token to the host/evidence, be revoked/rejected solely by generic disable/normal refresh for an otherwise eligible live frozen-config view, or survive/resurrect after the independent 60-second/selection/visibility/activation/input-owner/lifecycle broker cancellation.
- OSC 52 cannot enforce immutable identity/opaque plugin-owned payload, immediate background/inactive denial before one-ask/session and eight/service counters without focus/selection, strict 30-second expiry, revoke-before-selection/policy/disable/refresh/teardown publication, effective deny while disabled, stale-token rejection, a sole clipboard replacement linearized under the same gate, one-shot consumption, and secure scrub.
- Any accepted profile/location/display identity would be truncated, hashed, aliased, or only partly persisted; Windows launch could not retain exact nonnull `lpApplicationName` plus fixed basename argv0; or WSL's complete identity and five-token command-line bounds could not reject independently before side effects.
- Draft 2020-12 `maxLength` would be treated as a UTF-16-unit or UTF-8-byte guarantee; an applicable property lacks the exact nonassertive annotation; runtime cannot enforce the frozen unit/byte ceilings; the closed BMP/astral/multibyte dispositions or serializer-both-layers proof is absent; or any schema-invalid/runtime-valid value exists.
- Any local/plugin-backed cmd could rely on initial `CreateProcessW` cwd without
  the verified embedded `.cmd`, six exact `/E:ON /V:OFF /S /K call` tokens,
  eight fixed canonical/cleaned environment keys, protected ready-before-create
  inbound pipe, exact root PID, or fixed-order `RSTCMD1` Ack through clean EOF;
  use HMAC/VT/cmd-side Unicode-cwd serialization or a transport fallback; restore
  an inherited mixed-case reserved key; accept a prefix/kept-open/second/trailing
  result; hide the trusted AutoRun/same-root limitation; accept an occupied,
  local/SUBST, over-cap, malformed, unterminated, DFS/provider/share-aliased, or
  complete/root/tail-inconsistent UNC reverse map; cancel a mapping while its
  root lives; force/cancel a changed/reused/shadowed mapping after root death; or
  assume process exit performs the unissued `popd`; or let any cmd reach
  Running/input/`CmdLaunchRootedRunning` memory before complete PID/capability/
  nonce/operation/generation/mode/target/status/drive/EOF proof, applicable UNC
  reverse mapping, and cleanup. AutoRun exit/hang/suppression/missing/late/wrong
  result must never be success.
- Cmd would require ongoing semantic/history/follow/insertion/activity/cwd integration; PowerShell/pwsh could expose a nonce through the environment/nested child, accept a second/live handshake after permanent revocation, or rekey without explicit Restart/new session; or nested integration would be advertised.
- A PowerShell semantic marker could use synchronous Write, more than one
  `WriteAsync` call/message, a Task callback/continuation, helper/background
  thread/runspace, alternate pipe/transport, unbounded wait/allocation, or an
  uncharged buffer; arm Read before Connect/PID validation; reconnect within an
  epoch; parse a chunk prefix or treat EOF as delimiter; omit the strict first-
  chunk receive clock; reject ordinary frames 2..N or accept a second Handshake;
  overlap an owned write; use anything but the bounded `IAsyncResult` wait on
  PowerShell 5.1; advance sequence before strict-before 50-ms completion; fail to
  cancel/close/latch inert and later observe/scrub Task/token/buffer state; make
  the valid staged-Handshake/Unavailable boundary race fatal or capability-
  bearing; or claim a real tuple without the loaded-runtime and 2-MiB real-
  message proof. STOP as well if semantic parsing adds a Terminal-local UTF
  converter instead of `Common/StringConversion.h`, or embedded PowerShell uses
  any encoding except `System.Text.UTF8Encoding(false, true)`.
- Control/revocation could treat write-only `TerminalSemantic` as the child read
  sentinel; fail to close/cancel both pipes at the specified fence; release any
  I/O/parser/dispatch/module quiet gate before actual completion; call
`TerminalControl` without either required `ReadAttributes|WriteAttributes`,
  grant `WriteData`, omit the exact noninheriting async/anonymous constructor or production-DACL message-
  mode proof; call
  `FlushFileBuffers` or a synchronous pipe Write/Flush; build mutable payload
  after `operationT0`; use another Write, a retry/timer/second wake, let
  `PollControlTransport(NoWait)` wait or dispatch, let it overwrite an occupied
  stage, let `HandlerWaitOnce` inspect/wait on a newly armed read before taking
  the staged frame, let key-handler-only
  `PollControlTransport(HandlerWaitOnce)` use anything but the sole nonalertable
  250-ms AsyncWaitHandle wait after owner/binding/epoch/lifecycle revalidation, or use an
  unbounded/UI/alertable/callback/Task/poll/spin wait;
  accept a short/equality/late OVERLAPPED completion; submit a control-pipe Write
  without first reserving the fixed-reserve descriptor/byte/final queue sequence;
  discover wake capacity only after submission; assign a second sequence, bypass
  later user descriptors, fail to consume a pre-submit placeholder tombstone, or
  release a submitted-failure inert fence without Discard/confirmed Close; issue wake without epoch-
  live validation; close/release the in-flight handle/event/OVERLAPPED/frame/pin
  before exact cancel plus quiet-bounded event drain and terminal completion;
  close the epoch-persistent pipe or release operation/input hold merely from a
  successful write; let that drain authorize an effect; let a BeginRead still
  unsignaled after the sole handler wait act/emit/revoke child-side; omit exact
  EndRead/rearm/final strict-deadline checks; fail to leave a no-operation physical
  press live after its at-most-250-ms root-shell wait; let a replacement binding
  receive more than one raw wake; replay a completed frame; admit another control
  operation after Probe non-Success, Expired, BindingChanged,
  ShellOperationFailed, or another failed/revoked/ambiguous outcome; or fail to
  return Idle after matching Success or Follow/Insert pre-effect
  RejectedState/BufferChanged; or
  release ambiguous queued input merely because 250 ms elapsed/disable occurred.
  A five-second drain miss must quarantine/fail fatally without freeing or
  unloading reachable state.
- A local/UNC PowerShell launch could skip the mandatory bootstrap; use any argv
  but the four-token fail-closed wrapper; admit an unsafe literal/target/semantic
  secret; bypass policy or reach prompt after exception/early/extra result;
  omit fixed merged-stream capture, exact typed-finalize validation, or the sole
  output-silent one-shot complete-frame CommitAck Write; allow Flush, retry, or
  any later pipe operation; gate from client Write completion; omit the full
  server read before gates; allow profile interception of startup Set/Get-Location;
  loosen the protected root-PID-bound pipe/HMAC/nine-kind/post-profile-rooting
  contract; fail to pre-reserve/admit both join slots without blocking; grant
  from one `StartupSemanticJoin` slot; close/fail either pre-CAS cleanup server or
  make either valid `RevokedTerminalOnly` Prepared fatal/capability-bearing; treat ReleaseAck/EOF as
  prompt commit; or publish prompt before the Running CAS or input/profile
  memory/capability before accepted CommitAck.
- The Running CAS cannot be the sole lifecycle-locked winner against close/
  restart/exit/timeout/shutdown; CommitPrompt disposition/generation does not
  accept exactly all four valid rows—original TerminalOnly plus Prepared
  TerminalOnly -> TerminalOnly at equal/later generation; SemanticEnable plus
  SemanticReady plus completed join -> SemanticActive only at unchanged
  generation; SemanticEnable plus SemanticUnavailable -> TerminalOnly only at
  unchanged generation; either valid SemanticEnable outcome after latched
  disable -> TerminalOnly at current strictly greater generation regardless of
  later re-enable—or accepts generation decrease/wrap, Active without Enable+
  Ready+join, TerminalOnly outside those rows, or any other pair; the script
  retains more than Decision generation plus Prepared outcome or cannot force
  inert/mapping cleanup; a post-CAS disable cannot close/cancel both pipes and
  make completed control EOF/error or semantic-write failure latch the
  interposer inert; matching cleanup Ack is keyed
  to live policy and rejected, ordinary input/integration-independent memory is
  withheld, or disable alone retires; an unrelated post-CAS failure cannot
  retire the already-Running incarnation with startup gates closed; or tests/
  docs would claim physical erasure of immutable OS/.NET strings.
- The PowerShell observer would require prompt wrapping, key/cell/native-history/control-wake inference, a custom/shadowed/nonallowlisted tuple or overload, anything before first-statement `$lastRunStatus = $?`, altered ReadLine status/return/native state, or fake acceptance/completion; parse failure cannot complete at only the next root read; replacement/failure cannot remain conservatively busy; or live disable cannot retain an inert same-recipe wrapper without fake `FunctionInfo` restoration.
- `Test-TerminalPowerShellAdapters.ps1 -AuditInstalled` mutates state, exact adapter/resource/license updates are not independently reviewable, or real Windows PowerShell 5.1 and every shipped pwsh tuple cannot pass without a fake-only result or environmental skip.
- Interactive confirmation, natural-exit/quiet policy, chooser/plugin removal, and forced close could not be linearized by one plugin per-view retirement-intent arbiter and one shared host `{instanceId,hostViewGeneration}` removal claim, allowing a late prompt/callback to cause a second host mutation.
- More than one interaction broker/visible prompt can exist per committed view; its exact four states, full token identity, or Close-over-Paste-over-OSC table cannot be enforced; a lower-priority request can resurrect; History/context/plugin-text-input can bypass `None`; or broker/domain ownership leaks into the host.
- A background close can issue UserTab before stable-ID select/reveal/layout/focus/revalidation, show an invisible confirmation, or roll back selection after it succeeds; a background/inactive OSC ask can count/retain/queue a token or steal focus; selection/deactivation can preserve/resurrect a prompt; or exact 60-/30-second equality, Cancel-default/restore/Escape/Enter, OSC `Alt+A`/`Alt+D`/UIA, lifecycle scrub, and posted-payload drain cannot be proved.
- A v1 WSL path would add or require `--exec`, `--user`, `-e`, a wrapper, another process, a bootstrap or `WSLENV`/environment/PTY/profile/staged-script injection; alter the configured default user/shell; create a semantic nonce/epoch; transition out of permanent `RequestedUnverified`; expose history/follow/insertion/activity/authenticated-cwd/profile-memory capability; package a dormant WSL integration payload; or make terminal usability depend on integration. STOP and require the separate approved WSL-integration design rather than weakening the five-token contract.
- Disable would retire the loaded generation or close a live terminal; refresh could be canceled after admission, abandon configuration settlement, time out/force-close the terminal-object barrier, use normal-runtime fatal/process-retention policy, silently resume after a timeout, reopen the old generation, map a replacement before resetting the old owning module handle, or permit overlapping module/service generations.
- `maxTabsPerPane` would be host-enforced, visible UI would commit before Open success, quota would use any result other than exact `ERROR_QUOTA_EXCEEDED`, live lowering would evict, or any failure/close/quarantine race could leak/double-release a per-host view slot.
- Runtime staging diverges from `RuntimeDependencies.props`, portable packaging cannot consume it, or any package stages a loose PowerShell startup/adapter/broker manifest/schema/script/wrapper/resource or startup/observer/broker/domain implementation outside `Plugins\Terminal.dll`.
- `Terminal.dll` is missing from a package, a dynamic engine closure is app-root/outside `Plugins\TerminalRuntime`, the EXE contains terminal implementation, plugin mutations do not degrade only Terminal, or module quiet cannot be proved.
- Signed MSIX production mutates the unsigned artifact, uses an unpinned/PATH/latest-SDK tool, passes PFX/password material to a child command line, cannot prove the exact locked certificate/Publisher/EKU/RFC-3161 timestamp, signs in place or by wildcard, leaves an owned certificate/staging artifact after failure, uploads signed/unsigned under one ambiguous artifact identity, or permits an unsigned MSIX to be attached/labeled as installable.
- Preferences can expose or invoke Reset outside the exact family status, omit the destructive confirmation, leak recovery/path/content details, or report a clear/forget/disable/reset success before every current-flavor active/recovery target is durably verified and stale writers are fenced.
- Any hard resource/security/privacy/perf gate is disabled to make final tests pass.
- The approved repository/source-input identity, toolchain lock, exact evidence companion/layout, containment quiet point, or disposable-VM cleanup disposition cannot be proved.
- ARM64 artifact/build, Release budget quality, notices/SBOM, package smoke, authoritative spec, or full regression evidence is missing.
