# Terminal Core, ConPTY Runtime, And Model

> **Status — Completed 2026-08-03.** Implementation and focused verification are complete. The durable contract is
> `../../Terminal/Terminal_EmbeddedPlugin.md`; this archived plan is non-normative, and later WIP/checkpoint language
> below records the route taken rather than unfinished v1 work. The test-enabled Full run completed 1,174 passes,
> 5 isolation-sensitive failures, and 53 skips; all five residuals passed unchanged-binary isolation reruns.

## First-implementation checkpoint (2026-08-01)

The selection blocker is resolved: the exact qualified Ghostty pin and private
runtime identities are recorded in the execution index and authoritative
terminal spec. The smallest product core is implemented in `Terminal.dll`:
unsuffixed `ITerminal`, private runtime loading, ConPTY, a tracked reader,
Ghostty formatting, DirectWrite/Direct2D presentation, keyboard/resize, and
quiet-point close. The larger service, immutable snapshot, Kitty, stress,
native matrix, and release resource-budget work below remains WIP.

> **Authority:** Execution plan 2 of 6 for `Terminal_EmbeddedConPtyLibGhosttyVt_IntegrationPlan_2026-07-13.md`. Start only after the reviewed Gate 0 decision has rewritten all generic names in this file to the exact selected engine.

## Status and predecessor gate

- **State:** WIP / first core slice implemented; advanced core/release gates open
- **Baseline:** RedSalamander `fbf6af76f373bf253488862426f4ea70a9ff25ba`
- Required evidence: the master's explicit ordered Gate-0 Done chain and sole
  `successfulGate0Handoff`; the authenticated winner-bearing Done blob and every
  prior no-winner lineage entry; approved
  `Specs/Terminal/TerminalEngineDecision.md`; passing x64/ARM64 offline
  verification; exact ABI/snapshot/resource contract; final capability ledger;
  recorded dependency-status/upgrade/lifecycle-test script digests; passing
  dependency-lifecycle Pester; and mutation-free `offlineLockedStatus=pass`.

If any predecessor artifact is missing, stale, or generic, STOP.

Before implementation, compare HEAD with the baseline across the master/child plans, the approved engine decision/finalization evidence/lock, `Common/PlugInterfaces/{Viewer,Factory,Host,Informations}.h`, `ViewerPluginManager.*`, plugin quiet-point contracts, `RuntimeDependencies.props`, `Directory.Build.targets`, `Tools/RuntimeDependencies.psm1`, `Common/MinimumOsVersion.*`, ConPTY/runtime code, terminal specs, build/test scripts, and installer consumers. Record the review even when no relevant files changed. Any material drift in engine identity, plugin ABI/module ownership, snapshot behavior, resource or lifecycle policy, centralized staging, test evidence, or owned durable specs requires reconciliation in the master and this plan, renewed predecessor verification/review where affected, and a new baseline before work continues; until then, STOP.

## Deliverables

- `Common/PlugInterfaces/Terminal.h` with the master's immutable `ITerminal`/`ITerminalCallback` IID and size-based records. Terminal is not a viewer/Preview association. The coordinated source-tree viewer record-header normalization is independent of this boundary.
- `Plugins/Terminal/Terminal.vcxproj` built as `Plugins\Terminal.dll` in required x64 Debug/Release/`ASan Debug` and ARM64 Debug/Release solution mappings. Add ARM64 `ASan Debug` only when the selected MSVC/toolchain proves support; otherwise preserve exact unsupported evidence and never fake/map/relabel Debug as ASan. The EXE has no static dependency on the terminal implementation.
- Selected-engine adapter and move-only RAII handles.
- One lazy plugin-module-owned `TerminalService`, strong active/retiring registries, plugin-owned per-complete-`PhysicalHostKey` view-reservation ledger for `maxTabsPerPane`, separate fixed 32-session admission ledger, two-worker snapshot executor, 512 MiB process text/snapshot plus Kitty resource ledgers, retirement/reaper, standard plugin callback drain, and module quiet point.
- Generic `TerminalPluginManager::BeginRefresh` and `BeginProcessShutdown` orchestration for both public-object classes (`ITerminal` plus the module-scoped `IPluginConfigurationOperations`/coordinator generation), plus nonblocking `PluginModuleLifecycle` begin/sample support for the existing standard shutdown/`CanUnloadNow` exports. Normal refresh and process shutdown have distinct timeout/outcome policies; the EXE receives neither a terminal-service accessor nor a terminal-implementation symbol.
- `ConPtySession`, bounded ordered input queue, terminal model/effects, immutable snapshot producer, perf seams.
- `Tests/TerminalTests` registered in `Tools/TestRunPlan.ps1` for CI/Full; `Tools/Run-AllTests.ps1` remains the product selftest runner. Add the shared, content-free `TerminalNative` evidence schema/runner/verifier defined below for direct native test executables used by plans 2–5.
- Add the master's exact content-free `Tools/TerminalCommandsEvidence.psm1`, `Tools/Run-TerminalCommandsEvidence.ps1`, and `Tools/Tests/TerminalCommandsEvidence.Tests.ps1`, plus narrow `Run-AllTests.ps1`/`SelfTestCommon` receipt-and-generic-archive-suppression seams and focused `Tools/Tests/TestRunArchive.Tests.ps1` changes. Freeze all seven scenario mappings, transactional exact file set, content allowlist, explicit `-RunPath` validation support, and `CommandsEvidenceSetSha256` algorithm before creating the Core baseline; the current generic `env.txt` repository archive is prohibited terminal evidence.
- Central manifest/task extension with `SourceRoot` and package-relative `PackagePath`; stage exact `Plugins\Terminal.dll`, only an applicable dynamic runtime/closure under `Plugins\TerminalRuntime\`, and Gate-0 notices in this plan. The state schema is created/staged by plan 5/6 work after it exists.

Before adding terminal entries, update `RuntimeDependencies.props`, `Directory.Build.targets`, `Tools/RuntimeDependencies.psm1`, and package consumers to resolve `SourceRoot=VcpkgTriplet|RepoRoot|TerminalEngineOutput` plus app-relative `PackagePath` identically. Legacy omissions default exactly to `VcpkgTriplet` and `Plugins\<OutputName>`; every new entry is explicit. `VcpkgTriplet` is canonical `$(VcpkgInstalledDir)$(RSVcpkgTriplet)\`—first-party `<RepoRoot>\.build\vcpkg_installed\x64-windows\` or `arm64-windows\`; canonical build/package entrypoints pass those already-resolved values to non-MSBuild consumers, which never consult ambient environment. `TerminalEngineOutput` resolves `.build\TerminalEngine\<Platform>\<Configuration>\`, including only lock-declared reuse, and destinations resolve `.build\<Platform>\<Configuration>\<PackagePath>`. Every path-bearing field on dependency/removal rows—`Source` where present, `OutputName`, and `PackagePath`—is literal-only with no MSBuild/item/environment tokens or wildcards; reject empty/rooted/dot/dot-dot/ADS/escaping forms, and require the single-filename `OutputName` to equal the `PackagePath` leaf. Split current `Flavor=Any` AWS rows into literal Debug `debug\bin\...`/Release `bin\...` rows. Merge the current Debug `zlib-debug`/`image-zlib-debug` and Release `zlib-release`/`image-zlib-release` pairs into one row per flavor with unioned Projects before enabling strict case-insensitive destination uniqueness; every later duplicate destination fails. Stage `Terminal.dll` exactly at `Plugins\Terminal.dll`; stage a dynamic selected-engine primary/closure only at lock-declared `Plugins\TerminalRuntime\<leaf>` paths; keep notices app-root. Add focused `Tools\Tests\BuildReproducibility.Tests.ps1` coverage for defaults/root resolution, every applicable path-field token/wildcard, token-free AWS migration, real zlib union rows, duplicate rejection, invalid/escaping paths, stale removal, missing-required cases, Terminal DLL presence, private runtime closure, and no app-root duplicate across x64 Debug/`ASan Debug`/Release and ARM64 Debug/Release, plus ARM64 ASan only when supported. A source/static winner is linked into `Terminal.dll` and has no runtime entry.

## Ownership and ABI sequence

1. The loaded `Terminal.dll` factory/module uniquely owns/dependency-injects one `TerminalService`; `ITerminal` objects get a non-owning reference plus session ID and never extend service lifetime. `ApplicationContext` owns only the generic plugin manager. No EXE terminal service, separate singleton, public global accessor, or direct terminal-implementation symbol exists.
2. The plugin service owns runtime/global callbacks, registries, executor, ledgers. Registry references keep sessions alive; sessions/workers/engine handles keep runtime and `Terminal.dll` pinned. Sessions never own the host, terminal object, or service. `RedSalamanderPluginCanUnloadNow` remains false until the last object/callback/worker/session/engine/runtime/child-HWND ownership reaches quiet; `RedSalamanderPluginShutdown` is idempotent and closes admission first.
3. Reserve one of 32 process-lifetime session slots before any engine/ConPTY handle and retain it through Starting/Running/Failed/Draining/Closing/retiring/quarantine until complete quiet. Slot exhaustion creates no engine/process state. Any quarantine closes launch/restart admission until every quarantine recovers; provisional/diagnostic and quiet retained exited views use no slot.
4. For a dynamic engine, lock/open the primary and complete non-system closure only below canonical `Plugins\TerminalRuntime\` paths while denying write/delete; validate identity/digest/PE/imports and retain handles. Reject preloaded same-basename identity mismatch; load primary with DLL-directory/System32 flags; enumerate actual non-system modules and match the locked set; only then call the asserted identity bootstrap, compare all fingerprints, and resolve callbacks/handles. Source/static selection links into `Terminal.dll` and provides equivalent generated assertions.
5. Every mismatch returns a localized terminal-unavailable result and creates zero callback or engine handle.
6. Callback userdata remains registry-owned; callbacks are nonblocking/bounded/non-reentrant and produce no UI, file, D2D, clipboard, pipe I/O, or unbounded allocation.
7. Runtime init is single-flight; creation/retirement/`BeginShutdown` serialize. The generic manager serializes module load, disable/re-enable, irreversible refresh, and process shutdown without allowing two mapped Terminal module/service generations. Plugin shutdown closes service admission first; racing startup retires immediately. Session callbacks detach at session quiet, while process-global callbacks remain until all active/retiring/quarantined/handle-owning sessions are gone and detach once.
8. Root-exit observation, explicit view close, and `BeginShutdown` serialize. `Closing` first suppresses later `Exited`; root-drain first retains its earlier `t0`/final-output path. A later explicit close only removes its view and switches any quiet-failure reporting to the content-free non-tab diagnostic; app shutdown suppresses `closeOnExit` and uses the earlier existing/session-wide fatal deadline. No race extends a deadline.
9. `ITerminal::SetCallback(nullptr,nullptr)` first closes callback admission and synchronously drains in-flight callback invocations; it is never called from inside a callback. `ITerminal::Close` is idempotent, nonblocking on the UI thread, destroys/gates the plugin child view, and transfers any session to retirement. Factory/module unload waits only at the plugin-manager quiet path, never in a tab destructor or `DllMain`.

### Plugin-owned `maxTabsPerPane` view admission

- `Terminal.dll` is the sole owner and reader of the current effective `maxTabsPerPane` value, closed to `1..32`. `ITerminal::Open` performs every synchronous context/configuration/plugin-admission validation first, then atomically reserves one view slot in the module-service counter keyed by the complete `PhysicalHostKey`, before creating the child HWND or starting asynchronous/profile/process work. A count at the limit returns exactly `HRESULT_FROM_WIN32(ERROR_QUOTA_EXCEEDED)`; no synchronous validation/quota failure retains a reservation, creates a child/process, or publishes a callback, and Open never delivers a callback synchronously before returning.
- The host may create only a hidden uncommitted record and register the callback before Open. That record is not a tab, changes no selection/zoom/MRU/chrome, and never counts. On quota or any other synchronous Open failure, the host synchronously nulls/drains the callback, calls idempotent `Close`, releases the object/record, and preserves the prior UI; quota alone shows localized `Terminal tab limit reached for this pane` with `Open Terminal settings`. On Open success, a later host-commit/child-attachment failure performs the same null-drain/Close/public-release rollback, releasing the reservation and leaving no tab/zoom/child/process leak.
- A successful Open retains its view slot through chooser, diagnostic, Starting, Running, Failed, Exited, natural-exit teardown timeout, and host view-removal admission. Release occurs exactly once only after callback null/drain, a valid `Close` has destroyed/detached the child view, and the host releases the public object. Session retirement or quarantine that continues after public-view release owns no transferred view slot. Folder, Preview, hidden uncommitted records, and content-free service diagnostics after view removal do not count; view slots and the separate 32 process-session slots are never conflated.
- Live lowering atomically publishes the new limit, evicts no existing object/view, and rejects later per-key reservations until that key's count is below the limit. Raising resurrects nothing. Erase a counter key at zero, and serialize simultaneous plus destructively reentrant Open attempts through the same reservation boundary. The host never reads or pre-enforces the setting.

### Size-based Open, path-insertion, and view-removal ABI

- `TerminalOpenContext` begins with `sizeBytes`; source and launch locations are independently named and retained. Consumers reject a value shorter than the required current prefix and accept unknown tail bytes. Runtime boundary tests cover undersized/current/oversized records without freezing a concrete size or offset table. The semantic field order remains parent window, instance ID, original source key, physical host key, source generation, source location, launch location, selection mode, open purpose, requested profile ID, and reserved fields.
- `TerminalOpenPurpose` is closed as `InteractiveTab=1` and `PendingPathInsertion=2`. `InteractiveTab` requires the second `Open` argument to be null and allows `ResolveDefault|ForceChooser|ExactProfile` subject to the ordinary selection-mode/profile rules. `PendingPathInsertion` requires a nonnull fully valid copied insertion, `ResolveDefault`, Windows-compatible source/launch/initiating locations, and the hidden two-phase commit-or-abort route; it rejects every WSL/unsupported/cross-family location, `ForceChooser`, `ExactProfile`, and null/non-null mismatch. Unknown purpose values or any invalid purpose/insertion/selection/location combination return `E_INVALIDARG` synchronously before child creation, process work, callback delivery, status mutation, or either view/pending-open slot reservation.
- ABI validation uses the master's exact untruncated UTF-16 component ceilings: Windows path 32,767, Linux path 32,768, WSL distribution 256, plugin short ID 128, requested/stable profile ID 32,772, display leaf 32,768, title 1,024, and status 2,048 code units. NUL in any span is invalid. The 32,772 profile ceiling deliberately admits the complete `pwsh:` plus a maximum 32,767-unit canonical Windows path; every accepted span is copied byte-for-byte before return, and no ABI, catalog, callback, retry, or restart layer may truncate, alias, or hash it into another identity.
- `sourceLocation` is the owning pane's exact committed follow baseline at the real nonzero `sourceGeneration`; `launchLocation` alone selects the initial cwd, the applicable Windows profile-memory key, WSL preflight, and immutable original-launch Restart fallback. WSL uses the launch location for deterministic exact-distro selection and preflight but never reads or writes folder-profile memory in v1. Both locations must be independently valid and profile-family compatible. Ordinary opens pass value-identical locations and freeze internal `launchTracksSourceUntilStart=true`, so a fully revalidated newer source update before process creation replaces both pending locations. Directory Edit is the sole v1 unequal route: source remains the real pane/current-generation baseline, launch is the revalidated focused child, and it freezes `launchTracksSourceUntilStart=false`; source navigation and Retry advance only the follow record and never replace that child launch. Process creation freezes launch. An explicit follow off/on may internally requeue the currently stored same-generation source baseline without manufacturing or accepting an ABI update and without lowering a newer generation.
- `TerminalPathInsertion` begins with `sizeBytes`; consumers validate the current required prefix and ignore unknown tail bytes. Its semantic order remains initiating source, initiating generation, initiating location, item generation, mode, reserved0, item location, parent location, display leaf, and reserved fields. Concrete compiler sizes and offsets are deliberately not part of the contract.
- `initiatingSource` is the command-originating live Folder source, not an alias for the target terminal's immutable `originalSource`. Immediately before `InsertPath`, the host revalidates the atomic `{initiatingSource, initiatingSourceGeneration, initiatingSourceLocation}` and, for item modes, the focused item plus `itemGeneration` against that Folder source; a pending insertion repeats both revalidations immediately before eventual delivery. The plugin requires a nonzero initiating key/generation, copies the entire request before return, treats generations only as host-owned correlation, has no pane/item catalog, and never compares the initiating source triple with the target object's immutable original source or follow record. For an insertion-capable Windows target it independently translates request locations into that target profile; an unsupported or cross-family initiating/item location returns `HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED)` and emits zero bytes. An exact-distribution WSL target accepts only the quoted full native Linux path, degrades contextual mode to full, and exposes no trusted prompt capability; a different/unknown distribution or cross-family location rejects before input admission with zero bytes.
- Every insertion mode requires a supported `initiatingSourceLocation`. `ContextualLeafOrFull` additionally requires nonzero `itemGeneration`, supported item and parent locations, and nonempty `displayLeaf`; that leaf is untrusted presentation-only data and never supplies identity, translation, quoting, or PTY bytes, so the plugin derives the native leaf from translated `itemLocation` and validates its translated parent. `AlwaysFull` requires nonzero `itemGeneration` and a supported item location, with Unsupported/empty parent and null/zero display leaf. `CurrentDirectoryFull` requires zero `itemGeneration`, Unsupported/empty item and parent, and null/zero display leaf; it translates and inserts the copied `initiatingSourceLocation`, never the target's followed source or authenticated cwd. Every other cross-field combination returns `E_INVALIDARG` and emits zero bytes.
- `TerminalViewRemovalReason` is closed as `ChooserCancelled=1`, `AutomaticCleanExit=2`, `AutomaticAnyExit=3`, and `PluginCloseAction=4`. `ITerminalCallback` v1 order is exactly `TerminalStateChanged`, `TerminalPendingOpenResolved`, `TerminalOpenSiblingRequested`, `TerminalCloseApproved`, `TerminalViewRemovalRequested`, `TerminalQuiet`; the second slot is `TerminalPendingOpenResolved(const TerminalInstanceId*, uint64_t pendingOpenGeneration, TerminalPendingOpenResult, void*) noexcept`, and the removal slot is `TerminalViewRemovalRequested(const TerminalInstanceId*, uint64_t viewRemovalGeneration, TerminalViewRemovalReason, void*) noexcept`. `TerminalPendingOpenResult` is closed as `ReadyToCommit=1`, `ProfileChoiceRequired=2`, `ProfileNotInsertionCapable=3`, `IntegrationDisabled=4`, `LaunchFailed=5`, `PromptUnavailable=6`, and `TimedOut=7`. Pending-open resolution is noncoalescing and carries exactly one result for the exact generation; every non-Ready result makes Commit permanently illegal, while Ready still requires the host's exact revalidation plus Commit. Chooser cancellation is legal after the required view-slot reservation but only before any process-lifetime session slot, session, or process exists; deferred removal releases the view slot. Automatic-clean requires exit code zero plus complete-final-snapshot natural-exit quiet and `closeOnExit=clean`; automatic-any requires complete-final-snapshot natural-exit quiet and `closeOnExit=always`; plugin Close is legal only in provisional diagnostic/Failed/Exited after the applicable no-session/retirement admission and never substitutes for `RequestClose(UserTab)` in Starting/Running. Incomplete final snapshot, quarantine, and `closeOnExit=never` do not independently qualify an automatic request.
- Every view owns one serialized retirement-intent arbiter and one nonzero monotonic `retirementIntentGeneration`; its only states are `Open`, `PendingUserConfirm`, `UserApproved`, `ForcedWindow`, `ForcedApplication`, `ChooserCancelled`, `AutomaticCleanExit`, `AutomaticAnyExit`, and `PluginCloseAction`. `UserTab` allocates a generation; when confirmation is required its one-shot prompt token is exactly `{instanceId,viewGeneration,retirementIntentGeneration}`. Allow may move only that exact pending token to `UserApproved` and emit one generation-matching `TerminalCloseApproved`; cancel may move only it back to `Open`, emitting nothing and changing no view/session state. Duplicate close while pending is `HRESULT_FROM_WIN32(ERROR_BUSY)`; stale, duplicate, wrong-view, and wrong-generation replies are no-ops.
- Forced window/application, chooser/plugin, and qualifying natural-exit transitions compete through that same CAS/lock. A takeover from `Open|PendingUserConfirm` allocates a newer generation first, dismisses/invalidates any prompt, and emits only its one owning forced path or at most one `TerminalViewRemovalRequested`; no parallel interactive/automatic booleans exist. If `UserApproved` won, every later path joins retirement and emits no second callback. If another path won, late allow/cancel and CloseApproved are suppressed. Cancel relinquishes only its user intent, so a later independent forced/automatic claim may still acquire `Open`.
- Natural exit while confirmation is pending follows the exact master matrix: before complete quiet, retain the pending token. At quiet, a complete snapshot qualifying under `always`, or under `clean` with exit code zero, takes over with only the applicable removal request. A nonqualifying quiet success (`never`, nonzero `clean`, or incomplete final snapshot) converts the still-exact pending token to `UserApproved` and emits its one CloseApproved because process confirmation is then meaningless; cancel-first retains the exited view. Quiet failure/quarantine leaves the prompt unresolved. Approval-first suppresses later automatic removal, while forced window/application close always owns or joins teardown.
- The plugin-to-host removal callback carries no teardown ownership. The generic host has exactly one atomic removal claim keyed by `{instanceId,hostViewGeneration}`; interactive CloseApproved, removal callbacks, and forced window/application teardown all enter it. The winner alone may hide/remove the record, apply MRU/else-Folder fallback, cancel pending insertion, null/drain the callback, call idempotent `Close`, and release the object; losers join/no-op. A callback only posts a stable-ID/generation task, and the deferred winner revalidates after callback return. It never calls `RequestClose`, parses Terminal settings, or infers removal from lifecycle state. Duplicate/stale/unknown/wrong-instance requests, queued callbacks losing to forced teardown, and host-generation reuse cannot mutate twice. Chooser/plugin-Close removal follows its last retained state event; automatic removal order is exactly `TerminalStateChanged(Exited) -> TerminalQuiet -> TerminalViewRemovalRequested`.

### Normal plugin refresh and disable/re-enable lifecycle

- Disabling `builtin/terminal` is only an availability change. It closes new Open/Edit/path-insertion/Restart admission but never requests close of an existing terminal, retires the module-scoped administration object, invokes a shutdown export, or starts a second module/service generation. Existing terminals continue ordinary input, rendering, and close under their frozen effective configuration. Disable does not change the paste-policy generation or revoke safe/pending unsafe paste on an existing live view; those transactions remain usable under that view's frozen configuration. Separately, disable is an OSC 52 sensitive-effect transition: under the OSC 52 authorization gate it increments that policy generation, revokes/scrubs outstanding asks or prompt-free allows, and makes every later OSC 52 request effectively `deny` until re-enable. While retirement is `Idle`, re-enable reuses that retained generation, reconciles the latest persisted configuration, increments OSC 52 policy generation under the same gate, then increments admission and reaches enabled state only after reconciliation succeeds; configured OSC 52 policy applies only to later requests and no revoked request is replayed.
- Explicit Manage Plugins Refresh is the only normal-runtime retirement route. `TerminalPluginManager::BeginRefresh(uint64_t refreshGeneration)` is implemented in the generic Terminal manager over shared `PluginModuleLifecycle` and contains no Terminal service/runtime knowledge. The generation is nonzero, manager-issued, and application-lifetime unique: `S_OK` atomically changes `Idle` to `RefreshSettlingConfiguration`; the same active generation returns `S_FALSE`; zero, stale, or unissued values return `E_INVALIDARG`; a different concurrent generation returns `HRESULT_FROM_WIN32(ERROR_BUSY)`. Once `S_OK` returns, retirement of that old module generation is irreversible, cannot be canceled, and cannot queue another refresh.
- The accepted refresh captures the old module generation, increments and closes terminal/configuration/deep-link/status/action/page admission, and retains only host-copied schema/metadata plus content-free status. New host-mediated calls return `HRESULT_FROM_WIN32(ERROR_SHUTDOWN_IN_PROGRESS)`. It does not call `RequestClose`, hide, disconnect, callback-clear, close, or release any existing `ITerminal`; their ordinary terminal behavior remains usable until natural user/lifecycle removal. Normal refresh does not change the paste-policy generation or revoke/reject safe or pending unsafe paste for an existing live frozen-configuration view. `BeginRefresh` does, however, install the OSC 52 revocation fence under its dedicated gate before publishing refresh admission closure, canceling tokens and securely scrubbing every outstanding OSC 52 ask/allow payload. `RefreshWaitingForTerminalObjects` has no elapsed-time failure because refresh may never terminate a user's shell.
- Concurrently, the retained configuration coordinator enters `Retiring`, suppresses page presentation, and admits only private settlement for work accepted before the fence. Cancel and take every pre-`CommitStarted` result; before Settings CAS, abort the whole participant group and issue exactly one `Finalize(Failed)` for every successful Prepare/Arm; after CAS starts, await its real outcome and issue exactly `Finalize(Committed|Failed)`. Complete commit-won work, typed-result ownership, required compensation save, and matching `HostCompensationSave` acknowledgement before callback clear/drain, nonblocking `Close`, and release. Missing/idle administration completes without manufactured calls. No preparation, CAS outcome, result, compensation, or continuation is abandoned or guessed; either public-object barrier may complete first.
- A separate five-second configuration-settlement clock starts at accepted `BeginRefresh`, independent of terminal lifetime. Expiry is the nonfatal `RefreshBlocked(ConfigurationSettlementTimeout)` outcome: retain the old module/coordinator, keep all new admission closed, show host-owned localized `Restart RedSalamander to finish plugin refresh`, and allow mandatory private settlement to finish safely. Late completion never silently resumes/unloads; it only makes the generation eligible for explicit host Retry. Normal refresh never calls the application fatal seam.
- Only after configuration settlement and natural release of every old-generation public terminal object does the manager invoke the standard idempotent shutdown export exactly once and enter `RefreshWaitingForQuiet`. Generation-checked dispatcher polling of `RedSalamanderPluginCanUnloadNow` runs no more often than every 50 ms. The quiet clock begins at that export, not while tabs remain open. A true vote unregisters the resource owner and resets the old owning `wil::unique_hmodule`; no replacement DLL may be opened or mapped before reset completes. A false vote at five seconds becomes nonfatal `RefreshBlocked(ModuleQuietTimeout)`, stops automatic polling, retains the mapped old generation, and never queries/honors `RedSalamanderPluginRetainModuleUntilProcessExit`.
- Host-owned `plugin.refresh.retry` is exposed only in `RefreshBlocked`. For a settlement timeout it returns `HRESULT_FROM_WIN32(ERROR_BUSY)` without a plugin call until mandatory settlement and both public-object barriers complete; otherwise one invocation starts exactly one new bounded five-second quiet-sampling window. A still-false vote returns to the same blocked state. There is no background/unbounded retry and no concurrent unload/reload path.
- Enable/disable accepted during refresh persists and updates only one latest-wins `desiredAvailabilityAfterRefresh`; it never calls the retiring coordinator, reopens old admission, cancels retirement, or queues a generation. After successful old-handle reset, resolve/verify/map at most one replacement, create one administration object, and reconcile the latest complete Settings snapshot before exposing configuration or terminal admission. Desired disabled reaches `DisabledRetained`; desired enabled reaches `EnabledReady|EnabledDegraded` only after success. Replacement absence/corruption/architecture/ABI/startup failure leaves no old executable generation and reaches visible `EnabledConfigurationBlocked` with reinstall/Retry guidance.
- `BeginProcessShutdown` supersedes every refresh state, joins completed barriers/export work without repeating it, force-closes remaining public terminal objects through the stricter application-shutdown path, and starts fresh outer shutdown deadlines. Only that path may use the closed passive-UIA retention result or fatal policy; a normal blocked refresh cannot weaken it.
- Implement the host lifecycle in `RedSalamander/TerminalHost/TerminalPluginManager.h/.cpp` over `RedSalamander/PluginModuleLifecycle.h/.cpp`, with project/filter wiring and focused native/generic lifecycle tests. Do not duplicate module loading/export lookup in the manager or add a Terminal-private service accessor.

### Generic application-shutdown barrier

- `UserTab` is the only interactive close reason. Cancellation leaves the object, session, callback, and view unchanged; an accepted close emits exactly one generation-matching `TerminalCloseApproved` while the callback remains registered. `WindowClosing` is noninteractive and object-local: it admits forced retirement for only that FolderWindow's object, emits no `CloseApproved`, and never closes service-wide launch/restart admission.
- Main application shutdown calls `TerminalPluginManager::BeginProcessShutdown(validatedMainTarget, nonzeroApplicationLifetimeGeneration)` before destroying any FolderWindow, the main posted-message target, or its pump. The first valid generation captures manager `shutdownT0`; closes terminal/configuration creation, Preferences-deep-link, status/Retry, and page-presentation admission; increments the admission generation; and snapshots both the module-scoped `IPluginConfigurationOperations`/coordinator generation when present and every public `ITerminal` object in stable `TerminalInstanceId` order.
- The terminal-public-object barrier calls `RequestClose(ApplicationShutdown)` on every snapshot entry while its callback is live; the first accepted object starts the plugin's one service epoch and later calls join it. It then hides the view, calls `SetCallback(nullptr,nullptr)` to close/drain callback delivery, calls idempotent `Close`, releases the object, and invalidates the generic host record. No approval callback is awaited.
- Concurrently, the configuration-public-object barrier marks the retained administration/coordinator generation `Retiring` and suppresses page output while keeping its callback/dispatcher/private settlement admission alive. It cancels every pre-commit accepted operation and takes/releases each result. If no global Settings CAS started, abort the complete participant group and consume every successful Prepare/Arm through exactly one `Finalize(Failed)`; if CAS started, await its actual outcome and issue exactly `Finalize(Committed)` or `Finalize(Failed)` to each successful participant. Commit-won reconcile/action/finalize work, every typed result, and every required fail-safe/compensation Settings save plus matching `HostCompensationSave` acknowledgement settle before retirement; query results are validated, taken, and discarded. No preparation, transaction-bearing result, compensation, or continuation is abandoned or guessed.
- Only when the configuration generation owns no accepted operation, unconsumed preparation, host CAS, compensation, retained result, or coordinator continuation may the manager call `IPluginConfigurationOperations::SetCallback(nullptr,nullptr)`, call its nonblocking idempotent `Close`, and release it. `Close` must return `HRESULT_FROM_WIN32(ERROR_BUSY)` with no retirement mutation while transaction-bearing state remains; clearing the callback before such a result is retrievable is a host invariant breach and takes the fatal path. Missing/idle administration completes without manufacturing calls.
- Only after both public-object barriers complete does the manager use shared `PluginModuleLifecycle` to invoke the existing idempotent `RedSalamanderPluginShutdown` export. With zero terminal objects or only object-free retiring work, that export begins the same service epoch. The EXE performs this sequence solely through public `ITerminal`, public configuration-administration, and generic module-lifecycle surfaces; it never resolves or calls `TerminalService` or another Terminal-private symbol.
- Extend `PluginModuleLifecycle` with generic nonblocking begin/sample operations for the existing shutdown, `RedSalamanderPluginCanUnloadNow`, and process-shutdown-only retention exports; do not duplicate `GetProcAddress` or unload policy in the Terminal manager. Its Terminal path returns only `Busy|Unloaded|RetainedUntilProcessExit`, and a false unload vote is `Busy` unless the plugin explicitly votes for retention. Retain the module/resource owner and sample only on generation-checked UI-dispatcher turns at intervals of at most 50 ms while the main target and message pump remain alive. Exports never wait. `Unloaded` and `RetainedUntilProcessExit` each post exactly one generation-bound continuation; `Busy` reschedules without blocking UI or destroying a message target.
- The manager's outer two-/five-second deadline is measured from `shutdownT0` before either public-object barrier: at two seconds it ensures every still-cancelable terminal/coordinator owner received its one cancellation, and at five seconds any incomplete barrier, uncalled export, or operational false quiet sample invokes `FatalShutdownPolicy`. When terminals exist, the first accepted `ApplicationShutdown` request in the initial dispatcher turn owns the service/session `t0`; that watchdog shares concurrent deadlines and the earlier manager deadline is a no-later bound. With zero terminals, administration retirement is still bounded before the export starts the otherwise-empty service epoch. Polling cannot reset or multiply either clock. The watchdog remains independent of a stuck retirement worker.
- The existing `RedSalamanderPluginRetainModuleUntilProcessExit` helper is a closed exception, not a general mapped-busy escape. Terminal may vote true only after the shutdown export ran, both public-object barriers completed, every provider was disconnected and every call captured before disconnect returned, and callbacks, workers, engines/renderers, process/pipe/HPCON handles, posts, continuations, in-flight calls, and all host/service-touching owners are gone. The sole remaining gates may be externally held disconnected UIA provider/range COM objects owning only a self-contained immutable snapshot and passive reference-counted ledger state; every post-disconnect method except `IUnknown` returns `UIA_E_ELEMENTNOTAVAILABLE` with null/zero output and touches no host/service. The lifecycle helper rechecks `CanUnloadNow` once after a false retention vote to admit a racing final external `Release`, then either unloads normally or retains the module/resource owner through OS teardown and posts the same single continuation. Normal refresh/disable never queries/honors this vote, and any other active-work false result still reaches the five-second fatal policy.
- Duplicate same-generation shutdown begins, repeated `ApplicationShutdown`, duplicate `WM_CLOSE`, and repeated standard exports join or no-op idempotently. A stale/zero generation, stale target, superseded callback/view generation, or late poll has no effect and cannot post a second continuation. The initiating UI path returns without waiting, joining workers, calling `DestroyWindow`, or blocking the UI thread.

## ConPTY startup and I/O

- Use WIL/custom RAII for pipes, process/thread handles, attribute storage, environment block, and `HPCON`; create ConPTY with flags `0`. Own the one `GetEnvironmentStringsW` snapshot with `wil::unique_environstrings_ptr`; never reinterpret it as ANSI or hand-clean it.
- Build one explicit UTF-16 environment block per launch. Parse the parent snapshot through its required double-NUL terminator with checked arithmetic. For ordinary entries, split at the first `=`, require a nonempty name, and make the later parent spelling/value win under `CompareStringOrdinal(..., TRUE)`. Preserve hidden drive-current-directory entries in a separate map only when they are exactly `=<ASCII drive letter>:=<nonempty absolute same-drive DOS path>`; canonicalize the drive key to uppercase and make the last parent entry for that drive win. Reject every other leading-`=` entry, malformed key/path/parent terminator, embedded NUL in a supplied delta, or arithmetic overflow.
- After parent deduplication, overwrite `SystemRoot`, `windir`, `ProgramFiles`, applicable `ProgramW6432`/`ProgramFiles(x86)`, `ComSpec`, and every permitted launch/integration environment key with canonical spelling and copied API-derived value; no case-variant parent key survives and no unrelated parent entry changes. The PowerShell semantic nonce is explicitly not such a key and is never present in initial argv or any private process argument: plan 5 creates and delivers it only after root authentication in the one post-root `SemanticEnable` startup-pipe decision. Combine hidden and ordinary entries, sort emitted `name=value` strings by ordinal-ignore-case with ordinal case-sensitive comparison as the total-order tie break, append one UTF-16 NUL per entry plus one additional terminal NUL, and represent an empty block as exactly two NULs.
- The mutable command line is at most 32,767 UTF-16 code units including its terminating NUL. The complete environment block is at most 524,288 UTF-16 code units including both terminal NULs. Exact limits succeed; +1, malformed input, or an over-cap canonical overwrite returns a localized retryable launch failure before process creation, with no truncation or partial block. These caps are fixed, not settings.
- Start output service before/as child creation, then call `CreateProcessW` with the explicit validated executable, mutable command line, canonical explicit UTF-16 environment block, explicit current directory, and creation flags exactly `EXTENDED_STARTUPINFO_PRESENT | CREATE_UNICODE_ENVIRONMENT`; close duplicate host ends after success. No launch may inherit a null environment pointer or omit either flag.
- Every Windows-shell launch passes its held/revalidated canonical executable path as exact nonnull `CreateProcessW.lpApplicationName`, while the mutable command line starts with exactly one unquoted fixed basename token—`cmd.exe`, `powershell.exe`, or `pwsh.exe`—as child `argv[0]`; it never repeats the canonical path. Only that profile's reviewed arguments follow. The shared `CommandLineToArgvW`-inverse helper must round-trip the exact child token array, and hostile current directory, `PATH`, or an app-local same-basename file cannot redirect creation because `lpApplicationName`, not `argv[0]`, selects the image.
- A WSL launch resolves only the reopened non-reparse `<GetSystemDirectoryW()>\wsl.exe` and passes it as `CreateProcessW.lpApplicationName`. Its exact logical argv is `[wsl.exe, --distribution, <catalog exact distro name>, --cd, <absolute profile-native Linux path>]`; there is no `--exec`, `--user`, `-e`, bootstrap/extra command or process, PATH/ShellExecute/command-processor resolution, home/default-distro retry, or other fallback. The configured distribution default user and default shell run unchanged. `--cd` is a supported-platform invariant under the existing Windows-build-22000.2600 minimum, not a help/version probe. Reject NUL/CR/LF, empty/leading-dash distro, non-absolute Linux path, and command-line overflow before creation. V1 performs no pre- or post-launch WSL integration injection: no bootstrap is concatenated into argv or environment, written into the PTY, installed in a Linux profile, staged through a guessed `/mnt/<drive>` path, or carried through `WSLENV`; no wrapper or second process is allowed. Every WSL shell, including bash, is terminal-only and creates no semantic-integration nonce or epoch.
- WSL identity representability and process-command representability are separate. The ABI may carry the complete 32,768-unit Linux path, but the launch builder quotes all five tokens with the same shared helper and rejects a result whose mutable command line including NUL exceeds 32,767 UTF-16 units before preflight, ConPTY, or process creation. There is no shortening, hash, environment indirection, response file, wrapper, or second process. For quote-free one-unit distro `D`, constants consume 31 units including NUL, so a 32,736-unit absolute Linux path is the exact accepted boundary and +1 fails with zero preflight/process/pending-insertion bytes.
- Before WSL process creation, a cancelable off-UI generation-bound `WslLaunchPathPreflight` freezes catalog generation, exact distro, Linux path, and launch generation; opens the exact canonical `\\wsl.localhost\<catalog exact distro name>\<path>` directory for attributes with WIL ownership and no delete sharing; verifies directory/catalog generation; and retains that directory handle through `CreateProcessW` plus startup-channel admission. Its fixed monotonic three-second deadline irrevocably fails the generation before invoking `CancelSynchronousIo` on its owned worker; a late completion can only release its handle. Failure/cancel/timeout/stale/missing/path loss creates a diagnostic and starts no terminal. The service owns the `std::jthread` through return under normal quiet/fatal policy; it is never detached, UI-joined, cached as proof, or run on the UI thread.
- WSL has no shell-independent post-create cwd acknowledgement. It may publish `Running` after launcher/preflight/ConPTY/process admission only with internal `launchCwdTrust=RequestedUnverified`; `Running` does not claim an authenticated Linux cwd. That value remains `RequestedUnverified` for the entire v1 incarnation, there is no WSL trusted-cwd event, and every attempted transition to `Verified|Diverged` is rejected. App history, authenticated activity/cwd, pane follow, contextual-leaf insertion, and WSL folder-profile reads/writes are disabled. Explicit insertion is limited to the quoted full native Linux path in the exact launch distribution. An asynchronous launcher failure or root exit may follow a briefly visible `Running` state after an external distro/path race, but it never retries, falls back, or writes memory; tests must not promise that every asynchronous `--cd` failure is observable before `Running`.
- A later WSL-integration feature requires a separate approved design before code changes. Its first gate must prove clean-default Bash plus profile-override/`exec`/early-exit behavior, bootstrap before user-input admission, zero visible PTY/native-history input, no Linux-profile edit, configured default-user/default-shell/startup preservation, no `/mnt/c` assumption, bounded nonce/path-state erasure, and safe fallback to a fully usable terminal-only session. If it needs `--exec`, a wrapper, `WSLENV`, or another process, that design must explicitly replace the five-token invariant and freeze exact argv/environment/path/ownership/timeout/cleanup semantics. V1 carries no dormant or best-effort experiment.
- One reader owns blocking reads and serialized engine mutation. Never drop PTY output; coalesce presentation.
- One writer consumes a single monotonic sequence across user input, atomic paste, integration operations, and engine responses. A foreground-control operation reserves its inert wake placeholder in that sequence before its pipe Write; later user descriptors therefore receive later sequence values and can never force the wake to bypass FIFO.
- Resize coalesces to newest nonzero cells and runs off UI.

### Queue contract

- Preallocated nonblocking MPSC admission; no callback waits or takes an unbounded allocation path.
- Normal capacity defaults to and cannot exceed 8 MiB; `inputQueueMaxMiB` may lower its effective byte ceiling live to 1 MiB. User descriptors stop at 3840. The fixed protocol reserve is a separate 256 KiB/256 descriptors shared only by required engine responses and at most one foreground-control wake placeholder; total descriptor ceiling is 4096. User traffic never consumes that reserve. Responses may also use free normal descriptors, but the control placeholder may use only the fixed reserve so later normal-capacity pressure cannot invalidate it.
- Assign a queue sequence only after admission. Under the generation/input-operation gate and before any `TerminalControl` pipe Write, atomically reserve one protocol-reserve descriptor plus its one eventual wake byte and assign that inert placeholder its final queue sequence. If either fixed-reserve unit is unavailable, fail before pipe submission with zero shell effect and no control-sequence advance. Later user descriptors receive later queue sequences and the writer stops at the inert head placeholder; no producer, callback, or UI thread waits.
- Exact successful pipe completion plus epoch/deadline revalidation activates that same placeholder in place with the one wake byte; it performs no second capacity check, admission, or sequence assignment. A mechanically proved pre-submit/no-action loss converts it to a consumed zero-byte tombstone before later user input becomes writable. Any submitted timeout, failure, disable race, or unresolved result leaves it as an inert held fence and retains later input until the existing explicit Discard/confirmed-Close ambiguity recovery; a safe authenticated result releases the separate post-wake input hold. Activating, tombstoning, holding, or retiring a placeholder never permits priority reordering or reuse of its sequence.
- Native event and complete bracketed paste are atomic. Validate final encoded size against `pasteMaxBytes` and remaining normal capacity; never truncate/partially wrap.
- Every paste action reads `CF_UNICODETEXT` from the clipboard exactly once on the UI thread. Copy it into plugin-owned storage with a bounded terminating-NUL scan; reject malformed UTF-16, embedded NUL, size overflow, or strict UTF-8 conversion failure; then normalize CRLF and lone CR to LF exactly once. The resulting immutable normalized UTF-8 payload—not the mutable clipboard—is the sole source for classification, bounded preview, selected-engine encoding, and final admission. Unsafe classification is frozen from those bytes and the captured policy: logical line count at least `pasteWarningLineThreshold`, or any C0/C1 control other than HT/LF. Neither preview nor payload enters a host posted-message payload, log, diagnostic, metric, crash annotation, or evidence record.
- Freeze one paste identity `{serviceGeneration,sessionId,sessionIncarnation,viewGeneration,pastePolicyGeneration,inputModeGeneration,pasteSequence}` plus an opaque one-shot token. `pasteSequence` is per incarnation, monotonic `uint64_t`, and may never wrap. Charge pending normalized payload plus preview storage before materialization against fixed caps of 8 MiB per terminal and 32 MiB service-wide; securely zero complete owned capacity on every release. Exactly one confirmation may be pending per terminal. A newer paste atomically cancels, invalidates, and scrubs the older request before publishing a replacement; service-cap failure rejects the new paste with localized content-free status.
- Safe paste and unsafe paste captured with `warnOnUnsafePaste=false` use this same identity/final-admission transaction without prompting. An unsafe warned paste publishes only its opaque token and bounded content-free metadata to the plugin-owned confirmation surface. Cancel, supersession, input-mode generation change, paste-policy generation change, view/session close, incarnation restart/retirement, or application shutdown invalidates the token under the paste-authorization gate and scrubs payload/preview before publication or teardown completes. A more-permissive setting never approves/replays an older classification. Generic disable and normal refresh are deliberately not paste revocations: an existing live view retains safe paste and its already-pending unsafe prompt under the frozen effective configuration until its own close/restart/application-shutdown fence.
- Approval consumes the exact token at most once and never rereads the clipboard. First snapshot the current selected-engine paste mode and encode only the frozen normalized bytes. Then acquire the paste-authorization gate and immediately revalidate the complete identity, unchanged paste-policy/input-mode generations, matching live view/incarnation, and `Running`/not-closing input capability. Revalidate the current `pasteMaxBytes`, effective normal-byte ceiling, 3840-user-descriptor ceiling, and exact remaining queue capacity against the complete encoded event including bracket wrappers. Only one linearized operation may atomically admit the whole event and assign its queue sequence; every failure admits zero bytes. Consume the token on allow, deny, stale/duplicate reply, encoding failure, or queue failure. Scrub source/preview immediately after a successful move; the queue-owned encoded RAII buffer scrubs after write or close/drop. There is no retry, partial wrapper, second admission, or send into a newer incarnation.
- User exhaustion rejects only the new request, posts one coalesced localized status, and clears below 50% of the current effective byte and descriptor ceilings. Live lowering atomically changes normal admission; accepted requests drain in order, nothing is discarded, and no new normal request is accepted until usage fits. The fixed protocol reserve remains unchanged and usable only by required responses and the one control placeholder. Follow/history retry or cancel truthfully.
- A required engine response that cannot fit the fixed reserve marks the session fatal and immediately returns from its callback; after parser return, the reader transitions to Failed/retirement. Never drop a required response. Failure to reserve the one control placeholder instead occurs before its pipe Write and follows the mechanically proved no-submit/no-action branch; it is never reclassified as a dropped response or allowed to fail after submission.
- Closing rejects admissions, cancels synchronous writer I/O, discards queued input, and joins writer only on retirement.

### Non-Kitty parser and effect bounds

All sizes below are exact byte counts (`1 KiB = 1024`, `1 MiB = 1,048,576`) measured on the original encoded untrusted transport span before UTF-8 decoding, normalization, sanitization, base64 decoding, or copying. For OSC, a record is the bytes after the OSC introducer and before BEL/ST; the introducer and terminator are not counted. For CSI/DCS/APC and other escape records, the body is everything after its introducer through the final byte, excluding the introducer. Printable stream data is not an escape body. A dedicated authenticated integration channel measures the exact complete encoded frame bytes defined by plan 5. Named protocol caps replace the generic body cap for that record and are still subject to the pending-effect aggregate cap; field caps are conjunctive with their named record cap. Parsers count incrementally with checked arithmetic and enter a bounded discard-until-final/terminator state as soon as a cap is exceeded, so an attacker cannot force allocation of the oversized body.

- An otherwise unnamed non-Kitty escape/effect body is limited to 64 KiB.
- OSC 8 is limited to 16 KiB for the full record, 1 KiB for the parameter field, and 8 KiB for the URI field.
- A title/status record is limited to 4 KiB, then sanitized and truncated to 256 Unicode scalar values internally and 80 displayed grapheme clusters; invalid encoding or sanitizer failure rejects the effect rather than retaining a partial title.
- A cwd/path effect is limited to 128 KiB.
- OSC 52 is limited to 1 MiB encoded and 768 KiB decoded. Clipboard-read requests are ignored. Streaming base64 decoding reserves the decoded upper bound before materialization and rejects malformed or over-limit data without a partial clipboard request.
- Every decoded supported OSC 52 write receives immutable identity `{serviceGeneration,sessionId,sessionIncarnation,viewGeneration,osc52PolicyGeneration,effectSequence}` and an opaque one-shot token. `effectSequence` is a per-incarnation monotonic `uint64_t` that never wraps; exhaustion permanently denies OSC 52 for that incarnation. The plugin owns decoded bytes in an RAII buffer that securely zeroes its complete capacity on every disposition. UI posts contain only opaque identity/token plus bounded content-free display metadata, never clipboard bytes.
- Effective `deny` scrubs immediately. Effective `allow` queues no prompt but still commits only through the final authorization gate. Effective `ask` returns promptly from the engine callback and permits exactly one outstanding request per session and eight service-wide. A newer same-session request atomically denies/scrubs the older request and replaces its token; this security-broker supersession is not pending-effect coalescing. A ninth distinct-session request is denied/scrubbed with one coalesced localized status. No callback waits.
- Any OSC 52 policy transition, newer same-session request, session/view close, incarnation retirement/restart, generic plugin disable/re-enable, normal-refresh `BeginRefresh`, or application shutdown first installs its revocation fence under the OSC 52 authorization gate, then invalidates UI tokens and securely scrubs every outstanding ask/allow payload. Effective-policy and disable/re-enable transitions increment `osc52PolicyGeneration` before publication. Disable makes later requests effectively `deny`; re-enable applies configured policy only to later requests. A more-permissive transition never upgrades/replays an older request. UI allow/deny consumes its exact token at most once; stale, duplicate, wrong-session/view/incarnation/generation, and post-revocation replies are no-ops, and deny never calls the clipboard.
- After UI approval, and for prompt-free `allow`, acquire the OSC 52 authorization gate and immediately revalidate the complete identity, exact payload ownership, current effective policy (`ask` for an approval, `allow` for prompt-free commit), enabled/non-refreshing service generation, live view generation, and a `Running`/not-closing session. Perform the sole system-clipboard replacement while that gate remains held, consume the token whether replacement succeeds or fails, and scrub local bytes before release. Policy publication and teardown use the same gate, so the clipboard write is linearly ordered before or after—not across—revocation. There is no retry, delayed write, partial clipboard mutation, or host callback after consumption.
- An authenticated integration frame is limited to 2 MiB, with nested command text limited to 1 MiB and nested cwd limited to 128 KiB. Exceeding any limit permanently invalidates that incarnation's sole semantic epoch/trust state, purges its pending semantic work/history, and rejects every later marker or second handshake. Only explicit Restart or a new session may create another epoch; there is no live reauthentication.
- Pending optional effects are limited per session to 256 descriptors and 2 MiB of aggregate encoded body bytes, including descriptor/payload storage charged before enqueue. Only the latest-value `title`, `status`, and unauthenticated cwd presentation effects are coalescible, keyed by their exact effect kind; bell, hyperlink/range, clipboard, and authenticated-integration effects are never coalesced. A coalescing replacement may publish only after reserving its complete new descriptor/byte charge; until then the older effect remains intact, and the swap releases the old charge only after the replacement is installed. It never temporarily exceeds either limit and never evicts an unrelated effect.

An oversized optional display/status/title/cwd/hyperlink/clipboard effect is ignored or denied in full, increments a content-free per-protocol rejection counter, and posts at most one coalesced localized status; ordinary parsing continues after the bounded discard state reaches the record terminator. The identical reject-in-full outcome applies when an individually valid non-integration effect would become descriptor 257 or make aggregate encoded bytes exceed 2 MiB: a failed coalescing replacement leaves the prior value unchanged, a noncoalescible effect is discarded, and neither case evicts or partially publishes queued work. If an individually valid authenticated-integration effect cannot reserve either aggregate ceiling, atomically and permanently invalidate that incarnation's sole integration epoch, discard the new effect, purge and release every not-yet-applied queued integration effect/history candidate from that epoch, and reject every later marker/second handshake until explicit Restart or a new session; do not wait for a later sequence gap to discover the loss. A protocol-specific rejection latch clears only after one later well-formed, within-limit ordinary record of that protocol is processed; this latch never restores semantic trust. The pending-effect aggregate latch clears only when both descriptor and encoded-byte usage fall below 50% of their ceilings. An oversized authenticated integration frame follows the same permanent epoch-invalidating rule. If the engine identifies a response as required by the terminal protocol and its encoded response cannot fit the existing fixed protocol/control reserve after any already-admitted control-placeholder charge, the session follows the queue contract's fatal path; it is never silently dropped or reclassified as optional.

## Model and immutable snapshots

- One mutation mutex and exactly one snapshot producer per session. Copy every borrowed datum within the Gate-0-proved engine lifetime.
- Service owns two snapshot `std::jthread`s. UI/render/reader never build full snapshots.
- Track latest requested model generation, single queued/running bit, latest immutable publication slot, stable session/control generation, one notification-in-flight bit, and one Kitty conversion generation queued/running for the entire session.
- Snapshot scheduling is release-publish generation, acquire clear-to-set ownership, then producer clear-and-acquire-recheck; a race after clear is queued only by the thread that reacquires. Presentation publication/notification follows the same adopt-clear-reread-reacquire rule. Active failed posts get one bounded service retry; retired posts discard; ownership never remains stuck.
- Maintain `fullSnapshotRequired` independently of engine dirty flags until successful publication. Gate 0 proves whether begin/update consumes dirty state and how failure/cancel after begin forces a later full snapshot.
- Visible mutations single-flight/coalesce; hidden sessions parse without presentation builds; selecting hidden requests full.
- Publish only complete snapshots. Before natural exit, failure/cancel publishes nothing, leaves dirty/full state, and may schedule only the coalesced ledger-release retry defined below; it never spins or recursively retries. Natural exit requests one final full snapshot and follows the bounded final-snapshot outcome below instead of the ordinary indefinite-dirty rule.
- A presentation post carries only stable session/control IDs and published generation. The receiver uses repository payload helpers to first adopt the latest immutable publication, then clears notification ownership, acquire-rereads the published generation, and reacquires/reposts if newer work raced the clear. Active post failure clears ownership and schedules one bounded retry; retired/stale IDs discard. HWND reuse and null takes for nonzero tokens are ignored.

## Resource accounting

- Hard maxima: 32 service session slots; 64 MiB Kitty decoded/transient CPU/session; 256 MiB Kitty CPU/process; 256 MiB Kitty GPU/process; 16M pixels/image; one conversion generation/session total; 64 MiB terminal/scrollback peak/session; 512 MiB terminal/scrollback/snapshot CPU/process. One generation may convert multiple changed images sequentially; newer generations coalesce/cancel unpublished work, with no per-image workers.
- Charge all engine-owned decoded storage, staging, BGRA blobs, distinct published/UI/in-flight allocations, and allocated D2D bitmap bytes exactly once. If any engine-owned bytes cannot be measured and bounded before untrusted allocation, STOP. Checked width/height/stride/byte arithmetic is mandatory.
- The terminal/scrollback ledger includes current/alternate grids, retained rows/metadata, immutable text snapshots, UIA document/provider snapshots, copied UIA text/range buffers, selection/copy extraction snapshots and ranges, and simultaneous reflow/compression old/new or source/destination staging. Each owning allocation has one RAII ledger token shared by all aliases. A UIA RPC/provider object or copy operation holding an older generation keeps that generation, token, and every derived charged buffer reserved until its last external/internal reference drops; publishing a newer generation or retiring HWND discoverability never releases storage still reachable by an outstanding RPC/provider/range object. Removing discoverability atomically marks the provider disconnected and removes lookup maps; a call that captured its immutable snapshot/token before that fence may complete once, while every later method except `IUnknown` returns `UIA_E_ELEMENTNOTAVAILABLE` with null/zero output and no selection/scroll/focus/copy side effect. An externally reachable disconnected provider/range is reduced to only its self-contained immutable snapshot and passive ref-counted ledger state; its final release cannot call any host, service, HWND, dispatcher, engine, or session object. Track disconnected external-object and captured-in-flight-call gates separately. A UIA call that cannot reserve returns exactly `E_OUTOFMEMORY`, initializes every out parameter to its empty/null value, publishes no partial string/range/array, and leaves the provider retryable. A local selection/copy export that cannot reserve returns `HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY)`, leaves the clipboard and selection unchanged, and posts one coalesced localized resource-limit status. No UIA/copy path borrows uncharged model memory.
- Prune oldest scrollback before growth/reflow; never prune visible cells. If a resize peak cannot fit, retain last safe dimensions and publish one bounded status. Optional compression reserves both sides and checks cancellation.
- Charge every such allocation to both the 64 MiB session and 512 MiB process text ledgers, including hidden and quiet-retained exited snapshots. On process pressure release obsolete/unpublished generations, then prune oldest scrollback from LRU hidden exited, hidden live, then visible sessions without invalidating an in-use publication or visible grid. Old publications, UIA/provider snapshots, and copied ranges stay charged until references drop; admission remains paused. If required current output/grid growth still cannot reserve, fail/retire only that session. A quiet exited view releases its service slot but not snapshot charges.
- Reserve through RAII before decode/copy/convert/upload; release on every failure/cancel/device-loss/teardown path.
- Session LRU recency changes only on a user action that selects/reveals that terminal as visible, never on output, parser mutation, snapshot publication, notification, paint, background UIA access, pane activation that leaves it hidden, or polling. `TerminalService` assigns a service-wide monotonic `uint64_t` activation sequence while holding the registry/ledger lock; eviction orders ascending `(lastActivationSequence, TerminalSessionId)`. A never-activated session has sequence zero. Before counter overflow, under the same lock renumber live/retained sessions to `1..N` in their existing `(lastActivationSequence, TerminalSessionId)` order and continue, preserving total order. `TerminalSessionId` is immutable and ordinal compared.
- Kitty CPU/GPU recency changes on the first successful user-visible present of a new image allocation/upstream generation, or when selecting/revealing a terminal causes that retained image to be presented after being hidden. Output, hidden publication, activation that does not reveal the image, and routine repaint of an already-visible unchanged generation do not advance it. Eligible images evict ascending `(imageRecencySequence, TerminalSessionId, KittyImageId, upstreamGeneration, deviceGeneration)` after the CPU/GPU eligibility class; every tuple field is immutable for that allocation and compared ordinally/unsigned, so equal presentation recency is deterministic. The service owns the checked `uint64_t` image-presentation sequence and, before overflow, renumbers retained allocations to `1..N` under the ledger lock in their existing full eviction-tuple order. For CPU pressure, evict eligible unreferenced decoded Kitty/blobs and obsolete generations in that order; for GPU pressure, evict hidden/unreferenced eligible images then obsolete generations. Never invalidate the current publication/paint.
- A failed ledger reservation atomically sets one per-session/per-ledger `retryRegistered` bit, records the exact checked `requestedBytes` for the latest coalesced request, the observed release epoch, and a monotonic registration sequence, and enters that ledger's waiter set once. Releasing charged bytes increments the ledger's monotonic release epoch and starts at most one release pass when none is active. The pass freezes the cohort of registrations whose observed epoch precedes that release, marks its release epoch, and transfers the single service-wide `retryOwner` to the first member in ascending `(observedReleaseEpoch, registrationSequence, TerminalSessionId)` order; release itself only schedules that owner and never executes snapshot/decode work inline. The owner runs one off-UI eligibility/prune/reservation attempt for its latest coalesced request and records that it was attempted for this pass. Success removes its registration. Failure may re-register against the now-current epoch, but that new registration is ineligible for the current pass. Cancellation/retirement removes its registration. After each success, failure, or cancellation, the owner releases ownership under the ledger lock and transfers it, one owner at a time, to the earliest not-yet-attempted frozen-cohort waiter whose checked `requestedBytes` fit the currently unreserved bytes; if none fits, the pass ends. A waiter whose request changes updates `requestedBytes` under the lock but never joins an already-frozen cohort. Thus one release may service multiple waiters sequentially when residual capacity permits, but each frozen member runs at most once, the pass performs at most its frozen cohort size in attempts, and it never spins, polls, or creates concurrent owners. Waiters left registered require a later release epoch. Other producers merely coalesce their requested generation. Release-epoch and registration-sequence counters are checked `uint64_t` values and never wrap: before overflow, under the ledger lock rebase registered observations/current epoch and renumber waiters while preserving the same total waiter order, then perform the triggering release/registration.
- Lowering a cap blocks new reservations, evicts/prunes off UI, and commits only once usage fits; failure retains the old cap, and no allocation is grandfathered. Rejection keeps terminal usable.

## Retirement and shutdown

- Root-exit observation, explicit view close, application shutdown, and the per-view retirement-intent arbiter serialize the path decision. `Closing` first suppresses later `Exited`; root-drain first keeps its earlier `t0` and final-output policy, while the pending-confirmation matrix above selects the sole host-removal owner at quiet. No race resets or extends a deadline.
- Explicit user/application retirement and observed root exit are separate. After writer join, reset the sole host input-write handle exactly once. Root exit enters `DrainingAfterRootExit` with its own monotonic `t0`, captures code, and drains descendants until output EOF or two seconds; root signaling alone never ends the drain. It then resolves the final-snapshot outcome below and creates `Exited`. `closeOnExit` is evaluated only after this natural-exit path reaches complete quiet and only when `finalSnapshotComplete=true`.
- For the natural-exit final full snapshot, first cancel/remove any ordinary snapshot ledger-retry registration, release obsolete/unpublished generations, perform the deterministic eligible eviction/scrollback pruning above, and attempt one complete reservation/build. The attempt returns exactly success, resource-unavailable, lifecycle-cancelled, or engine-contract-failure; resource-unavailable includes a failed ledger reservation or a physical allocation failure after a valid reservation. On resource-unavailable, wait off UI until the text-ledger release epoch changes or 250 ms elapses, whichever occurs first, but never beyond `t0+5s`; if time remains, retry the complete reservation/build exactly once even when no epoch changed. Reaching `t0+5s` before that retry or a second resource-unavailable result ends resource snapshot work; there is no further timer, output-triggered, paint-triggered, ordinary-ledger, or release-triggered retry for that exited generation. Lifecycle-cancelled follows the already-serialized explicit-close/application-shutdown path and publishes no resource-limit `Exited` state. Engine-contract-failure transitions to the normal localized terminal-failed retirement outcome and is never mislabeled as resource pressure.
- On the bounded resource-unavailable outcome, retain the last successfully published immutable snapshot as `lastSafeSnapshot`; if none exists, expose only a content-free empty diagnostic surface. Set retained metadata `finalSnapshotComplete=false`, show the localized visible status exactly `Exited — final output unavailable (resource limit)`, and permanently suppress `closeOnExit` for that exited view. Continue process/engine teardown immediately: final-snapshot resource failure alone never quarantines, never changes the existing quiet deadline, and never extends handle/runtime lifetime. `Restart` stays disabled until complete quiet, then becomes enabled while the incomplete marker/status remains. Complete quiet releases the service slot even though the retained last-safe snapshot remains charged until its references/view are released.
- On successful final publication set `finalSnapshotComplete=true`; retain the published generation through the normal `Exited`/`closeOnExit` policy. The flag and status outcome are immutable for that exited incarnation and are not changed by late reference release.
- If root-exit drain reaches EOF, close ConPTY off-thread. At its two-second deadline, below 26100 cancel/join reader and close output before final close; 26100+ close while reader drains EOF. Direct-root termination is a no-op for already-signaled root.
- UI close transfers active to retiring first, gates posts/input/snapshot, invalidates view generation, drains/destroys HWND, removes view, and returns without join/wait/engine destruction/`ClosePseudoConsole`.
- Off-UI retirement starts monotonic `t0`, closes/discards input and joins writer.
- Reuse `Common/MinimumOsVersion`: below 26100 continue draining until output EOF or the two-second deadline; root-process signaling alone never ends the drain. At the deadline cancel/join any remaining reader and close host output before final off-thread `ClosePseudoConsole`; 26100+ close immediately on teardown while the reader drains EOF.
- Graceful deadline `t0+2s`; then cancel remaining I/O, terminate only the direct root if alive, and complete any pending legacy close sequence. No production kill-on-job-close/descendant enumeration; detached GUI/background processes survive.
- Explicit `Closing` publishes no final snapshot and never transitions to `Exited`; its view is already retired. If its later quiet deadline fails, retain strong ownership/runtime and emit a content-free application lifecycle diagnostic/status that is not a terminal tab and does not count toward `maxTabsPerPane`. Only `DrainingAfterRootExit` resolves the bounded final-snapshot outcome/`Exited`, then tears down process/engine resources while retaining immutable snapshot/restart metadata. Its `Restart` action remains disabled until quiet; five-second teardown failure keeps/converts a still-present tab to teardown-timeout quarantine and suppresses automatic removal. If the user already closed that view, failure uses the same content-free diagnostic. Application shutdown instead uses the app-global fatal policy.
- After I/O/close workers join, atomically gate new per-session snapshot/conversion scheduling. Natural exit first resolves its one bounded final-snapshot outcome; explicit close/application shutdown publish none. Cancel queued/unpublished work and wait off UI for a per-session fence covering every accepted queued/running snapshot and Kitty conversion. Each unit decrements through scope-exit only after releasing every engine borrow/model lock and ledger reservation. The service's shared executor remains available to other sessions and teardown never waits for their work. Only after this fence reaches zero detach session callbacks under the model lock, destroy engine/render state, drain posts, close handles, and release runtime. A miss shares the existing `t0+5s` quarantine/application-fatal outcome; it never authorizes destruction or unload.
- Every normal or recovered complete-quiet transition releases its service slot exactly once; entering `Exited`, retaining a view, or failing the final snapshot never releases it early.
- Quiet deadline `t0+5s`. Explicit-close failure uses the content-free non-tab diagnostic; natural-exit failure uses the visible teardown-timeout view; both quarantine strong ownership/runtime. App shutdown invokes injectable production fail-fast policy. Never detach or unload reachable code.
- If quarantined work later reaches complete quiet, atomically remove its registry ownership and release last runtime/handles. A removed view stays absent and its content-free diagnostic clears. A still-present natural-exit view enables Restart/Close, retains its snapshot, and never re-evaluates `closeOnExit`; it becomes `Exited — cleanup recovered` only when `finalSnapshotComplete=true`, otherwise it returns to the exact `Exited — final output unavailable (resource limit)` status. Cover both late-recovery paths and both final-snapshot outcomes with deterministic barriers.
- Release the service slot on that late quiet transition and reopen launch/restart admission only when no quarantine remains and capacity exists.
- App shutdown uses one service-wide `t0` and starts all teardown concurrently. A nonblocking watchdog enforces deadlines without joining/calling ConPTY close and can still escalate a stuck worker. Main `WM_CLOSE` defers `DestroyWindow`; a stable service-quiet message continues destruction while target/pump stay alive.

## Tests and performance

Add ABI/module tests that create `builtin/terminal` only through `RedSalamanderCreate(IID_ITerminal,...)`, reject `IID_IViewer`, validate every v1 struct size/version/reserved/pointer-length case including the exact 336-byte Open layout/offsets and source/launch cross-product plus the exact 416-byte insertion layout/offsets and every mode cross-field combination, require one ownership-marked direct child, and prove copied input lifetimes. The shared boundary corpus accepts every exact UTF-16 component/profile maximum and rejects +1/NUL with no truncation, alias, digest, callback, child, or process. Insertion cases use an initiating source both equal to and different from the target terminal's immutable original source, prove that the latter is accepted for an insertion-capable Windows target when its copied locations translate, prove `CurrentDirectoryFull` uses only the initiating source location, prove exact-distribution WSL accepts only a quoted full native Linux path, and assert different-distro/cross-family/unsupported cases fail with the specified HRESULT and zero bytes. Exercise the exact callback vtable order/removal enum and reentrant state/close/removal delivery, forbid callback clear inside callback, drain callback delivery, and cover stale/duplicate/wrong-instance/unknown-reason removal plus automatic `Exited -> Quiet -> ViewRemovalRequested` ordering. Deterministic CAS barriers prove the single per-view arbiter for UserTab allow/cancel, every forced/chooser/plugin/automatic takeover, natural exit before/at/after quiet with every `closeOnExit`/snapshot/quiet result, stale/duplicate prompt replies, and approval/cancel/takeover ties. A shared `{instanceId,hostViewGeneration}` removal claim admits exactly one host mutation/null-drain/Close/release/MRU fallback when CloseApproved, a queued removal callback, forced teardown, and host-generation reuse race. Close during Starting/Running, call idempotent `RedSalamanderPluginShutdown`, and assert `RedSalamanderPluginCanUnloadNow` stays false through active, retiring, quarantined, callback, child-HWND, worker, selected-engine, runtime, captured UIA call, and externally held disconnected provider/range ownership; the last case may use only the closed process-exit-retention contract, while all ordinary owners must reach quiet. Add `Tools/Tests/TerminalPluginBoundarySourceContracts.Tests.ps1` with the master's exact reviewed allowlist; it rejects terminal implementation includes/symbols in the EXE/Common host bridge beyond `PlugInterfaces/Terminal.h` and the generic manager/pane host allowlist, rejects Terminal use of `IViewer`, and rejects direct terminal implementation/selected-engine linkage into the EXE.

The Open ABI/golden corpus must explicitly assert `openPurpose@300`, both valid purpose values, and every unknown, missing, null/non-null-mismatched, selection-mode-mismatched, or location-family-mismatched purpose/insertion combination. Every invalid synchronous Open returns `E_INVALIDARG` with zero child/process, zero callback/status delivery, and zero view/pending-open reservations. The loaded-DLL callback oracle must explicitly assert the six slots in order: `TerminalStateChanged`, `TerminalPendingOpenResolved`, `TerminalOpenSiblingRequested`, `TerminalCloseApproved`, `TerminalViewRemovalRequested`, `TerminalQuiet`. It covers one noncoalescing pending-open result for the exact generation, every closed result, Ready followed by exact Commit or Abort, and stale/duplicate/wrong-generation Ready/Commit/Abort without mutation.

Cover exact boundary/+1 queue/resource cases, including 32 held slots and denied 33rd with zero engine/process state; every slot-holding lifecycle state; rapid retiring sessions; quarantine-wide admission closure and late-quiet release/reopen; quiet retained exited snapshots using no slot; 64 MiB session and 512 MiB process text/snapshot limits; old publications, UIA provider/document generations, copied text/ranges, and outstanding RPC references staying charged; deterministic hidden-exit/hidden-live/visible scrollback pruning; retained exited-snapshot charging; unserviceable-session-only failure; and complete release on close/restart. Use deterministic barriers for retirement with a queued snapshot, a snapshot paused before and inside the selected-engine lifetime, a running Kitty conversion, cancellation before publication, natural-exit final-snapshot completion before the fence closes, explicit-close/application-shutdown races, and two sessions proving one retirement waits only for its session fence while the shared executor keeps serving the other session. Drive equal-recency sessions and Kitty images in reverse construction order to prove stable `TerminalSessionId`/`KittyImageId` tie-breaks; prove only user-visible activation changes session recency, paint/output/UIA do not, and counter renumbering preserves order. Register multiple differently sized waiters, release enough bytes for more than one, and prove one concurrent retry owner hands off deterministically through the same frozen release cohort while residual capacity fits later requests; each cohort member runs at most once, newly registered/failed waiters wait for a later epoch, a too-large request cannot prevent a later ordered fitting request, and there is no polling, recursive retry, or retry storm.

For every non-Kitty bound, test exactly limit and limit+1 encoded bytes, fragmented at every introducer/terminator and cap boundary: generic 64 KiB; OSC 8 16 KiB record/1 KiB params/8 KiB URI; title/status 4 KiB plus 256/257 sanitized Unicode scalar values and 80/81 displayed grapheme clusters; cwd/path 128 KiB; OSC 52 1 MiB encoded/768 KiB decoded; integration 2 MiB frame/1 MiB command/128 KiB cwd; and pending effects 256 descriptors/257 plus 2 MiB/2 MiB+1 aggregate. At both aggregate boundaries cover successful and failed latest-value replacement, noncoalescible optional rejection with the old queue unchanged, and authenticated-integration admission failure purging that incarnation's sole epoch and invalidating trust immediately. Verify checked arithmetic, bounded discard recovery into the following valid printable/escape record, no partial effect/clipboard/title/path, content-free counters and exact latch-clear behavior, parser continuation for optional effects, permanent integration-epoch invalidation with every later marker/second handshake rejected and restart/new-session-only rekey, and fatal retirement only for a required response that cannot fit its fixed reserve.

Add deterministic unsafe-paste authorization tests with an instrumented clipboard, allocator/scrubber, selected-engine mode fake, queue barriers, and injected generations. Prove exactly one UI-thread clipboard read even when clipboard contents change before approval; bounded terminating-NUL scan; malformed UTF-16/embedded-NUL/overflow rejection; strict UTF-8 conversion; one-time CRLF/CR normalization; threshold-1/exact-threshold line classification; HT/LF versus every other C0/C1 control; and preview/classification/final bytes derived only from the frozen payload. Verify the complete immutable identity and no-wrap sequence; opaque one-shot token; one pending prompt per terminal; exact 8 MiB/session and 32 MiB/service reservation boundaries/+1; reserve-before-materialize; newer-paste supersession; every cancel/policy/input-mode/view/session/incarnation/restart/retirement/app-shutdown revocation; more-permissive-setting no-upgrade; and secure zero of full source/preview capacity on every branch. Safe and captured-no-warning unsafe requests use the same final path without a prompt. Approval never rereads the clipboard, uses only current matching paste mode, and under the authorization gate atomically revalidates identity/generations/liveness plus current `pasteMaxBytes`, normal-byte, 3840-descriptor, and remaining-capacity bounds before one complete wrapper/event admission. Stale/duplicate/wrong-generation, encoding, cap, queue, close, and newer-incarnation cases emit zero bytes and consume/scrub once; the queue-owned encoded buffer scrubs after write/drop. Crucially, generic disable and accepted normal refresh leave a live view's safe paste and already-pending unsafe paste usable under its frozen configuration; continuity barriers before/after both transitions prove neither changes paste-policy generation, revokes the token, nor rejects final admission.

Add deterministic ordered-control placeholder tests at fixed-reserve descriptor and byte boundary and +1. With barriers before reservation, after reservation/before pipe submission, during immediate and pending completion, and before activation, enqueue later user input and prove: user-before-reservation cancels the undispatched operation; reservation assigns the earlier immutable queue sequence; later input remains behind the inert head; exact completion activates that same descriptor without another admission or sequence; and the writer emits the wake before any later user byte. Cover pre-submit cancel/tombstone, fixed-reserve admission failure with zero Write/effect, post-submit timeout/disable/failure held-fence behavior, safe authenticated result release, ambiguous Discard/Close recovery, required-response competition, and close/retirement. Assert 4096 total/3840 user/256 fixed-reserve accounting, no sequence reuse/gap-induced stall, no priority bypass, and no wake-capacity failure after pipe submission.

Add deterministic OSC 52 authorization tests for `deny|ask|allow`, ignored reads, immutable service/session/incarnation/view/policy/effect identity, effect-sequence exhaustion, opaque content-free UI posts, and full-capacity secure scrub on every disposition. Exercise exact one/session and eight/service outstanding-ask limits, ninth distinct-session denial, newer same-session supersession, and nonblocking callbacks. Place barriers before/inside/after policy changes, generic disable/re-enable, normal-refresh `BeginRefresh`, close/restart/retirement, newer request, and application shutdown; the revocation fence must win before policy/availability/teardown publication, invalidate tokens, scrub both ask and prompt-free-allow payloads, make disabled behavior effective deny, and never replay after re-enable or a more-permissive transition. Stale/duplicate/wrong-identity/post-revocation replies make no clipboard call. For approved ask and prompt-free allow, linearize the sole clipboard replacement while holding the same gate after revalidating exact ownership, current effective policy, enabled/non-refreshing generation, live view, and Running/not-closing state; success and injected clipboard failure both consume once and scrub, with no retry/delayed/partial write.

Force the natural-exit final snapshot to fail before reservation, with a resource-allocation failure during build, without any release epoch, with a release before 250 ms, at 250 ms, and after `t0+5s`. Prove at most one retry, no infinite release/output/paint retry, last-safe-snapshot retention and its continued ledger charge, the no-prior-snapshot empty diagnostic, exact `finalSnapshotComplete=false` metadata/status, permanent `closeOnExit` suppression, unchanged teardown deadline, no quarantine solely from snapshot pressure, Restart disabled until quiet/enabled afterward, and service-slot release exactly at quiet. Separately prove lifecycle cancellation publishes no resource-limit exit and a non-resource engine snapshot failure uses `Failed`. Also cover the successful `finalSnapshotComplete=true` path, teardown-timeout/late-recovery status precedence for both outcomes, live queue/cap lowering with accepted-data drain and no eviction of visible grid, callback nonblocking failure, strict order, atomic paste, ABI mismatch before handles, deterministic barriers immediately before/after snapshot/notification clears, post failure without later output, failure after engine dirty consumption, HWND reuse, malformed/split input, allocation failure, lifecycle ordering/stress, and no terminal content in diagnostics.

Add deterministic shutdown-coordinator tests for the three close reasons; stable-ID ordering across many terminal objects; callbacks remaining live through every `ApplicationShutdown` request and draining before `Close`/release/export; zero-terminal and object-free-retiring export start; immediate/deferred quiet; startup versus first shutdown; duplicate/stale generation, target, `WM_CLOSE`, callback, and poll delivery; exactly one continuation; and normal refresh/disable remaining separate from the process-wide continuation. Cover administration absent/idle and every accepted Prepare/Arm/pre-CAS/in-CAS/post-CAS/Finalize/compensation/result state, cancellation-wins/loses, exact Failed/Committed Finalize, compensation acknowledgement, retired-page suppression, result take/retry/release, callback-clear/Close/release ordering, `ERROR_BUSY` invariant breach, and export only after both public barriers. Inject the clock and dispatcher to prove samples are generation-checked and never more than 50 ms apart, the main target/pump remains alive, manager/service deadlines have the exact no-later relationship, and the initiating UI turn performs no wait, join, or `DestroyWindow`. Prove an externally held disconnected UIA provider/range with zero captured in-flight calls and only passive immutable snapshot/ledger ownership returns `UIA_E_ELEMENTNOTAVAILABLE`, touches no host/service, and may retain the module through process exit; a racing final Release permits ordinary unload, every active-work false vote is fatal at five seconds, and refresh/disable cannot use retention. Source-contract tests reject an EXE terminal-service accessor, direct Terminal-private symbol, duplicated export lookup, or any process-exit-retention vote outside that closed passive tail.

Add deterministic normal-refresh tests for disable with running terminals and no retirement/export; idle re-enable reuse; every `BeginRefresh` generation/result case and irreversibility; terminals outliving both five-second clocks without `RequestClose`/hide/disconnect; configuration work paused before/after `CommitStarted` and CAS with exact Finalize/compensation/result settlement; settlement timeout plus late safe completion; last terminal release before/after settlement; immediate/deferred module quiet; `ModuleQuietTimeout` plus one host Retry window; enable/disable churn in every refresh state; replacement validation/load/configuration failure; old `wil::unique_hmodule` reset before the sole replacement map; absence of overlapping mapped module/service generations; and process shutdown superseding every state. Assert normal refresh never calls retention or fatal seams, never abandons callbacks/private transactions, admits no post-fence public request, and never silently resumes after a blocked timeout. Explicitly prove disable/refresh continuity for safe and pending unsafe paste on an existing live frozen-config view, while both transitions revoke/scrub outstanding OSC 52 authorization and prevent any post-fence OSC 52 clipboard write.

Add plugin-owned view-admission tests at effective limits 1 and 32 plus the rejected +1; independent complete physical-host keys; simultaneous and callback-reentrant Open; quota after all synchronous validation but before child/async work; ordinary Open failure; failure after Open during host commit; chooser/diagnostic/Failed/Exited/teardown-timeout retention; live lowering below the current count without eviction; raising; callback-drain/Close/public-release exactly-once release; internal retirement/quarantine after view removal retaining no slot; zero-key erasure; and release-then-reopen. Assert exact quota HRESULT and no visible rejected tab, zoom/selection change, child, process, callback publication, or leaked reservation.

Add WSL launch tests for exact reopened System32 `wsl.exe`, exact five-token argv with spaces/quotes/backslashes/metacharacters/leading-option-like and non-BMP distro/path data, configured default-user/default-shell preservation, no `--exec`/`--user`/`-e`/wrapper/bootstrap/extra process/fallback, and zero process on preflight failure. Deterministic preflight barriers cover exact UNC directory identity, retained handle through process/startup admission, catalog/launch-generation staleness, missing/renamed/path loss, cancellation before/during/after directory open, and just-before/equal/after the fixed three-second deadline including irrevocable failure before `CancelSynchronousIo` and late-handle-only release. Launch/environment/transport fakes and source contracts prove no WSL-specific argv/environment/`WSLENV` delta, no PTY/profile/staged-script injection, no guessed `/mnt/<drive>` path, and no WSL semantic nonce/epoch. State tests hold `RequestedUnverified` permanently, reject every attempted `Verified|Diverged` transition, keep the raw terminal usable, expose no app-history/follow/activity/authenticated-cwd/profile-memory capability and allow only exact-distribution full-native-path explicit insertion, and allow a post-Running launcher failure with no fallback or false memory; no test requires all asynchronous `--cd` failures to precede `Running`.

Add a fake child-recorder launch suite that asserts the exact creation-flag pair and byte-for-byte UTF-16 argv/environment round trip for non-ASCII values; later-parent ordinal-ignore-case duplicate resolution including retained winning spelling; valid `=X:` entries and rejection of malformed/relative/wrong-drive hidden entries; canonical OS/permitted-integration overwrite spelling and precedence while rejecting a semantic-nonce environment entry; deterministic total order; the exact two-NUL empty block; 32,767-unit command-line and 524,288-unit environment inclusive boundaries and +1; malformed/missing parent double-NUL; embedded-NUL/checked-overflow rejection; and zero process on every failure. For Windows profiles, require exact nonnull held/revalidated `lpApplicationName` plus fixed unquoted `cmd.exe|powershell.exe|pwsh.exe` `argv[0]`, exact child-token round trip, and immunity to hostile current directory/PATH/app-local same-basename files. For WSL, separately accept the exact five-token quoted 32,767-unit boundary (32,736-unit quote-free Linux path for distro `D`) and reject +1 before preflight/process. Assert no truncation, hash/alias substitution, ANSI conversion, null inherited block, or case-variant duplicate. Reuse `Tests/TestSupport/ChildProcess.h` recorder primitives and keep Terminal-specific overwrite policy in `Terminal.dll`; extract a shared builder only if its complete semantics match and update the shared-helper catalog.

Instrument the aggregated parse/read/queue/snapshot/scrollback/Kitty/session-close metrics defined by the master plus content-free service-slot held/denied/quarantine-admission gauges; session/process text bytes split by grid/scrollback/publication/UIA/copy; outstanding provider generations/ranges; reservation release epochs/wakeups/retry ownership; final-snapshot complete/incomplete/retry latency; deterministic pruning/Kitty eviction; paste pending-storage-cap admission/denial and OSC 52 ask-cap/revocation counts without clipboard length, payload, token, or identity values; and per-protocol oversized/rejected counters. Add a deterministic fake-transport Commands selftest `terminal_perf_core_model_baseline` before optimization; build test-enabled Release and archive its ungated baseline in the runner's `Commands` area.

## Canonical native-test evidence for plans 2–5

Plan 2 adds `Specs/Testing/TerminalNativeEvidence.schema.json`, `Tools/Run-TerminalNativeTests.ps1`, and `Tools/Verify-TerminalNativeEvidence.ps1`, and documents the `Specs/TestRuns/<MachineHash>/TerminalNative/<RunId>/` archive area in `Specs/TestRuns/README.md`. It also extends the existing `Tools/TerminalEvidence.psm1` with `ConvertTo-RSCanonicalTestRunPath` and `Resolve-RSCanonicalTestRunPath` exactly as the master defines, with focused tests. Plans 2–5 use this runner for every direct `TerminalTests`, `DxUiTests`, and `SettingsSchemaTests` invocation; raw console output and local absolute paths are never archived.

One run writes canonical RFC 8785 JCS UTF-8 without BOM or trailing newline as `terminal-native-evidence.v1.json` plus `terminal-native-evidence.v1.sha256` containing exactly its 64-lowercase-hex SHA-256 and LF. `runId` and its directory leaf are identical `yyyy-MM-dd_HHmmss.fffffffZ_<32-lowercase-hex-UUID-without-hyphens>` values, so concurrent/same-tick runs never overwrite; directory creation is fail-if-exists and a UUID collision regenerates before any evidence write. The v1 schema requires `recordKind=terminal-native-run`, schema/tool versions and digests, `runId`, `machineHash`, 40-lowercase-hex `repositoryCommit`, selected-engine-lock digest, slice ID, native platform/configuration, start/end UTC and duration, and a nonempty ordered invocation array. Each invocation records only the project/executable basename, executable SHA-256, PE machine, stable filter/suite ID, sanitized argument IDs with no filesystem/user/terminal data, exit code, and structured discovered/passed/failed/skipped counts emitted by the shared native test harness. A blocking invocation passes only with exit zero, discovered greater than zero, failed zero, and skipped zero. The runner rejects a dirty plan-owned source/test/tool set, architecture/configuration mismatch, missing structured result, duplicate invocation ID, or content-bearing output field; it may use a private temporary console capture but deletes it and never copies it into `Specs/TestRuns`.

`Run-TerminalNativeTests.ps1 -Slice Core|Renderer|PaneProfiles|ShellState -Platform <x64|ARM64> -Configuration <Debug|Release|ASan Debug> -EvidenceRoot .\Specs\TestRuns -PassThruManifest` owns the frozen invocation list for each slice and returns the one exact manifest path. `Verify-TerminalNativeEvidence.ps1 -Manifest <exact-path> -ExpectedSlice <id> -ExpectedRepositoryCommit <commit> -ExpectedEngineLockSha256 <digest>` validates schema, canonical bytes, companion digest, directory run ID/machine hash, tool identity, native PE identities, counts, and semantic pass; it never searches for a latest run. Immediately after verification, a child converts the resolved manifest to exact canonical repo-relative NFC `/` form `Specs/TestRuns/<MachineHash>/TerminalNative/<RunId>/terminal-native-evidence.v1.json`, records only that path plus both file digests, and its successor requires that spelling before resolving it for I/O. Absolute, alternate-root, wrong-case, backslash, dot-segment, reparse, outside, or alias paths fail. Tests copy the same pair under two different checkout roots and prove identical path-bearing identity-set digests, then exercise every rejection.

The v1 slice lists are exact and schema-tested: `Core` invokes `TerminalTests` filters `terminal_engine`, `conpty`, `terminal_snapshot`, and `terminal_resource_limits`; `Renderer` invokes `TerminalTests` filters `terminal_renderer`, `terminal_input`, and `terminal_accessibility`; `PaneProfiles` invokes the complete `DxUiTests` suite plus `TerminalTests` filters `terminal_profiles` and `terminal_working_directory`; `ShellState` invokes `TerminalTests` filters `terminal_shell_integration`, `terminal_pane_follow`, and `terminal_state_store` plus the complete `SettingsSchemaTests` suite. Unknown slice/configuration/platform combinations fail rather than silently shrinking the list. `Core` is required for x64 Debug/Release/`ASan Debug` and native ARM64 Debug/Release; the later slices require native x64 Debug here and are rerun by plan 6's full native closeout lanes.

## Verification

Native x64 lane:

```powershell
. .\Tools\TestRunPlan.ps1
$expectedRepositoryCommit = (& git rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0 -or $expectedRepositoryCommit -cnotmatch '^[0-9a-f]{40}$') { throw 'Cannot resolve exact repository commit.' }
$expectedEngineLockSha256 = (Get-FileHash -LiteralPath '.\External\terminal-engine.lock.json' -Algorithm SHA256).Hash.ToLowerInvariant()
foreach ($testFile in @(
    '.\Tools\Tests\BuildReproducibility.Tests.ps1',
    '.\Tools\Tests\TestRunArchive.Tests.ps1',
    '.\Tools\Tests\TerminalCommandsEvidence.Tests.ps1',
    '.\Tools\Tests\TerminalDependencyLifecycle.Tests.ps1',
    '.\Tools\Tests\TerminalPluginBoundarySourceContracts.Tests.ps1'
)) {
    $pesterParameters = New-RSPesterInvokeParameters -Path (Resolve-Path $testFile).Path
    $pesterResult = Invoke-Pester @pesterParameters
    $failedProperty = $pesterResult.PSObject.Properties['FailedCount']
    $failed = if ($failedProperty) { $failedProperty.Value } else { $pesterResult.PSObject.Properties['Failed'].Value }
    if ([int]$failed -ne 0) { exit 1 }
}
.\build.ps1 -ProjectName TerminalTests -Platform x64 -Configuration Debug
$x64DebugNativeManifest = .\Tools\Run-TerminalNativeTests.ps1 -Slice Core -Platform x64 -Configuration Debug -EvidenceRoot .\Specs\TestRuns -PassThruManifest
.\Tools\Verify-TerminalNativeEvidence.ps1 -Manifest $x64DebugNativeManifest -ExpectedSlice Core -ExpectedRepositoryCommit $expectedRepositoryCommit -ExpectedEngineLockSha256 $expectedEngineLockSha256
.\build.ps1 -ProjectName TerminalTests -Platform x64 -Configuration Release
$x64ReleaseNativeManifest = .\Tools\Run-TerminalNativeTests.ps1 -Slice Core -Platform x64 -Configuration Release -EvidenceRoot .\Specs\TestRuns -PassThruManifest
.\Tools\Verify-TerminalNativeEvidence.ps1 -Manifest $x64ReleaseNativeManifest -ExpectedSlice Core -ExpectedRepositoryCommit $expectedRepositoryCommit -ExpectedEngineLockSha256 $expectedEngineLockSha256
.\build.ps1 -ProjectName TerminalTests -Platform x64 -Configuration "ASan Debug"
$x64AsanNativeManifest = .\Tools\Run-TerminalNativeTests.ps1 -Slice Core -Platform x64 -Configuration "ASan Debug" -EvidenceRoot .\Specs\TestRuns -PassThruManifest
.\Tools\Verify-TerminalNativeEvidence.ps1 -Manifest $x64AsanNativeManifest -ExpectedSlice Core -ExpectedRepositoryCommit $expectedRepositoryCommit -ExpectedEngineLockSha256 $expectedEngineLockSha256
try {
    $env:RSBuildEnableTests = 'true'
    .\build.ps1 -ProjectName RedSalamander -Platform x64 -Configuration Release
} finally {
    Remove-Item Env:RSBuildEnableTests -ErrorAction SilentlyContinue
}
$coreCommandsRun = .\Tools\Run-TerminalCommandsEvidence.ps1 -Scenario CoreModel -Platform x64 -Configuration Release -RepositoryCommit $expectedRepositoryCommit -EvidenceRoot .\Specs\TestRuns -PassThruRunPath
.\Tools\Test-TestRunArchive.ps1 -RunPath $coreCommandsRun
```

Native ARM64 lane using the same Gate-0 lock/artifact identities:

```powershell
$expectedRepositoryCommit = (& git rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0 -or $expectedRepositoryCommit -cnotmatch '^[0-9a-f]{40}$') { throw 'Cannot resolve exact repository commit.' }
$expectedEngineLockSha256 = (Get-FileHash -LiteralPath '.\External\terminal-engine.lock.json' -Algorithm SHA256).Hash.ToLowerInvariant()
.\build.ps1 -ProjectName TerminalTests -Platform ARM64 -Configuration Debug
$arm64DebugNativeManifest = .\Tools\Run-TerminalNativeTests.ps1 -Slice Core -Platform ARM64 -Configuration Debug -EvidenceRoot .\Specs\TestRuns -PassThruManifest
.\Tools\Verify-TerminalNativeEvidence.ps1 -Manifest $arm64DebugNativeManifest -ExpectedSlice Core -ExpectedRepositoryCommit $expectedRepositoryCommit -ExpectedEngineLockSha256 $expectedEngineLockSha256
.\build.ps1 -ProjectName TerminalTests -Platform ARM64 -Configuration Release
$arm64ReleaseNativeManifest = .\Tools\Run-TerminalNativeTests.ps1 -Slice Core -Platform ARM64 -Configuration Release -EvidenceRoot .\Specs\TestRuns -PassThruManifest
.\Tools\Verify-TerminalNativeEvidence.ps1 -Manifest $arm64ReleaseNativeManifest -ExpectedSlice Core -ExpectedRepositoryCommit $expectedRepositoryCommit -ExpectedEngineLockSha256 $expectedEngineLockSha256
```

Expected: all x64 and native-ARM64 tests pass while loading the matching `Plugins\Terminal.dll` through `IID_ITerminal`; five exact content-free `TerminalNative` manifests verify; `IID_IViewer` is rejected; malformed ABI/plugin/runtime mismatch creates no engine state; the exact untruncated component/profile caps and +1 cases pass; plugin-owned per-host view admission returns the exact quota result and rolls back without visible/child/process state; the exact Unicode creation flags, fixed-basename Windows argv0, WSL five-token command boundary, and canonical bounded environment/command-line recorder matrix pass with zero-process rejection and no truncation/hash; the single per-view retirement arbiter plus shared host removal claim admits exactly one callback/forced path and host mutation through every confirmation/natural-exit/forced-close race; the immutable one-read paste transaction proves frozen normalization/classification, complete identity/generations, fixed pending-storage caps, one-shot final current-mode/cap/queue admission and scrub, with safe and pending unsafe paste surviving generic disable/normal refresh on a live frozen-config view; OSC 52 proves immutable identity, one/session and eight/service ask limits, revoke-before-policy/availability/teardown publication, effective deny while disabled, opaque tokens, one linearized clipboard write, and scrub; WSL launches the exact distro/path with only the five frozen tokens and configured default user/shell, remains permanently `RequestedUnverified`, and exposes zero integration/history/follow/activity/cwd/profile-memory capability or launch integration, while allowing only exact-distribution full-native-path explicit insertion while the raw terminal stays usable; callback drain/module quiet pass; callbacks never block; the UI-thread seam records no blocking operation; every shutdown and normal-refresh lifecycle test reaches its specified distinct outcome; disable never retires; refresh never force-closes a terminal or uses fatal/retention policy; the old module handle is reset before exactly one replacement; and the content-free Release core baseline is archived under `Specs/TestRuns/<MachineHash>/Commands/<RunId>/`. ARM64-native runner absence or an unverifiable native manifest blocks completion.

## Child completion and move

- **Completion record:** `TERMINAL-HANDOFF-UNRESOLVED` — replace this line with
  the complete ordered Gate-0 lineage, exact successful-handoff Done path/
  FDone closeout-commit/blob pair, every Round 5+ routing-only FPublish and
  durable FSeal commit/path/raw digest, TerminalEngine
  decision/finalization/lock/five-lane
  evidence repository commit/run/digest identities, implementation evidence
  commits, native result identities, and Commands archive path/file SHA-256
  values required below.

This child is complete only when its predecessor identity, implementation evidence, durable contracts, and move are all closed as one review. No implicit newest-run or directory scan may select evidence.

1. Consume the exact winner-bearing Done path named by the master's
   `successfulGate0Handoff`, never a hard-coded legacy Plan-1 path and never a
   WIP path. Copy the master's explicit ordered Gate-0 basenames through that
   winner and record, for every entry, its outcome, FDone closeout commit, Done-plan
   Git blob object ID, commit-bound decision path/SHA-256, and RecoveryArchive
   FA commit/path/SHA-256 (or the exact Round-4 legacy N/A triplet), routing-only
   FP, and durable FS/seal path/raw digest. Also record
   the winner's routing-only FPublish plus FSeal commits, require FS to be the
   one-file direct child of FP, FP to be the direct child of FDone, and FDone to
   be the direct child of FA. Require the winner
   to be the final listed round, every earlier outcome to be `no-winner`, and
   exact marker-byte occurrences across all supplied Done plans: zero in each
   no-winner, one in the sole final winner, and one total. Record every TerminalEngine
   evidence record's `repositoryCommit`. From the approved
   `Specs/Terminal/TerminalEngineDecision.md`, copy the exact finalization-
   manifest path, its approved 64-lowercase-hex digest, `runId`, winner candidate
   ID, three referenced collection-manifest paths/digests,
   `External/terminal-engine.lock.json` digest, and five-lane verification-set
   path/digest into the plan-2 review record. Revalidate the original finalization
   record's schema, exact bytes/companion digest, semantic selection, input
   references, directory `runId`, and winner; rerun Finalize only with those
   three explicit collection paths and require the same winner/frozen identities.
   A newly emitted validation record does not replace the two-reviewer-approved
   predecessor digest. Verify the lock and verification-set digests and rerun the
   native x64/ARM64 artifact verification above against that lock. Any chain gap,
   mismatch, second marker, later round after the winner, or unavailable native
   lane returns this plan to blocked.

   Use this fail-closed identity check after substituting only the exact values copied from the decision; placeholders, “latest,” and wildcard discovery are forbidden:

   ```powershell
   $gate0RoundNames = @(
       'Terminal_EngineSelectionAndDependencyProof_2026-07-22.md',
       'Terminal_EngineSelectionAndDependencyProof_Round5_WezTermSynchronousCallerOwned_2026-07-31.md'
       # Append each later consecutive Gate-0 basename explicitly before plan 2 starts.
   )
   $approvedGate0Lineage = @(
       [pscustomobject]@{
           Ordinal = 4
           PlanName = 'Terminal_EngineSelectionAndDependencyProof_2026-07-22.md'
           PredecessorHandoffPublishCommit = 'none'
           PredecessorCloseoutSealCommit = 'none'
           PredecessorCloseoutSealSha256 = 'none'
           CloseoutCommit = '<40-lowercase-hex Round-4 closeout commit>'
           DoneBlob = '<40-lowercase-hex Round-4 Done blob>'
	           DecisionPath = 'Specs/Terminal/TerminalEngineDecision.md'
	           DecisionSha256 = '<64-lowercase-hex Round-4 commit-bound decision digest>'
	           RecoveryArchiveCommit = 'N/A-legacy-round-no-recovery-archive'
	           RecoveryArchivePath = 'N/A-legacy-round-no-recovery-archive'
	           RecoveryArchiveSha256 = 'N/A-legacy-round-no-recovery-archive'
	           HandoffPublishCommit = 'N/A-legacy-round-no-split-closeout'
	           CloseoutSealCommit = 'N/A-legacy-round-no-split-closeout'
	           CloseoutSealPath = 'N/A-legacy-round-no-split-closeout'
	           CloseoutSealSha256 = 'N/A-legacy-round-no-split-closeout'
	           Outcome = 'no-winner'
       },
       [pscustomobject]@{
           Ordinal = 5
           PlanName = 'Terminal_EngineSelectionAndDependencyProof_Round5_WezTermSynchronousCallerOwned_2026-07-31.md'
           PredecessorHandoffPublishCommit = 'N/A-legacy-round-no-split-closeout'
           PredecessorCloseoutSealCommit = 'N/A-legacy-round-no-split-closeout'
           PredecessorCloseoutSealSha256 = 'N/A-legacy-round-no-split-closeout'
           CloseoutCommit = '<40-lowercase-hex Round-5 FDone closeout commit>'
           DoneBlob = '<40-lowercase-hex Round-5 Done blob>'
	           DecisionPath = 'Specs/Terminal/TerminalEngineDecision.Round5.md'
	           DecisionSha256 = '<64-lowercase-hex Round-5 decision digest>'
	           RecoveryArchiveCommit = '<40-lowercase-hex Round-5 FA commit>'
	           RecoveryArchivePath = '<explicit tracked Round-5 RecoveryArchive path>'
	           RecoveryArchiveSha256 = '<64-lowercase-hex Round-5 RecoveryArchive digest>'
	           HandoffPublishCommit = '<40-lowercase-hex Round-5 FP routing commit>'
	           CloseoutSealCommit = '<40-lowercase-hex Round-5 FS durable handoff commit>'
	           CloseoutSealPath = 'Specs/Terminal/TerminalEngineRound5CloseoutSeal.v1.json'
	           CloseoutSealSha256 = '<64-lowercase-hex Round-5 closeout-seal digest>'
	           Outcome = '<no-winner|winner>'
       }
       # Append one exact object for every later round; replace the sole final outcome with winner.
   )
   $successfulGate0HandoffPlanName = '<exact master-named winner-bearing Done basename>'
   $approvedGate0HandoffPublishCommit = '<40-lowercase-hex routing-only FP commit>'
   $approvedGate0CloseoutSealCommit = '<40-lowercase-hex durable FS handoff commit>'
   $approvedGate0CloseoutSealPath = '<fixed winner closeout-seal path>'
   $approvedGate0CloseoutSealSha256 = '<64-lowercase-hex closeout-seal digest>'
   $approvedEvidenceRepositoryCommit = '<40-lowercase-hex TerminalEngine evidence repositoryCommit>'
   $approvedCurrentDecisionSha256 = '<64-lowercase-hex approved current TerminalEngineDecision.md digest>'
   $finalizationManifestRelativePath = '<decision-recorded canonical terminal-engine-evidence.v1.json path>'
   $finalizationManifest = Resolve-Path -LiteralPath $finalizationManifestRelativePath
   $approvedManifestSha256 = '<decision-recorded 64-lowercase-hex sha256>'
   $approvedRunId = '<decision-recorded yyyy-MM-dd_HHmmss runId>'
   $approvedWinnerCandidateId = '<decision-recorded winner candidate ID>'
   $approvedLockSha256 = '<decision-recorded External/terminal-engine.lock.json sha256>'
   $approvedVerificationSetRelativePath = '<decision-recorded canonical five-lane verification-set manifest path>'
   $approvedVerificationSetManifest = Resolve-Path -LiteralPath $approvedVerificationSetRelativePath
   $approvedVerificationSetSha256 = '<decision-recorded five-lane verification-set sha256>'
   $collectionManifests = @(
       '<decision-recorded native x64 Release collection manifest>',
       '<decision-recorded native x64 ASan Debug collection manifest>',
       '<decision-recorded native ARM64 Release collection manifest>'
   )

   $masterWip = Resolve-Path -LiteralPath '.\Specs\Plans\WIP\Terminal_EmbeddedConPtyLibGhosttyVt_IntegrationPlan_2026-07-13.md'
   $masterText = Get-Content -LiteralPath $masterWip.Path -Raw
   $handoffStart = $masterText.IndexOf('- **successfulGate0Handoff:**', [StringComparison]::Ordinal)
   $goalHeadingMatches = [regex]::Matches($masterText, '(?m)^## Goal\r?$', [Text.RegularExpressions.RegexOptions]::CultureInvariant)
   $handoffEnd = if ($handoffStart -ge 0 -and $goalHeadingMatches.Count -eq 1) { $goalHeadingMatches[0].Index } else { -1 }
   if ($goalHeadingMatches.Count -ne 1 -or $handoffStart -lt 0 -or $handoffEnd -le $handoffStart) { throw 'Cannot locate the unique bounded master Gate-0 handoff region.' }
   $handoffRegion = $masterText.Substring($handoffStart, $handoffEnd - $handoffStart)
   function Assert-MasterGate0Token([string]$Label, [string]$Value) {
       $tick = [char]0x60
       $linePrefix = "  - $tick$Label="
       $exactLine = "$linePrefix$Value$tick"
       $matchingLines = @([regex]::Split($handoffRegion, '\r?\n') | Where-Object {
           $_.StartsWith($linePrefix, [StringComparison]::Ordinal)
       })
       if ($matchingLines.Count -ne 1 -or $matchingLines[0] -cne $exactLine) {
           throw "Master Gate-0 token is missing, duplicated, or mismatched: $Label"
       }
   }
   $masterRoundPlanTokens = [regex]::Matches(
       $handoffRegion,
       '(?m)^  - `gate0Round\[(\d+)\]\.plan=[^`\r\n]+`\r?$')
   if ($masterRoundPlanTokens.Count -ne $approvedGate0Lineage.Count) { throw 'Master Gate-0 round count differs from the supplied lineage.' }
   for ($index = 0; $index -lt $approvedGate0Lineage.Count; $index++) {
       $entry = $approvedGate0Lineage[$index]
       if ($entry.Ordinal -ne (4 + $index)) { throw "Gate-0 ordinal gap at index $index." }
       $predecessorPlan = if ($index -eq 0) { 'none' } else { [string]$approvedGate0Lineage[$index - 1].PlanName }
       $predecessorCommit = if ($index -eq 0) { 'none' } else { [string]$approvedGate0Lineage[$index - 1].CloseoutCommit }
       $predecessorBlob = if ($index -eq 0) { 'none' } else { [string]$approvedGate0Lineage[$index - 1].DoneBlob }
       $predecessorPublishCommit = if ($index -eq 0) { 'none' } else { [string]$approvedGate0Lineage[$index - 1].HandoffPublishCommit }
       $predecessorSealCommit = if ($index -eq 0) { 'none' } else { [string]$approvedGate0Lineage[$index - 1].CloseoutSealCommit }
       $predecessorSealSha256 = if ($index -eq 0) { 'none' } else { [string]$approvedGate0Lineage[$index - 1].CloseoutSealSha256 }
       if ([string]$entry.PredecessorHandoffPublishCommit -cne $predecessorPublishCommit -or
           [string]$entry.PredecessorCloseoutSealCommit -cne $predecessorSealCommit -or
           [string]$entry.PredecessorCloseoutSealSha256 -cne $predecessorSealSha256) {
           throw "Gate-0 predecessor FP/FS carry mismatch at index $index."
       }
       Assert-MasterGate0Token "gate0Round[$index].ordinal" ([string]$entry.Ordinal)
       Assert-MasterGate0Token "gate0Round[$index].plan" ([string]$entry.PlanName)
       Assert-MasterGate0Token "gate0Round[$index].predecessorPlan" $predecessorPlan
       Assert-MasterGate0Token "gate0Round[$index].predecessorCloseoutCommit" $predecessorCommit
       Assert-MasterGate0Token "gate0Round[$index].predecessorDoneBlob" $predecessorBlob
       Assert-MasterGate0Token "gate0Round[$index].predecessorHandoffPublishCommit" $predecessorPublishCommit
       Assert-MasterGate0Token "gate0Round[$index].predecessorCloseoutSealCommit" $predecessorSealCommit
       Assert-MasterGate0Token "gate0Round[$index].predecessorCloseoutSealSha256" $predecessorSealSha256
       Assert-MasterGate0Token "gate0Round[$index].closeoutCommit" ([string]$entry.CloseoutCommit)
       Assert-MasterGate0Token "gate0Round[$index].doneBlob" ([string]$entry.DoneBlob)
	       Assert-MasterGate0Token "gate0Round[$index].decisionPath" ([string]$entry.DecisionPath)
	       Assert-MasterGate0Token "gate0Round[$index].decisionSha256" ([string]$entry.DecisionSha256)
	       Assert-MasterGate0Token "gate0Round[$index].recoveryArchiveCommit" ([string]$entry.RecoveryArchiveCommit)
	       Assert-MasterGate0Token "gate0Round[$index].recoveryArchivePath" ([string]$entry.RecoveryArchivePath)
	       Assert-MasterGate0Token "gate0Round[$index].recoveryArchiveSha256" ([string]$entry.RecoveryArchiveSha256)
	       Assert-MasterGate0Token "gate0Round[$index].outcome" ([string]$entry.Outcome)
   }
   Assert-MasterGate0Token 'successfulGate0HandoffPlan' $successfulGate0HandoffPlanName
   Assert-MasterGate0Token 'successfulGate0HandoffCloseoutCommit' ([string]$approvedGate0Lineage[-1].CloseoutCommit)
   Assert-MasterGate0Token 'successfulGate0HandoffDoneBlob' ([string]$approvedGate0Lineage[-1].DoneBlob)
   Assert-MasterGate0Token 'successfulGate0FinalizationManifest' $finalizationManifestRelativePath
   Assert-MasterGate0Token 'successfulGate0FinalizationSha256' $approvedManifestSha256
   Assert-MasterGate0Token 'successfulGate0WinnerCandidateId' $approvedWinnerCandidateId
   Assert-MasterGate0Token 'successfulGate0EngineLockSha256' $approvedLockSha256
   Assert-MasterGate0Token 'successfulGate0VerificationSetManifest' $approvedVerificationSetRelativePath
   Assert-MasterGate0Token 'successfulGate0VerificationSetSha256' $approvedVerificationSetSha256
   $expectedMasterPointer = '- **successfulGate0Handoff:** `{0}`' -f $successfulGate0HandoffPlanName
   if ([regex]::Matches($handoffRegion, [regex]::Escape($expectedMasterPointer)).Count -ne 1) {
       throw 'The bold master successfulGate0Handoff pointer differs from its machine-readable record.'
   }

   if ($approvedGate0Lineage.Count -ne $gate0RoundNames.Count -or
       (@($approvedGate0Lineage.PlanName) -join '|') -cne ($gate0RoundNames -join '|')) {
       throw 'Approved Gate-0 lineage does not equal the master explicit round order.'
   }
   $priorGate0Lineage = @($approvedGate0Lineage | Select-Object -First ($approvedGate0Lineage.Count - 1))
   if ($gate0RoundNames[-1] -cne $successfulGate0HandoffPlanName -or
       $approvedGate0Lineage[-1].Outcome -cne 'winner' -or
       @($priorGate0Lineage | Where-Object { $_.Outcome -cne 'no-winner' }).Count -ne 0) {
       throw 'The successful Gate-0 handoff must be the final round after only no-winner lineage.'
   }
   if ($approvedEvidenceRepositoryCommit -cnotmatch '^[0-9a-f]{40}$') { throw 'TerminalEngine repositoryCommit is not exact 40-lowercase-hex.' }

   function Get-CommitBlobSha256([string]$Commit, [string]$RelativePath) {
       $blobObjectId = (& git rev-parse --verify ("{0}:{1}" -f $Commit, $RelativePath)).Trim()
       if ($LASTEXITCODE -ne 0 -or $blobObjectId -cnotmatch '^[0-9a-f]{40,64}$') { throw "Cannot resolve commit-bound blob: $RelativePath" }
       $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
       $startInfo.FileName = 'git'
       $startInfo.WorkingDirectory = (Get-Location).Path
       $startInfo.UseShellExecute = $false
       $startInfo.RedirectStandardOutput = $true
       $startInfo.RedirectStandardError = $true
       [void]$startInfo.ArgumentList.Add('cat-file')
       [void]$startInfo.ArgumentList.Add('blob')
       [void]$startInfo.ArgumentList.Add($blobObjectId)
       $process = [System.Diagnostics.Process]::Start($startInfo)
       $hasher = [System.Security.Cryptography.SHA256]::Create()
       try { $hashBytes = $hasher.ComputeHash($process.StandardOutput.BaseStream) }
       finally { $hasher.Dispose() }
       $errorText = $process.StandardError.ReadToEnd()
       $process.WaitForExit()
       if ($process.ExitCode -ne 0) { throw "git cat-file failed for $RelativePath`: $errorText" }
       return [Convert]::ToHexString($hashBytes).ToLowerInvariant()
   }

   $gate0DonePaths = @()
   for ($index = 0; $index -lt $approvedGate0Lineage.Count; $index++) {
       $entry = $approvedGate0Lineage[$index]
       foreach ($value in @($entry.CloseoutCommit, $entry.DoneBlob)) {
           if ($value -cnotmatch '^[0-9a-f]{40}$') { throw "Invalid Gate-0 Git identity: $($entry.PlanName)" }
       }
       if ($entry.DecisionSha256 -cnotmatch '^[0-9a-f]{64}$' -or
           [string]::IsNullOrWhiteSpace($entry.DecisionPath)) { throw "Invalid Gate-0 decision identity: $($entry.PlanName)" }
       $relativePath = "Specs/Plans/Done/$($entry.PlanName)"
       $donePath = (Resolve-Path -LiteralPath $relativePath).Path
       $gate0DonePaths += $donePath
       $actualBlob = (& git rev-parse --verify ("{0}:{1}" -f $entry.CloseoutCommit, $relativePath)).Trim()
       if ($LASTEXITCODE -ne 0 -or $actualBlob -cne $entry.DoneBlob) { throw "Gate-0 Done blob/commit mismatch: $($entry.PlanName)" }
       $workingBlob = (& git hash-object -- $donePath).Trim()
       if ($LASTEXITCODE -ne 0 -or $workingBlob -cne $entry.DoneBlob) { throw "Gate-0 Done bytes differ from accepted blob: $($entry.PlanName)" }
	       if ((Get-CommitBlobSha256 -Commit $entry.CloseoutCommit -RelativePath $entry.DecisionPath) -cne $entry.DecisionSha256) {
	           throw "Gate-0 commit-bound decision digest mismatch: $($entry.PlanName)"
	       }
	       if ($entry.PlanName -ceq 'Terminal_EngineSelectionAndDependencyProof_2026-07-22.md') {
	           $legacyArchiveNa = 'N/A-legacy-round-no-recovery-archive'
	           $legacyCloseoutNa = 'N/A-legacy-round-no-split-closeout'
	           if ($entry.RecoveryArchiveCommit -cne $legacyArchiveNa -or
	               $entry.RecoveryArchivePath -cne $legacyArchiveNa -or
	               $entry.RecoveryArchiveSha256 -cne $legacyArchiveNa) {
	               throw 'Round 4 must use the exact legacy RecoveryArchive N/A triplet.'
	           }
	           if ($entry.HandoffPublishCommit -cne $legacyCloseoutNa -or
	               $entry.CloseoutSealCommit -cne $legacyCloseoutNa -or
	               $entry.CloseoutSealPath -cne $legacyCloseoutNa -or
	               $entry.CloseoutSealSha256 -cne $legacyCloseoutNa) {
	               throw 'Round 4 must use the exact legacy split-closeout N/A quartet.'
	           }
	       } else {
	           $archiveFileName = 'TerminalEngineRound{0}RecoveryArchive.v1.json' -f $entry.Ordinal
	           $archiveRunIdPattern = 'round{0}-(?<timestamp>[0-9]{{8}}T[0-9]{{6}}Z)-[0-9a-f]{{16}}' -f $entry.Ordinal
	           $archivePathPattern = '^Specs/TestRuns/TerminalEngineRecovery/' + $archiveRunIdPattern + '/' + [regex]::Escape($archiveFileName) + '$'
	           $archivePathMatch = [regex]::Match([string]$entry.RecoveryArchivePath, $archivePathPattern, [Text.RegularExpressions.RegexOptions]::CultureInvariant)
	           if ($entry.RecoveryArchiveCommit -cnotmatch '^[0-9a-f]{40}$' -or
	               -not $archivePathMatch.Success -or
	               $entry.RecoveryArchiveSha256 -cnotmatch '^[0-9a-f]{64}$' -or
	               $entry.HandoffPublishCommit -cnotmatch '^[0-9a-f]{40}$' -or
	               $entry.CloseoutSealCommit -cnotmatch '^[0-9a-f]{40}$' -or
	               $entry.CloseoutSealPath -cne ("Specs/Terminal/TerminalEngineRound{0}CloseoutSeal.v1.json" -f $entry.Ordinal) -or
	               $entry.CloseoutSealSha256 -cnotmatch '^[0-9a-f]{64}$') {
	               throw "Invalid Gate-0 RecoveryArchive identity: $($entry.PlanName)"
	           }
	           $archiveTimestampText = $archivePathMatch.Groups['timestamp'].Value
	           [datetime]$archiveTimestamp = [datetime]::MinValue
	           $archiveTimestampStyles = [Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal
	           if (-not [datetime]::TryParseExact($archiveTimestampText, "yyyyMMdd'T'HHmmss'Z'", [Globalization.CultureInfo]::InvariantCulture, $archiveTimestampStyles, [ref]$archiveTimestamp) -or
	               $archiveTimestamp.ToString("yyyyMMdd'T'HHmmss'Z'", [Globalization.CultureInfo]::InvariantCulture) -cne $archiveTimestampText) {
	               throw "RecoveryArchive timestamp is not canonical UTC: $($entry.PlanName)"
	           }
	           $repositoryRoot = [IO.Path]::GetFullPath((& git rev-parse --show-toplevel).Trim())
	           if ($LASTEXITCODE -ne 0) { throw 'Cannot resolve the canonical repository root.' }
	           $recoveryRoot = [IO.Path]::GetFullPath((Join-Path $repositoryRoot 'Specs\TestRuns\TerminalEngineRecovery'))
	           $archiveIoPath = [IO.Path]::GetFullPath((Join-Path $repositoryRoot ($entry.RecoveryArchivePath.Replace([char]'/', [IO.Path]::DirectorySeparatorChar))))
	           $recoveryPrefix = $recoveryRoot.TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
	           if (-not $archiveIoPath.StartsWith($recoveryPrefix, [StringComparison]::OrdinalIgnoreCase)) {
	               throw "RecoveryArchive path escapes Specs/TestRuns/TerminalEngineRecovery: $($entry.PlanName)"
	           }
	           $walkPath = $repositoryRoot
	           foreach ($part in $entry.RecoveryArchivePath.Split('/')) {
	               $walkPath = Join-Path $walkPath $part
	               $item = Get-Item -LiteralPath $walkPath -Force -ErrorAction Stop
	               if ($item.Name -cne $part -or ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
	                   throw "RecoveryArchive path is a case alias or crosses a reparse point: $($entry.PlanName)"
	               }
	           }
	           $closeoutParentRow = (& git rev-list --parents -n 1 $entry.CloseoutCommit).Trim()
	           [string[]]$closeoutParentParts = @($closeoutParentRow.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries))
	           if ($LASTEXITCODE -ne 0 -or $closeoutParentParts.Count -ne 2 -or
	               $closeoutParentParts[0] -cne $entry.CloseoutCommit -or
	               $closeoutParentParts[1] -cne $entry.RecoveryArchiveCommit) {
	               throw "Gate-0 FDone closeout must have RecoveryArchive FA as direct parent: $($entry.PlanName)"
	           }
	           $archiveParentRow = (& git rev-list --parents -n 1 $entry.RecoveryArchiveCommit).Trim()
	           [string[]]$archiveParentParts = @($archiveParentRow.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries))
	           if ($LASTEXITCODE -ne 0 -or $archiveParentParts.Count -ne 2 -or
	               $archiveParentParts[0] -cne $entry.RecoveryArchiveCommit -or
	               $archiveParentParts[1] -cnotmatch '^[0-9a-f]{40}$') {
	               throw "Cannot resolve the E parent of RecoveryArchive FA: $($entry.PlanName)"
	           }
	           $decisionCutoffCommit = $archiveParentParts[1]
	           & git merge-base --is-ancestor $entry.RecoveryArchiveCommit HEAD
	           if ($LASTEXITCODE -ne 0) { throw "RecoveryArchive commit is not in the accepted lineage: $($entry.PlanName)" }
	           [string[]]$archiveChanges = @(& git diff-tree --no-commit-id --name-status --no-renames -r $decisionCutoffCommit $entry.RecoveryArchiveCommit)
	           if ($LASTEXITCODE -ne 0 -or $archiveChanges.Count -ne 1 -or
	               $archiveChanges[0] -cne ("A`t{0}" -f $entry.RecoveryArchivePath)) {
	               throw "RecoveryArchive FA is not an archive-only direct child of E: $($entry.PlanName)"
	           }
	           [string[]]$archiveTreeEntry = @(& git ls-tree $entry.RecoveryArchiveCommit -- $entry.RecoveryArchivePath)
	           $archiveTreePattern = '^100644 blob [0-9a-f]{40,64}\t' + [regex]::Escape($entry.RecoveryArchivePath) + '$'
	           if ($LASTEXITCODE -ne 0 -or $archiveTreeEntry.Count -ne 1 -or
	               $archiveTreeEntry[0] -cnotmatch $archiveTreePattern) {
	               throw "RecoveryArchive is not one ordinary 100644 Git blob: $($entry.PlanName)"
	           }
	           $roundWipPath = "Specs/Plans/WIP/$($entry.PlanName)"
	           $roundDonePath = "Specs/Plans/Done/$($entry.PlanName)"
	           [string[]]$doneChanges = @(& git diff-tree --no-commit-id --name-status --no-renames -r $entry.RecoveryArchiveCommit $entry.CloseoutCommit)
	           [string[]]$expectedDoneChanges = @(("A`t{0}" -f $roundDonePath), ("D`t{0}" -f $roundWipPath))
	           if ($LASTEXITCODE -ne 0 -or $doneChanges.Count -ne 2 -or
	               @(Compare-Object -ReferenceObject $expectedDoneChanges -DifferenceObject $doneChanges -CaseSensitive).Count -ne 0) {
	               throw "Gate-0 FDone is not the exact WIP-to-Done-only direct child of FA: $($entry.PlanName)"
	           }
	           [string[]]$doneTreeEntry = @(& git ls-tree $entry.CloseoutCommit -- $roundDonePath)
	           $doneTreePattern = '^100644 blob [0-9a-f]{40,64}\t' + [regex]::Escape($roundDonePath) + '$'
	           if ($LASTEXITCODE -ne 0 -or $doneTreeEntry.Count -ne 1 -or
	               $doneTreeEntry[0] -cnotmatch $doneTreePattern) {
	               throw "Gate-0 Done plan is not one ordinary 100644 Git blob: $($entry.PlanName)"
	           }
	           if ((Get-CommitBlobSha256 -Commit $entry.RecoveryArchiveCommit -RelativePath $entry.RecoveryArchivePath) -cne
	               $entry.RecoveryArchiveSha256) {
	               throw "Gate-0 RecoveryArchive commit/path/digest mismatch: $($entry.PlanName)"
	           }
	           $currentArchiveSha256 = (Get-FileHash -LiteralPath $entry.RecoveryArchivePath -Algorithm SHA256).Hash.ToLowerInvariant()
	           if ($currentArchiveSha256 -cne $entry.RecoveryArchiveSha256) {
	               throw "Current Gate-0 RecoveryArchive bytes differ: $($entry.PlanName)"
	           }
	           $archiveSchemaPath = Resolve-Path -LiteralPath (".\Specs\Terminal\TerminalEngineRound{0}RecoveryArchive.schema.json" -f $entry.Ordinal)
	           $archiveText = [IO.File]::ReadAllText((Resolve-Path -LiteralPath $entry.RecoveryArchivePath).Path, [Text.UTF8Encoding]::new($false, $true))
	           if (-not (Test-Json -Json $archiveText -SchemaFile $archiveSchemaPath.Path -ErrorAction Stop)) {
	               throw "Gate-0 RecoveryArchive fails its ordinal schema: $($entry.PlanName)"
	           }
	           $archive = $archiveText | ConvertFrom-Json -Depth 100
	           if ($archive.recordKind -cne 'terminal-engine-recovery-archive' -or
	               [int]$archive.roundOrdinal -ne [int]$entry.Ordinal -or
	               [string]$archive.decisionCutoffCommit -cne $decisionCutoffCommit) {
	               throw "Gate-0 RecoveryArchive does not bind its round and E cutoff: $($entry.PlanName)"
	           }
	           & .\Tools\Test-TerminalEngineRecoveryArchive.ps1 -ArchivePath $entry.RecoveryArchivePath -RoundOrdinal $entry.Ordinal -ExpectedRawSha256 $entry.RecoveryArchiveSha256 -ExpectedDecisionCutoffCommit $decisionCutoffCommit
	           if ($LASTEXITCODE -ne 0) { throw "Gate-0 RecoveryArchive semantic verification failed: $($entry.PlanName)" }
	           if ($entry.Outcome -ceq 'no-winner') {
	               if ($index + 1 -ge $approvedGate0Lineage.Count) {
	                   throw 'A no-winner Gate-0 round must have an explicit successor before product admission.'
	               }
	               $expectedRoute = "Specs/Plans/WIP/$($approvedGate0Lineage[$index + 1].PlanName)"
	               $successorArchivePath = [string]$approvedGate0Lineage[$index + 1].RecoveryArchivePath
	               $successorArchiveSha256 = [string]$approvedGate0Lineage[$index + 1].RecoveryArchiveSha256
	           } elseif ($entry.Outcome -ceq 'winner') {
	               if ($index -ne $approvedGate0Lineage.Count - 1) {
	                   throw 'Only the final Gate-0 round may be the winner.'
	               }
	               $expectedRoute = 'Specs/Plans/WIP/Terminal_CoreConPtyRuntimeAndModel_2026-07-22.md'
	               $successorArchivePath = 'N/A-winner-routes-product-plan-2'
	               $successorArchiveSha256 = 'N/A-winner-routes-product-plan-2'
	           } else {
	               throw "Invalid Gate-0 outcome: $($entry.PlanName)"
	           }
	           $publishedCloseoutVerifier = ".\Tools\Test-TerminalEngineRound{0}PublishedCloseout.ps1" -f $entry.Ordinal
	           if (-not (Test-Path -LiteralPath $publishedCloseoutVerifier -PathType Leaf)) {
	               throw "Missing ordinal published-closeout verifier: $($entry.PlanName)"
	           }
	           & $publishedCloseoutVerifier `
	               -Mode RetrospectiveCarry `
	               -DecisionCutoffCommit $decisionCutoffCommit `
	               -RecoveryArchiveCommit $entry.RecoveryArchiveCommit `
	               -DoneCommit $entry.CloseoutCommit `
	               -HandoffPublishCommit $entry.HandoffPublishCommit `
	               -CloseoutSealCommit $entry.CloseoutSealCommit `
	               -WipPlanPath $roundWipPath `
	               -DonePlanPath $roundDonePath `
	               -RecoveryArchivePath $entry.RecoveryArchivePath `
	               -RecoveryArchiveSha256 $entry.RecoveryArchiveSha256 `
	               -CloseoutSealPath $entry.CloseoutSealPath `
	               -CloseoutSealSha256 $entry.CloseoutSealSha256 `
	               -Outcome $entry.Outcome `
	               -ExpectedRoute $expectedRoute `
	               -SuccessorRecoveryArchivePath $successorArchivePath `
	               -SuccessorRecoveryArchiveSha256 $successorArchiveSha256
	           if ($LASTEXITCODE -ne 0) { throw "Gate-0 published closeout verification failed: $($entry.PlanName)" }
	       }
	   }
   if ($approvedGate0HandoffPublishCommit -cne [string]$approvedGate0Lineage[-1].HandoffPublishCommit -or
       $approvedGate0CloseoutSealCommit -cne [string]$approvedGate0Lineage[-1].CloseoutSealCommit -or
       $approvedGate0CloseoutSealPath -cne [string]$approvedGate0Lineage[-1].CloseoutSealPath -or
       $approvedGate0CloseoutSealSha256 -cne [string]$approvedGate0Lineage[-1].CloseoutSealSha256) {
       throw 'Final winner FP/FS identities differ from the explicit Gate-0 lineage.'
   }
   if ($approvedGate0HandoffPublishCommit -cnotmatch '^[0-9a-f]{40}$') {
       throw 'Invalid routing-only Gate-0 FPublish commit.'
   }
   $publishParentRow = (& git rev-list --parents -n 1 $approvedGate0HandoffPublishCommit).Trim()
   [string[]]$publishParentParts = @($publishParentRow.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries))
   if ($LASTEXITCODE -ne 0 -or $publishParentParts.Count -ne 2 -or
       $publishParentParts[0] -cne $approvedGate0HandoffPublishCommit -or
       $publishParentParts[1] -cne [string]$approvedGate0Lineage[-1].CloseoutCommit) {
       throw 'Winner FPublish must have the final round FDone closeout as its direct parent.'
   }
   $publishParent = $publishParentParts[1]
   & git merge-base --is-ancestor $approvedGate0HandoffPublishCommit HEAD
   if ($LASTEXITCODE -ne 0) { throw 'Winner FPublish is not in the accepted lineage.' }
   $masterRelativePath = 'Specs/Plans/WIP/Terminal_EmbeddedConPtyLibGhosttyVt_IntegrationPlan_2026-07-13.md'
   $wipReadmeRelativePath = 'Specs/Plans/WIP/README.md'
   [string[]]$publishChanges = @(& git diff-tree --no-commit-id --name-status --no-renames -r $publishParent $approvedGate0HandoffPublishCommit)
   [string[]]$expectedPublishChanges = @(("M`t{0}" -f $masterRelativePath), ("M`t{0}" -f $wipReadmeRelativePath))
   if ($LASTEXITCODE -ne 0 -or $publishChanges.Count -ne 2 -or
       @(Compare-Object -ReferenceObject $expectedPublishChanges -DifferenceObject $publishChanges -CaseSensitive).Count -ne 0) {
       throw 'Winner FPublish must be the routing-only direct child of FDone.'
   }
   $sealParentRow = (& git rev-list --parents -n 1 $approvedGate0CloseoutSealCommit).Trim()
   [string[]]$sealParentParts = @($sealParentRow.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries))
   if ($LASTEXITCODE -ne 0 -or $sealParentParts.Count -ne 2 -or
       $sealParentParts[0] -cne $approvedGate0CloseoutSealCommit -or
       $sealParentParts[1] -cne $approvedGate0HandoffPublishCommit) {
       throw 'Winner FSeal must have the routing-only FPublish commit as its sole parent.'
   }
   & git merge-base --is-ancestor $approvedGate0CloseoutSealCommit HEAD
   if ($LASTEXITCODE -ne 0) { throw 'Winner FSeal is not in the accepted lineage.' }
   $winnerSealPath = [string]$approvedGate0Lineage[-1].CloseoutSealPath
   [string[]]$sealChanges = @(& git diff-tree --no-commit-id --name-status --no-renames -r $approvedGate0HandoffPublishCommit $approvedGate0CloseoutSealCommit)
   if ($LASTEXITCODE -ne 0 -or $sealChanges.Count -ne 1 -or
       $sealChanges[0] -cne ("A`t{0}" -f $winnerSealPath) -or
       (Get-CommitBlobSha256 -Commit $approvedGate0CloseoutSealCommit -RelativePath $winnerSealPath) -cne $approvedGate0CloseoutSealSha256) {
       throw 'Winner FSeal is not the exact one-file durable handoff.'
   }
   if (-not (Test-Path -LiteralPath $winnerSealPath -PathType Leaf) -or
       (Get-FileHash -LiteralPath $winnerSealPath -Algorithm SHA256).Hash.ToLowerInvariant() -cne $approvedGate0CloseoutSealSha256) {
       throw 'Current winner CloseoutSeal bytes differ from the accepted FS blob.'
   }
   [string[]]$masterAtPublishLines = @(& git show --no-textconv ("{0}:{1}" -f $approvedGate0HandoffPublishCommit, $masterRelativePath))
   if ($LASTEXITCODE -ne 0) { throw 'Cannot read the master blob at FPublish.' }
   $masterAtPublishText = $masterAtPublishLines -join "`n"
   $publishedHandoffStart = $masterAtPublishText.IndexOf('- **successfulGate0Handoff:**', [StringComparison]::Ordinal)
   $publishedGoalHeadingMatches = [regex]::Matches($masterAtPublishText, '(?m)^## Goal\r?$', [Text.RegularExpressions.RegexOptions]::CultureInvariant)
   $publishedHandoffEnd = if ($publishedHandoffStart -ge 0 -and $publishedGoalHeadingMatches.Count -eq 1) { $publishedGoalHeadingMatches[0].Index } else { -1 }
   if ($publishedGoalHeadingMatches.Count -ne 1 -or $publishedHandoffStart -lt 0 -or $publishedHandoffEnd -le $publishedHandoffStart) {
       throw 'Cannot locate the unique FPublish Gate-0 handoff region.'
   }
   $publishedHandoffRegion = $masterAtPublishText.Substring($publishedHandoffStart, $publishedHandoffEnd - $publishedHandoffStart)
   $normalizedPublishedHandoff = [regex]::Replace($publishedHandoffRegion, '\r\n?', "`n")
   $normalizedCurrentHandoff = [regex]::Replace($handoffRegion, '\r\n?', "`n")
   if ($normalizedPublishedHandoff -cne $normalizedCurrentHandoff) {
       throw 'Current bounded Gate-0 handoff region differs from the accepted FPublish.'
   }
   [string[]]$readmeAtPublishLines = @(& git show --no-textconv ("{0}:{1}" -f $approvedGate0HandoffPublishCommit, $wipReadmeRelativePath))
   if ($LASTEXITCODE -ne 0) { throw 'Cannot read WIP README at FPublish.' }
   $expectedWinnerRouteLine = '- **terminalGate0Route:** `Terminal_CoreConPtyRuntimeAndModel_2026-07-22.md`'
   if (@($readmeAtPublishLines | Where-Object { $_ -ceq $expectedWinnerRouteLine }).Count -ne 1) {
       throw 'FPublish WIP README does not contain exactly the reviewed winner route to Plan 2.'
   }
   [byte[]]$successMarkerBytes = [Text.Encoding]::ASCII.GetBytes('- **successfulGate0Handoff:** `TERMINAL-GATE0-SUCCESS`')
   function Get-ExactByteSequenceCount([string]$LiteralPath, [byte[]]$Needle) {
       [byte[]]$bytes = [IO.File]::ReadAllBytes($LiteralPath)
       if ($Needle.Length -eq 0 -or $bytes.Length -lt $Needle.Length) { return 0 }
       $count = 0
       for ($index = 0; $index -le $bytes.Length - $Needle.Length; $index++) {
           $equal = $true
           for ($offset = 0; $offset -lt $Needle.Length; $offset++) {
               if ($bytes[$index + $offset] -ne $Needle[$offset]) { $equal = $false; break }
           }
           if ($equal) { $count++; $index += $Needle.Length - 1 }
       }
       return $count
   }
   $successRecords = @(
       for ($index = 0; $index -lt $approvedGate0Lineage.Count; $index++) {
           [pscustomobject]@{
               PlanName = [string]$approvedGate0Lineage[$index].PlanName
               Outcome = [string]$approvedGate0Lineage[$index].Outcome
               Count = Get-ExactByteSequenceCount -LiteralPath $gate0DonePaths[$index] -Needle $successMarkerBytes
           }
       }
   )
   foreach ($record in $successRecords) {
       $expectedCount = if ($record.Outcome -ceq 'winner') { 1 } elseif ($record.Outcome -ceq 'no-winner') { 0 } else { throw "Invalid Gate-0 outcome: $($record.PlanName)" }
       if ($record.Count -ne $expectedCount) { throw "Gate-0 success-marker occurrence mismatch: $($record.PlanName)" }
   }
   $successTotal = ($successRecords | Measure-Object -Property Count -Sum).Sum
   $successOwner = @($successRecords | Where-Object { $_.Count -eq 1 })
   if ($successTotal -ne 1 -or $successOwner.Count -ne 1 -or
       $successOwner[0].PlanName -cne $successfulGate0HandoffPlanName) {
       throw 'Gate-0 Done bytes do not contain exactly one marker in the master-named sole winner.'
   }
   $predecessorPlan = Resolve-Path -LiteralPath (".\Specs\Plans\Done\{0}" -f $successfulGate0HandoffPlanName)
   $approvedPredecessorCloseoutCommit = [string]$approvedGate0Lineage[-1].CloseoutCommit
   $approvedPredecessorBlobObjectId = [string]$approvedGate0Lineage[-1].DoneBlob

   foreach ($value in @($approvedCurrentDecisionSha256, $approvedManifestSha256, $approvedLockSha256, $approvedVerificationSetSha256)) {
       if ($value -cnotmatch '^[0-9a-f]{64}$') { throw 'An approved digest is not exact lowercase SHA-256.' }
   }
   if ($approvedRunId -cnotmatch '^\d{4}-\d{2}-\d{2}_\d{6}$') { throw 'Approved TerminalEngine runId is invalid.' }

   $currentDecisionSha256 = (Get-FileHash -LiteralPath '.\Specs\Terminal\TerminalEngineDecision.md' -Algorithm SHA256).Hash.ToLowerInvariant()
   if ($currentDecisionSha256 -cne $approvedCurrentDecisionSha256) { throw 'Current terminal-engine decision digest mismatch.' }

   $manifestPath = $finalizationManifest.Path
   if ((Split-Path -Leaf (Split-Path -Parent $manifestPath)) -cne $approvedRunId) { throw 'Finalization directory/runId mismatch.' }
   $manifestSha256 = (Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash.ToLowerInvariant()
   if ($manifestSha256 -cne $approvedManifestSha256) { throw 'Approved finalization manifest digest mismatch.' }
   $companionPath = Join-Path (Split-Path -Parent $manifestPath) 'terminal-engine-evidence.v1.sha256'
   $companionText = [System.Text.Encoding]::ASCII.GetString([System.IO.File]::ReadAllBytes($companionPath))
   if ($companionText -cne ($approvedManifestSha256 + "`n")) { throw 'Finalization companion is not exact lowercase digest plus LF.' }
   $manifestText = [System.IO.File]::ReadAllText($manifestPath, [System.Text.UTF8Encoding]::new($false, $true))
   $schemaPath = (Resolve-Path -LiteralPath '.\Specs\Terminal\TerminalEngineEvidence.schema.json').Path
   if (-not (Test-Json -Json $manifestText -SchemaFile $schemaPath -ErrorAction Stop)) { throw 'Finalization manifest fails its v1 schema.' }
   $manifest = $manifestText | ConvertFrom-Json -Depth 100
   if ($manifest.recordKind -cne 'finalization' -or $manifest.runId -cne $approvedRunId) { throw 'Wrong finalization record identity.' }
   if ($manifest.repositoryCommit -cne $approvedEvidenceRepositoryCommit) { throw 'TerminalEngine evidence repositoryCommit mismatch.' }
   if ($manifest.selection.status -cne 'winner' -or $manifest.selection.winnerCandidateId -cne $approvedWinnerCandidateId) { throw 'Approved winner mismatch.' }
   $revalidationPath = .\Tools\Evaluate-TerminalEngines.ps1 -Mode Finalize -Candidate All -EvidenceManifest $collectionManifests -EvidenceRoot .\Specs\TestRuns -RequireAllEvidence -PassThruManifest
   if (-not (Test-Path -LiteralPath $revalidationPath -PathType Leaf)) { throw 'TerminalEngine semantic revalidation returned no exact manifest.' }
   $revalidationText = [System.IO.File]::ReadAllText((Resolve-Path -LiteralPath $revalidationPath).Path, [System.Text.UTF8Encoding]::new($false, $true))
   if (-not (Test-Json -Json $revalidationText -SchemaFile $schemaPath -ErrorAction Stop)) { throw 'Revalidation record fails its v1 schema.' }
   $revalidation = $revalidationText | ConvertFrom-Json -Depth 100
   if ($revalidation.selection.status -cne 'winner' -or $revalidation.selection.winnerCandidateId -cne $approvedWinnerCandidateId) { throw 'Semantic revalidation changed the winner.' }
   foreach ($propertyName in @('repositoryCommit', 'schemaSha256')) {
       if ($revalidation.$propertyName -cne $manifest.$propertyName) { throw "Semantic revalidation changed $propertyName." }
   }
   foreach ($propertyName in @('version', 'sourceSha256')) {
       if ($revalidation.harness.$propertyName -cne $manifest.harness.$propertyName) { throw "Semantic revalidation changed harness.$propertyName." }
   }
   foreach ($propertyName in @('requirementsSha256', 'scoreAnchorsSha256', 'fixtureCorpusSha256', 'measurementProtocolSha256', 'candidateSetSha256')) {
       if ($revalidation.frozen.$propertyName -cne $manifest.frozen.$propertyName) { throw "Semantic revalidation changed frozen.$propertyName." }
   }
   $lockSha256 = (Get-FileHash -LiteralPath '.\External\terminal-engine.lock.json' -Algorithm SHA256).Hash.ToLowerInvariant()
   if ($lockSha256 -cne $approvedLockSha256) { throw 'Terminal engine lock digest mismatch.' }
   $verificationSetSha256 = (Get-FileHash -LiteralPath $approvedVerificationSetManifest.Path -Algorithm SHA256).Hash.ToLowerInvariant()
   if ($verificationSetSha256 -cne $approvedVerificationSetSha256) { throw 'Five-lane verification-set digest mismatch.' }
   $successfulDoneText = Get-Content -LiteralPath $predecessorPlan.Path -Raw
   $successfulDoneTokens = [ordered]@{
       successfulGate0FinalizationManifest = $finalizationManifestRelativePath
       successfulGate0FinalizationSha256 = $approvedManifestSha256
       successfulGate0WinnerCandidateId = $approvedWinnerCandidateId
       successfulGate0EngineLockSha256 = $approvedLockSha256
       successfulGate0VerificationSetManifest = $approvedVerificationSetRelativePath
       successfulGate0VerificationSetSha256 = $approvedVerificationSetSha256
   }
   foreach ($token in $successfulDoneTokens.GetEnumerator()) {
       $label = "$($token.Key)="
       $labelPattern = '(?m)^[ \t]*-[ \t]+`' + [regex]::Escape($label) + '[^`\r\n]*`[ \t]*\r?$'
       $exactPattern = '(?m)^[ \t]*-[ \t]+`' + [regex]::Escape("$label$($token.Value)") + '`[ \t]*\r?$'
       if ([regex]::Matches($successfulDoneText, $labelPattern).Count -ne 1 -or
           [regex]::Matches($successfulDoneText, $exactPattern).Count -ne 1) {
           throw "Successful Gate-0 Done token mismatch: $($token.Key)"
       }
   }
   ```

2. Create the successful test-enabled Release Core Commands evidence only with `Run-TerminalCommandsEvidence.ps1 -Scenario CoreModel -Platform x64 -Configuration Release -RepositoryCommit <exact-commit> -EvidenceRoot .\Specs\TestRuns -PassThruRunPath`; retain its one exact returned `Specs/TestRuns/<MachineHash>/Commands/<RunId>/` path and validate it with `Tools/Test-TestRunArchive.ps1 -RunPath <that-path>`. Its canonical `results.json`, `trace.txt`, runner-created content-free `run-all-tests-results.json`, and required `perf/perf_metrics.jsonl` must show `terminal_perf_core_model_baseline` passed and contain only the master's content-free fields/counters plus repository/engine-lock identity. Also retain the exact x64 Debug, x64 Release, x64 `ASan Debug`, native ARM64 Debug, and native ARM64 Release `TerminalNative` manifest/companion pairs emitted above. Record every path and lowercase SHA-256 in this plan and rerun `Tools/Verify-TerminalNativeEvidence.ps1` with the recorded repository/lock/slice values for all five. Do not move the plan with generic `env.txt`, console output, implicit latest discovery, or only a local `.build` artifact.

3. Merge the implemented lasting contract into `Specs/Terminal/Terminal_EmbeddedPane.md`: new-IID/not-Viewer decision, `builtin/terminal` factory/configuration/child-HWND/callback/module-quiet ABI, plugin-owned atomic per-`PhysicalHostKey` view reservation and host rollback contract, the exact disable/re-enable and irreversible normal-refresh state machine with its two nonfatal blocked outcomes and process-shutdown supersession, plugin-only implementation boundary, selected adapter/runtime ownership and validation, ConPTY I/O/queueing, the immutable one-read paste identity/storage/supersession/revocation/final-admission/scrub transaction including disable/refresh continuity, the OSC 52 identity/ask-cap/revocation/effective-deny/final-write/scrub transaction, exact explicit Unicode launch flags plus canonical environment construction/caps/failure behavior, the exact WSL five-token/default-user/default-shell launch and permanent terminal-only/`RequestedUnverified` boundary with no injection, semantic epoch, capability, or profile memory, exact non-Kitty caps and aggregate-overflow outcomes, snapshot algorithm, UIA/copy token accounting and exact failure results, deterministic resource/LRU/retry-pass policy, and every retirement/final-snapshot outcome. Update `Specs/Plugins/Plugins_PluginAPI.md` for the shared lifecycle helper's distinct normal-refresh versus process-shutdown policies, and `Specs/Plugins/Plugins_ViewerPlugins.md` to state why Terminal is not a viewer and the applicable generic factory/quiet reuse; update installer specs with exact `Plugins\Terminal.dll` and private `Plugins\TerminalRuntime\` centralized staging; update testing specs/archive docs with the ABI/module/core/view-admission/refresh/launch-recorder cases and evidence. Update `Specs/Core/Core_SharedHelpers.md` for every shared helper added or extended. Normative behavior may not remain only in this plan.

4. Resolve every STOP condition and review the diff for generic engine names or stale Gate-0 assumptions. Change this file's `**State:**` to `Done`, then move it exactly:

   ```powershell
   git mv -- 'Specs/Plans/WIP/Terminal_CoreConPtyRuntimeAndModel_2026-07-22.md' 'Specs/Plans/Done/Terminal_CoreConPtyRuntimeAndModel_2026-07-22.md'
   ```

5. Update `Specs/Plans/WIP/README.md` N3 to route next to `Specs/Plans/WIP/Terminal_RendererControlAndAccessibility_2026-07-22.md`, naming the plan-2 Done path, all five exact `TerminalNative` manifest/digest identities, and the exact Commands archive/run ID as predecessor evidence. Plan 3 must rerun the applicable engine-lock, native-manifest, and core-evidence verification before implementation.

## STOP conditions

- Gate 0 contract changes or is insufficient for exact ABI/snapshot/resource enforcement.
- `IViewer` must be modified/overloaded, terminal-specific runtime code must enter the EXE/Common, `Terminal.dll` cannot remain pinned through quiet, or a source/static winner cannot be contained in the plugin DLL.
- Any UI path joins, blocks, or destroys a worker-owning object.
- PTY bytes must be dropped, queue/callback must block, protocol order/responses cannot be preserved, aggregate-effect overflow cannot reject/invalidate exactly as specified, or a release pass can starve a fitting frozen waiter or create concurrent retry owners.
- Paste would require a second clipboard read, classify/preview/encode mutable clipboard data, omit any frozen identity/policy/input-mode/lifecycle generation, retain more than one prompt per session or exceed the fixed 8 MiB/session and 32 MiB/service pending-storage caps, authorize without one final current-mode/cap/descriptor/queue atomic check, fail to scrub, or let generic disable/normal refresh revoke or reject safe/pending unsafe paste for an existing live frozen-config view.
- OSC 52 cannot keep payloads plugin-owned and opaque, enforce one outstanding ask/session and eight/service, install revocation before every policy/disable/refresh/teardown publication, behave as effective deny while disabled, linearly order the sole clipboard replacement under the same gate, consume tokens once, or securely scrub every disposition.
- A launch would use an inherited/null/ANSI environment, omit either exact process-creation flag, accept malformed hidden/ordinary entries, truncate command line/environment, or exceed either inclusive UTF-16 cap.
- Any accepted ABI/profile/location identity would be truncated, shortened, aliased, or hashed; a Windows-shell command line would repeat the canonical executable instead of using the fixed basename argv0 with exact nonnull `lpApplicationName`; or WSL identity and five-token command-line representability could not remain separately bounded with zero-process +1 rejection.
- Close confirmation, forced teardown, chooser/plugin removal, and natural-exit quiet would use independent ownership flags; a pending confirmation could authorize a late second path; or the generic host could not linearize every callback/forced route through one `{instanceId,hostViewGeneration}` removal claim.
- A PowerShell semantic-protocol failure could accept another handshake/live rekey in the same incarnation, or cmd/WSL could acquire an ongoing semantic epoch.
- A v1 WSL launch would add or require `--exec`, `--user`, `-e`, a wrapper, another process, a bootstrap or `WSLENV`/environment/PTY/profile/staged-script injection; alter the configured default user/shell; transition out of permanent `RequestedUnverified`; expose history/follow/insertion/activity/authenticated-cwd/profile-memory capability; or make terminal usability depend on integration. STOP and require a separately approved WSL-integration design rather than weakening the frozen five-token contract.
- The host would read/enforce `maxTabsPerPane`, expose/zoom/select a tab before Open succeeds, or the plugin cannot reserve/release a per-complete-host view slot atomically and return the exact quota HRESULT without child/process state.
- Five-second quiet needs detach/unload or process-exit mapped retention of reachable work, the generic manager cannot keep its target/pump alive while polling nonblockingly, or engine/scrollback/image memory cannot be bounded pre-allocation.
- Normal disable would retire a module/administration object, refresh would force-close a live terminal, cancel after admission, abandon configuration settlement, use fatal/process-retention policy, reopen the old generation, map a replacement before old-handle reset, or permit concurrent module/service generations.
- Baseline review finds unreconciled material drift; the explicit ordered Gate-0
  lineage, sole successful handoff, approved finalization manifest/digest/run ID/
  winner/lock/verification set cannot be reproduced and tied to the exact
  predecessor evidence; or a required `TerminalNative` manifest/schema/companion/
  artifact identity cannot be generated and independently verified.
