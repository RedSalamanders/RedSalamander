# Plan 003: Implement the ConPTY, Ghostty model, renderer, input, and accessibility core

> **Status — Completed 2026-08-03.** The v1 renderer, input, IME, Kitty graphics, selection, clipboard, and immutable
> UI Automation implementation is complete. Durable behavior is in `../../Terminal/Terminal_EmbeddedPlugin.md`;
> advanced mutation/history APIs are explicit non-goals. Later WIP/checkpoint language below is historical.

## First-implementation outcome (2026-08-01)

**ADVANCED CORE SLICE COMPLETE.** `Terminal.dll` owns ConPTY
creation, the shell process and pipes, a tracked reader `std::jthread`,
serialized Ghostty ingestion, a reusable two-phase structured render state,
DirectWrite/Direct2D styled-cell presentation, terminal-mode-aware Ghostty key
encoding, bounded asynchronous input, safe/bracketed paste, Ghostty selection
and copy, application mouse reporting/local scroll, basic IME composition and
commit, grid resize, and pinned asynchronous quiet-point teardown. The host
owns neither a Ghostty handle nor terminal rendering, clipboard policy, or VT
input synthesis.
Deterministic lifecycle coverage opens a real ConPTY terminal, verifies the
plugin-owned child and Running state, and verifies immediate child removal.
The live DLL contract suite additionally proves styled-cell extraction and
application-cursor-mode key changes against the staged private runtime.

The advanced slice also implements bounded direct Kitty image presentation in
all raw formats plus WIC-decoded PNG, the nonblocking unsafe-paste confirmation
overlay, tracked scrollback selection/autoscroll, capture-safe application
mouse release, and a passive immutable UI Automation Document/Text provider
with retirement-safe DLL retention. Complete IME/AltGr/dead-key/surrogate and
input matrices, device-loss/DPI/WARP matrices, high
contrast, and complete UIA selection/caret/geometry/scroll/quota behavior
remain mandatory release gates. The plan therefore remains WIP.

The 2026-08-03 hyperlink slice resolves OSC 8 metadata through the required
private Ghostty export, enforces an 8 KiB strict-UTF-8 bound and safe-scheme
allowlist, and gates Ctrl+click activation behind a localized
`disabled`/`ask`/`open` policy owned entirely by `Terminal.dll`.

The same hardening pass completes normalized OSC 52 text writes. The Ghostty
effect callback accepts only one bounded standard `text/plain` request, copies
borrowed bytes, and posts an opaque payload; UI-thread strict conversion and
clipboard mutation are gated by plugin-owned deny/ask/allow policy. Payloads
are scrubbed and the HWND payload registry is initialized/drained across create
and teardown.

> **Executor instructions**: Execute only after Plan 002 passes its boundary,
> loader, lifecycle, and package gates. Keep all files in this plan inside
> `Plugins/Terminal` except the test/project/spec registration explicitly named
> below. Run every verification and honor every STOP condition. Update
> `Specs/Plans/WIP/Terminal_FirstImplementationExecutionIndex_2026-08-01.md` when done.
>
> **Drift check (run first)**:
> `git diff --stat aac3ed260..HEAD -- Plugins/Terminal Common/PlugInterfaces/Terminal.h Common/DxUi Common/WindowMessages.h Common/Helpers.h Tests/TerminalTests Tests/DxUiTests RedSalamanderMonitor/ColorTextView* Specs/Plans/WIP Specs/Testing Tools/TestRunPlan.ps1`
> Expected Plan 002 changes are allowed. STOP if `ITerminal` ownership or
> size-based record contract,
> the winner lock, runtime function table, callback drain, or package root has
> changed without an updated reviewed contract.

## Status

- **Execution state**: ADVANCED CORE SLICE COMPLETE; renderer/input/UIA hardening gates open
- **Priority**: P1
- **Effort**: L (30-45 owner-days)
- **Risk**: HIGH
- **Depends on**: `Specs/Plans/WIP/Terminal_PluginRuntimeBoundary_2026-08-01.md`
- **Category**: migration
- **Planned at**: commit `aac3ed260`, 2026-08-01

## Why this matters

The plugin boundary is useful only if the terminal session, engine mutation,
render snapshots, GPU resources, input translation, accessibility objects, and
teardown all remain behind it without blocking the UI or retaining borrowed
Ghostty data. This plan implements that core in layers and makes resource and
lifetime failures deterministic before host-pane UX is added.

## Current state and frozen constraints

- Plan 002 provides `Terminal.dll`, `TerminalService`, the exact private runtime
  loader/function table, a diagnostic child HWND, and the public `ITerminal` ABI.
- `Terminal.dll` owns one lazy service/runtime. Each successful public Open owns
  one direct child HWND and at most one session; the host never destroys or
  renders that child.
- `maxTabsPerPane` is closed to `1..32`. `ITerminal::Open` validates all copied
  inputs and atomically reserves a service slot before creating child/process
  state. A denied 33rd returns `HRESULT_FROM_WIN32(ERROR_QUOTA_EXCEEDED)` and
  creates nothing.
- Text/grid/scrollback/publication/UIA/copy allocations are hard-capped at
  64 MiB per session and 512 MiB process-wide. Old immutable publications stay
  charged until their last reference drops.
- Kitty/image allocations are independently capped at 64 MiB per session,
  256 MiB process-wide CPU/GPU, 16 million pixels per decoded image, and one
  conversion generation. Checked transient allocations count before admission.
- Ghostty render/Kitty memory is borrowed. No pointer or handle may survive the
  session lock or the next mutating call.
- Shutdown order is stop producers -> cancel/close ConPTY and queues -> drain
  posts/callbacks -> destroy engine/render objects -> clear global callbacks ->
  release runtime -> allow module unload. UI code never waits or joins.
- Natural/explicit/application close arbitration, two-second process escalation,
  five-second operational quiet, and one 250 ms final-snapshot retry are frozen
  by the master/core WIP plans.

## Commands you will need

| Purpose | Command | Expected on success |
|---|---|---|
| Plugin build | `.\build.ps1 -ProjectName Terminal -Configuration Debug -Platform x64` | exit 0 |
| Native core suite | `.\.build\x64\Debug\TerminalTests.exe --suite core-runtime` | all cases pass |
| WARP renderer suite | `.\.build\x64\Debug\TerminalTests.exe --suite renderer --adapter warp` | all cases pass |
| Input/UIA suite | `.\.build\x64\Debug\TerminalTests.exe --suite input-uia` | all cases pass |
| ASan lane | `.\build.ps1 -ProjectName TerminalTests -Configuration 'ASan Debug' -Platform x64; .\.build\x64\ASanDebug\TerminalTests.exe --suite stress` | exit 0; no sanitizer finding; use actual mapped output name from build if different |
| Full gate | `.\Tools\Run-AllTests.ps1 -Suite Full` | exit 0 |

Use the repository's test-run/evidence scripts when recording retained Release
results; never invent a second archive format.

## Scope

**In scope**:

- `Plugins/Terminal/TerminalService.h/.cpp`
- `Plugins/Terminal/TerminalPlugin.h/.cpp`
- `Plugins/Terminal/TerminalSession.h/.cpp` (new)
- `Plugins/Terminal/ConPtySession.h/.cpp` (new)
- `Plugins/Terminal/TerminalReaper.h/.cpp` (new)
- `Plugins/Terminal/TerminalEngineSession.h/.cpp` (new)
- `Plugins/Terminal/TerminalGhosttyAllocator.h/.cpp` (new)
- `Plugins/Terminal/TerminalEffectsQueue.h/.cpp` (new)
- `Plugins/Terminal/TerminalPngDecoder.h/.cpp` (new)
- `Plugins/Terminal/TerminalResourceLedger.h/.cpp` (new)
- `Plugins/Terminal/TerminalSnapshot.h/.cpp` (new)
- `Plugins/Terminal/TerminalImageCache.h/.cpp` (new)
- `Plugins/Terminal/TerminalRenderer.h/.cpp` (new)
- `Plugins/Terminal/TerminalControl.h/.cpp` (new)
- `Plugins/Terminal/TerminalInput.h/.cpp` (new)
- `Plugins/Terminal/TerminalUiaProvider.h/.cpp` (new)
- `Plugins/Terminal/TerminalMetrics.h/.cpp` (new)
- `Plugins/Terminal/TerminalResources.rc` and terminal satellite resources
- `Plugins/Terminal/Terminal.vcxproj` and `.filters`
- `Tests/TerminalTests/` native tests and fixtures
- `Tests/DxUiTests/` only for reusable native-control/UIA test infrastructure
- `Tools/TestRunPlan.ps1` for explicit Terminal native/evidence registration
- `Tools/Run-TerminalCommandsEvidence.ps1` (new if absent)
- `Tools/Verify-TerminalNativeEvidence.ps1` (new if absent)
- `Specs/Testing/TerminalPerfBudgets.json5` only after a measured Release baseline
- `Specs/Plans/WIP/Terminal_CoreConPtyRuntimeAndModel_2026-07-22.md`
- `Specs/Plans/WIP/Terminal_RendererControlAndAccessibility_2026-07-22.md`
- `Specs/Plans/WIP/Terminal_EmbeddedConPtyLibGhosttyVt_IntegrationPlan_2026-07-13.md`

**Out of scope**:

- pane tabs, Open Command Shell routing, folder Edit routing, profile chooser,
  persistent history, navigation follow, full Preferences UI, and final package
  closeout;
- host-side escape sequence generation, ConPTY ownership, rendering, UIA, or
  terminal state;
- file/shared-memory Kitty transmission, full Unicode BiDi paragraph reorder,
  arbitrary shell integration injection, or remote terminals in v1;
- changing public `ITerminal` methods, ownership, or required record prefixes
  without returning to Plan 002 review.

## Git workflow

- Commit by layer: resource/queue/session, snapshot/renderer, input/UIA, then
  lifecycle/stress/evidence. Keep each commit buildable.
- Do not push without instruction. Do not commit live shell output or content-
  bearing traces; evidence must use the repository privacy schema.

## Steps

### Step 1: Implement resource ledgers and bounded queues first

Before opening a process or engine, implement reusable plugin-private RAII
admission tokens:

- aggregate Ghostty allocator current/peak/hard cap with reserve-before-
  allocate, rollback on failure, resize/remap delta accounting, and checked
  sizes;
- text/snapshot session and process ledgers (64/512 MiB);
- Kitty/image CPU/GPU session/process ledgers (64/256 MiB) and 16M-pixel cap;
- bounded PTY-read, input-write, effect-response, snapshot request, UI-post, and
  paste queues with explicit item/byte limits;
- deterministic waiter cohorts: a release epoch selects one retry owner, tries
  each frozen waiter at most once, skips oversized requests so later fitting
  requests can proceed, and never polls/retries until a later release;
- stable LRU tie-breaks by immutable session/image IDs. Only explicit user
  visibility changes session recency; output, paint, snapshot, and UIA do not.

Every owned allocation has exactly one RAII token shared by aliases. Quota
failure publishes no partial state. Use atomics/locks with an explicit order;
never read shared non-atomic state unlocked.

**Verify**:

```powershell
.\.build\x64\Debug\TerminalTests.exe --suite resource-ledger
```

Expected: boundary and +1, overflow, rollback, resize, held-old-generation,
stable LRU, deterministic waiter, simultaneous staging, and teardown cases pass
with current bytes returning to zero.

### Step 2: Implement ConPTY creation, I/O, resize, and teardown

`ConPtySession` owns all process, pipe, attribute-list, job, and `HPCON`
resources with WIL RAII/custom unique types:

1. Create overlapped input/output pipes and a pseudoconsole at a validated,
   nonzero cell size. Build `STARTUPINFOEXW` without inheriting unrelated
   handles and start the selected command/working directory off the UI thread.
2. Put terminal-owned process trees in the planned kill-on-close containment
   policy while respecting the frozen detached-GUI behavior. Record the exact
   Windows build-dependent close ordering already specified in the core plan.
3. One serialized session mutation lane consumes bounded read chunks and calls
   Ghostty `vt_write`. A separate bounded write path sends user/effect bytes to
   ConPTY only after the engine lock is released.
4. Debounce resize to the latest nonzero cell size. Execute
   `ResizePseudoConsole` and engine resize off the UI thread under the session
   mutation order. Zero-pixel/intermediate sizes are normal control flow.
5. Natural exit drains final output, attempts the final immutable snapshot, and
   performs at most one resource-failure retry after 250 ms. Then close pipes,
   process/job/HPCON in the frozen OS-order contract.
6. Explicit close sends the configured graceful action, waits asynchronously
   up to two seconds, then escalates. Application shutdown uses its stricter
   coordinator and five-second operational quiet policy. No UI thread wait,
   `join`, `WaitForSingleObject`, or `DestroyWindow` is allowed.
7. Use `std::jthread`/stop tokens or tracked threadpool work; no detached plugin
   threads. Threadpool callbacks transfer module pins through
   `FreeLibraryWhenCallbackReturns` at entry.

**Verify**:

```powershell
.\.build\x64\Debug\TerminalTests.exe --suite conpty-lifecycle
```

Expected: echo, split UTF-8, resize coalescing, natural exit, graceful close,
escalation, process tree, detached GUI, final output, failed start, rapid close,
and app-shutdown races pass with no UI wait or leaked handle/process.

### Step 3: Bind Ghostty session mutation and effects safely

`TerminalEngineSession` owns the Ghostty terminal and per-screen configuration:

- instantiate through the validated `GhosttyApi` and custom allocator;
- set the per-screen Kitty storage cap and disable every non-direct medium;
- serialize feed, resize, reset, search/model actions, and render-begin access;
- effect callback copies only the minimal typed response/title/clipboard/bell/
  cwd data into the bounded effects queue and returns immediately;
- drain effects after unlocking. PTY responses preserve engine order. Policy-
  sensitive title/OSC 52 effects pass through plugin state gates before any UI
  publication or clipboard mutation;
- `vt_write` returning no failure is not treated as proof that all semantic
  effects succeeded; expose resource/protocol counters separately;
- clear Ghostty dirty state only after a successful snapshot/frame commit.

Implement the process-global Ghostty log and PNG callbacks in `TerminalService`.
The WIC PNG decoder initializes COM as MTA on worker threads, reads dimensions
before decode, checks every multiplication/addition, enforces width/height/
pixel/decoded byte limits, and returns memory allocated with the provided
Ghostty allocator. Never decode on the UI thread.

**Verify**:

```powershell
.\.build\x64\Debug\TerminalTests.exe --suite ghostty-session
```

Expected: functional corpus, effect ordering, reentrancy rejection, queue
overflow policy, cap/+1, malformed PNG, all pixel formats, disabled media,
callback clearing, and same-process runtime reload pass.

### Step 4: Build immutable text and Kitty snapshots

Implement single-flight, generation-coalesced snapshots:

1. Request publication by generation; at most one builder and one coalesced
   successor exist per session.
2. Pre-size candidate buffers outside the session lock where possible.
3. Under the session lock, call render-state begin, copy row/cell/style/cursor/
   hyperlink metadata, snapshot Kitty generations/placements, and copy pixels
   only for new/changed image generations. Retain no borrowed Ghostty object or
   address.
4. If capacity/generation changes, unlock, reserve/resize, and retry from a new
   generation; do not mutate borrowed data or hold the lock during allocation,
   DWrite shaping, GPU upload, UIA publication, or host callback.
5. End the render update, build immutable publication and ledger ownership, then
   atomically publish. On failure retain dirty state and the last safe snapshot.
6. Natural exit publishes a complete final snapshot or the frozen incomplete
   last-safe/empty outcome after the single retry. Incomplete final state
   suppresses automatic close-on-exit.
7. Copy/display ordering is deterministic; hidden views can publish model/UIA
   state but do no product paint work.

**Verify**:

```powershell
.\.build\x64\Debug\TerminalTests.exe --suite snapshots --repeat 500
```

Expected: dirty retry, mutation races, generation coalescing, borrowed-lifetime
poison tests, final snapshot, hidden view, cap/+1, and two-session fairness pass;
measured lock hold is emitted for every case.

### Step 5: Replace the placeholder with the plugin-owned renderer/control

`TerminalControl` remains the direct plugin child HWND. Implement:

- D3D11/DXGI device and texture cache for Kitty images; D2D/DirectWrite for
  cell backgrounds, glyphs, decorations, cursor, selection, hyperlinks, and
  status overlays;
- DWrite shaping/fallback of complete grapheme clusters into Ghostty's fixed
  cell order, clipped to their cells; handle combining, wide cells, surrogate
  pairs, emoji, color glyphs, bold/italic/underline/strike/faint/inverse;
- no claim of full paragraph BiDi in v1. RTL/Arabic fixtures must be truthfully
  shaped/clipped and documented;
- Kitty raw RGB, RGBA, grayscale, gray-alpha, and decoded PNG textures; straight
  alpha; crop/source rectangle, scaling, placement, virtual placement,
  replacement/delete, stable generation-keyed cache, and protocol z-order;
- negative-z images behind text and nonnegative image layers according to the
  locked fixture contract;
- per-monitor DPI, theme/high contrast, font metric/cell recalculation, hidden
  view behavior, resize, occlusion, WARP injection, device loss/recreation, and
  no stale texture across generation/device reset;
- manual global/row dirty clearing only after the corresponding frame commits.
  Device-loss frames preserve dirty state.

`WM_NCCREATE/WM_CREATE` initializes posted-payload ownership;
`WM_NCDESTROY` closes/drains it. All D2D/DWrite/D3D/COM resources use WIL RAII.

**Verify**:

```powershell
.\.build\x64\Debug\TerminalTests.exe --suite renderer --adapter warp
```

Expected: golden hash/geometry cases for text style, three Kitty layers,
formats/crop/scale/z-order, DPI/theme/high contrast, hidden view, dirty rows,
rapid resize, and repeated device loss all pass without a hardware GPU.

### Step 6: Implement keyboard, mouse, IME, clipboard, and terminal actions

Translate Win32 input entirely in `Terminal.dll`:

- key down/up with modifiers, AltGr distinction, dead keys, layout changes,
  surrogate pairs, application cursor/keypad modes, function/navigation keys,
  Ctrl combinations, and bracketed paste;
- IME composition/start/update/commit/cancel with DPI-aware candidate position;
- mouse selection, rectangular/word/line extension, wheel scrolling, autoscroll,
  hyperlink activation, and reporting modes only when the engine requests them;
- copy and paste with exact newline/NUL normalization and maximum byte limits;
- unsafe multiline paste and OSC 52 are typed deferred interactions. Never
  modify/read the clipboard from the Ghostty effect callback or worker lock;
- public `ITerminal` actions (`Copy`, `Paste`, `SelectAll`, `Search`, `Restart`,
  etc.) marshal to the owning thread and validate instance/view generation.

Host code must only route generic command IDs/actions; it never synthesizes VT
bytes or quotes a path.

**Verify**:

```powershell
.\.build\x64\Debug\TerminalTests.exe --suite input-clipboard
```

Expected: US/international/AltGr/dead-key/IME/surrogate matrix, VT modes,
selection, mouse reporting, paste limit/+1, bracketed paste, unsafe-paste
cancel/allow, OSC 52 allow/deny/ask, clipboard failure, and teardown races pass.

### Step 7: Implement immutable base terminal-document UI Automation

Expose UIA from plugin-owned objects using immutable snapshot/ledger state:

- document text/ranges, visible ranges, selection, caret/focus, scroll,
  supported text attributes, native window/status, and selection mutation;
- provider map keys include HWND, session/control generation, and immutable
  snapshot generation. Every owner-thread mutation revalidates all keys;
- removing discoverability atomically retires providers. A call that captured a
  snapshot/token before retirement may finish once. Any later method except
  `IUnknown` returns `UIA_E_ELEMENTNOTAVAILABLE` with null/zero output and no
  side effect;
- externally held disconnected providers/ranges retain only self-contained
  immutable snapshot plus passive reference-counted ledger state. Destruction
  must not call host, HWND, dispatcher, engine, session, or mutable service;
- allocation/reservation failure returns `E_OUTOFMEMORY`, initializes all out
  values empty, publishes no partial BSTR/SAFEARRAY/range, and remains retryable;
- local Copy reservation failure returns
  `HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_MEMORY)`, leaves clipboard/selection
  unchanged, and posts one coalesced localized status;
- password/nonce/private interaction payloads never enter UIA or diagnostics.

Only this passive disconnected UIA tail may vote for process-exit module
retention, and only after every operational gate and captured call is zero.

**Verify**:

```powershell
.\.build\x64\Debug\TerminalTests.exe --suite input-uia
```

Expected: document/range/selection/scroll/focus, worker reads, stale HWND reuse,
held range through close/RPC timeout, post-retirement errors, quota/no-partial,
privacy, racing final Release, ordinary unload, and passive-only retention pass.

### Step 8: Complete core close arbitration and stress evidence

Implement the single per-view retirement arbiter and service registry states:
Starting, Running, Exited, Retiring, Quiet, Quarantined. Cover user allow/cancel,
natural exit, explicit close, restart, forced takeover, and application shutdown
with stable IDs/generations and exactly one host removal claim. Admission remains
closed while a session is quarantined and reopens only after late quiet.

Run stress for output storm, 100,000 bounded lines, 32 sessions plus denied
33rd, rapid resize/feed/render/close, Kitty replace/evict, held UIA generations,
device loss, shell exit, app shutdown, and 1,000 complete session cycles. Run
x64 ASan and native ARM64 Debug/Release. Archive the exact ungated Release
baseline using the repository Commands/TestRuns format and content-free metrics.

**Verify**:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Full
.\Tools\Run-TerminalCommandsEvidence.ps1 -Scenario RendererControl -Platform x64 -Configuration Release -RepositoryCommit (git rev-parse HEAD) -EvidenceRoot .\Specs\TestRuns -PassThruRunPath
```

Expected: full suite exits 0; evidence path is returned and passes
`Tools/Test-TestRunArchive.ps1` plus `Tools/Verify-TerminalNativeEvidence.ps1`;
all required TerminalNative cases are present, passed, and unskipped.

## Test plan

- Deterministic native seams for clocks, dispatcher, allocator, queues, ConPTY,
  process containment, engine API, WIC decoder, D3D/WARP, clipboard, and UIA.
- Exact boundary/+1 tests for every byte/item/pixel/session/process limit.
- Functional Ghostty fixtures plus poisoned borrowed-memory tests.
- Concurrency barriers at pre-lock/in-lock/post-lock publication, callback,
  close, retirement, HWND destruction/reuse, device loss, and UIA disconnect.
- Native x64 Debug/Release/ASan and ARM64 Debug/Release. Record unsupported
  ARM64 ASan honestly; never substitute Debug.
- Source-contract scans for UI waits, detached threads, raw posted payloads,
  raw COM ownership, manual Win32 cleanup, `catch (...)`, and terminal logic
  outside the plugin.

## Done criteria

- [ ] ConPTY/session/engine mutation/render/input/UIA implementation exists only
      in `Terminal.dll`.
- [ ] No UI path waits, joins, destroys plugin HWND ownership incorrectly, or
      handles Ghostty borrowed memory.
- [ ] All hard resource limits, +1 cases, LRU/waiter ordering, and final-snapshot
      outcomes pass deterministically.
- [ ] Direct Kitty media works; disabled media fail safely; decoder limits and
      formats pass.
- [ ] WARP renderer covers text/Kitty/DPI/theme/device-loss/dirty behavior.
- [ ] Keyboard, IME, mouse, clipboard, and unsafe-paste matrices pass.
- [x] Bounded OSC 8 Ctrl+click activation and disabled/ask/open policy are
      implemented and covered at the private runtime and source boundaries.
- [x] Bounded standard-text OSC 52 writes use deny/ask/allow policy, UI-thread
      clipboard mutation, opaque payload ownership, scrubbing, and teardown
      drain, with private-runtime and source-boundary coverage.
- [ ] UIA is immutable, quota-safe, private, retirement-safe, and supports the
      closed passive-retention exception only.
- [ ] Close/arbitration/quarantine/quiet and 1,000-cycle stress pass on required
      lanes with no sanitizer finding or leaked operational gate.
- [ ] Full suite and retained Release evidence pass.
- [ ] WIP plans and `Specs/Plans/WIP/Terminal_FirstImplementationExecutionIndex_2026-08-01.md` status are updated.

## STOP conditions

Stop and report if:

- a public ABI change is needed after Plan 002;
- any Ghostty pointer/handle must survive unlock/mutation;
- effects cannot be delivered without blocking/re-entering under engine lock;
- any allocation category cannot be hard-bounded before allocation;
- ConPTY creation/resize/close requires a UI-thread wait or untracked thread;
- terminal rendering/input/UIA or process ownership must move to the host;
- WARP cannot represent the production renderer path;
- required text/Kitty/input/UIA fixture cannot be represented truthfully with
  the selected Ghostty API;
- module unload can race a worker/callback/window/provider/runtime;
- process-exit retention would include any active worker, callback, session,
  engine, HWND, dispatcher, or captured UIA call;
- a required native lane or deterministic test fails twice after a reasonable
  correction;
- implementation needs full BiDi or non-direct Kitty media for v1. Record the
  product limitation/change request instead of silently expanding scope.

## Maintenance notes

- Ghostty API changes first affect `GhosttyApi` and `TerminalEngineSession`; no
  other layer should include its headers or call its functions.
- Dirty-state ownership is manual. Review every early return/device-loss path so
  failed frames remain dirty and successful rows/global state clear exactly once.
- Resource ledgers are correctness/security boundaries, not telemetry. Metrics
  observe them but never replace admission.
- Keep UIA passive-retention code minimal and separately audited; it is the only
  acceptable reason for the plugin module to remain mapped at process teardown.
