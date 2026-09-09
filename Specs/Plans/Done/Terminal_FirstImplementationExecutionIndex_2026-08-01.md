# Terminal / libghostty first-implementation execution plan

> **Status — Completed 2026-08-03.** Plans 001–005 reached their v1 outcome. Implementation, authoritative
> Terminal/installer documentation, exact dependency locks/notices, offline x64/ARM64 reproduction, and focused
> lifecycle/source/native/package contracts are complete. Durable authority is under `../../Terminal/`; later status
> tables and WIP instructions below are retained as non-normative execution history.

This WIP index is the implementation entry point for the terminal
specifications under `Specs/`. It corrects the invalid assumption that an
engine must add zero non-system runtime DLLs without rewriting the protected
Round 4 historical record. There is no parallel `advisor-plans/` hierarchy;
all active terminal planning and status live under `Specs/Plans/WIP/`.

Read every plan completely before starting it. Execute in order, run every
verification gate, and stop on the named STOP conditions. Update the status
column after each plan.

## Executive decision

Use a separately shipped, privately loaded `ghostty-vt.dll` behind
`Plugins\Terminal.dll`, subject to a new bounded private-runtime policy and a
native Gate 0 proof. This is the best architectural option because it:

- preserves RedSalamander's stronger boundary: the host sees only the new
  `ITerminal` ABI; no Ghostty types, headers, handles, callbacks, or libraries
  cross into `RedSalamander.exe` or `Common`;
- isolates Ghostty's currently evolving C API and Zig toolchain inside one
  replaceable plugin/runtime closure;
- permits terminal-only degraded behavior when the runtime is missing or
  invalid instead of making the application fail to start;
- avoids a permanent private fork unless the bounded proof demonstrates a
  concrete upstream gap;
- makes the runtime, adapter, lock, notices, and rollback one reviewable unit.

Plan 001 has now recorded a passing, native, reproducible Ghostty winner for
the changed bounded-private-runtime policy. The first x64 product vertical
slice is implemented; later release gates remain WIP and are not implied by
that first working result.

## Decision table

Scores are 1 (poor) to 5 (best). The recommendation assumes the policy frozen
in Plan 001: at most four private non-system runtime leaves and 64 MiB of
uncompressed private runtime payload per architecture.

| Option | Product fit | Boundary isolation | Windows integration cost | Upgrade/rollback | Long-term ownership | Main failure mode | Decision |
|---|---:|---:|---:|---:|---:|---|---|
| Private dynamic `ghostty-vt.dll` loaded only by `Terminal.dll` | 5 | 5 | 3 | 5 | 4 | Windows/API proof may fail or exceed owner-day cap | **Recommended qualification path** |
| Ghostty linked statically into `Terminal.dll` | 5 | 5 | 2 | 3 | 3 | Zig/static/SIMD dependency closure and rebuild coupling | Fallback only if dynamic loading is impossible |
| RedSalamander-maintained Ghostty fork/wrapper DLL | 5 | 5 | 2 | 3 | 1 | Permanent ABI/security/upstream merge burden | Reject unless one small upstreamable patch is the sole blocker |
| WezTerm synchronous caller-owned proof | 3 | 4 | 1 | 2 | 2 | No stable embeddable C library target; large Rust adaptation | Keep only as contingency evidence, not the active first choice |
| New in-house VT engine | 2 | 5 | 1 | 4 | 1 | Years of conformance, security, graphics, bidi, and accessibility work | Reject for v1 |

## Execution order and status

| Plan | Title | Priority | Effort | Depends on | Status |
|---|---|---|---|---|---|
| [001](Terminal_QualifyPrivateLibGhosttyWinner_2026-08-01.md) | Qualify private libghostty and record a real winner | P1 | L | - | COMPLETE — x64/ARM64 engine winner |
| [002](Terminal_PluginRuntimeBoundary_2026-08-01.md) | Establish the Terminal plugin boundary, private loader, and package topology | P1 | L | 001 | FIRST SLICE COMPLETE |
| [003](Terminal_CoreRendererInputAccessibility_2026-08-01.md) | Implement the ConPTY, Ghostty model, renderer, input, and accessibility core | P1 | L | 002 | ADVANCED CORE SLICE COMPLETE; renderer/input/UIA hardening gates open |
| [004](Terminal_UserInteractionsAndSettings_2026-08-01.md) | Integrate panes, commands, profiles, path insertion, shell reuse, history, follow, and Preferences | P1 | L | 003 | WORKING INTERACTION SLICE COMPLETE; advanced history/semantic-close gates open |
| [005](Terminal_ReleaseUpgradeCloseout_2026-08-01.md) | Prove security, performance, packaging, dependency upgrades, and closeout | P1 | L | 004 | FIRST IMPLEMENTATION VERIFIED; RELEASE HARDENING IN PROGRESS |

## First implementation checkpoint

Implemented and passing on x64 Debug:

- unsuffixed `ITerminal` ABI, `Terminal.dll`, exact private Ghostty loader, and
  centralized nested runtime staging;
- ConPTY plus Ghostty ingestion, structured styled-cell/cursor rendering,
  terminal-mode-aware navigation/function keys, resize, and nonblocking
  pinned close cleanup;
- bounded asynchronous ConPTY input; safe/bracketed paste; Ghostty selection,
  bounded copy, application mouse reporting/local scroll, and basic IME;
- opposite-pane Folder/Preview/Terminal tabs, Windows and WSL folder launch,
  live terminal reuse, directory Edit, `Ctrl+Enter`, `Ctrl+Shift+Enter`,
  `Ctrl+Space`, `Ctrl+Shift+Space`, `Ctrl+Alt+T`, and `Alt+7`;
- an exact-root-PID-bound PowerShell prompt/cwd channel, authenticated
  contextual insertion, history-free idle follow, and conservative cmd/WSL
  fallback; the obsolete FolderWindow pseudo prompt is removed;
- Terminal-typed Preferences discovery and validated plugin settings;
- deterministic `cmd_pane_embedded_terminal_plugin_lifecycle` coverage.

The authoritative current behavior and explicit release-open gates are in
`Specs/Terminal/Terminal_EmbeddedPlugin.md`.

### Validation checkpoint (2026-08-02)

- Full Debug x64 run
  `20260802T190401Z-4616-084da5bd053c43418460a8bf7a4ee84a`: 1,180 passed,
  0 failed, 53 justified skips, all 17 suites passed, and zero classification,
  quarantine, or disk-audit finding.
- Focused real terminal lifecycle run
  `20260802T204347Z-2000-ced799a18ae64b8ea52d28e82e290865`: 10/10 passed
  after the final x64 dependency-state restoration and Terminal rebuild.
- Focused runtime/build reproducibility contract: 10/10 passed.
- Debug x64 build `msbuild-20260802_205528_655.log`: 0 warnings, 0 errors.
- Full Release x64 build `msbuild-20260802_220332_128.log`: 0 errors; 13 known
  warnings from Release-disabled non-Terminal selftest helpers.
- Terminal performance evidence remains at
  `Specs/TestRuns/7d3a1247382a/Commands/2026-08-02_005913/`.

### ABI-normalization checkpoint (2026-08-03)

- Every extensible Terminal record and the coordinated in-tree viewer records
  now use one leading `sizeBytes`; redundant per-record version fields and
  frozen `sizeof`/offset assertions were removed.
- `PluginContractTests.exe` passed the live DLL boundary matrix: undersized
  input rejection, current and oversized input acceptance, nested output
  `sizeBytes`, and unknown oversized-output tail preservation.
- Debug x64 solution build `msbuild-20260803_000454_554.log`: 0 warnings,
  0 errors.
- Full-suite run
  `20260802T220452Z-23076-70c6fb5304ee4281b24fbba833be01d7`
  reached every suite. The two reported failures were isolated from this ABI
  change: the ViewerText clipboard-warning case passed on immediate focused
  rerun, while Tools Pester reports the already checked-in 7.0 MiB historical
  `Specs/TestRuns/...terminal-full-suite-mixed-failures` archive above the
  repository's 5.0 MiB archive limit. The historical archive is outside this
  plan and remains untouched.
- A complete `Tools/Tests` rerun produced 582 passed, 1 failed, 5 skipped; its
  sole failure was the same historical archive-size contract. No
  `Specs/Plans/Done/` or `Specs/TestRuns/` content changed in this layer.

### Structured renderer/key/lifecycle checkpoint (2026-08-03)

- `Terminal.dll` now consumes reusable Ghostty render-state rows/cells instead
  of flattening the production frame. It draws resolved colors, style,
  decorations, selection, hyperlink color, grapheme clusters, and cursor, and
  clears dirty state only after a successful D2D frame commit.
- Navigation and F1-F24 events use Ghostty's key encoder after importing the
  current terminal modes; application cursor mode is therefore engine-owned.
- The UI-facing `Close` path removes the child and resources without joining or
  waiting. Pinned threadpool cleanup owns reader join, remaining process
  termination, engine teardown, and runtime unload.
- Debug x64 Terminal build `msbuild-20260803_093852_048.log` and
  PluginContractTests build `msbuild-20260803_093305_436.log` completed with
  zero warnings and zero errors. `PluginContractTests.exe`, the real terminal
  lifecycle selftest, and all four focused Terminal source contracts passed.

### Native interaction checkpoint (2026-08-03)

- All input-pipe writes now run through one bounded pinned worker. Complete
  requests are admitted atomically under an 8 MiB aggregate/3,840-request cap;
  rejected, completed, and close-drained buffers are scrubbed.
- Safe and bracketed paste, Ghostty-owned character/rectangular/word/select-all
  selection, bounded copy, application mouse reporting, local scroll, and IME
  composition/commit/candidate positioning now live wholly in `Terminal.dll`.
- At this checkpoint, unsafe paste with warnings enabled deliberately failed
  closed with a beep; confirmation and tracked selection were completed by the
  advanced-core checkpoint below.
- Debug x64 Terminal build `msbuild-20260803_100220_573.log` completed with
  zero warnings and zero errors. All nine focused Terminal source contracts,
  `PluginContractTests.exe` (including 12 Terminal runtime selftests), and the
  real pane lifecycle selftest passed after the IME integration.

### Kitty, confirmation, tracked-selection, and UIA checkpoint (2026-08-03)

- Direct Kitty RGB, RGBA, grayscale, gray-alpha, and WIC-decoded PNG images are
  copied behind `Terminal.dll`'s session lock, bounded, generation-cached, and
  rendered in all three Ghostty z layers. File, temp-file, and shared-memory
  transmission stay disabled.
- Unsafe paste now uses a localized nonblocking plugin overlay, while selection
  uses tracked Ghostty grid references plus viewport-edge autoscroll. Mouse
  capture loss sends the matching application button release.
- The child exposes a plugin-owned immutable UI Automation Document/Text
  provider. Retirement disconnects providers, passive externally-held objects
  return `UIA_E_ELEMENTNOTAVAILABLE`, and their count participates in the DLL
  unload quiet point.
- Debug x64 Terminal build `msbuild-20260803_103837_505.log` and host build
  `msbuild-20260803_104037_931.log` completed with zero warnings and errors.
  All 11 focused source contracts, `PluginContractTests.exe` (16 Terminal
  runtime selftests), and the real lifecycle/UIA case passed.
- Process-tree containment now creates the shell suspended, assigns it to a
  plugin-owned kill-on-close job, and resumes only after successful admission.
  Terminal build `msbuild-20260803_105621_008.log`, all 12 source contracts,
  and the real lifecycle case passed.
- Ctrl+click OSC 8 activation now uses Ghostty cell metadata, strict bounded
  conversion, a safe HTTP/HTTPS/mail allowlist, and a plugin-owned
  disabled/ask/open preference with a localized nonblocking confirmation.
  Terminal build `msbuild-20260803_110802_127.log`, all 13 source contracts,
  `PluginContractTests.exe` with 17 Terminal runtime selftests, and the real
  lifecycle/UIA case passed.
- Normalized OSC 52 standard-text writes now copy bounded borrowed data in the
  engine callback and cross to UI-thread clipboard mutation through the shared
  opaque payload registry. Deny/ask/allow policy, a configurable decoded byte
  cap, one-pending-request admission, privacy scrubbing, and teardown drain are
  plugin-owned. Terminal build `msbuild-20260803_112028_638.log`, all 14 source
  contracts, `PluginContractTests.exe` with 18 Terminal runtime selftests, and
  the real lifecycle/UIA case passed.

### Trusted prompt, follow, and insertion checkpoint (2026-08-03)

- PowerShell and pwsh launch with an encoded, profile-independent integration
  script. A random duplex named pipe accepts only the exact root process PID;
  strict bounded absolute Windows/UNC cwd reports authenticate the idle prompt.
- `Ctrl+Enter` uses a leaf only for authenticated same-parent cwd; full and
  current-directory insertion remain unconditional no-Enter operations. Open
  Command Shell, directory Edit, and every insertion command reuse a live
  opposite-pane terminal. The former bottom pseudo prompt and `cmd.exe /C`
  execution route are removed.
- Follow is plugin-owned, off by default, source-generation ordered, paused by
  pending input, and applied from `PowerShell.OnIdle` through the duplex reply.
  It sends no ConPTY bytes and creates neither app nor PSReadLine history.
- The ConPTY process attribute now passes the `HPCON` value directly. This
  fixes the child loader's `0xc0000142` failure and is source-contract guarded
  against the invalid `&pseudoConsole` form.
- Debug x64 Terminal build `msbuild-20260803_132018_753.log` and host build
  `msbuild-20260803_132057_106.log`, plus PluginContractTests build
  `msbuild-20260803_132451_393.log`, completed with zero warnings and errors.
  All 16 Terminal source contracts and `PluginContractTests.exe`, including 23
  Terminal runtime selftests, passed. Real lifecycle run
  `20260803T112544Z-25140-2aea4b0d55f04337a513ce6cb00c265c` passed 10/10 with
  zero Application Popup events, zero history additions, and zero sandbox
  finding. Focused Commands runs
  `20260803T124415Z-6636-58afd6f794e644668f4b2f6fa3e619e9` and
  `20260803T124425Z-48120-80e7d3e3e60d4bdababfc135e53bb59e` passed the trusted
  prompt/reuse/follow/path-insertion and embedded-plugin lifecycle cases.
- Post-change Full run
  `20260803T112657Z-19196-94e839e9ea0a4f2aae00dcd0c2e5e32a` reached the full
  matrix but is not a green closeout record. Terminal source/runtime/plugin and
  focused lifecycle coverage passed. Its remaining failures were outside this
  slice: long-run desktop/clipboard and Viewer scheduler contamination, the
  independently reproducible Monitor ETW-latency exit, and the immutable
  7.0-MiB historical TestRuns archive violation already recorded above. The
  Item Properties case at which Commands stopped passed alone in run
  `20260803T124135Z-10076-d1d91e322db345c98791eeb218a721f7` (1/1). Per the
  explicit repository-history constraint, no `Specs/Plans/Done/` or
  `Specs/TestRuns/` file is changed to manufacture a green aggregate.

This closes the first-working-implementation goal, not Plan 005. The plan stays
in WIP until the advanced renderer/input/accessibility, exact accepted-command
history and semantic-close state, native ARM64 product, ASan, security, and
final package mutation gates listed
in the authoritative Terminal spec are complete.

## Estimated ownership

These are owner-days, not elapsed calendar days. Parallel work is allowed only
after the ABI and ownership contracts are frozen.

| Work | Estimate | Exit gate |
|---|---:|---|
| Corrected Gate 0 and Ghostty proof | 15-20 | Native five-lane evidence and recorded winner |
| Plugin ABI, runtime loader, build/package foundation | 20-30 | Terminal-only degraded load and clean unload |
| ConPTY/model/render/input/UIA | 30-45 | Native functional and teardown suites |
| Pane/user interaction/settings integration | 25-40 | Deterministic command/profile/follow/history tests |
| Security/performance/package/upgrade closeout | 15-25 | Full suite and retained release evidence |
| **Total** | **105-160** | Authoritative specs updated; WIP plans moved to Done |

The estimate excludes a material Ghostty fork. If the Gate 0 proof needs more
than 20 owner-days or a long-lived fork, stop and reconsider the requirements
using the decision table rather than silently expanding the project.

## Non-negotiable architecture

```text
RedSalamander.exe
  -> generic plugin discovery/configuration/module lifecycle
  -> Common/PlugInterfaces/Terminal.h (ITerminal only)
  -> Plugins/Terminal.dll
       -> ConPTY/session/model/render/profile/history/follow/settings
       -> exact absolute private load
          Plugins/TerminalRuntime/ghostty-vt.dll
```

- Terminal does not implement or extend `IViewer`; it has its own `ITerminal`
  boundary. The source-tree viewer ABI received one coordinated normalization:
  `ViewerOpenContext` and `ViewerTheme` now use a leading `sizeBytes` only.
  That change does not make Terminal a viewer or alter viewer associations.
- All terminal-specific behavior lives in `Terminal.dll`. Host changes are
  limited to the public ABI and generic plugin, pane/tab, command, source
  marshaling, theme, and Preferences plumbing.
- `Terminal.dll` resolves a closed Ghostty function table dynamically. The EXE
  and `Common` do not include Ghostty headers, link its import library, or name
  Ghostty symbols.
- No runtime path is user-configurable. The only valid engine location is the
  package-relative, locked `Plugins\TerminalRuntime\` closure.
- Engine replacement is atomic and takes effect after application restart or a
  proven complete plugin quiet point. It is never hot-swapped under live
  sessions.

## Public API naming rule

Do not suffix ordinary public C++ structs, enums, interfaces, callbacks, or
methods with `V1`, `V2`, and so on. Versioning belongs in the value/ABI contract,
not in the name used at every call site.

| Concern | Required form | Example |
|---|---|---|
| Public type/interface/method | unsuffixed | `TerminalOpenContext`, `TerminalViewState`, `ITerminal`, `Open` |
| Extensible record boundary | one leading byte count | `sizeBytes` |
| ABI compatibility proof | behavioral boundary test | undersized rejected; current/oversized accepted |
| Persisted/evidence format | versioned schema/file | `schemaVersion: 1`, `*.v1.json` |
| Migration implementation | names both concrete formats | `MigrateConfigurationV1ToV2` |
| Incompatible interface successor | repository COM convention | `ITerminal2` with a new IID, not `ITerminalV2` |

All extensible public ABI records use `sizeBytes` only. A consumer rejects a
record shorter than the prefix it needs, accepts larger values without reading
unknown tail bytes, and preserves unknown caller tails on output. Append-only
optional fields may remain on the same method. Incompatible ownership or
semantics require a new IID/interface or method, never reinterpretation of an
existing layout. Persisted/evidence formats keep independent schema versions.

## Highest-risk problems to solve

| Risk | Probability | Impact | Required control | Stop threshold |
|---|---|---|---|---|
| libghostty C API and build conventions are not version-stable | High | High | Exact source/toolchain/output pin, runtime function table, ABI exercise probe | Cannot freeze one reviewed export/layout contract |
| Windows shared build or native ARM64 lane fails | Medium-high | High | Reproducible five-lane build before product code | No native ARM64 result or private closure exceeds cap |
| Ghostty PNG/log callbacks are process-global | Medium | High | One module runtime generation, startup-only registration, strict quiet-point clearing | Callbacks cannot be cleared before unload |
| Borrowed render/Kitty pointers are invalidated by mutation | Medium | High | Snapshot under the session lock; copy changed image data; retain no borrowed pointer | Any pointer/handle survives unlock or a mutating call |
| Exact typed internal memory budgets are not exposed | High | High | Approved requirement change to a hard aggregate Ghostty allocator cap plus separate plugin queues/caches | Aggregate cap cannot reject before allocation or stay observable |
| PNG/image decompression bombs | Medium | High | WIC dimension/pixel/byte preflight, checked arithmetic, direct medium only in v1 | Decoder must allocate before limits are known |
| Runtime DLL search/hijack or time-of-check/time-of-use | Medium | High | Exact absolute path, reparse/PE/hash/import/export checks, held file handle, loaded identity recheck | Loader falls back to basename/PATH or accepts an unbound file |
| `RuntimeDependencies.props` only supports flat `Plugins\` output today | High | Medium | Add validated package-relative destination semantics used by build and every packager | Build, ZIP, MSI, and MSIX cannot consume one manifest |
| UI-thread stalls from parsing/render snapshots or teardown | Medium | High | Serialized worker mutation lane, short render snapshot lock, async close/reaper, perf counters | UI thread waits on worker/process or p95 budgets fail |
| ConPTY shutdown races and callbacks after module unload | Medium | High | stop producers -> cancel -> drain posts/callbacks -> release sessions -> unload | Any callback/thread/module pin remains at quiet point |
| DirectWrite/Kitty geometry, DPI, device loss, or color glyph mismatch | Medium | Medium-high | Golden fixtures, WARP tests, dirty-state discipline, device recreation tests | Required fixture cannot be represented without upstream data |
| IME, AltGr, dead keys, surrogate pairs, paste, and VT key encoding conflict | Medium | High | Translate Win32 input in plugin; native keyboard/IME matrix | Host must synthesize terminal escape sequences |
| WSL path/profile detection is ambiguous or stale | Medium | Medium | Typed source/launch locations and authenticated cwd; deterministic matrix | Launch requires guessing an unvalidated distro/path |
| Individual DLL signing is introduced later | Low today | Medium | Hash final shipped bytes; sign runtime before manifest generation or use signed detached manifest | Post-manifest signing changes the runtime hash |
| License/SBOM closure differs from the whole Ghostty repository | Medium | High | Derive notices and dependency closure from the actual lib-vt build | Any shipped leaf lacks provenance/license evidence |

## Dependency notes

- Plan 001 must end in either a recorded Ghostty winner or an explicit no-winner
  decision. Plans 002-005 must not start on a no-winner result.
- Plan 002 freezes the public ABI, loader, runtime manifest, nested staging, and
  degraded behavior consumed by every later plan.
- Plan 003 creates the terminal core and public actions but intentionally does
  not add the full RedSalamander UX.
- Plan 004 owns the user-visible interactions. It may add generic host plumbing,
  but shell/profile/history/follow/quoting logic remains in `Terminal.dll`.
- Plan 005 is not optional cleanup. Repository policy requires performance
  evidence, clean package extraction, authoritative specs, and WIP closeout.

## Findings considered and rejected

- **Stretch `IViewer`**: rejected. Its file-viewer lifecycle cannot express a
  terminal session, running close, profile, authenticated cwd, path insertion,
  restart, or quiet-point contract without an ABI break.
- **Put Ghostty directly in the EXE**: rejected. It leaks engine/toolchain churn
  into the host and prevents terminal-only degradation/rollback.
- **Forbid all private runtime DLLs**: rejected. It conflicts with the existing
  product intent that already permits a locked `Plugins\TerminalRuntime\`
  closure and does not improve security over a strictly verified private load.
- **Expose a runtime path in Preferences**: rejected. It defeats package
  identity, digest validation, supportability, and safe rollback.
- **Require a typed accounting patch inside Ghostty for v1**: rejected unless
  the aggregate custom-allocator proof is insufficient. A private fork is a
  worse default than an enforceable whole-engine cap plus separate plugin
  budgets.
- **Use Ghostling as Windows integration proof**: rejected. It is a useful API
  sample, not evidence for RedSalamander's Windows ABI, ConPTY, D2D/DWrite,
  D3D11, UIA, package, or teardown requirements.
