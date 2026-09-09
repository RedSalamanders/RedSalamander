# Terminal Shell Integration, History, And Pane Follow

> **Status — Completed 2026-08-03.** The v1 trusted PowerShell prompt protocol, accepted-line tracking, bounded
> history/state, source-generation-safe pane follow, path insertion admission, and fallback behavior are implemented.
> Durable behavior is in `../../Terminal/Terminal_EmbeddedPlugin.md`; broader shell semantics are future work. Later
> WIP/checkpoint language below is historical.

## Current checkpoint (2026-08-03)

The working first implementation now has production PowerShell/pwsh
authentication for the subset it needs: exact root PID, current provider path,
conservative idle prompt, contextual insertion, and history-free pane follow.
The plugin uses one random duplex pipe and a bounded `PowerShell.OnIdle`
`Get-Location`/`Set-Location` action; no command bytes are injected into ConPTY
and PSReadLine history is unchanged. Cmd and WSL remain untrusted. Exact
accepted input/completion, child/password/TUI state, app history, and semantic
close policy remain wholly owned by this WIP plan.

The extensive adapter/protocol below is the reviewed release-hardening design,
not a description of the current working channel. It must replace that channel
atomically if a remaining gate requires its additional capabilities; do not
partially combine both protocols.

> **Authority:** Execution plan 5 of 6 for `Terminal_EmbeddedConPtyLibGhosttyVt_IntegrationPlan_2026-07-13.md`. Plans 1–4 and the master are predecessors.

## Status and entry gate

- **State:** WIP / authenticated cwd-idle-follow slice implemented; advanced history and semantic close not implemented
- **Baseline:** `fbf6af76f373bf253488862426f4ea70a9ff25ba`
- Enter only with the exact `Specs/Plans/Done/Terminal_PaneTabsProfilesAndLifecycle_2026-07-22.md` closeout commit/blob, its predecessor identities, the exact PaneProfiles `TerminalNative` manifest/companion paths and digests, and exact terminal/Preview Commands archives/digests. Recompute the Done blob and rerun `Tools/Verify-TerminalNativeEvidence.ps1`; a missing/mismatched path, implicit latest-run lookup, skipped invocation, or repository/engine-lock mismatch blocks implementation.
- Do not begin until the `builtin/terminal` DLL/ABI is green; reused open command, directory Edit, contextual/full path insertion routes, and pseudo-command-line removal pass; a clean terminal launches for cmd, Windows PowerShell, pwsh, and exact matching WSL distro profile; stable source/session IDs and ordered input operations are available.
- The accepted plan-4 boundary is intentionally terminal-only for production
  PowerShell: `TerminalOnly` and `SemanticEnable -> SemanticUnavailable` complete
  the generic startup/rooting protocol with zero hook, semantic handshake, or
  capability; only deterministic fake adapters prove `SemanticReady` and both
  `StartupSemanticJoin` arrival orders. If plan 4 instead staged/installed a
  production adapter, claimed real `SemanticReady`, or added its manifest/
  observer/control resources, STOP and repair the predecessor boundary before
  this plan begins.
- A launchable shell may remain terminal-only. This plan never weakens terminal usability to claim integration.
- Before implementation, compare the master and plans 1–4 plus integration, SettingsStore, state-file, watcher, and test-helper surfaces with the baseline. Any material ownership, ABI, protocol, generation, persistence, privacy, or harness drift is a STOP: update/re-review this plan and record a new baseline before product edits.

## Implementation and build surfaces

- This plan is the first implementation owner for the allowlisted PowerShell
  adapter manifest, observer/control scripts and RCDATA resources, versioned
  owner/reparse/hash-verified extraction/install path, real semantic handshake,
  production `SemanticReady`, and real-shell semantic closeout. Replace plan 4's
  production `SemanticUnavailable` seam only after the exact tuple and extracted
  bytes validate; a fake Ready result never activates production capability.
- Implement the observer/control protocol only in `Plugins/Terminal/TerminalShellIntegration.h/.cpp` and the plugin-owned shell-adapter/resource support it calls. Add `Plugins/Terminal/ShellAdapters/PowerShellReadLineAdapters.json`, `TerminalSemanticProtocol.v1.json`, each closed schema and valid/invalid corpus, the deterministic `Tools/Generate-TerminalSemanticProtocol.ps1` plus its checked-in generated C++/PowerShell outputs, and the exact versioned wrapper/control script resources to `Terminal.rc`/`resource.h`; list every source/resource in `Terminal.vcxproj` and `.filters`. The two canonical JSON inputs, generated PowerShell constants, wrapper template, and reviewed scripts are build inputs compiled into `Terminal.dll` RCDATA with exact resource ID/version/SHA-256 identities; the generated C++ constants compile into that same DLL. None is a loose runtime dependency or host/EXE/Common implementation file, and MSBuild runs the generator in `-Check` mode before compilation.
- Add `Tools/Test-TerminalPowerShellAdapters.ps1` for closed schema/corpus validation and read-only `-AuditInstalled`, plus `Tools/Test-TerminalPowerShellControl.ps1` for the mandatory real-shell product matrix through the public Terminal ABI. Register deterministic observer/control cases in `Tests/TerminalTests`, the ShellState native slice, and `terminal_shell_state_` Commands coverage. Source-contract and extracted-package tests reject loose adapter manifests/scripts, a second parser, host-side semantic interpretation, and any product implementation outside `Terminal.dll`.
- The adapter manifest is exact, not a compatibility hint. Each tuple binds PowerShell edition and host-version interval, PSReadLine version, module path/file SHA-256, exact UTF-8 `FunctionInfo.Definition` SHA-256, reflected static `Microsoft.PowerShell.PSConsoleReadLine.ReadLine` overload signature, strong assembly identity, one audited recipe, wrapper/manifest resource identities, a clean-default real-shell test, and an upstream-license record. Reject duplicate/overlapping/floating tuples, unhashed inputs, nearest-version/range fallback, unknown overload/recipe, and function-text patching.

## Trusted integration protocol

- Every local or UNC Windows PowerShell/pwsh launch, including an explicit
  Restart and a launch with `shellIntegrationEnabled=false`, first consumes Plan
  4's exact plugin-owned protected-HMAC/nine-kind
  `PowerShellStartupBootstrap`, including its canonical fixed-`*>&1` fail-closed
  wrapper. The initial child token array is exactly
  `[fixedBasename,"-NoExit","-Command",
  bootstrapWrapperCommand]`; normal profiles run exactly once before the
  verified embedded script. The requested target and every semantic nonce/
  control-pipe value are absent from the initial argv and child environment.
  The startup bootstrap key is a distinct launch-only key and is never the
  ongoing semantic epoch.
- After authenticated post-profile module-qualified rooting and a matching
  `RootResult`, the plugin samples the latest effective integration value and
  policy generation exactly once and sends the sole `SemanticDecision`.
  `TerminalOnly` creates no semantic nonce/ongoing pipe and installs no hook.
  `SemanticEnable` alone adds one fresh 128-bit `BCryptGenRandom` semantic nonce,
  one newly assigned nonzero no-wrap `channelGeneration:uint64`, plus the
  already-created canonical UTF-16 pipe leaves
  `RedSalamander.TerminalSemantic.<32-lowercase-hex-random>` and
  `RedSalamander.TerminalControl.<32-lowercase-hex-random>` for the bounded
  inbound event pipe and distinct outbound operation pipe to that authenticated
  decision. Each is validated exactly and passed to a local-server
  `NamedPipeClientStream`; neither is an arbitrary path. The script receives
  those values only through the bootstrap pipe,
  copies the nonce into the root adapter module's private closure, and removes/
  zeroes mutable decision copies before prompt commit. The requested target
  arrives only in authenticated RootRequest and semantic nonce/channel-generation/
  pipe values only in SemanticEnable; neither appears in initial argv or the
  child environment. The
  wrapper's seven internal bootstrap literals remain exactly Plan 4's fixed
  launch contract and contain neither target nor semantic material. No semantic
  value is inherited by a child/nested process.
- The decision-delivered nonce is the incarnation's sole ongoing semantic
  epoch. Its unsigned 64-bit semantic sequence starts at 1 with the sole
  authenticated handshake and wrap is fatal. There is no distinct epoch created
  after Running, no live rekey, and no second semantic handshake in one
  incarnation. Cmd and WSL create no ongoing semantic nonce or epoch.
- On `SemanticEnable`, the script connects the exact semantic pipe, installs the
  allowlisted adapter before Running while all input/history/activity/follow/
  insertion/capability gates remain closed, and reports exactly
  `Prepared(SemanticReady|SemanticUnavailable,decisionPolicyGeneration)`.
  The handshake uses the bounded asynchronous semantic-emission contract below.
  A handshake write throw/fault/cancel/timeout closes the event stream, leaves
  any partial wrapper process-lifetime inert pass-through, and reports
  `SemanticUnavailable` without retrying or failing an otherwise rooted
  terminal-only launch. The plugin's bounded
  `StartupSemanticJoin` has two pre-reserved optional slots for Prepared Ready
  and the independently transported handshake; either may arrive first, neither
  grants capability alone, and only their fully validated matching conjunction
  strictly before the common deadline may advance. TerminalOnly requires no
  Handshake and no occupied join slot; any paired Handshake is fatal.
  Unavailable normally has no Handshake, but one structurally/authentically valid
  matching sequence-1 Handshake staged before the server deadline and paired
  with child Unavailable under the exact client-observation race is discarded/
  revoked terminal-only nonfatally. Duplicate/conflicting/gapped/wrong-identity/
  malformed/extra/oversize/queue-admission records are fatal and permanently
  revoke the staged epoch.
- After accepted Prepared, `ReleasePrompt` makes the script remove target/
  decision/semantic variables, zero mutable buffers, retain only the bootstrap
  HMAC key as mutable secret state plus the minimal immutable Decision generation
  and Prepared outcome needed for commit, send `ReleaseAck`, and wait without
  returning to a prompt. `ReleaseAck` is not commit. The plugin then wins the sole lifecycle-
  locked `Starting/BootstrapReleaseAcked -> Running/PromptCommitPending` CAS or
  sends no CommitPrompt and retires. A winning CAS makes Running public while
  user input, profile memory, and semantic capabilities remain gated.
- After that CAS, authenticated `CommitPrompt` carries the CAS-frozen final
  disposition and policy generation. Its only valid rows are: original
  TerminalOnly plus Prepared TerminalOnly -> TerminalOnly at equal/later
  generation; SemanticEnable plus SemanticReady plus a completed join ->
  SemanticActive only at unchanged generation; SemanticEnable plus
  SemanticUnavailable -> TerminalOnly only at unchanged generation; and either
  valid Enable outcome after a latched disable -> TerminalOnly at the current
  strictly greater generation regardless of later re-enable. Every other pair,
  decrease/wrap, or Active without Enable+Ready+join is fatal. After applying
  the disposition, the script prebuilds/HMACs CommitAck, zeroes the last key,
  and returns exactly one two-property object whose first PSTypeName is
  `RedSalamander.TerminalBootstrapFinalize.V1`, whose ordered properties are
  exactly `Sentinel,Finalize`, whose Sentinel is
  `RSTBOOT-WRAPPER-READY-1`, and whose Finalize is a scriptblock. Only after
  validating that sole object does the wrapper invoke its private output-silent
  one-shot finalizer, which sends the complete CommitAck in exactly one
  synchronous message-mode Write with no Flush/retry/later pipe operation. Only
  the server's full valid CommitAck read opens ordinary input/profile-memory
  admission and publishes independently proved semantic capability.
- One monotonic `launchT0 + 5 seconds` deadline covers bootstrap connection and
  every receipt through the server's complete CommitAck read; every receipt must
  be strictly before it and equality is timeout. Close, WindowClosing, Restart,
  root/profile exit, timeout, ApplicationShutdown, and protocol/root/wrapper-
  finalization failure race the one lifecycle CAS. A pre-CAS winner publishes no
  Running/profile memory; an unrelated failure after the CAS retires the already-
  Running incarnation with input/profile-memory/capability gates still closed.
  The deliberate live-disable cleanup path below is the sole nonfatal policy
  exception and disable alone never retires startup.
- Every child-to-plugin handshake, accepted-line/root-read marker, activity/cwd
  update, and foreground-control result uses exactly the incarnation's
  `\\.\pipe\RedSalamander.TerminalSemantic.<32-lowercase-hex-random>` inbound
  server. It is `PIPE_ACCESS_INBOUND | FILE_FLAG_OVERLAPPED |
  FILE_FLAG_FIRST_PIPE_INSTANCE`, message-mode, one-instance, has a 64 KiB
  inbound buffer and uses `PIPE_REJECT_REMOTE_CLIENTS`. Its noninherited protected
  descriptor is current-user-owned and permits only LocalSystem full access and
  that SID write/synchronize access. The root module connects write-only;
  `GetNamedPipeClientProcessId` must equal the exact root PowerShell/pwsh PID.
  Before `SemanticEnable`, the plugin charges the session/process memory ledgers,
  reserves exactly one 2 MiB semantic-frame staging buffer plus fixed Connect/
  Read OVERLAPPED state, and has one `ConnectNamedPipe` pending. It arms the
  first `ReadFile` only after connection completion and exact-PID acceptance;
  no pre-connect pending read is claimed. One client/server connection persists
  for the epoch and is never reconnected. One read remains pending while the
  connected epoch is idle; after a complete message, staging is reset and the
  next read is armed before dispatch. Reservation/connect/arm/rearm failure
  closes/revokes; parser admission never allocates or waits. The
  outbound `TerminalControl` pipe below is only app-to-root request transport,
  never an alternate marker path.
- One semantic frame is one pipe message emitted by exactly one reviewed-module
  `NamedPipeClientStream.WriteAsync(byte[], 0, frameLength,
  CancellationToken)` call. The client uses local server name `.`, the exact
  decision-delivered `RedSalamander.TerminalSemantic.<32-lowercase-hex-random>`
  leaf, `PipeDirection.Out`, `PipeOptions.Asynchronous`, and
  `TokenImpersonationLevel.Anonymous`. One call is exactly one message. The
  server continues only that message across `ERROR_MORE_DATA`; the final
  successful `ReadFile` delimits it, never EOF. EOF with partial staging is
  truncation, clean unexpected EOF revokes, lifecycle-fenced EOF is normal, and
  2 MiB filled with more data pending is the fatal +1 case. It never parses a
  prefix. The little-endian layout is exactly
  `magic[8]="RSTSEM01"`, `frameSize@8:uint32`, `version@12:uint32=1`,
  `kind@16:uint32`, `flags@20:uint32=0`, `sequence@24:uint64`,
  `semanticNonce@32:uint8[16]`, `payloadBytes@48:uint32`,
  `reserved@52:uint32=0`, payload at offset 56, then a 32-byte HMAC-SHA256 over
  every preceding byte using the semantic nonce. `frameSize` includes the tag,
  equals `56 + payloadBytes + 32`, and is at most 2 MiB. Sequence starts at 1
  with the sole Handshake; ordinary frames 2 through N on the same connection
  are required/valid and advance exactly once, but a second Handshake or wrap is
  fatal. Before the startup join, any extra semantic frame is fatal. Unknown
  kind/flags/reserved/trailing byte, bad length/tag/nonce, partial/over-cap byte,
  or sequence error permanently invalidates the epoch. Plan 4 bootstrap frames
  remain a separate launch-generation/operation-
  keyed schema with no semantic-nonce field. No engine tagged effect, PTY/OSC/
  APC output, second semantic pipe, or second VT parser can authorize state.

  The `RSTSEM01` payload language is closed at exactly four V1 kinds:
  `Handshake=1`, `RootReadReady=2`, `AcceptedInput=3`, and `ControlResult=4`.
  Every offset below is relative to payload byte zero. All integers are exact
  little-endian unsigned byte sequences; both implementations use checked
  byte-copy readers/writers and never cast payload bytes to a C/C# structure,
  depend on padding, or accept TLV/extensions. Each payload's computed exact
  size must equal `payloadBytes`, with checked additions and no trailing byte.
  Unknown kind, enum, flag bit, reserved value, status, or cross-field
  combination is a fatal epoch error.

  - `Handshake=1` is exactly `payloadVersion@0:uint32=1`,
    `psEdition@4:uint32` (`Desktop=1|Core=2`), `editMode@8:uint32`
    (`Windows=1|Emacs=2`), `wakeCodePoint@12:uint32` (exactly one of
    `0x1c|0x1d|0x1e|0x1f|0x07`), nonzero
    `channelGeneration@16:uint64`, `capabilityFlags@24:uint64`,
    `tupleIdBytes@32:uint32`, `psVersionBytes@36:uint32`,
    `psReadLineVersionBytes@40:uint32`, `reserved@44:uint32=0`, and
    `adapterDefinitionSha256@48:uint8[32]`; tuple ID, PowerShell version, and
    PSReadLine version bytes are concatenated in that order at offset 80. Its
    exact size is `80 + tupleIdBytes + psVersionBytes +
    psReadLineVersionBytes`. Tuple ID is 1..128 ASCII bytes matching only
    `[A-Za-z0-9._-]`. Each version is 3..43 ASCII bytes in canonical
    `System.Version.ToString()` form: two through four dot-separated canonical
    decimal components, each `0..2147483647`, with no sign or leading zero
    except the single digit `0`; reparsing and `ToString()` must reproduce the
    bytes exactly. Edition, both versions, tuple ID, and definition SHA-256 must
    equal the one exact allowlisted manifest record. `channelGeneration` must
    equal the nonzero semantic-channel generation issued with this
    `SemanticDecision`. Allowed capability bits are exactly
    `TrustedRootBoundary=0x1`, `TrustedFilesystemCwd=0x2`,
    `ExactAcceptedInput=0x4`, `FollowControl=0x8`, and
    `InsertControl=0x10`; all other bits are fatal. A control bit requires
    `TrustedRootBoundary`, and `FollowControl` additionally requires
    `TrustedFilesystemCwd`. These are tuple claims only: runtime publication
    still intersects them with the accepted startup join, current state, and
    successful real control Probe.
  - `RootReadReady=2` is exactly `payloadVersion@0:uint32=1`, `flags@4:uint32`,
    nonzero `readCycle@8:uint64`, `completedAcceptedSequence@16:uint64`,
    `cwdKind@24:uint32`, `cwdBytes@28:uint32`, and cwd bytes at offset 32;
    exact size is `32 + cwdBytes`. Allowed flags are exactly
    `CompletedAccepted=0x1`, `PriorStatusAvailable=0x2`, and
    `PriorCommandSucceeded=0x4`. Success requires both other bits;
    availability is forbidden without CompletedAccepted.
    `completedAcceptedSequence` is zero iff CompletedAccepted is absent;
    otherwise it equals the still-pending immediately prior `AcceptedInput`
    frame sequence and consumes that candidate once. `readCycle` starts at one,
    increments once for every owned root-wrapper entry, and may not wrap.
    `cwdKind=0` means unavailable/non-filesystem and requires `cwdBytes=0`;
    `cwdKind=1` means Windows filesystem and requires 1..131,072 strict UTF-8
    bytes containing the byte-identical canonical Windows drive/UNC logical
    path produced by this plan's Location-identity canonicalizer.
  - `AcceptedInput=3` is exactly `payloadVersion@0:uint32=1`, `cwdKind@4:uint32`,
    `readCycle@8:uint64`, `commandBytes@16:uint32`, `cwdBytes@20:uint32`,
    command bytes at offset 24, then cwd bytes; exact size is
    `24 + commandBytes + cwdBytes`. `readCycle` must equal the RootReadReady
    cycle that opened this exact still-active ReadLine invocation, even when
    intervening ControlResult frames exist. Command is 0..1,048,576 strict
    UTF-8 bytes from the exact returned `System.String`; empty and multiline
    are legal and CR/LF are preserved without normalization. Cwd is captured at
    acceptance and obeys the same `cwdKind`/131,072-byte rules as RootReadReady.
  - `ControlResult=4` is exactly `payloadVersion@0:uint32=1`,
    `controlKind@4:uint32` (`Probe=1|Follow=2|Insert=3`),
    `status@8:uint32` (`Success=0|RejectedState=1|Expired=2|
    BindingChanged=3|BufferChanged=4|ShellOperationFailed=5`),
    `flags@12:uint32=0`, `controlSequence@16:uint64`,
    `operationId@24:uint8[16]`, `cwdBytes@40:uint32`,
    `reserved@44:uint32=0`, `bufferUtf16CodeUnits@48:uint64`,
    `bufferSha256@56:uint8[32]`, and cwd bytes at offset 88; exact size is
    `88 + cwdBytes`. It must match the sole awaiting `RSTCTL01` kind, nonzero
    sequence, operation ID, epoch, and deadline. At epoch creation the plugin
    generates one nonzero 64-bit operation namespace with `BCryptGenRandom` and
    never changes it in that epoch; operation ID bytes are exactly namespace
    little-endian at `[0..7]` plus that operation's nonzero control sequence
    little-endian at `[8..15]`. No-wrap sequence makes IDs unique within the
    epoch without an unbounded replay set; nonce plus ID is the cross-epoch
    identity. Probe requires all result-data fields zero. Follow Success
    requires a nonempty canonical cwd and zero buffer length/digest. Insert
    Success requires zero cwd, `bufferUtf16CodeUnits` in `1..1,048,576`, and
    SHA-256 of the exact complete post-insert .NET buffer serialized as raw
    UTF-16LE code units with no BOM or terminator. Every non-Success result has
    zero cwd/length/digest. `RejectedState`, `Expired`, `BindingChanged`, and
    `BufferChanged` may be emitted only before any shell-effect API call and are
    mechanically no-action; `ShellOperationFailed` means an API call was
    attempted and is conservatively effect-ambiguous. A BeginRead still
    incomplete after the sole handler `WaitOne` emits no result. The result
    advances the ordinary semantic-frame sequence exactly
    once only after its bounded write succeeds.

  All payload text uses the PowerShell-side
  `System.Text.UTF8Encoding(false,true)` and the Terminal.dll side reuses
  `Common/StringConversion.h` strict UTF-8/UTF-16 helpers before applying these
  protocol-specific rules; a terminal-local converter is forbidden. No text
  slice may contain U+FEFF, U+0000, or an unpaired surrogate. Tuple IDs and paths
  reject every control character. Command text permits only HT, LF, and CR from
  the C0/C1 control ranges and rejects every other C0/C1 code point. There is no
  fifth Activity/Cwd kind: accepted AcceptedInput enters
  `RunningRootCommand`; RootReadReady returns the root to idle, publishes the
  cwd classification, and optionally completes exactly one pending acceptance.
  Sequence one is Handshake and every later kind consumes the next exact frame
  sequence.

  The canonical machine-readable source is
  `Plugins/Terminal/ShellAdapters/TerminalSemanticProtocol.v1.json`, validated
  by its closed schema and corpus and compiled into `Terminal.dll` RCDATA.
  `Tools/Generate-TerminalSemanticProtocol.ps1` deterministically emits
  `Plugins/Terminal/Generated/TerminalSemanticProtocolV1.generated.h` and
  `Plugins/Terminal/Generated/TerminalSemanticProtocolV1.generated.ps1.txt`;
  the latter is also compiled as RCDATA, never staged as a loose runtime file.
  The build runs the generator in `-Check` mode and fails on stale output.
  Source-contract tests decode the C++ constants and the embedded PowerShell
  constants and require byte-identical enums, offsets, masks, bounds, and
  cross-field tables against the canonical JSON. The corpus covers one valid
  minimum/maximum of every kind, multibyte and astral text, plus unknown kind/
  status/flag, nonzero reserved, bad version/manifest identity, illegal
  capability combination, length overflow/underflow/trailing bytes, invalid
  UTF-8/control/path/version, command 1-MiB+1, cwd 128-KiB+1, full frame
  2-MiB+1, sequence/read-cycle wrap, wrong completed sequence, and wrong
  control ID/kind/data combination.

  Each exact supported Windows PowerShell 5.1/pwsh tuple must prove from its
  loaded runtime API and a real message-mode pipe that this four-argument
  overload exists, returns without a synchronous large-write stall, and one
  call preserves one message boundary through the 2 MiB cap. Failure makes
  production integration unavailable for that tuple; there is no synchronous
  Write, helper process, background runspace, alternate runtime leaf, or second
  transport. For every message the server captures `receiveT0` with the first
  completed read chunk and requires the final successful `ReadFile` strictly
  before `checked(receiveT0 + ceil(qpcFrequency/20))`; equality/overflow closes/
  cancels/revokes. Absence of the first Handshake chunk uses the existing strict
  `launchT0 + 5 seconds` deadline because adapter installation follows connect.
- Every emission owns a fresh cancellation-token source, frame byte array, and
  returned Task. A module-injected monotonic Stopwatch starts before frame
  serialization/HMAC and defines `writeT0 + 50 ms`. Success requires the sole
  Task to be `RanToCompletion`, synchronously observed, and followed by a clock
  sample strictly before the deadline; equality is timeout. The foreground hook
  waits only through
  `((IAsyncResult)task).AsyncWaitHandle.WaitOne(nonnegativeWholeMsRemainder)`
  with the remainder rounded down, then samples clock/status. `Task.Wait`,
  `.Result`, blocking `GetAwaiter`, `SpinWait`, callback, and continuation are
  forbidden. Only success advances the semantic sequence and permits release.

  A module-local synchronous throw, task fault/cancel, timeout, observed EOF/
  break, root exit, or module teardown cancels the token, immediately closes the
  client stream, and latches all adapter hooks process-lifetime inert. There is
  no retry or replacement frame for that failed emission, fallback transport,
  Task continuation/callback, helper thread, or background runspace.
  `AcceptedInput` returns the exact user line, `RootReadReady` proceeds to the
  underlying shell read, and other failed emissions return ordinary raw shell
  control. The module privately retains the Task, token source, and byte array
  until a later owned foreground entry or teardown sees completion, then
  accesses a faulted Task's `Exception` only to observe it content-free, zeroes
  mutable bytes, and only after settlement disposes/releases Task, wait handle,
  token source, and buffer. It admits no overlapping semantic write while that record exists.
  Teardown uses the same cancel/close rule and never waits in a prompt hook; the
  existing bounded root-process shutdown is final cleanup if the Task does not
  settle. Until Prepared the child's outcome controls capability: a valid
  matching sequence-1 Handshake staged before the receive deadline but paired
  with Prepared Unavailable—including server read before and client observation
  at/after the write deadline—is scrubbed/revoked terminal-only nonfatally.
  Ready still requires the staged match; malformed/wrong identity/nonce/HMAC/
  sequence, extra pre-Prepared frame, or Handshake with TerminalOnly is fatal.
- Any ongoing semantic nonce/framing/bounds/replay/duplicate/gap/order failure, sequence wrap, oversized integration record, or aggregate-effect admission failure permanently invalidates that incarnation's sole epoch. Atomically revoke every semantic capability, discard the pending history candidate, cancel/purge pending not-yet-writer-admitted follow/insertion/integration work, reject every later marker and second handshake, and show localized `Shell integration unavailable — Restart terminal`. Already writer-admitted bytes are not recalled but grant no later state. Only explicit Restart or a new session may create a new nonce/handshake.
- Ordinary OSC 7/133/semantic cells may render but never authorize persistence/synthetic input. Untagged nested/remote/SSH/tmux/screen output cannot authorize anything and does not itself invalidate a valid top-level epoch. While a nested program owns input, top-level capabilities pause and resume only at the next correctly ordered root marker in that same valid epoch after return. Nested integration is unavailable in v1: no child inherits the root nonce and no nested shell may negotiate a fresh/second handshake.
- The pipe/PID/nonce/HMAC/sequence contract mitigates ordinary PTY, nested/
  remote, stale-launch, unrelated-process, and replay spoofing, not a same-user
  local attacker or earlier code in the same root. User-authored PowerShell/pwsh
  profiles run before the reviewed wrapper and are explicitly trusted; correct-
  key/correct-nonce same-root preemption is indistinguishable and outside the
  threat model. Same-user command-line/.NET-heap inspection is also outside;
  mutable buffers are zeroed and immutable references dropped, but no physical-
  erasure claim is allowed. Limitation probes document that boundary while
  wrong PID/key/nonce/sequence/replay and PTY-marker attempts fail closed. State
  it in the threat model/capability UI.

The pipe-pair revocation/unload contract is exact. A pre-CAS live disable keeps
only the `RevokedTerminalOnly` cleanup transports through Prepared and the
Running CAS, then closes/cancels both `TerminalControl` and `TerminalSemantic`.
A disable at/after that CAS, Close, Restart, root exit, shutdown, or independent
protocol retirement installs its generation/admission fence and immediately
closes/cancels both. The plugin then cancels/drains every accepted Connect/Read
completion, fixed staging/parser unit, and semantic/control dispatch. Handles,
OVERLAPPED storage, reservations, and parser/dispatch gates remain counted until
actual release; `CanUnloadNow` stays false. The existing off-UI five-second quiet
deadline ends in session quarantine or fatal application-shutdown policy on a
miss, never reachable-code unload.

The root module owns exactly one bounded validated-control-frame staging slot in
addition to its current BeginRead and invokes its one `PollControlTransport`
routine in two exact modes. Every ordinary owned wrapper entry uses `NoWait`:
only while the slot is empty, it may call `EndRead` for an already-completed
BeginRead, validate the complete frame into that slot, and arm the next read; it
never waits or dispatches. While the slot is occupied it leaves the current read
untouched. The key-handler uses `HandlerWaitOnce` only after the owner/binding/
epoch/lifecycle checks below and first atomically takes an occupied slot. That
invocation performs zero `WaitOne` and zero `EndRead` on the newly armed read and
dispatches or scrubs the staged frame once. Only with no staged frame may it use
the sole bounded `WaitOne` and call `EndRead` for an already-complete or signaled
read; that path validates into the same slot, rearms, then takes the slot before
dispatch. Either mode latches all hooks
process-lifetime inert on zero-byte EOF or named read failure before new emission/
dispatch. Closing `TerminalSemantic` breaks a pending
WriteAsync, or the next emission when none is pending, and follows the same
inert/no-replacement-frame rule. A completed control read that already won the
pre-fence writer-admission point below may execute once; close never pretends to
erase it. All resulting semantic frames/results are rejected after the plugin
fence and authorize no state/history/capability.

Runtime capabilities are independent flags: trusted prompt, trusted cwd, exact accepted input, non-executing insertion, integration-owned cwd change, native-history suppression. History/follow/insertion each require their full subset.

For an authenticated integration-capable Windows PowerShell/pwsh session, the plugin publishes one atomic root-shell activity snapshot with `HasRunningCommandOrChild`, `IdleAtPrimaryPrompt`, password, alternate-screen/TUI, editable-input-empty, and insertion-at-cursor capability. Close confirmation, Restart, pane follow, history insertion, and explicit path insertion consume that same generation; no surface independently guesses process activity from titles, process enumeration, or terminal cells. Cmd and WSL publish no authenticated activity/cwd snapshot in v1; explicit close still follows the conservative Starting/Running confirmation contract without guessing activity.

Plan 4's `LaunchCwdBootstrap` applies to every local/plugin-backed/UNC cmd launch
and is a correlated launch proof, not an ongoing authenticated semantic protocol.
This plan must revalidate its unchanged exact predecessor contract: reopened/
hash-verified extracted RCDATA `.cmd`; child tokens
`["cmd.exe","/E:ON","/V:OFF","/S","/K",
"call \"<verified-extracted-script-path>\""]`; the eight case-insensitive
reserved names `REDSALAMANDER_TERMINAL_CMD_V1_{PIPE,CAPABILITY,NONCE,
OPERATION_ID,LAUNCH_GENERATION,MODE,TARGET,TARGET_UTF16_SHA256}` after removal of
every inherited casing; and the protected inbound first-instance message pipe
ready before CreateProcess with exact-root-PID validation. After normal AutoRun,
the fixed script uses `setlocal EnableExtensions DisableDelayedExpansion`, the
literal local `cd /d` or UNC `pushd`, a safe-ASCII emit subroutine, pre-Ack clear,
`endlocal`, and a second clear that never restores inherited values.

The sole Ack remains fixed-case/order ASCII
`RSTCMD1;CAP=<64-lowerhex>;NONCE=<32-lowerhex>;OP=<32-lowerhex>;
GEN=<canonical-uint64>;MODE=<LOCAL|UNC>;TARGET=<64-lowerhex>;
STATUS=<0|1|2>;DRIVE=<-|[A-Z]:>\r\n`, at most 512 bytes. The server accumulates
complete message writes in a 513-byte staging buffer until clean EOF, then
parses; `ERROR_MORE_DATA`, byte 513, prefix/partial/extra/second-line, held-open
deadline, wrong PID/capability/nonce/operation/generation/digest, or replay fails
closed. There is no HMAC or Unicode cwd serialization: local success is the
reviewed primitive result bound to immediate target revalidation; UNC also
uses the master's exact proof: normalized `\\server\share` root plus empty-or-
backslash-prefixed tail, prelaunch `GetLogicalDrives` free-letter snapshot,
reported uppercase drive, `DRIVE_REMOTE`, and bounded
`WNetGetUniversalNameW(REMOTE_NAME_INFO_LEVEL)` with 4,096 initial bytes/one
exact `ERROR_MORE_DATA` resize/262,144-byte cap. Universal/connection/remaining
fields must be in-buffer/aligned/NUL-terminated. Normalize connection to
`\\server\share` without trailing slash. Remaining is empty or a relative
component sequence with zero/one leading slash; reject drive/UNC/two-leading-
slash/root-escape/above-share/NUL forms, canonicalize nonempty to one leading
slash while removing `.`/legal `..`, and require connection+remaining ==
universal == requested full plus separate root/tail ordinal-ignore-case matches.
Malformed pointers/sizes and provider/DFS aliases fail.
All child/plugin staging is absent before
`CmdLaunchRootedRunning`/input/profile memory, and cmd can never publish ongoing
activity/prompt/history/follow/insertion/native-history-suppression flags.
AutoRun cwd change is corrected; exit/hang/wrong result/cleanup failure never
reaches Running or memory. User AutoRun is trusted; correct-capability/correct-
nonce same-root preemption is explicitly indistinguishable/outside threat, while
wrong exact pipe/root PID/capability/nonce/stale/PTY attempts fail. Every cmd
remains terminal-only after launch.

UNC success retains only
`{drive,expectedRoot,expectedTail,launchGeneration,prelaunchLetterWasFree}` until
the exact root dies and never unmaps it live. The counted post-death worker
revalidates the same bounded universal-name tuple plus bounded
`WNetGetConnectionW`; already `ERROR_NOT_CONNECTED` succeeds. Only an unchanged
owned mapping receives one non-forcing
`WNetCancelConnection2W(L"D:",0,FALSE)`, followed by a required disconnected
recheck. Changed/reused/in-use/error mappings remain untouched with one content-
free cleanup failure and no retry/broader unmap. The plugin never issues `popd`;
cmd documentation guarantees removal only when `popd` executes, so root exit is
never assumed to remove the mapping and the retained-owner post-death
revalidation/cancel path is mandatory. Another same-user process reusing the
same prelaunch-free letter for the exact same logical target before cleanup is
indistinguishable through WNet and remains inside the declared same-user
limitation; changed-target reuse is left untouched, exact-target reuse is
documented rather than falsely proved, and covering it later requires an owned
mapping mechanism/handle.

PowerShell never uses the cmd bootstrap. Its mandatory local/UNC
`PowerShellStartupBootstrap` both proves post-profile rooting and, only after
that proof, makes the one decision that can deliver this plan's sole semantic
epoch. No launch-only frame is promoted into a semantic handshake; no second
post-Running epoch exists. This plan reruns Plan 4's exact startup matrix against
the durable settings path and requires correct local/UNC cwd with integration
enabled or disabled, zero app/native history for startup, Running CAS before
prompt evaluation, CommitAck before input/profile memory/capability, and zero
semantic capability for TerminalOnly/Unavailable/revoked outcomes.

### Foreground control admission and fixed expiry

This phase consumes the master's exact persistent protected `TerminalControl`
message pipe, `RSTCTL01` frame, one pending root-module BeginRead, and one
integration-owned PSReadLine wake key. Its current-user ACE and exact
  noninherited local-server client constructor request only
  `PipeAccessRights.ReadData|PipeAccessRights.ReadAttributes|
  PipeAccessRights.WriteAttributes|PipeAccessRights.Synchronize` with
`PipeOptions.Asynchronous`, `TokenImpersonationLevel.Anonymous`, and
  `HandleInheritability.None`; `WriteData` is absent, `ReadAttributes` is the
  least additional right required by the real PS5.1/pwsh constructor/connect
  path, and `WriteAttributes` permits the one `ReadMode=Message` change. Each
  supported PS5.1/pwsh tuple must prove that exact granted mask, absence of
  `WriteData`, successful Connect, and `ReadMode=Message` through that
  constructor/API path or control is unavailable. `FlushFileBuffers` and every synchronous
pipe Write/Flush are forbidden. The plugin first reserves the fixed frame buffer
plus immutable bounded payload without serializing the frame. Its off-UI control
worker then holds only the generation/input-operation gate, performs final state
validation, and atomically reserves one inert descriptor plus its eventual one-
byte wake from the fixed protocol reserve, assigning that placeholder its final
input-queue sequence. Failure of either reserve unit is a mechanically proved
pre-submit/no-action result. The worker samples `operationT0`, serializes/HMACs, sets
`deadlineQpc = checked(operationT0 + ceil(qpcFrequency/4))`, and submits exactly
one message-mode overlapped `WriteFile`. Nonpositive frequency/overflow rejects
before Write. `writer-admitted` is exactly submission returning `TRUE` or
`FALSE/ERROR_IO_PENDING` while that gate remains held, not placeholder
reservation; the gate is then released.

Before Handshake, the root module uses only the public PSReadLine key APIs to
select the first unbound exact wake candidate named by the schema and installs
the one integration-owned mapping without replacing a user binding. If none is
unbound, install/revalidation changes a prior mapping, or the mapping cannot be
proved exact, the adapter reports `SemanticUnavailable`; the closed nonzero wake
field forbids a partial semantic/history-only epoch.

If user input wins before placeholder reservation, it cancels the undispatched
operation. Once reservation wins, every later user descriptor receives a later
queue sequence and the writer stops at the inert placeholder. If reserve failure
or disable wins before submission, the placeholder becomes a consumed zero-byte
tombstone before later input becomes writable, no Write occurs, `controlSequence`
does not advance, and the shell effect is zero. A submission that wins is the sole pre-fence winner and may cause
at most one shell action strictly before its 250-ms embedded expiry even if
disable wins before completion/wake. The worker waits only for cancellable
OVERLAPPED completion off UI. The OVERLAPPED owns a noninherited manual-reset
event plus module pin. `WriteFile(TRUE)` is immediate and requires the exact
frame byte count; `ERROR_IO_PENDING` performs exactly one
`GetOverlappedResultEx(handle,&ov,&bytes,
floor(max(0,deadline-now) milliseconds),FALSE)`. No UI/INFINITE/alertable wait,
poll, or spin exists; a floor-rounded timeout with less than 1 ms remaining is
 conservative timeout. Success requires exact bytes and a post-return QPC sample
 strictly before the deadline before the already-sequenced placeholder is
 atomically activated in place with exactly one wake byte. Activation performs no
 second capacity check, admission, sequence assignment, or priority bypass; there
 is no retry/timer/second wake. Write
success never asserts client read. `WAIT_TIMEOUT`, equality, short count, or
  error revokes/classifies timeout, activates no wake, retains the placeholder as
  an inert held fence, and calls
  `CancelIoEx(handle,&ov)` without closing the handle; `ERROR_NOT_FOUND` in that
  cancel race does not alter disposition. The prohibition above is on a second
  admission/deadline wait. For retirement only, the off-UI drain waits exactly
  once on the owned manual-reset OVERLAPPED event for no longer than the remaining
  session `t0 + 5 seconds` quiet budget, then calls
  `GetOverlappedResult(handle,&ov,&bytes,FALSE)` exactly once to observe the
  terminal normal or `ERROR_OPERATION_ABORTED` state. This drain never admits a
  wake/effect/result or changes the already-failed disposition. Only after that
  observation may it close the persistent pipe handle and release the event,
  OVERLAPPED, frame, module pin, and any non-ambiguous input owner. Equality with or a miss of the
  quiet deadline quarantines/fatals and retains every owner and handle; it never
  frees/reuses/unloads reachable storage. A normal-success write already has
  terminal completion and releases only that write's event, OVERLAPPED, frame,
  and write module pin without a retirement wait. It retains the sole connected
  server handle for the entire epoch and retains the operation record, activated-
  placeholder/post-wake input hold, and later queued input
  until a matching authenticated `ControlResult` or ambiguity recovery. Disable
after submission activates no placeholder that had not already won the input
gate. A submitted failure/disable/timeout leaves the placeholder as an inert
held fence; explicit Discard/confirmed Close retires it without a wake. Input
hold release follows only the exact status classification below: validated
Success and the named safe pre-effect statuses release, a mechanically proved
no-submit/no-action branch releases as failure, and every ambiguous status/result
retains. Expiry prevents future action but cannot prove whether a submitted
Follow/Insert already executed; ambiguous effect/result—including after disable—
keeps input held, latches terminal-only/untrusted, and exposes only `Discard
queued input` or confirmed Close. Disable never auto-releases it. Cancel/close
  calls `CancelIoEx` for the exact OVERLAPPED and uses that same one-wait/one-
  `GetOverlappedResult` retirement sequence as counted quiet work independently;
  it never closes the handle before terminal status observation.

`Discard queued input` is plugin-owned and one-shot with exact identity
`{sessionId,incarnation,inputHoldGeneration,operationId}`. It is invokable only
after the strict expiry/admission fence has entered Retiring and made wake
activation irrevocably impossible. Under the generation/input-operation gate it
revalidates that identity/state, captures the current tail, makes any remaining
placeholder a permanent zero-byte no-wake tombstone at its original sequence,
tombstones and securely scrubs every held normal-user/paste/integration
descriptor through that tail, preserves required engine responses in FIFO,
releases the post-wake hold, and increments a no-wrap input-recovery generation.
The epoch remains terminal-only/untrusted; only newly admitted raw input may
flow. Late completion/result and stale/duplicate/wrong-generation actions are
inert, Close wins or follows through the existing retirement arbiter, and this
action releases no pipe/OVERLAPPED/module quiet owner.

The persistent channel states are exactly `Connecting -> Idle -> WritePending ->
AwaitingResult -> Idle`. Only `Idle` admits an operation. On exact successful
write completion, one input-operation-gate transaction first publishes
`AwaitingResult` plus the complete expected epoch/operation ID/control sequence/
deadline/result identity, then as its final release step activates the
placeholder and makes it runnable. The writer never emits the wake from
`WritePending`; an immediate child result therefore sees the awaiting record even
before the completion callback returns. A fence loss activates nothing and
retires the submitted ambiguity. A fully validated matching `Success`
result, or for Follow/Insert only a matching pre-effect
`RejectedState|BufferChanged`, releases the operation/input hold, advances the
no-wrap control sequence, and returns to `Idle`, where a later operation is
admissible; the latter two report safe no-action failure. A mechanically proven
pre-submit/no-action loss also returns to `Idle` without advancing the sequence.
Probe non-Success, `Expired`, `BindingChanged`, `ShellOperationFailed`, timeout,
malformed/mismatched result, protocol/policy failure, or unresolved ambiguity
moves to `Retiring`, closes through the exact cancellation/drain
contract, latches terminal-only, and bars every later operation; quiet success
reaches `Retired`, while a quiet miss reaches `Quarantined`/fatal shutdown.

Ordinary owned wrapper entries use nonwaiting `PollControlTransport` and may only
validate/stage an already-complete frame into the sole slot and rearm; they never
dispatch. On key-handler entry only, after exact owner-token, key-binding, epoch,
and lifecycle revalidation, an occupied slot is taken first and dispatched or
scrubbed once with no wait or `EndRead` on the newly armed read. Only with no
staged frame does an incomplete current BeginRead perform exactly one nonalertable
`((IAsyncResult)read).AsyncWaitHandle.WaitOne(250)` and no other wait, callback,
`Task.Wait`, `.Result`, blocking `GetAwaiter`, poll, or spin. A signaled/already-
complete read is consumed by exactly one `EndRead`; zero-byte EOF/read error
latches the child inert, while one complete frame is validated into the sole
slot and the next BeginRead is armed before the handler takes and dispatches it.
Unsignaled timeout performs zero action,
emits zero semantic/result frame, and does not itself revoke child integration.
A physical press therefore can stall only this root shell for at most 250 ms and
leaves integration live. Separately, after the plugin sends its single wake, its
missing-result/strict embedded 250-ms deadline retires/closes both pipes and
wakes the read wait. After a complete read, the final QPC sample and one-shot
consume must be
strictly before the embedded deadline; equality/later scrubs without action.
A completed pre-fence frame may be consumed once by the admitted wake or a
physical press of the integration-owned key; exact operation ID/sequence, strict
handler pre-effect deadline, and one-shot consume prevent replay. A staged frame
at/equality-after expiry is scrubbed without action before mapping removal/pass-
through. Exactly one injected wake bounds a mapping-replacement race: replacement
after child/plugin validation may invoke the later foreign handler at most once,
then missing result retires the epoch; no retry may invoke it again. After the
fence, every result/semantic frame is rejected and authorizes
no plugin state/history/capability. Close/disable tests must prove delayed/no-read,
BeginRead before/after wake, physical-key race, every pre-submit/submitted/
completed/pre-wake/post-wake boundary, input hold/release, `CancelIoEx`
completion, and module-unload gates.

### Live integration-policy transition

- A committed `shellIntegrationEnabled: true -> false` increments the one
  service-wide integration-policy generation and installs a non-failing fence
  before the configuration result publishes. If it wins before
  `SemanticDecision`, that launch simply samples false and chooses TerminalOnly.
  If it wins after SemanticEnable while startup is pre-CAS, mark
  `StartupSemanticJoin` `RevokedTerminalOnly`, permanently close semantic
  operation/capability admission, and scrub/ignore staged and late semantic
  records. Retain the already-created `TerminalSemantic` and `TerminalControl`
  pipes only as exact-root-PID/no-write cleanup transports; they may accept the
  client connections but the plugin submits no control Write, and disable must
  not make connect/setup fail merely to signal policy. Require an otherwise-
  valid Prepared Ready or Unavailable as cleanup proof, but no Handshake is
  required; missing Prepared still times out. The Running CAS closes both pipes
  (or independent retirement/protocol failure closes them
  earlier), and CommitPrompt selects TerminalOnly
  at the current strictly greater generation so the script latches any installed
  wrapper inert and removes the control mapping iff still exactly app-owned
  before Ack.
- A disable that wins after the Running CAS installs the generation/admission
  fence, immediately closes/cancels both pipes, drains all counted Connect/Read/
  parser/dispatch/OVERLAPPED/staging units, and suppresses capability publication.
  Every owned child entry polls the completed control BeginRead; EOF/read failure
  latches inert, while closing the semantic server breaks a pending/next
  WriteAsync. The sole already-submitted control operation may follow the exact
  pre-fence-winner rule above; no later wake/control admission occurs and every
  result is rejected. This remains true even if an already-built CommitPrompt
  said SemanticActive. Matching CommitAck acceptance
  is keyed to bootstrap/session/incarnation/deadline, not live policy generation:
  a valid Ack still closes/zeroes bootstrap
  state, opens ordinary input, and permits integration-independent Windows
  profile memory, while publishing zero semantic capability and rejecting every
  old-generation event. Disable alone never retires an otherwise valid startup.
- For an already committed Running incarnation, the same fence permanently
  invalidates its sole epoch under normal model/input serialization, clears the
  activity/cwd/input capability snapshot, stops history capture, discards the
  pending accepted-command candidate, and cancels pending/not-yet-writer-
  admitted follow and path insertion. Every callback, marker, completion,
  debounce, chooser, or worker result carries and revalidates the old policy
  generation and becomes inert.
- Ordered input whose admission already won before the revocation fence is not
  removed from the writer or reinterpreted. It may drain as pre-fence bytes, but
  its later output/result is untrusted for capability, follow completion,
  insertion authorization, or history. No post-fence command candidate can be
  created. The setting UI may truthfully report disabled once the generation
  fence is installed; it does not claim that already-written shell bytes were
  rolled back.
- A committed `false -> true` that wins before a Starting launch's one Decision
  may make that Decision SemanticEnable. Once the Decision was TerminalOnly, the
  incarnation remains terminal-only, consumes its CommitAck normally, and
  exposes content-free Restart required; there is no second Decision, live
  resource injection, hook installation, or re-handshake. A
  true-to-false-to-true sequence therefore cannot clear an incarnation's
  revocation latch.
- Cmd and WSL are outside this semantic policy state machine in v1. The setting
  creates, revokes, or restarts no cmd/WSL epoch because none exists; every
  current and future cmd/WSL incarnation remains terminal-only with all ongoing
  semantic capability flags false.
- Cmd's one-shot local/plugin-backed/UNC `LaunchCwdBootstrap` has its separate generation/nonce and
  remains allowed to finish initial cwd proof across either transition.
  PowerShell's mandatory RootRequest/RootResult phase likewise remains
  integration-setting independent. Neither can recreate semantic capability,
  preserve a pending history candidate, or bypass the revocation latch. Root
  result, the one Decision, policy commit, Running CAS, CommitPrompt, CommitAck,
  and Restart serialize through `TerminalService`/the session lifecycle fence so
  every equality has one deterministic winner.

## V1 matrix

| Profile | Launch | History | Follow | Insert |
|---|---:|---:|---:|---:|
| System cmd | yes, every local/plugin-backed/UNC launch uses the correlated one-use exact-cwd bootstrap; user AutoRun is trusted/outside same-root authentication | no — ongoing terminal-only in v1 | no | no |
| Windows PowerShell / validated pwsh | yes | exact `AcceptedInput` then matching next-root `RootReadReady` observer | trusted root activity/cwd plus native-history suppression | proven non-executing buffer operation |
| WSL bash/zsh/fish/elvish/nushell/unknown | yes, exact distro plus `--cd` | no — terminal-only in v1 | no | no |
| Unknown/unchainable/blocked or epoch-revoked PowerShell | yes | no; Restart required after revocation | no | no |
| Git Bash/MSYS2/Cygwin/custom Windows/first-class SSH | not v1 profiles | out | out | out |

### PowerShell accepted-line and root-activity observer

V1 observes the real root interactive read boundary. It never reconstructs a
command from keypresses, mutable cells, native history, prompt output, or the
app-to-shell control wake.

- After the user profile and PSReadLine module load, resolve the single effective
  `PSConsoleHostReadLine`. Install one global interposer with a random in-memory
  owner token only when it is the unshadowed `FunctionInfo` exported by
  PSReadLine and the complete exact adapter tuple above matches. A preexisting
  custom/replaced/ambiguous function or alias, missing member, unknown tuple, Vi
  mode, or unsupported host stays usable but terminal-only; never wrap, rename,
  invoke, patch, or repair it.
- The interposer's first executable statement is exactly
  `$lastRunStatus = $?`. Parameter binding, tracing, logging, prompt work, or an
  integration call may not precede it. Afterward only output-silent/nonthrowing
  private helpers run. Apply the tuple's exact `Set-StrictMode -Off` recipe and
  invoke exactly once either
  `ReadLine($host.Runspace,$ExecutionContext)` for
  `StrictModeOffThenReadLine2`, or
  `ReadLine($host.Runspace,$ExecutionContext,$lastRunStatus)` for
  `CaptureStatusStrictModeOffThenReadLine3`, passing the captured Boolean
  unchanged. No unlisted overload or call back through the audited function
  after integration work is allowed.
- The observer never wraps, replaces, evaluates, captures, appends to, or
  normalizes the user's `prompt`. It does not modify `$LASTEXITCODE`, `$Error`,
  current provider, prompt output, key bindings, `Get-History`, PSReadLine
  options, or the PSReadLine history file.
- At wrapper entry, after the host has evaluated the user's prompt and before
  blocking for input, emit exactly one authenticated `RootReadReady` with the
  next semantic sequence, preceding accepted sequence or zero, supported
  captured prior-command success, and classified filesystem cwd or explicit
  non-filesystem/unavailable value. With a matching pending acceptance it is
  that command's completion boundary; otherwise it is only a root-idle
  boundary. A matching event atomically changes
  `RunningRootCommand -> IdleAtPrimaryPrompt` and only then persists an eligible
  immutable candidate.
- Call the allowlisted static ReadLine overload exactly once. Propagate its
  exception or null result exactly and emit no acceptance. A non-null result
  must be exactly one `System.String` retained as the same object. After it
  returns and before returning that object to the host, capture classified cwd,
  strictly encode the complete string to UTF-8 once, emit exactly one
  authenticated `AcceptedInput` carrying exact bytes/cwd/profile/incarnation/
  epoch/sequence, and return the same string exactly once. Embedded CR/LF in a
  multiline return is payload data, not generic input, and is never normalized.
- The wrapper owns no acceptance until `AcceptedInput` enters the bounded
  protected `TerminalSemantic` pipe lane. Ordinary Enter at a trusted empty root prompt pessimistically
  makes app state non-idle while that event is pending; admitted
  `AcceptedInput` establishes `RunningRootCommand`. Missing/malformed/over-1-MiB
  UTF-8/effect-admission/sequence/epoch failure still returns the exact user
  line but permanently revokes semantics and remains conservatively busy until
  root exit or explicit Restart. It never guesses idle state or history.
- Empty, whitespace, leading-space, recalled/edited, Unicode, and multiline
  returns use the same acceptance/completion ordering. Blank text is not
  persisted and `historyIgnoreLeadingSpace` filters only the app candidate;
  neither changes what PSReadLine executes/stores. A syntactically invalid line
  that returns to the next root read completes as a failed accepted command and
  may persist under the same filters; the adapter does not parse or rewrite it.
  A native process or nested shell remains busy until root read resumes. Null/
  EOF, `exit`, process loss, replacement, epoch loss, or restart before the
  matching `RootReadReady` discards the candidate.
- `RedSalamanderControlV1` Probe/follow/non-executing-insertion wake runs inside
  the one outstanding ReadLine and must not make it return. It therefore cannot
  emit `AcceptedInput`, `RootReadReady`, or a history candidate. Only a later
  ordinary user PSReadLine AcceptLine binding and the resulting actual ReadLine
  return count as acceptance.
- The closure retains the exact installed scriptblock and owner token and
  validates integration state plus its exact control-key mapping on every owned
  entry. User replacement is never overwritten or restored. Replacement or any
  observer failure leaves the session conservatively busy with no history,
  follow, or insertion until Restart/new session; replacement while a line is
  being read cannot make that line a trusted completion.
- Live `shellIntegrationEnabled` disable or pipe/epoch revocation makes the
  already-installed wrapper a process-lifetime inert pass-through: it continues
  to invoke the same exact audited ReadLine recipe but emits/dispatches nothing.
  V1 never fabricates a `FunctionInfo` restoration. At the next safe handler/
  entry remove the unbound control mapping only when it remains exactly app-
  owned; preserve a later user mapping. A new disabled session installs no
  wrapper, and close/root exit destroys the runspace.

`Tools/Test-TerminalPowerShellAdapters.ps1 -AuditInstalled` is read-only and
reports exactly `supported|terminal-only-unknown-version|custom-hook|missing`;
it never edits a profile or manifest. A new PSReadLine/pwsh version is an
explicit adapter dependency upgrade: inspect the upstream entry definition and
assembly/API diff, generate comparison only under disposable `.build`, review
and deliberately apply one exact tuple/resource change, then rerun the complete
real Windows PowerShell/pwsh matrix before it can become `autoEligible`.

### Frozen WSL v1 terminal-only boundary

- Resolve only the reopened non-reparse System32 `wsl.exe` and launch exactly
  `[wsl.exe, --distribution, <catalog exact distro name>, --cd, <absolute
  profile-native Linux directory>]`. The configured distro default user and
  default shell run unchanged. Do not add `--exec`, `--user`, `-e`, a wrapper,
  shell command, second process, home/default-distro fallback, or any other
  token.
- V1 performs no pre- or post-launch WSL integration injection: no bootstrap is
  concatenated into argv or environment, carried through `WSLENV`, written into
  the PTY, installed in a Linux profile, or staged through a guessed
  `/mnt/<drive>` path. Do not stage/source/inject a script or infer a prompt from
  PTY bytes. This plan implements adapters only for the independently proven
  Windows PowerShell/pwsh rows above.
- A WSL session reaches `Running` only as the truthful launcher/preflight/
  ConPTY/process-admission state and remains
  `launchCwdTrust=RequestedUnverified` for its entire v1 incarnation. There is
  no WSL trusted-cwd event; every attempted `Verified|Diverged` transition is
  rejected. App history, authenticated activity/cwd, pane follow, history
  insertion, every explicit path-insertion mode, and WSL folder-profile
  reads/writes are unavailable while the raw terminal remains fully usable.
  A launcher error/root exit may follow briefly visible `Running`; it never
  retries, falls back, or writes memory, and tests must not require every
  asynchronous `--cd` failure to precede `Running`.
- A later WSL-integration feature requires a separate approved design before
  code changes. Its first gate must prove clean-default Bash plus profile-
  override/`exec`/early-exit cases; bootstrap before user-input admission; zero
  visible PTY/native-history input; no Linux-profile edit; preservation of the
  configured default user/shell and startup order; no `/mnt/c` assumption;
  bounded nonce/path-state erasure; and fallback to a fully usable terminal-only
  session. If it needs `--exec`, a wrapper, `WSLENV`, or another process, that
  later spec must explicitly replace the five-token invariant and freeze exact
  argv/environment/path/ownership/timeout/cleanup semantics. V1 carries no
  dormant or best-effort version of that experiment.

## Pane-follow transaction

- The following transaction exists only for a currently `FollowCapable`
  authenticated Windows PowerShell/pwsh adapter. Cmd and WSL sessions are never
  `FollowCapable` in v1: navigation may update ordinary source/view bookkeeping,
  but it queues no semantic operation, emits zero PTY bytes, and truthfully
  reports ongoing terminal-only capability regardless of
  `followPathWhenIdle`.
- `OriginalSourceKey{FolderWindowInstanceId, PaneSourceId}` navigation publishes committed generation/location only; both components are application-lifetime unique/non-reused, generation is monotonic within the key, and latest wins after debounce/off-UI resolution. The controller separately retains immutable `PhysicalHostKey`; follow never changes it. Detach tombstones before cancellation; stale two-window/destroyed-window results never bind.
- Apply only at trusted idle empty primary prompt, not password/alternate/TUI/running/partial input. Recheck immediately before dispatch and serialize against user bytes until trusted completion/failure.
- User input admitted before control-placeholder reservation cancels that undispatched attempt. Otherwise reserve the inert one-byte wake descriptor and its immutable FIFO sequence from the fixed protocol reserve before the pipe Write. Later user descriptors may queue only with later sequences; the writer stops at the inert placeholder and then the separate post-wake follow hold. Exact pipe completion activates that same descriptor with no second admission/sequence/capacity check, so the wake is written before every later user byte without bypassing FIFO. Only matching authenticated Success or the named safe pre-effect no-action result releases the post-wake hold. The embedded/missing-result 250-ms strict-before deadline alone governs action/result disposition and equality/timeout retires control immediately; the session `t0+5s` clock is only the cancellation/ownership quiet bound and never delays disposition or authorizes action. On timeout/lost trust the inert fence and held input are neither sent nor discarded automatically. Expose `Discard queued input` and `Close terminal`; Close uses the ordinary confirmed Running-session retirement path and held input is never replayed.
- Adapter generates operation-ID-tagged cwd change with shell-specific literal quoting/translation. It cannot change visible input, app history, native shell history, hooks, or user history policy.
- Suppress/confirm by operation ID and canonical trusted cwd, never command-text comparison. Until confirmation show pending; mismatch/failure pauses at actual cwd. Unsupported later locations pause and later supported navigation resumes.
- `Retry pane following` retries only the latest source generation after all capability/location/idle/input checks; disable it with reason while prerequisites fail. Later committed navigation may trigger the same revalidation.
- PowerShell/pwsh must prove native-PSReadLine-history-free cwd mutation without
  replacing the user's hooks or policy. Never alter DOSKEY, PSReadLine,
  `HISTCONTROL`, or user hooks to manufacture capability; cmd has no v1 follow
  transaction.
- Manual shell cwd changes update terminal context only; never navigate file manager. Source close detaches, never redirects.

The persisted plugin setting is `followPathWhenIdle` and defaults to `false`; the session action is `Follow folder navigation when idle`. A follow operation is eligible only when the shared activity snapshot says `IdleAtPrimaryPrompt=true` and `HasRunningCommandOrChild=false`. Busy means coalesce the latest source generation and wait; it is not an error and does not inject.

## Location identity

- Windows logical absolute normalization handles separators, extended spelling, lexical dots/root/trailing separator and ordinal-ignore-case. Never collapse symlink, junction, mapped/SUBST drive, or UNC alias to physical identity.
- Cmd local `cd /d`-equivalent and UNC `pushd` state exist only inside Plan 4's
  mandatory one-use `LaunchCwdBootstrap`; each result is revalidated against its
  exact expected logical location, with UNC additionally reverse-mapping the
  assigned drive. All launch state is erased before Running/input. It never
  becomes an ongoing cwd identity, history, follow, insertion, or Restart source.
- Matching `\\wsl.localhost\<distro>\...` and `wsl:<distro>:/...` unify; distro ignores case, Linux path preserves case/no Unicode folding. Windows drive and WSL `/mnt/<drive>` stay distinct.
- Plugin identity includes normalized `pluginShortId` plus validated logical backing launch directory. Full plugin ID/display-only path is never a cwd key. `displayPath` does not participate in equality.
- A trusted PowerShell non-filesystem provider pauses history/follow/restart fallback. Launch-folder fallback is allowed only before any explicit cwd report.

## Exact history capture and insertion

- Exact bounded input comes only from authenticated `AcceptedInput` carrying the
  exact `System.String` returned by the allowlisted root ReadLine interposer,
  strictly encoded once before return to the host. Under the model lock, copy
  one immutable candidate with exact text, trusted acceptance cwd, profile,
  incarnation/epoch, and sequence; persist nothing yet. Only its matching later
  authenticated `RootReadReady` persists it, using acceptance cwd so `cd` is
  filed where entered. A parse failure that returns to the next root read is a
  completed failed acceptance; overlap, hook replacement, observer/epoch loss,
  ambiguity, null/EOF, or process/shell exit first discards it. Generic OSC,
  keypress reconstruction, model anchors/cells, prompt text, control wake, later
  discovered ranges, or native history are untrusted and record nothing. Validate
  strict UTF-8/control/size bounds without wrapper or line-wrap normalization.
- Ignore blank, leading-space per setting, integration operation, alternate/password/ambiguous/unfinished commands. Never reconstruct from keypresses or import native history.
- Chooser filters current trusted cwd/profile and must remain the active owned surface. Insert closes it, restores terminal focus, then atomically rechecks tab/session/model/cwd/prompt/input/password/screen/follow generations; failed focus/recheck emits zero bytes.
- `InsertTextWithoutAccept` sends no Enter. Never send raw CR/LF generically. A
  PowerShell/pwsh insertion uses only its proven non-executing line-buffer API
  after the full recheck; unsupported multiline text is Copy-only. Cmd has no
  v1 history-insertion or explicit-path-insertion path. Partial input is untouched.
- Choosing/inserting records nothing until a later real ReadLine return emits
  authenticated `AcceptedInput` and its next matching root read emits
  `RootReadReady`. The control wake itself cannot create either event. Native
  Up/Down/Ctrl+R behavior is unchanged.

History insertion and Folder path insertion are separate commands:

- `ITerminal::InsertPath` accepts only a copied `TerminalPathInsertion`; it never accepts a host-quoted command string. The record begins with `sizeBytes`; consumers validate the current required prefix, accept larger records, and ignore unknown tails. Its semantic order is initiating source, initiating generation, initiating location, item generation, mode, reserved0, item location, parent location, display leaf, and reserved words.
- The initiating source triple is the command-originating live Folder source's key, committed generation, and copied current logical location. Immediately before a direct call, the host atomically revalidates `{initiatingSource, initiatingSourceGeneration, initiatingSourceLocation}` against that source and, for item modes, the focused item plus `itemGeneration`; a pending insertion repeats both revalidations immediately before eventual delivery. The plugin requires a nonzero initiating key/generation, copies the entire request before returning, treats generations only as host-owned correlation, and has no pane/item catalog.
- The target terminal may have a different immutable `originalSource`: the plugin never compares that target source or follow record with the initiating key, generation, or location. For an insertion-capable Windows target it independently translates each request location into that target's active profile namespace. An unsupported or cross-family initiating/item location returns `HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED)` with zero bytes; there is no substitution from the target followed location, launch location, authenticated cwd, home, or default distro. An exact-distribution WSL target admits only the quoted full native Linux path and degrades contextual mode to full without claiming prompt trust; every different/unknown distro and cross-family case rejects before admission with the localized terminal-only capability status and zero bytes.
- Every mode requires a supported initiating source location. `ContextualLeafOrFull` additionally requires nonzero item generation, supported item and parent locations, and nonempty display leaf; the display leaf is presentation-only and never supplies identity, translation, quoting, or PTY bytes. After translation the plugin derives the native leaf from item location, validates its parent against translated parent location, and inserts the adapter-quoted leaf only when that parent identity equals the authenticated profile-native filesystem cwd, otherwise the adapter-quoted full item path. `AlwaysFull` requires nonzero item generation and supported item location, with Unsupported/empty parent and null/zero display leaf, and always inserts the translated full item path. `CurrentDirectoryFull` requires zero item generation, Unsupported/empty item and parent, and null/zero display leaf; it translates and inserts the copied `initiatingSourceLocation`, never the target terminal's followed source or authenticated cwd. Every other combination returns `E_INVALIDARG` with zero bytes.
- For an existing terminal, revalidate `InsertCapable`, primary prompt, not-running, password/TUI/follow state, target instance/session/activity generations, integration-policy generation, and focus immediately before admission, in addition to the host's request-source/item checks. If typed input exists, insertion is allowed only when that adapter proves non-executing insertion at the current edit cursor; otherwise reject. Busy existing terminals never defer to a later prompt.
- A command-created new terminal may own exactly one pending insertion until its first authenticated input-capable prompt. The request is bound to target instance/session, initiating source/item, view, and integration-policy generations and is destroyed on chooser cancel, diagnostic/failure, initiating source/item change, tab replacement, close, or epoch loss.
- A Folder/Preview source must not open a WSL terminal for a pending insertion,
  because no v1 WSL incarnation can become insertion-capable. Existing WSL
  targets never defer the request and never create a replacement tab.
- Emit exactly one atomic ordered-input descriptor containing adapter-produced bytes, no CR/LF/Enter. Reject NUL/newline/unrepresentable/oversized input and queue exhaustion in full. Do not add the path to app/native history merely because it was inserted; only a later ordinary AcceptLine causing the audited ReadLine call to return, followed by exact `AcceptedInput`/matching `RootReadReady`, may capture the eventual command.

## State file and deletion fences

Use the master v1 schema and SettingsStore-derived isolated Debug/Release filenames. Required root fields include command-history generation/capture and shell-choice generation/capture UUID/Boolean pairs. A preferred profile requires update UTC and mutation UUID. Release state is one same-`AppId` family of at most four exact canonical `<AppId>-<Major>.<Minor>.terminal-state.json` active targets and at most 512 MiB across active targets plus their owned recovery siblings. Debug is a separate one-target family with a 256 MiB data cap. The bounded privacy journal is separately reserved overhead so cap recovery can always create it. Check every count/size with overflow-safe arithmetic before import, recovery, or publication; never create an unbounded fifth Release target.

Runtime validation is authoritative for the hard ceilings independent of user
settings: 32,772 UTF-16 code units for a complete profile ID, 33,037 for
canonical `locationKey`, 33,040 for `displayPath`, and 1 MiB of strict UTF-8 for
an accepted command. `locationKey` uses only the exact canonical `file:win:`,
length-prefixed `file:wsl:`, or length-prefixed `file:plugin:` form; malformed
lengths, noncanonical spelling, component +1, or key/location mismatch reject
the record. Exact maxima must survive catalog, chooser/`ExactProfile`, Open,
the applicable profile-specific eligible commit, history save/load/merge, missing-profile UI, and Restart
byte-for-byte. No layer truncates, hashes, aliases, substitutes a physical
target/display spelling, or silently drops only persistence. Aggregate
count/file/family byte caps may reject the complete mutation before any write
but never shorten an admitted identity.

`TerminalState.schema.json` is JSON Schema Draft 2020-12. Its `maxLength`
keywords count Unicode code points, so they are only coarse structural
backstops and are not proof of a UTF-16-code-unit or UTF-8-byte ceiling. Add the
exact nonassertive annotations `x-redsalamander-maxUtf16CodeUnits` and
`x-redsalamander-maxUtf8Bytes` to every applicable property. The runtime
validator enforces those annotations' frozen values, and the serializer admits
only values that pass both runtime validation and the schema. For each
applicable ceiling the closed corpus names and freezes
`bmp-boundary-both-pass` and `bmp-plus-one-both-reject`; each UTF-16-annotated
ceiling also has `astral-utf16-plus-one-schema-pass-runtime-reject`, and each
UTF-8-annotated ceiling has
`multibyte-utf8-plus-one-schema-pass-runtime-reject`. A
`schema-invalid-runtime-valid` case is forbidden. Tests assert each
named disposition; they do not claim blanket schema/runtime parity or that
Draft 2020-12 `maxLength` rejects every UTF-16/UTF-8 +1 case.

- Implement Plan 4's unchanged Windows-only `ITerminalPreferenceStore` with the plugin-service-owned lazy `TerminalStateStore` inside `Terminal.dll`, not `ApplicationContext` or a separately reachable global singleton. The immutable read and conditional write remain cancelable/asynchronous and generation-fenced; persistent/cross-process failures never fall back to the Plan-4 fake. Rerun Plan 4's complete Windows source-compatible profile-precedence and closed eligibility suite against this durable implementation: cmd records only after its correlated exact-cwd launch result plus Running CAS; ordinary PowerShell/pwsh only after accepted matching post-Running CommitAck; hidden PowerShell/pwsh only after no-fail host finalize; every earlier/failure/abort path writes zero. Rerun the WSL no-read/no-write cases before adding history or follow behavior. Exact WSL distro selection is source-derived and never reads or writes this store in v1.
- Use one same-user, cross-Windows-session named mutex per physical state target plus `Common::Files::LocalFileTransaction`; reuse shared child-process/test helpers. Its exact name is `Global\RedSalamander.TerminalState.<sha256>`, with no literal path and `<sha256>` exactly 64 lowercase hex. `Local\` or unnamed fallback is forbidden because two console/RDP sessions of the same SID can write the same LocalAppData target. Normalize the SettingsStore-derived target to absolute DOS spelling, `\` separators, lexical dot removal, and no trailing separator except a root, then obtain its ordinal-ignore-case key through the shared `TryMakePathKey(FileSystemPathIdentity::OrdinalIgnoreCaseForLocalFileSystem(), ...)`; do not add a terminal-local fold. An unrepresentable key makes persistent state unavailable/read-only rather than selecting a different lock. Hash with BCrypt SHA-256 over UTF-16LE `path key`, one UTF-16 NUL, the exact `ConvertSidToStringSidW` result, one UTF-16 NUL, and exact build flavor `Debug` or `Release`, with no terminating NUL.
- Create/open the global mutex with a protected, non-inherited security descriptor whose owner is the current user SID and whose DACL grants only that SID `SYNCHRONIZE|MUTEX_MODIFY_STATE|READ_CONTROL`; there is no Everyone, Users, Anonymous, or other-user allow ACE. Request those exact access rights, then query and validate the owner/protected DACL even when `ERROR_ALREADY_EXISTS`; a mismatched pre-existing object's owner/DACL makes persistence read-only/unavailable. Creating a global mutex does not require enabling `SeCreateGlobalPrivilege`; never enable that privilege. If the OS nevertheless returns `ERROR_PRIVILEGE_NOT_HELD`, `ERROR_ACCESS_DENIED`, an unsupported-namespace result, or DACL validation failure, expose a content-free persistent-state-unavailable diagnostic and STOP durable history/preferences for that physical target—never fall back to `Local\`, a different name, or unlocked I/O. The same-user local-attacker limitation from the integration threat model still applies; the DACL prevents accidental cross-user sharing, not a hostile process already running as that user.
- Every active-state publication runs only on the store worker and acquires the flavor-family mutex first, enumerates the complete bounded active/recovery family, then acquires only its target mutex under one hard two-second deadline; target-first publication is forbidden. `WaitForMultipleObjects` receives the cancellation event before each mutex, and after each mutex result the worker rechecks cancellation before touching state, so cancellation wins an acquisition race. Under both locks it reloads/merges the target, computes the exact serialized replacement delta with checked arithmetic, and revalidates journal absence plus every active/recovery count and byte cap. It publishes only when the post-replacement family remains within cap; otherwise it returns `HRESULT_FROM_WIN32(ERROR_QUOTA_EXCEEDED)` with no bytes changed. Active-target growth, recovery growth, import, Reset, and absent Arm therefore serialize on one family owner, exactly one boundary grower can commit, and no conforming writer creates even a transient over-cap family. Reads may remain target-local. Timeout makes that mutation fail and exposes generic Retry; it never mutates the in-memory durable view or blocks UI. An abandoned acquisition discards the loaded family/target snapshots, re-enumerates and reloads under the ordered locks, then may apply only a still-generation-valid pending mutation; it is never treated as a clean continuation.
- Maintain pending mutations tagged with loaded generations. Under the ordered locks reload disk, reject stale-generation mutations, apply current pending changes, deterministically prune, compute/recheck the exact family replacement equation, and publish. Never union a stale full snapshot. Same-generation preference conflicts use newest valid update UTC then mutation UUID ordinal; generation change dominates.
- Treat `entryId` as globally unique within a physical state file. Identical duplicate records may coalesce once; the same UUID with divergent folder/profile/text/timestamp is file corruption and follows preservation recovery. An owned recovery sibling is exactly `<leaf>.bad.<yyyyMMddTHHmmss.fffffffZ>.<lowercase-UUID>` and must be a same-directory, non-reparse, regular, single-hardlink file with actual-case canonical leaf spelling. Per canonical target leaf derived by removing that final exact suffix, allow at most four recovery siblings, each at most 64 MiB and at most 256 MiB aggregate, regardless of whether the corresponding active file exists, while also enforcing the flavor-family cap. Admit count/bytes `N` only when adding the current file leaves every bound at or below its cap; reject `N+1`. A safe fifth orphan sibling or safe orphan-derived aggregate above 256 MiB is `recovery-cap` and may be handled only by confirmed Reset within its separate repair ceilings; an unsafe alias or 257th complete-family recovery entry is `manual-cleanup-required`. Active-file absence never relaxes admission. Recovery files are never age-pruned, overwritten, renamed over, or silently evicted.
- Preserve a malformed/schema-invalid current-version file of at most 64 MiB only after acquiring the family mutex and then its state-file mutex, revalidating the complete family count/byte baseline, and creating a unique fail-if-destination-exists recovery name; a UUID collision chooses another UUID within that transaction. Corruption discovered under a file mutex releases it and restarts in family-then-file order; acquiring family after file is forbidden. Only after preservation succeeds may the store publish empty v1 state with both capture Booleans false and fresh generations. If the active file is oversized, unsafe/reparse/multi-link/nonregular, any owned-looking sibling is unsafe/aliased, preservation/publication fails, or adding the recovery would exceed a count/byte/family cap, leave the active bytes untouched and place the whole flavor family capture-off/read-only with content-free reason `incompatible|recovery-cap|family-cap`. A divergent current-process pending collision fails save; never choose by timestamp.
- A newer unknown schema, unknown format, or wrong flavor is byte-preserved and puts the whole flavor family capture-off/read-only. Ordinary saves, shutdown flushes, migration, toggles, clear/forget, and enable cannot touch or claim deletion of it. A compatible newer app may recover it; otherwise only the confirmed `state.reset-incompatible` transaction below may delete it.
- One-time import into a new Release version acquires family then the canonical
  ordered union of current target, exact newest verified-compatible source, and
  exact oldest numerically ordered verified-compatible retirement targets; each
  retirement candidate must have zero recovery siblings. Under those locks it
  re-enumerates/revalidates the complete family and computes the exact serialized
  imported bytes with checked arithmetic. Admission succeeds only when the
  complete current target plus every still-present old active/recovery byte fits
  the four-target/512-MiB caps before any retirement; old state is never deleted
  merely to make room. Failure returns
  `HRESULT_FROM_WIN32(ERROR_QUOTA_EXCEEDED)`, writes/deletes nothing, preserves
  every source byte, and exposes content-free `family-cap` Reset/Retry. After
  atomic durable publication it may make one non-gating best-effort pass over
  only the prevalidated oldest targets; automatic retention never removes an
  opaque recovery sibling or unsafe/unknown active bytes. A crash before/within
  publication leaves either current absent with all sources untouched or a
  complete current target plus all sources, both within cap. A crash/delete
  failure during later retention leaves the complete current target and only a
  prefix-subset retired, remains within cap, and is an ordinary capture-enabled
  family. It creates no durable cleanup marker or Retry/Reset debt. A later pass
  must freshly enumerate and independently select its then-oldest safe set; it
  never resumes a dead process's intent. A later absent-target import that no
  longer fits fails before mutation with `family-cap`. No crash state has a
  partial current target or transient fifth/over-byte family. Debug never
  imports, reads, or mutates Release.
- Add and corpus-test
  `Specs/Terminal/TerminalStatePrivacyTransaction.schema.json`. Its only runtime
  journal is exact
  `<AppId>-release.terminal-privacy-transaction.v1.json` for Release or
  `<AppId>-debug.terminal-privacy-transaction.v1.json` for Debug: JCS UTF-8,
  content-free, non-reparse, single-link, fail-if-unexpected, and at most
  256 KiB. The exact root is
  `{formatId="red-salamander-terminal-privacy-transaction",schemaVersion=1,
  transactionId,buildFlavor,kind,phase,scopeKeySha256,createdUtc,targets,
  recoveryLeaves}`. `kind` is exactly
  `clear-folder-history|clear-all-history|disable-history|
  forget-folder-shell-choice|forget-all-shell-choices|disable-shell-choice|
  disable-history-and-shell-choice|reset-incompatible`. The three disable kinds
  target respectively `{command-history}`, `{shell-choice}`, or the ordinal set
  `{command-history,shell-choice}`; the combined kind is one atomic family
  transaction and never two nested journals. Only a folder-scoped kind carries
  the lowercase SHA-256 of the canonical logical folder key in
  `scopeKeySha256`, otherwise it is null.
  `phase` is exactly `deleting|awaiting-settings-false`. Clear/forget remains
  `deleting`. A disable starts `deleting`, performs its complete family scrub,
  and removes the journal immediately only when its bound persisted plugin
  configuration already has every Boolean in the kind's targeted set false;
  otherwise it atomically
  changes to `awaiting-settings-false` until the source-specific persistence/
  restart rule proves every target false. Reset starts `awaiting-settings-false`
  iff either bound persisted privacy Boolean is true and otherwise starts
  `deleting`. A live continuation requires an exact newer manager-local
  observation matching its retained compensation expectation and every targeted
  field false
  (false/false for Reset). After process loss, `StartupLoad`—or a fresh Retry
  bound to the latest current status when no live expectation remains—never
  compares manager-local sequences across processes and may settle from the
  current canonical persisted all-targeted-false value only after full lock/journal/
  family revalidation. The journal is the cross-version false fence in either
  phase. `targets` is ordinal by canonical state leaf and contains at most four
  Release active-state records or one Debug active-state record for normal
  actions. Reset repairs a pre-existing cap violation and may contain at most 64
  active-state records; when the canonical current output target is absent, it
  may add exactly one synthetic current-output record with null expected
  generations. Its complete family enumeration admits at most 256 exact owned
  recovery siblings. Each target record is exactly
  `{stateLeaf,expectedCommandGeneration,nextCommandGeneration,
  finalCommandCaptureEnabled,expectedShellChoiceGeneration,
  nextShellChoiceGeneration,finalShellChoiceCaptureEnabled}`. Generations are
  UUIDs, except an incompatible Reset active target or the synthetic initially
  absent current output may have null expected generations; a nontargeted domain
  repeats its expected generation. Reset names every retained active file for
  deletion and assigns fresh next generations only to the current empty file it
  recreates. `recoveryLeaves` is the ordinal unique actual-case array of every
  post-lock-enumerated owned recovery leaf, empty or at most 256 entries. Each
  exact-grammar leaf is at most 512 UTF-16 units and uniquely derives its
  canonical target state leaf by stripping the final exact
  `.bad.<UTC>.<UUID>` suffix. The derived leaf may match a recorded active target
  or be an orphan whose active file is absent; no active-state record is
  invented. No path/content is stored and no recovery file may be deleted unless
  recorded. A 65th active target, 257th
  recovery sibling, unsafe/case-aliased owned-looking entry, enumeration error/
  deadline, or journal-size overflow is `manual-cleanup-required`; mutate
  nothing and never claim Reset.
- Add the family mutex `Global\RedSalamander.TerminalStateFamily.<sha256>`. Derive its 64-lowercase-hex digest with the per-file UTF-16LE/NUL/current-user-SID/build-flavor recipe but substitute the canonical Settings-directory/`AppId` family key; apply and revalidate the identical protected same-user DACL with no weaker fallback. New-version import/retention, recovery creation, reset, every privacy action, absent Arm, and every ordinary active-state publication acquire the family mutex first and enumerate/identity-check the exact bounded family under the shared two-second deadline. A multi-target family transaction then acquires the unique union of recorded active/synthetic target mutexes and target mutexes derived from every recovery leaf in canonical ordinal path order; a one-target publication/recovery/Arm acquires only that target mutex. Equality fails. An orphan recovery is therefore locked through its derivable absent-active target without fabricating an active record. Recovery creation and ordinary replacement revalidate the complete family baseline and exact checked replacement delta under both locks, serializing cross-target cap admission. Failure releases acquired locks in reverse order and mutates nothing. Every publisher rejects/reschedules while the family journal exists; no conforming publication may bypass the family lock or use target-first order. After a disable scrub or awaiting Reset journal is durably committed, release every per-file mutex in reverse order and then the family mutex before returning a staged result or allowing the Settings CAS; named locks are never held across host persistence. On the qualifying exact newer persisted Settings observation, reacquire the family mutex and then the ordinal unique union of journal-target and recovery-leaf-derived target mutexes under the same shared deadline, re-enumerate, and revalidate the unchanged `transactionId`, `kind`, `phase`, target records, recovery leaves, and safe identity set before a phase advance, deletion, or journal removal. Only journal-authorized already-completed deletion absences are admissible. Any mismatch, unrecorded/new owned entry, unsafe identity, repair-cap/manual-cleanup condition, or lock contention mutates nothing, retains the journal, and leaves the family degraded/capture-off.
- Execute history clear/disable, shell-choice forget/disable, combined history-and-shell-choice disable, and reset as one family transaction. First revalidate configuration/action/page generations, acquire the ordered locks, re-enumerate and identity-check every active/recovery file, require the post-lock set to equal the pre-lock set, and prove every active file is valid/migratable for a domain-specific operation; any incompatible member requires Reset and is never partly scrubbed. Reset instead proves its complete scope is safe and within the separate repair ceilings without parsing incompatible bytes. Gate every affected domain false in the initiating process (both domains for combined disable and Reset), cancel its pending mutations and chooser data, choose fresh generations, and durably publish the one journal before touching user-state bytes. Other processes gate before their next read/use/write when watcher or preflight detects the journal.
- For folder/all history clear or history disable, rewrite every retained active state with the requested history removed, a fresh history generation, and the journal's final history-capture Boolean while preserving shell choices. For folder/all shell-choice forget or shell-memory disable, do the symmetric rewrite while preserving valid command history. Combined disable atomically removes both domains from every retained active state, rotates both generations, leaves both capture Booleans false, and uses only `disable-history-and-shell-choice`; it never sequences or nests two single-domain journals. Every one of these purge-claiming operations deletes **all** exact owned `.bad.*` recovery siblings across the flavor family because opaque recovery contents cannot be selectively scrubbed. Reopen and verify every resulting/deleted identity and flush the directory through the shared transaction seam where supported. Clear/forget then removes the journal last, may restore its pre-action true capture setting, and reports success only after lock release. A disable whose bound persisted plugin value already has every Boolean in its targeted set false also verifies, removes the journal while locked, releases, and may then report success. Otherwise it atomically rewrites the journal from `deleting` to `awaiting-settings-false`, releases all locks in the order above, and returns `DisableFenceCommitted` plus the complete canonical all-targeted-false configuration compensation; this is staged, not action success. After the exact newer persisted all-targeted-false observation, it reacquires/revalidates through the common protocol, verifies that the recorded family remains scrubbed and every recorded recovery leaf remains absent, removes the journal last, releases, and only then reports success. Settings CAS/access failure, generation loss, acknowledgement failure, or any deletion, rewrite, verification, flush, reacquisition, or journal-removal failure keeps every affected capture gated false, hides all family data in every UI, retains the journal/content-free `deletion-incomplete` diagnostic, and cannot report success or re-expose cached data.
- `Reset incompatible terminal state` is plugin Preferences action ID `state.reset-incompatible` with context version 0 and the schema-declared destructive confirmation. Expose it only for a repairable incompatible, corrupt, oversized, recovery-cap, family-cap, or deletion-incomplete state when persisted Terminal configuration is current/representable; future-configuration quarantine instead keeps Reset hidden with its upgrade/reinstall diagnostic because v1 must not rewrite opaque future JSON. It is the only v1 path allowed to delete those bytes. Release Reset covers every retained Release active/recovery file; Debug Reset covers Debug only. It stages the journal, durably compensates both main-Settings privacy values false when needed, then deletes/verifies the full selected family and publishes only the current empty v1 target with both captures false and two fresh generations. It never touches the other build flavor, unrelated/unsafe files, or rewrites Settings itself. Only a later explicit false-to-true Preferences Prepare/Arm/Finalize fence may re-enable. A valid existing deletion-incomplete journal is first resumed to its verified terminal state and removed, then the serialized action begins a separate Reset journal; any resume failure ends without Reset success. Unsafe/malformed journal or repair-ceiling overflow disables Reset and shows manual cleanup.
- A crash or failure after journal commit is deletion-incomplete, never success. Startup makes one bounded off-UI idempotent resume attempt using the journal; a compatible older binary observes the same journal fence and cannot read, capture, write, or recreate family data around it. `awaiting-settings-false` never mutates state while the latest coordinator-supplied persisted plugin configuration has any Boolean in the disable kind's targeted set true, or either privacy Boolean true for Reset; it returns/retains the corresponding canonical all-targeted-false compensation and waits with no locks held. A live continuation settles only from its exact newer local observation and retained expectation. After process loss, current canonical all-targeted-false `StartupLoad`, or a fresh latest-status Retry with no live expectation, may settle only after the common family/file reacquisition and complete journal/identity revalidation; it never compares `hostLoadSequence` with the dead process. A restored A/B/A false snapshot is therefore safely settleable, while any restored/current targeted true retains the journal and canonical false compensation. A missing recorded target/recovery leaf is admissible only where the journal proves that deletion step already completed; any unrecorded owned active/recovery leaf is an unsafe post-fence identity change and blocks resume. If resume cannot complete safely, keep the family capture-off/read-only, hide cached data, expose content-free `Stored terminal data could not be deleted`, Retry Reset, and a Settings link. An unsafe/malformed journal or unsafe file identity requires manual removal and is never claimed fixed.
- Each process owns a cancelable overlapped `ReadDirectoryChangesW` watcher for the state-file directory. It filters every retained active leaf, exact owned recovery-name pattern, and the family journal ordinal-ignore-case across atomic replacement/create/delete/rename actions. A matching change invalidates affected caches and pending mutations before queuing a worker reload. Overflow/rearm failure/unavailable watching hides family chooser/action data immediately and uses a one-second off-UI generation/capture/journal poll until watch/reload recovery. Before History Copy/Insert and remembered-profile resolution/start, asynchronously reload/revalidate the active generation, capture Boolean, and absence of a blocking journal under lock, then marshal a matching result back and revalidate UI/session generations; stale/disabled state emits zero or chooses no profile without blocking UI.
- A stale same-version or older retained Release writer that observes the journal, a changed generation, or target absence under lock discards its pending mutation. Absence never authorizes replay into or recreation of a deleted target. Partial multi-file completion therefore cannot leak data into UI/capture or be reported as a family-wide purge.
- WIL/RAII owns the watcher handle/overlapped state. `TerminalStateStore` cancels with `CancelIoEx`, drains completion, and joins/releases on its worker; UI teardown never waits, and late reload results are store-generation checked.
- Orderly application shutdown stops store admission and captures monotonic `flushT0`, a precommit decision deadline `flushT0 + 2 seconds`, and a worker-quiet deadline `flushT0 + 5 seconds`. The plugin-owned store worker serializes/prepares a unique sibling, acquires family then target and completes the bounded enumeration/replacement-cap preflight within the remaining precommit time, reloads/merges/validates, writes and flushes staging, then reaches one explicit `Preparing -> CommitStarted` atomic fence immediately before `LocalFileTransaction::Commit`. Cancellation is checked first under the same serialization. The fence may succeed only while monotonic `now < flushT0 + 2 seconds`; exact equality belongs to timeout. If cancellation/deadline wins before the fence, set `PrecommitTimedOut`, forbid every later commit for that shutdown generation, unwind/delete staging through RAII, preserve the prior durable target byte-for-byte, discard pending mutations, release target then family, and emit the one content-free not-saved warning/status below.
- Once `CommitStarted` wins before the two-second fence, timeout/cancellation cannot claim that the old file was preserved and cannot abandon the call. The noncancelable atomic commit runs to success or failure under both ordered mutexes: success records the new durable identity; failure follows `LocalFileTransaction`'s atomic contract and reports the retained prior durable target. A return after `flushT0 + 2 seconds` is still the post-fence success/failure outcome, never a timeout. The UI/window teardown never waits, but `ApplicationContext` retains the store, its code, both mutexes, staging transaction, and worker in the application shutdown barrier until the worker releases target then family and posts a generation-checked quiet continuation. No UI-owned `std::jthread` destructor joins it and it is never detached. If store quiet has not occurred by `flushT0 + 5 seconds`, invoke the existing injectable `FatalShutdownPolicy` (production `RaiseFailFastException`, tests record it); unloading reachable worker code or blocking the UI is forbidden. This five-second ownership bound is independent of the two-second precommit decision and does not retry publication.
- Precommit timeout/failure emits exactly one content-free warning and posts localized generic `Terminal history changes were not saved` only if an owning surface still exists. Post-fence failure emits the same outcome; post-fence success emits neither. No command, cwd, profile, path, nonce, state payload, staging name, or mutex digest enters the log/status.
- State data, commands, paths, malformed JSON, state/recovery filenames, and mutex digests never enter logs, metrics, dumps, or test archives. Query/status may expose only bounded recovery count, aggregate bytes, and content-free reason `normal|incompatible|deletion-incomplete|recovery-cap|family-cap|manual-cleanup-required`. When conditions overlap, select exactly one reason with strict highest-first precedence `manual-cleanup-required > deletion-incomplete > family-cap > recovery-cap > incompatible > normal`; unsafe/repair-ceiling overflow never appears repairable, while a valid resumable journal dominates ordinary cap/incompatibility diagnoses.
- An absent Arm rejected only by its exact projected active-count/family-byte equation retains one content-free, in-memory `ProspectiveFamilyCapObservation` scoped to the same module-administration generation. It stores only the immutable failed-preflight family identity/generation, build flavor, projected active delta, exact projected replacement-byte delta, and one/two-domain bitmask—never a path, command, profile, configuration bytes, or Settings digest. The failed preparation still receives normal Finalize-Failed. A subsequent QueryStatus for that administration generation reacquires the ordered locks, revalidates current family identity, and recomputes the checked stored equation; while it still fails, it binds `state.family-reason="family-cap"` and repair-ceiling/configuration-qualified `state.reset-available=true` into that status token even when disk alone remains at its legal boundary. Reset revalidates both the token and prospective equation before mutation. Any family, configuration, Settings, availability, or administration-generation change, successful Arm/Reset, different failed-Arm shape, or `Close` clears/supersedes the observation. It is never persisted or restored after restart; restart reports only disk state, and retrying the same enable recreates it deterministically when the projection still cannot fit.
- Ordinary saves cannot change capture Booleans. Folder/all clear and forget use the family transaction's temporary false gate but finish with the pre-action Boolean; disable finishes every targeted field false and Reset finishes both false. During one Preferences Apply, one true-to-false privacy-field change selects its single-domain disable kind and two true-to-false changes select the one combined kind. A candidate that simultaneously changes one field false-to-true and the other true-to-false is rejected before any journal/state mutation with `HRESULT_FROM_WIN32(ERROR_INVALID_STATE)` and localized field guidance to Apply the disable first, then enable the other field; v1 never creates two journals or a circular pending-enable. Only explicit plugin-Preferences Prepare/Arm/Finalize with exact expected generation may request a false-to-true privacy transition; external hot reload cannot re-enable. Disable completes and verifies its family-wide durable purge/rotate/false transaction before requesting plugin configuration false, retains its cross-version journal until the exact newer persisted all-targeted-false observation is reacquired/revalidated, and only then completes the action.
- Enable uses the master `pending-enable` fence. Prepare snapshots exactly `Present` plus both current domain-generation UUIDs or the explicit `Absent` sentinel with no UUID; it retains that closed identity through Finalize, and only Arm may replace Absent once with generated identities. Arm acquires the family mutex and then the validated current-target mutex under the shared two-second deadline, enumerates/freezes the complete bounded flavor family, and proves/rechecks journal absence plus the exact Present generations or still-absent sentinel. For Absent, every privacy domain whose effective target configuration is true must be listed or Arm returns revision mismatch. Arm generates two distinct baseline UUIDs and computes the exact serialized empty/pending-state size with checked arithmetic. It may create the absent target only when post-create Release active count is at most four (Debug exactly one) and complete-family bytes are at most 512 MiB Release or 256 MiB Debug. It never retires an older or recovery-bearing target. A four-old-plus-absent Release family or byte-cap breach closes Arm before mutation with `HRESULT_FROM_WIN32(ERROR_QUOTA_EXCEEDED)`, content-free `family-cap` Retry/Reset status, no preparation consumption, and no Settings save. After that preflight, Arm creates one empty v1 state with both captures false, gives each listed domain its own next UUID distinct from both baselines and every other listed next UUID, leaves each unlisted domain at its baseline UUID/false, atomically includes the pending record in that first publication, and binds all generated identities into the armed preparation. A crash exposes either continued absence or one complete within-cap pending state, never a fifth/over-cap partial target. For Present it uses the revalidated generations. It writes/flushes/verifies the coordinator's target Settings digest, releases target then family, and only then reports Arm success or permits Settings CAS. No mutex, thread, module pin held solely for that mutex, or other live transaction owner spans host persistence.
- Finalize-Committed, Finalize-Failed, startup reconciliation, Retry, and `HostCompensationSave` continuation independently reacquire the family mutex and then the current-target mutex under the shared cancelable two-second rule, enumerate/freeze the complete family, reload/validate the exact durable state, and recheck family-journal absence. Before any pending completion/cleanup publication they compute the exact serialized replacement delta with checked arithmetic and revalidate every family byte/count cap; publication is atomic and releases current then family. An exact pending transition requires the retained transition ID, domains, expected/next generations, and target Settings digest. Committed with the exact host-observed current/representable target configuration true for every listed domain atomically rotates those generations, sets their captures true, and clears pending. If another process already performed exactly that transition, Committed recognizes no pending record, every prepared next generation/listed capture true, unchanged unlisted baseline generation/capture, journal absence, and the exact persisted target/configuration as idempotent `Enable` bit-0 success. Failed clears only the exact matching pending record and leaves capture false.
- No other missing/mismatched pending transition or unexpected true state may directly request canonical-false compensation. Finalize first releases both current-target and family mutexes, then restarts from family-first acquisition for the canonical ordinal target union required by the journaled disable transaction for every affected domain; it never expands the lock set or inverts order while holding current. Only after the cross-process journal/fence is durable and state is gated/scrubbed may it return bit 1 plus canonical false. `Finalize(Failed)` carries zero persisted identity and its Prepare-captured baseline may be stale: if pending already became true or mismatched, that family false journal remains `awaiting-settings-false` with bit 1 until a later exact persisted observation settles it; Failed never removes the journal or infers current Settings false from its baseline. Bit-0 Failed cleanup is limited to clearing the still-exact pending record whose captures never became true. If the durable false fence cannot be established, return the ordinary noncompensating failure, close admission/require reload, and never claim rollback. `HostCompensationSave` may consume a false expectation only while the matching journal/pending false fence—or exact completed false state derived from it—exists; it never consumes false while an affected capture is true with no durable family fence.
- Every process observing pending enable keeps the named captures false and waits only for the family mutex and then target mutex, never an in-memory owner. Under both locks, after the same complete-family freeze/replacement-cap checks, an exact latest matching current/representable target Settings/configuration true for every listed domain may complete through the same publication path; the initiating Finalize then takes the exact observer-completed idempotent-success path. An exact latest nonmatching observation conservatively clears the pending record while retaining false through that same family-gated publication, then accepts already-false policy or requests canonical false compensation for persisted true. A concurrent in-flight CAS must then conflict, or its Finalize releases both locks, restarts the family/ordinal-union false-fence transaction, and only then compensates. Stale/foreign observations return reload-required and select neither branch. UI success requires every stage; after crash/disagreement, an outstanding privacy journal, pending recovery, or persisted false wins, and stale Settings true never overrides disk false. The host treats plugin JSON opaquely and never performs these transactions.
- Keep the exact schema/runtime relationship above, the authoritative runtime UTF-16/UTF-8 checks, the absolute 64 MiB preparse ceiling, configured bounds, and two-second debounced off-UI ordinary saves. The debounce, ordinary family-then-target shared lock deadline, multi-target family-transaction shared lock deadline, shutdown precommit decision, and shutdown worker-quiet deadline are distinct bounds and use the exact outcomes above.

## Tests

- Consume and rerun plan 4's generic local/UNC PowerShell startup protocol as a
  ShellState prerequisite: exact four-token `-NoExit -Command` argv, normal
  profiles and module-qualified rooting before the script result, target and
  semantic material absent from initial argv/environment, one post-root latest-
  policy `TerminalOnly|SemanticEnable` Decision, original TerminalOnly ->
  Prepared TerminalOnly, and the production predecessor's mandatory Enable ->
  Prepared Unavailable terminal-only result. Plan-4 fake-only evidence must also
  prove both pre-reserved `StartupSemanticJoin` arrival orders and Enable ->
  Prepared Ready, but it is not evidence that a real adapter was staged or
  installed. This plan then installs the first validated production adapter and
  reruns the same state machine with real Enable -> Prepared Ready plus the real
  semantic handshake. In both predecessor and upgraded paths, retain ReleaseAck
  as noncommit, the sole Running CAS, exact four-row CommitPrompt matrix with
  only Decision generation/Prepared outcome retained, typed-finalizer validation
  before the one CommitAck Write, and no input/profile memory/capability before
  the server reads that complete Ack. Exercise every receipt just before/equal/
  after the one `launchT0 + 5 seconds` deadline and race Close/WindowClosing/
  Restart/root exit/timeout/ApplicationShutdown against ReleaseAck validation,
  Running CAS, CommitPrompt, finalizer, and Ack acceptance. Pre-CAS failure
  publishes no Running/memory; unrelated post-CAS lifecycle/protocol/root/
  deadline/wrapper-finalization failure retires already-Running with every
  startup gate closed. Prove the one decision-delivered epoch can be installed
  before Running only behind closed gates; there is no post-Running epoch,
  second Decision, or rekey. Cover unchanged-generation Unavailable terminal-
  only success and every Ready/Unavailable disable timing from Decision through
  Ack acceptance, including `RevokedTerminalOnly` no-write cleanup transport,
  inert wrapper/mapping cleanup, ordinary input/integration-independent memory,
  zero semantic capability, and no policy-only retirement.
- Correct Windows PowerShell/pwsh handshakes enable only proved flags. Fake-child/
  source-contract tests require Decision-only nonce/channel-generation transport, one copy into the
  root module closure, decision/local mutable cleanup before prompt commit, no
  initial argv/environment entry, and no nested-child inheritance. Prove the
  exact protected inbound `TerminalSemantic` pipe flags/DACL/root PID, charged
  pre-Decision 2-MiB server reservation plus pending Connect, post-connect exact-
  PID first Read/rearm-before-dispatch, one persistent connection, and allocation-
  free parser admission. For Windows PowerShell 5.1 and
  every shipped pwsh tuple, source/runtime and real-pipe tests prove the exact
  local-server Out/Asynchronous/Anonymous constructor, availability of
  `WriteAsync(byte[],0,length,CancellationToken)`, prompt-thread return without a
  synchronous large-write stall, one-WriteAsync message boundary/chunk assembly
  across a 64-KiB pipe, exact `RSTSEM01` envelope/HMAC/2-MiB cap, and the four
  closed payload layouts. Require canonical JSON, generated C++ constants, and
  embedded PowerShell constants to be byte-identical; cover every enum/flag/
  status/cross-field row, direct exact PS/PSReadLine versions and definition
  digest, channel generation/capability dependencies, root read cycles/completed
  sequence, namespace-plus-control-sequence operation IDs, no fifth Activity/Cwd
  kind, Common codec reuse, and the complete valid/invalid boundary corpus. Cover empty,
  64-KiB, 64-KiB+1, 2-MiB, and +1 frames; back-to-back small/2-MiB/small frames;
  a valid second ordinary frame versus fatal second Handshake; final-successful-
  Read message delimiter versus complete unexpected EOF and mid-frame EOF; server
  read delayed/stalled/absent/closed; first-chunk-to-final-read and client
  completion just before/equal/after their independent injected 50-ms deadlines;
  startup first-chunk absence at the strict five-second deadline; synchronous
  throw/task fault/cancel; close while pending; Task
  completion after close; retained-buffer zero/release only after observation;
  no overlapping write, retry, callback, helper thread, background runspace, or
  synchronous fallback; and startup-handshake failure becoming Prepared
  Unavailable. Explicitly stage a matching valid Handshake before the server
  deadline while the child observes WriteAsync at/equality-after its deadline:
  Prepared Unavailable must scrub/revoke it terminal-only, whereas malformed/
  mismatched/extra/TerminalOnly pairing is fatal and Ready still requires the
  match. Disable/Close/Restart/root-exit/teardown races prove bounded raw-
  shell return and no post-failure emission/capability. Any failed real-shell
  runtime/API/message-boundary proof classifies that tuple Unavailable rather
  than weakening the transport. Engine tags/PTY/OSC/APC/second pipe cannot
  authorize. Wrong/
  missing/replay/duplicate/gap/order/size/malformed, sequence wrap, and aggregate-
  effect admission failure permanently revoke the incarnation's sole epoch,
  purge pending work/history, reject every later marker and second handshake,
  show Restart required, and recover only after explicit Restart/new session.
  Ordinary OSC/SSH/nested fixtures fail closed. Nested ownership pauses top-level
  capability and only the next ordered authenticated root marker in the same
  valid epoch resumes it; nested integration and a nested fresh handshake are
  unavailable. Fake semantic input cells inserted by ordinary output between
  authenticated `AcceptedInput`/`RootReadReady` events cannot change persisted
  command text.
- Run the product observer through real Windows PowerShell 5.1 and every pwsh tuple shipped as `autoEligible`; a fake-only pass or environmental skip cannot close this plan. Prove the wrapper's first executable statement captures `$?`, both audited two-/three-argument static ReadLine recipes, unchanged true/false status, exact same-object non-null `System.String` return, nonempty/blank/leading-space/Unicode/multiline/parse-failure/native/nested/null/EOF/exit cases, unchanged `$LASTEXITCODE`/`$Error`/prompt/PSReadLine options/keybindings/native history, `AcceptedInput` admission followed only by matching next-root `RootReadReady`, and no acceptance from the separate control wake. Cover replacement before/during a read, unknown/custom/missing/Vi tuples, event loss/malformed/over-1-MiB/queue failure, and live disable/revocation: each required failure remains conservatively busy with no history/follow/insertion, while an owned installed wrapper becomes process-lifetime inert pass-through using the same recipe and never fabricates a `FunctionInfo` restore. `Test-TerminalPowerShellAdapters.ps1 -AuditInstalled` must classify every installed shell read-only and leave profiles, modules, manifests, and repository state byte-identical.
- Plan-4 local/plugin-backed/UNC cmd `LaunchCwdBootstrap` stays one-shot/
  nonsemantic: enabled and disabled launches reach the same correlated cwd;
  exact six-token argv/RCDATA hash/eight fixed key overwrite, setlocal plus post-
  endlocal clear, protected root-PID pipe, fixed semicolon Ack, 512/513-byte and
  EOF/held-open/`ERROR_MORE_DATA`/extra-message cases, local primitive binding,
  and the exact UNC matrix: prelaunch free-letter bitmap, remote drive, bounded
  universal/connection/remaining pointer-size-NUL validation, normalized full/
  root/tail equality, connection/remaining split variants, dot/dot-dot and every
  rooted/alias/DFS/boundary/+1 rejection, root-live no-unmap, post-death already-
  disconnected/exact cancel/recheck, changed/reused/in-use/error no-touch, no
  `popd`/exit-cleanup assumption, cleanup races, and documented exact-target
  same-user indistinguishability. AutoRun cwd-change/exit/hang cases prove failure/timeout/
  wrong-PID/capability/nonce/op/generation/digest/cwd never reaches Running or
  preference persistence; mixed-case inherited names never return and later
  launch-only frames are inert. Include the trusted-AutoRun correct-capability/
  correct-nonce
  limitation probe. Every PowerShell local/UNC launch instead completes
  `PowerShellStartupBootstrap`; only a SemanticEnable Decision can create/install
  its one epoch, and only accepted CommitAck can publish it. Cmd proves complete
  launch-state erasure and zero ongoing nonce/history/follow/insertion/activity/
  authenticated-cwd capability without modifying DOSKEY/AutoRun/prompt.
- WSL fake/live tests launch the exact distro/path with only the five frozen tokens and preserve the configured default user/shell. Assert no `--exec`/`--user`/`-e`/wrapper/extra process/fallback; no argv/environment/`WSLENV`/PTY/Linux-profile/staged-script injection or guessed `/mnt/<drive>` path; no semantic nonce/epoch or trusted marker parsing; permanent `RequestedUnverified` with every `Verified|Diverged` attempt rejected; raw terminal usability; zero app history/follow/activity/authenticated-cwd/profile-memory capability/read/write plus exact-distribution full-native-path-only explicit insertion; truthful UI; and post-Running launcher error/root exit without fallback or false memory. Do not require every asynchronous `--cd` failure to precede `Running`.
- Race committed integration disable before Decision construction/delivery,
  semantic-pipe creation/connect, adapter install, either join slot, Prepared,
  ReleaseAck, Running CAS, CommitPrompt construction/send/receipt, wrapper
  validation/finalizer invocation, Ack Write/server receipt/acceptance, and after
  acceptance, then through nested ownership, pending `AcceptedInput`, matching
  `RootReadReady`, debounce/path chooser, follow/insertion admission, worker
  callback, and shutdown. Repeat startup timings for Ready and Unavailable.
  Assert disable never makes semantic connect/setup fail merely to signal
  policy, both no-write cleanup pipes close at CAS or earlier only for independent
  retirement/protocol failure, and the revocation fence publishes before
  disabled success. Matching cleanup Ack remains keyed to bootstrap/session/
  incarnation/deadline rather than live policy generation; input and
  integration-independent memory open without semantic capability, pending work
  disappears, admitted bytes are not recalled and grant no later trust/history,
  installed owned wrappers are inert same-recipe pass-through with no fake
  restoration and control mappings are removed iff still exactly app-owned, and
  rooting remains independent. For the control channel, prove the current-user
  ACE and client request are exactly
  `ReadData|ReadAttributes|WriteAttributes|Synchronize`, contain no WriteData,
  and Connect plus `ReadMode=Message` work on real Windows PowerShell 5.1 and
  every shipped pwsh. Race delayed/no client read, BeginRead completing before/
  during the handler's sole nonalertable `AsyncWaitHandle.WaitOne(250)`, physical-
  key/no-operation dispatch, NoWait completing/staging/rearming in an ordinary
  wrapper immediately before the wake followed by handler stage-first one-shot
  dispatch with zero wait/EndRead on the new read, close/disable waking that wait, mapping replacement
  between validation/wake causing at most one foreign invocation, and
  disable at pre-submit/submitted/completed/pre-wake/post-wake boundaries plus
  Close/root exit/shutdown. Prove exact 250-ms before/equal/after outcomes,
  fixed pre-serialization reservation, fixed-protocol-reserve descriptor/one-
  byte placeholder plus final queue sequence before Write, no synchronous Write/
  Flush, one submission admission point, immediate TRUE versus IO_PENDING, the sole floor-
  rounded nonalertable `GetOverlappedResultEx`, deadline-1/equality/<1-ms floor,
  exact/short byte counts, same-descriptor activation with no second capacity/
  admission/sequence or priority bypass, exactly one wake/no retry/timer,
  `CancelIoEx` normal/
  aborted/ERROR_NOT_FOUND followed by the one retirement event wait and one
  `GetOverlappedResult` before close, persistent Idle/WritePending/AwaitingResult/
  Retiring transitions and successful multi-operation handle reuse, five-second
  drain equality/miss retaining every owner/handle and quarantining, no Write/
  zero shell effect when the fence
  wins pre-submit, user-before-reservation cancellation, fixed-reserve byte/
  descriptor boundary and +1, later-user FIFO barriers before/during/after
  completion, AwaitingResult identity publication, and activation; prove an
  immediate child result arriving before the completion callback returns sees
  AwaitingResult; cover pre-submit tombstone, post-submit held fence, at-
  most-one pre-fence action when submission wins, no wake
  after disable, result rejection, input release only on validated Success, the
  named safe pre-effect statuses, or proven pre-submit no-action, ambiguous/
  retiring statuses holding through Discard/Close, staged-at-expiry scrub,
  and `CanUnloadNow` false through every
  pipe/OVERLAPPED/staging/parser/dispatch quiet gate. Race enable and true-false-true
  at the same points: enable before Decision may select SemanticEnable; enable
  after TerminalOnly never sends a second Decision/handshake and reports Restart
  required. Repeat both toggles with cmd and WSL and prove they create/revoke/
  restart no ongoing epoch, inject zero semantic bytes, and remain terminal-only.
- Each advertised FollowCapable adapter leaves app/native history unchanged; unproved adapters label unavailable.
- Shared activity-state generation gives identical running/idle answers to close, Restart, follow, history insertion, and explicit path insertion. Cover empty/busy/partial/password/alternate prompt, user-byte race, rapid latest navigation, metacharacter quoting, trusted cwd match/mismatch, unsupported recovery, two-window source collision/tombstoned detach/stale destroyed-window result, and manual cd.
- The one end-to-end runtime/serializer identity corpus accepts a 32,767-unit canonical pwsh path as the complete 32,772-unit profile ID and the exact 33,037-unit location key/33,040-unit display path, rejects every runtime +1/malformed length, and proves byte-for-byte survival through chooser/`ExactProfile`, Open, preference, history save/load/merge, missing profile, and Restart with no truncation/hash/partial persistence. Separately cover the cmd launch-bootstrap-only local logical-cwd check and UNC pushd reverse map, WSL source/launch identity without authenticated-cwd publication, and non-filesystem PowerShell.
- Two-phase history uses admitted exact `AcceptedInput` followed by its matching next-root `RootReadReady`; cover parse-failure completion, overlap/replacement/epoch-loss/ambiguity/null/EOF/exit discard, acceptance-cwd attribution for `cd`, multiline history insert/no Enter/no control-wake acceptance/zero-byte stale cases/copy-only fallback.
- Explicit focused-path insertion: exact `416/8` ABI/offsets and mode cross-fields; contextual same-cwd leaf versus different/untrusted-cwd full; always-full; current-directory full; Windows/UNC success; exact-distribution WSL inserting only the quoted full native Linux path, with different-distro/cross-family targets rejecting with zero bytes; cross-family/provider rejection; adapter literal quoting for spaces/Unicode/metacharacters; partial-input capability; busy/incompatible/unsafe selected Terminal zero-byte/no-defer/no-new-tab; one-shot new-session creation only from selected Folder/Preview when a compatible target can become insertion-capable; queue/cap rejection; and proof that no Enter/history acceptance occurs. Include an initiating source equal to and different from the target's immutable original source, `SwapPanes` and source-navigation races before pending delivery, spoofed display-leaf cases that cannot affect inserted bytes, and Ctrl+Space with deliberately different initiating location, followed source, and authenticated cwd proving that only the copied/revalidated initiating source location is translated and inserted.
- Schema corpus covers both `TerminalState.schema.json` and `TerminalStatePrivacyTransaction.schema.json`, Debug/Release migration/isolation, bounds/pruning, identical/divergent duplicate UUIDs, corruption preservation with unique existing `.bad` collisions, and atomic failure. Every annotated string ceiling includes the closed named `bmp-boundary-both-pass` and `bmp-plus-one-both-reject` classes; each UTF-16 annotation also includes `astral-utf16-plus-one-schema-pass-runtime-reject`, and each UTF-8 annotation includes `multibyte-utf8-plus-one-schema-pass-runtime-reject`. The corpus asserts that no `schema-invalid-runtime-valid` case exists and proves the serializer emits only values passing both layers. Exercise recovery count/byte boundaries at `N` and rejected `N+1`, 64-MiB per-file, 256-MiB per-target aggregate, four-target/512-MiB Release-family data, one-target/256-MiB Debug-family data, separately reserved 256-KiB journal, overflow arithmetic, unsafe/reparse/multi-link/actual-case identities, preservation failure, no auto-eviction, concurrent corrupt recovery on different targets proving serialized family admission, exact suffix stripping and rejection of ambiguous/malformed recovery names, recovery-only families whose derived active targets are absent, Reset with an absent current output and its sole synthetic null-expected record, Reset through the 64th active/256th recovery object, and `manual-cleanup-required` at the 65th/257th or oversized journal. Race active growth on two targets, recovery creation against active growth, and absent Arm against growth of an old target at the exact family boundary; every permutation proves family-first serialization, exactly one grow commit, `ERROR_QUOTA_EXCEEDED`/no-write for the loser, and no transient over-cap observation. Cover every overlapping family-reason pair and representative three-way overlaps to prove exact precedence `manual-cleanup-required > deletion-incomplete > family-cap > recovery-cap > incompatible > normal` and the resulting Reset availability.
- Release-import tests prove
  family-before-canonical-current/source/retirement lock order, exact newest
  compatible source and zero-recovery retirement selection, checked
  imported-byte computation, and admission with every old source still counted.
  Cover active-count and family-byte boundary success plus +1
  `ERROR_QUOTA_EXCEEDED` with current absent, byte-identical sources, zero
  deletion, and `family-cap` Reset/Retry. Crash before/within atomic publication
  yields only all sources/current absent or all sources/complete current, within
  cap. Inject crash/delete failure after each best-effort retirement candidate
  and require the complete current target, only a prefix-subset retired, family
  still within cap, ordinary capture-enabled operation, no persisted cleanup
  intent or Retry/Reset debt, fresh-selection-only later cleanup, later-import
  no-write quota failure when retained state does not fit, and no opaque removal,
  partial current, or transient over-cap state.
- Real second processes with stale history/choice state cannot resurrect data after folder/all clear, forget, or disable in either the current or an older retained Release target. Every purge deletes all exact owned `.bad.*` siblings across the Release family, preserves only the valid opposite domain, restores the prior capture Boolean only for clear/forget, and leaves disable false. Watch and pre-action checks revoke cached chooser/action state on active/recovery/journal create/replace/delete/rename even with missed/overflowed notifications.
- Exercise upgrade `N -> N+1`, purge/disable/reset in `N+1`, then launch a still-supported `N` writer: no command or shell choice reappears, target absence never recreates old state, and stale pending work is discarded. For single-domain history/shell-choice disable and the combined history-and-shell-choice disable, start with every targeted persisted field true and cover one atomic scrub/journal commit, reverse file-lock then family-lock release before `DisableFenceCommitted`, Settings CAS/access failure, plugin generation/quarantine failure, process crash before/after persistence acknowledgement, older-binary startup, acknowledgement-time family/file lock contention, full reacquisition/journal-and-identity revalidation, identity mismatch/manual cap, journal-removal failure, and success; the combined case uses one journal, scrubs/rotates both domains atomically, and requires both persisted fields false before restart/HCS/Retry settlement. Inject failure/compensation after every combined-domain scrub boundary and prove no partial success or nested journal. Reject an opposing one-false-to-true plus other-true-to-false Apply before any file/journal/configuration mutation with `ERROR_INVALID_STATE` and the localized two-step guidance. No result is full action success until the source-specific all-targeted-false rule has settled and removed the already-scrubbed journal. Exercise a live newer exact expectation, post-crash `StartupLoad`, and fresh latest-status Retry with no live expectation; restored false and A/B/A false settle only after full revalidation without cross-process sequence comparison, while any restored/current targeted true retains the journal and false compensation. Also cover the direct path where every persisted target was already false: verify the scrub, remove the journal while locked, release, then succeed without compensation. Reset starts with both persisted Settings privacy values true and covers journal-before-save, reverse lock release before compensation, CAS failure, plugin generation/quarantine failure, process crash after save before acknowledgement, older-binary startup, acknowledgement-time lock contention/reacquisition, journal/identity mismatch, acknowledgement/delete/verify failure, and success; immediate startup of deleted older target `N` always sees Settings false or the still-present journal and never initializes capture true. Cover a partial-family failure at every active rewrite, recovery delete, identity verification, directory flush, current-target publication, phase rewrite, and journal removal boundary; no branch reports success, exposes cached data, or enables capture before idempotent resume removes the journal.
- Pending-enable tests cover Prepare `Present` plus two UUIDs versus UUID-free `Absent`; present and initially absent current targets; family-then-current shared-deadline acquisition; complete-family freeze; Arm-first versus family-action-first ordering; absent-to-present race rejection; target-config true-domain completeness; two distinct false/false baseline UUIDs; distinct listed next UUIDs; unlisted baseline UUID/false preservation; binding generated IDs into the armed preparation; exact first-publication size calculation with checked arithmetic; post-create Release active-count four/Debug-count one and 512/256-MiB family-byte boundary admission; four-old-plus-absent-current and byte-cap +1 rejection before mutation with `ERROR_QUOTA_EXCEEDED`, `family-cap` Retry/Reset status, unconsumed preparation, and no Settings save; no eviction of a recovery-bearing would-be retirement target; atomic crash outcome of absence or one complete within-cap pending state; creation of the exact content-free prospective observation, same-admin QueryStatus ordered-lock/equation revalidation and token-bound Reset, clearing on family/configuration/Settings/availability/admin-generation change, success, new failed shape, and Close, plus no persistence across restart and deterministic retry recreation; journal/generation revalidation; atomic first pending-false create/write/flush; reverse target-then-family release before Arm success and Settings CAS; absence of any retained lock/live transaction owner; independent Finalize-Committed, Finalize-Failed, startup, Retry, observer, and HCS family-then-current reacquisition, complete-family freeze, exact replacement-delta/cap check, atomic publication, and current-then-family release; exact transition/domain/generation/digest re-read; matching completion; Failed cleanup; release-both/restart-family-then-ordinal-union for mismatch false fencing; process crash before/after CAS; target/family lock timeout/access/abandonment; family-journal appearance before continuation; and unsafe/missing/changed state. Race observer, Finalize, and HCS completion/cleanup against exact-boundary growth and family privacy work; prove one-way lock order, no transient cap excess, no partial pending publication, and deterministic retry/fence outcome. Assert exact result bits: Arm and successful enable use bit 0; an observer completing the exact transition before initiator Finalize is idempotent `Enable` bit 0; only clearing a still-exact never-true pending record may use Failed bit 0; `Finalize(Failed)` supplies zero persisted identity, so a pending-already-true or mismatched branch durably establishes and retains the family false journal plus bit 1 until a later exact persisted observation settles it, never using the stale Prepare baseline to remove the journal; a failed independently reacquired cleanup when Settings already equals canonical false is bit 0 with `KeepPrior|OpenDegraded` only for the still-exact never-true pending record; persisted mismatched true may yield bit 1 only after the family false journal/fence is durable. Inject failure to establish that fence and require the ordinary noncompensating failure with no `HostCompensationSave` expectation. Race exact matching and nonmatching observers against the in-flight CAS; require observer-wins success or pending-false abort plus CAS conflict/fenced Finalize compensation respectively, never a capture-true/no-fence false acknowledgement. Stale/foreign observers mutate nothing.
- Emit and archive the content-free `terminal.history.family_lock_wait_us` metric for every active-state publication path; the bounded-flush perf case covers ordinary saves plus the active-active, recovery-growth, and absent-Arm/old-growth contenders without emitting paths, commands, profile data, mutex digests, or state leaves.
- Reset tests require the exact `state.reset-incompatible` context-0 action and destructive confirmation; cover future schema, wrong format/flavor, corrupt, oversized, recovery-cap, family-cap, deletion-incomplete, orphan-recovery-only, and absent-current-output families, including resume-old-journal-then-new-reset ordering. Prove the ordinal unique active/synthetic/recovery-derived mutex union, orphan-derived absent-target lock ownership, complete orphan deletion, crash/resume after partial orphan deletion, and rejection of a recovery identity/suffix change before resume. Successful Release Reset removes every retained Release active/recovery byte and leaves exactly one current empty false/false state without touching Debug; successful Debug Reset is symmetric and Release-isolated. Failure/unsafe identity retains capture-off/read-only state and content-free diagnostics; malformed journal, active/recovery repair-ceiling +1, and manual-removal cases disable Reset and never claim repair.
- Plan-4 source-compatible Windows-default preference precedence, disabled-memory, missing/incompatible-choice, and profile-specific eligible-commit cases rerun unchanged against `TerminalStateStore`: cmd exact-cwd launch result plus Running CAS; ordinary PowerShell/pwsh accepted matching post-Running CommitAck; hidden PowerShell/pwsh no-fail host finalize; and zero writes for every earlier/failure/abort branch. Include stale generation, watcher invalidation, global mutex timeout, abandoned-owner reload, cancel/UI responsiveness, and cross-process write races. Separately prove exact WSL source-derived distro selection performs no store read/write in enabled, disabled, Running, exit, restart, and stale-generation cases and cannot create or resurrect WSL profile memory. Validate the exact `Global\` name digest, same-SID protected DACL/owner, existing-object revalidation, two same-SID helper processes serializing one target, the namespace abstraction's cross-session identity, injected access/privilege/unsupported/DACL failures entering read-only state, and a hard assertion that no `Local\`/unnamed/unlocked fallback occurs.
- Validate the family-mutex digest/DACL independently, family-before-canonical unique-union lock order for recorded active/synthetic and recovery-derived targets including orphan-derived absent targets, one shared two-second deadline, reverse unlock on failure and before every staged Settings operation, no named lock held across Settings CAS, ordinal union reacquisition under that deadline, ordinary-writer journal rejection/reschedule, concurrent import/retention/purge/reset serialization, one bounded startup resume attempt after each crash point, and stale older-process invalidation. Inject initial and acknowledgement-time family-mutex/file-mutex access, timeout, abandonment, malformed/aliased/changed journal, target/recovery identity mismatch, orphan recovery create/rename/delete races, and partial acquisition; no weaker lock, unlocked I/O, mutation before full revalidation, partial mutation, or unbounded retry is allowed.
- Orderly shutdown tests place deterministic barriers before, exactly at, and after the `Preparing -> CommitStarted` fence. Before/equal timeout must win, forbid later commit, retain byte-identical old state, discard pending work, and log/status once without content. Commit started strictly before two seconds must run to an injected success or atomic failure even when it returns after two seconds; success yields the new durable identity with no timeout warning, failure retains old durable bytes and reports once. In every branch UI/window teardown remains nonblocking, the app retains worker/code/transaction ownership through quiet, quiet releases by five seconds, and an injected stuck precommit or commit operation reaches only the five-second `FatalShutdownPolicy` seam—never a detach, blocking destructor, late publication after precommit timeout, or unload.

## Verification

```powershell
$expectedRepositoryCommit = (& git rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0 -or $expectedRepositoryCommit -cnotmatch '^[0-9a-f]{40}$') { throw 'Cannot resolve exact repository commit.' }
$expectedEngineLockSha256 = (Get-FileHash -LiteralPath '.\External\terminal-engine.lock.json' -Algorithm SHA256).Hash.ToLowerInvariant()
.\build.ps1 -ProjectName TerminalTests -Platform x64 -Configuration Debug
.\build.ps1 -ProjectName SettingsSchemaTests -Platform x64 -Configuration Debug
$shellStateNativeManifest = .\Tools\Run-TerminalNativeTests.ps1 -Slice ShellState -Platform x64 -Configuration Debug -EvidenceRoot .\Specs\TestRuns -PassThruManifest
.\Tools\Verify-TerminalNativeEvidence.ps1 -Manifest $shellStateNativeManifest -ExpectedSlice ShellState -ExpectedRepositoryCommit $expectedRepositoryCommit -ExpectedEngineLockSha256 $expectedEngineLockSha256
.\build.ps1 -ProjectName RedSalamander -Platform x64 -Configuration Debug
.\Tools\Test-TerminalPowerShellAdapters.ps1 -AuditInstalled
.\Tools\Test-TerminalPowerShellControl.ps1 -PluginPath .\.build\x64\Debug\Plugins\Terminal.dll -RequireWindowsPowerShell51 -RequirePwsh
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter terminal_integration_ -TimeoutMultiplier 2
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter terminal_pane_follow_ -TimeoutMultiplier 2
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter terminal_state_ -TimeoutMultiplier 2
.\Tools\Run-AllTests.ps1 -Suite Commands -SkipBuild -CaseFilter terminal_history_ -TimeoutMultiplier 2
```

Expected: the exact ShellState `TerminalNative` manifest verifies with nonzero
unskipped shell-integration/pane-follow/path-insertion/state-store/schema cases
and all Commands cases pass. Every local/UNC Windows PowerShell 5.1 and shipped
pwsh tuple uses the exact mandatory startup bootstrap: initial argv/environment
contain no target or semantic secret, the post-root Decision alone delivers the
one epoch when enabled, the adapter is staged before Running behind closed
gates, the order-independent join and Prepared outcome validate, ReleaseAck is
noncommit, Running CAS precedes CommitPrompt/prompt evaluation, and only the
exact first-PSTypeName `RedSalamander.TerminalBootstrapFinalize.V1`, exact
ordered `Sentinel,Finalize` properties, exact `RSTBOOT-WRAPPER-READY-1` Sentinel,
and scriptblock Finalize's single complete CommitAck Write
with no Flush/retry/later operation—followed by the server's full valid read—
opens input/profile memory and publishes proved capability. Every receipt is
strictly before `launchT0 + 5 seconds`; equality times out, and unrelated post-
CAS failure retires already-Running with gates closed. TerminalOnly, same-
generation Unavailable, and every live-disable
timing remain usable terminal-only; no separate post-Running epoch, second
Decision/handshake, or policy-only retirement occurs.

Ordinary output cannot authorize. The same real shells pass the observer matrix
with the exact audited static ReadLine recipe, first-statement `$?` capture,
same-object return, `AcceptedInput` then matching `RootReadReady`, parse-failure
completion, zero control-wake acceptance, unchanged prompt/native state,
conservative failure, and inert-pass-through disable; the read-only installed-
adapter audit changes nothing. Permanent protocol revocation rejects every
second/live handshake and only Restart/new session rekeys. Nested integration is
unavailable while root capability resumes only at the next valid in-sequence
marker. Advertised PowerShell follow is native-history-free and confirmed;
eligible authenticated Windows contextual/full/current-directory and history
insertions write the exact quoted text but never send Enter and never create
`AcceptedInput` or history acceptance solely from insertion; busy, unsafe, and
cross-namespace requests emit zero bytes. Cmd remains usable but ongoing terminal-only after its
local/plugin-backed/UNC one-use launch bootstrap is erased. Complete profile/location/display identities
persist without truncation/hash and every schema/runtime boundary passes both
layers. WSL remains usable with exact distro/`--cd` and configured default user/
shell but has no injection/epoch/profile memory/semantic capability and remains
`RequestedUnverified`. Stale processes cannot resurrect and evidence contains
no sensitive content. A fake-only observer result or skipped required real shell
is failure.

## Child completion and move

- **Completion record:** `TERMINAL-HANDOFF-UNRESOLVED` — replace this line with the exact plan-4 closeout-commit/blob pair, evidence repository commits, predecessor PaneProfiles native/pane/Preview archive identities, this plan's ShellState native manifest/companion identities, and integration/history/follow/state Commands archive path/file SHA-256 values required below.

Do not mark this child complete until all of the following are true:

1. Consume the exact predecessor `Specs/Plans/Done/Terminal_PaneTabsProfilesAndLifecycle_2026-07-22.md`. In this plan's final completion record, name/recompute the accepted plan-4 closeout commit, that commit's Done-plan Git blob object ID, every predecessor evidence `repositoryCommit`, the exact PaneProfiles `TerminalNative` manifest/companion paths/digests, the exact terminal and Preview `Specs/TestRuns/<MachineHash>/Commands/<RunId>/` archives, and SHA-256 for each canonical `results.json`, `trace.txt`, and runner-created `run-all-tests-results.json`. Rerun `Tools/Verify-TerminalNativeEvidence.ps1` with the recorded repository/engine-lock identities; any missing required file, blob/digest mismatch, skipped selected case, non-pass, or implicit newest-run selection blocks entry.
2. Create exactly one x64 Release Commands run with `Run-TerminalCommandsEvidence.ps1 -Scenario ShellState -Platform x64 -Configuration Release -RepositoryCommit <exact-commit> -EvidenceRoot .\Specs\TestRuns -PassThruRunPath`, validate that exact path with `Tools/Test-TestRunArchive.ps1 -RunPath`, and retain the exact ShellState `TerminalNative` manifest/companion pair emitted above. Record the run ID, evidence `repositoryCommit`, path, SHA-256 for canonical content-free Commands `results.json`, `trace.txt`, and `run-all-tests-results.json`, and the literal no-perf reason where applicable, plus both native-evidence files. Rerun `Tools/Verify-TerminalNativeEvidence.ps1`; the named native shell-integration/follow/path-insertion/state-store/schema invocations and `terminal_shell_state_` Commands cases must be present, passed, and unskipped, with no sensitive payload in either archive.
3. Merge the durable plugin-owned all-local/UNC PowerShell startup-to-semantic
contract into the authoritative specs: target/semantic-secret-free initial argv/
environment, post-profile rooting, sole latest-policy SemanticDecision,
Decision-only nonce/control delivery, pre-Running adapter staging behind closed
gates, pre-reserved order-independent `StartupSemanticJoin`, Prepared,
ReleaseAck noncommit, lifecycle-locked Running CAS, exact four-row CommitPrompt,
exact typed object (first PSTypeName
`RedSalamander.TerminalBootstrapFinalize.V1`, ordered `Sentinel,Finalize`, exact
`RSTBOOT-WRAPPER-READY-1`, scriptblock Finalize) plus one-Write/no-Flush-retry-
later-op/server-full-read CommitAck,
the strict five-second deadline plus pre/post-CAS failure lifecycle, CommitAck-
before-input/memory/capability, one installed
epoch with no post-Running replacement, and live-disable cleanup/inert/EOF/no-
retirement semantics. Include the canonical closed four-kind RSTSEM01 JSON,
generated C++/embedded-PowerShell identity, exact layouts/state transitions/
bounds, and Common-codec rule; its capability/threat/activity contract; exact
manifest-gated `PSConsoleHostReadLine` interposer and audited ReadLine recipes;
first-statement `$?`; `AcceptedInput`/matching `RootReadReady`; prompt/control-
wake noninterference; conservative replacement/failure; inert-pass-through
disable; read-only adapter upgrade audit; permanent revocation/restart-only
rekey; and no nested adapter. Also merge truthful cmd ongoing terminal-only
behavior after every local/plugin-backed/UNC launch bootstrap, including AutoRun
cwd-change/exit/hang outcomes; two-phase history; contextual/full/
current-directory insertion; follow; untruncated identity caps; the exact-five-
token WSL terminal-only/`RequestedUnverified`/zero-capability boundary and future
reopener; TerminalState persistence/generation/per-file-and-family-lock/DACL/
bounded-recovery/family-privacy-journal/reset/precommit-fence/worker-quiet; opaque
per-plugin configuration/privacy fencing; and privacy contracts into
`Specs/Terminal/Terminal_EmbeddedPane.md`, `Specs/Terminal/TerminalState.schema.json`,
`Specs/Terminal/TerminalStatePrivacyTransaction.schema.json`,
`Specs/Core/Core_SettingsStore.md`, `Specs/SettingsStore.schema.json`,
`Specs/UI/UI_CommandMenuKeyboard.md`, and `Specs/UI/UI_FolderWindow.md`. Draft
2020-12 code-point `maxLength`, exact nonassertive UTF-16/UTF-8 annotations,
authoritative runtime checks, the closed boundary corpus, serializer-both-layers,
Release-family upgrade-purge-downgrade, and no command/cwd/path/profile/nonce in
evidence are completion gates.
4. Change this document's State from WIP to Done, retain the exact evidence paths/digests in it, and perform exactly:

```powershell
git mv Specs/Plans/WIP/Terminal_ShellIntegrationHistoryAndPaneFollow_2026-07-22.md Specs/Plans/Done/Terminal_ShellIntegrationHistoryAndPaneFollow_2026-07-22.md
```

5. In the same change, update `Specs/Plans/WIP/README.md` N3 to route next to `Terminal_SettingsSecurityPerfPackagingCloseout_2026-07-22.md`. A move or index update without its paired durable-spec, schema/corpus, and evidence updates is incomplete.

## STOP conditions

- Trusted markers cannot be distinguished from child output without a second VT parser.
- `TerminalSemanticProtocol.v1.json`, its closed schema/corpus, generated C++
  constants, or embedded PowerShell constants are missing or differ; a fifth
  semantic kind/TLV/trailing-byte extension, local UTF codec, unbounded field,
  implicit structure layout, or noncanonical payload can be accepted; or the
  exact four-kind state machine cannot be proved on real PS5.1/every shipped pwsh.
- The PowerShell observer would need to wrap or infer from `prompt`, keys, cells, native history, the control wake, or any nonallowlisted tuple/overload; `$lastRunStatus = $?` cannot be the first executable wrapper statement; the exact returned `System.String` cannot be preserved; parse failure cannot complete only at the next root read; or an observer failure/replacement cannot remain conservatively busy with history/follow/insertion disabled.
- Live disable/revocation would require reconstructing or restoring a fake `FunctionInfo`, or an already-installed owned wrapper cannot remain process-lifetime inert while calling the same exact audited ReadLine recipe and emitting/dispatching nothing.
- The read-only installed-adapter audit or the real Windows PowerShell 5.1 and every shipped pwsh tuple matrix cannot pass without a fake-only substitution or skip.
- A local/UNC PowerShell/pwsh launch could skip the exact mandatory bootstrap;
  place target or semantic material in initial argv/environment; deliver target
  outside authenticated RootRequest or nonce/channel-generation/pipe outside post-root
  SemanticEnable; alter the wrapper's seven fixed bootstrap literals; stage or
  install the sole adapter/epoch after Running, or create/install any second
  post-Running epoch; publish from one join slot, Prepared, ReleaseAck, or the
  Running CAS; run a prompt before Running; accept anything outside the exact
  four CommitPrompt rows; deviate from the exact typed object or single complete
  Write; allow Flush/retry/later pipe operation or gate from client Write
  completion; or open input/profile memory/capability before the full valid Ack
  is read and accepted.
- The PowerShell startup deadline could differ from `launchT0 + 5 seconds`, treat
  equality as success, fail to race Close/WindowClosing/Restart/root exit/
  timeout/ApplicationShutdown at the lifecycle CAS, publish Running/memory on a
  pre-CAS failure, or fail to retire an already-Running gated incarnation for an
  unrelated post-CAS lifecycle/protocol/root/deadline/wrapper-finalization
  failure. Live policy disable is the sole nonfatal exception and cannot itself
  retire startup.
- A semantic nonce could remain in a root environment, be inherited by a nested
  child, permit a second/live handshake after permanent revocation, or rekey
  without explicit Restart/new session; live disable could fail an otherwise
  valid cleanup startup, publish old capability, or reactivate after re-enable;
  or nested integration would be advertised in v1.
- The real PS5.1/shipped-pwsh control client cannot connect/set message mode with
  exactly `ReadData|ReadAttributes|WriteAttributes|Synchronize` and no WriteData;
  the channel needs synchronous Write/Flush, a second wake/retry/timer, lets
  NoWait wait/dispatch/overwrite its sole occupied stage, lets the handler inspect
  or wait on the newly armed read before consuming that stage, uses a handler
  wait other than the sole nonalertable 250-ms AsyncWaitHandle wait, handle close
  before terminal OVERLAPPED status observation, or release/unload before every
  persistent-channel/operation/input/quiet owner reaches its exact terminal state.
- Cmd would require ongoing trusted prompt/activity/cwd, history, follow, insertion, or a semantic nonce after its one-use launch bootstrap is erased.
- A complete admitted profile/location/display identity would be truncated, hashed, aliased, or silently omitted from persistence rather than rejected as one whole over-cap mutation.
- Runtime cannot authoritatively enforce the frozen UTF-16-code-unit and UTF-8-byte ceilings; Draft 2020-12 `maxLength` is treated as anything stronger than a code-point backstop; the exact nonassertive annotations or named schema-pass/runtime-reject corpus classes are absent; a schema-invalid/runtime-valid value exists; or the serializer can emit a value that fails either layer.
- An adapter is advertised as `FollowCapable` despite failing marker authentication, native-history suppression, input serialization, literal cwd transport, or trusted-cwd confirmation.
- History text would be inferred from keys or inserted with an accept terminator.
- Explicit path insertion would require host-side shell quoting, a cross Windows/WSL/distro guess, deferred insertion into a busy existing terminal, or any CR/LF/accept byte.
- A v1 WSL path would add or require `--exec`, `--user`, `-e`, a wrapper, another process, a bootstrap or `WSLENV`/environment/PTY/profile/staged-script injection; alter the configured default user/shell; create a semantic nonce/epoch; transition out of permanent `RequestedUnverified`; expose history/follow/insertion/activity/authenticated-cwd/profile-memory capability; or make terminal usability depend on integration. STOP and require the separately approved WSL-integration design and gate above rather than weakening the five-token contract.
- Cross-process or older-Release clear/forget/disable/reset data can be resurrected; any exact owned recovery sibling remains after a purge claim; the family transaction cannot prove durable all-target deletion; or partial-family completion can report success.
- Integration/history/follow/path/state logic cannot remain in `Terminal.dll`, or the host must parse/validate terminal plugin configuration instead of treating it opaquely.
- Any material baseline drift is unresolved; either exact same-user per-file/family `Global\` lock identity/owner/DACL/order/deadline cannot be created and validated without a weaker fallback; the bounded recovery/journal/reset contracts cannot be enforced without overwriting, silently evicting, or exposing opaque bytes; the precommit fence can publish after losing timeout/cancel; or orderly shutdown would wait on UI, detach/unload store work, or lack the five-second fatal quiet seam.
