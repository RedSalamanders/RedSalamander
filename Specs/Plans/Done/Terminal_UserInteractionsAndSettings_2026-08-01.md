# Plan 004: Integrate panes, commands, profiles, insertion, shell state, history, follow, and Preferences

> **Status — Completed 2026-08-03.** Ctrl+Enter/Ctrl+Shift+Enter path insertion, opposite-pane Terminal reuse, folder
> Edit, Windows/WSL profile selection, navigation follow, trusted prompt/history behavior, and all v1 Terminal
> settings/Preferences integration are implemented. Durable behavior is in
> `../../Terminal/Terminal_EmbeddedPlugin.md`; later WIP/checkpoint language below is historical.

## First-implementation outcome (2026-08-01)

**WORKING INTERACTION SLICE COMPLETE.** Folder/Preview/Terminal tabs,
opposite-pane open and live-session reuse, directory Edit, Windows/UNC and WSL launch mapping, Terminal-typed
Preferences, `Ctrl+Enter`, `Ctrl+Shift+Enter`, `Ctrl+Alt+T`, and `Alt+7` are
implemented. `Ctrl+Space` and `Ctrl+Shift+Space` are also migrated. The stable
`cmd/pane/bringFilenameToCommandLine` id is retained
with its new Terminal insertion meaning; the full-path action has a distinct
id. `Ctrl+Space`/`Ctrl+Shift+Space` retain the current-directory command id but
now insert through Terminal, and the obsolete bottom pseudo prompt is removed.

PowerShell/pwsh now publish authenticated live cwd/idle state through a random
exact-root-PID-bound duplex pipe. Contextual insertion uses that proof and
follow-when-idle applies out of band without ConPTY injection or native history.
Cmd/WSL and every unknown state use the conservative full-path behavior. Exact
accepted-command history and semantic close policy remain release-open gates;
this plan therefore remains WIP.

The working slice also requires an insertion request to match the target
terminal's original source key, latest source generation, namespace, and source
path. Cross-source insertion into a compatible selected terminal after the
future multi-tab/SwapPanes work remains unchecked below. The current stricter
failure is intentional and fail-closed, not the final Plan-4 interaction
contract.

The first interaction-hardening slice also exposes `hyperlinkPolicy` through
the existing generic plugin Preferences pipeline. `Terminal.dll` validates and
persists `disabled`, `ask`, or `open`; the host owns no hyperlink default or
activation logic.

`osc52Policy` and `osc52MaxBytes` now use that same opaque configuration
pipeline. `Terminal.dll` alone validates the deny/ask/allow policy and decoded
byte cap, owns the localized confirmation, and performs the UI-thread clipboard
mutation; the host never sees terminal clipboard content.

> **Executor instructions**: Execute after Plan 003 passes core/render/input/UIA
> evidence. This plan adds generic host integration but all terminal-specific
> process/profile/quoting/history/follow/settings behavior remains in
> `Terminal.dll`. Run every verification. Stop rather than duplicating terminal
> policy in the host. Update `Specs/Plans/WIP/Terminal_FirstImplementationExecutionIndex_2026-08-01.md` when done.
>
> **Drift check (run first)**:
> `git diff --stat aac3ed260..HEAD -- RedSalamander/FolderWindow* RedSalamander/Command* RedSalamander/Shortcut* RedSalamander/Preferences* RedSalamander/ManagePlugins* RedSalamander/TerminalHost Common/DxUi Common/SettingsStore.h Common/Common/SettingsStore.cpp Common/PlugInterfaces Plugins/Terminal Specs/UI Specs/Core Specs/SettingsStore.schema.json Specs/Plans/WIP Tests Tools`
> Expected earlier-plan drift is allowed. STOP if the command IDs, Folder/
> Preview content model, Settings Store schema version, or `ITerminal`
> IID/vtable and size-based record
> contract materially changed without reconciliation.

## Status

- **Execution state**: WORKING INTERACTION SLICE COMPLETE; advanced history/semantic-close gates open
- **Priority**: P1
- **Effort**: L (25-40 owner-days)
- **Risk**: HIGH
- **Depends on**: `Specs/Plans/WIP/Terminal_CoreRendererInputAccessibility_2026-08-01.md`
- **Category**: migration
- **Planned at**: commit `aac3ed260`, 2026-08-01

## Why this matters

The terminal becomes a product feature only when it fits RedSalamander's pane,
command, settings, accessibility, and persistence models. This plan implements
the requested interactions exactly, including opposite-pane tabs, contextual
path insertion, directory Edit, Windows/WSL profile selection, shell reuse,
navigation follow, history, and a complete Preferences surface, while removing
the obsolete pseudo command line.

## Current state

- Plan 003 provides a functioning `ITerminal` control with ConPTY, Ghostty,
  renderer, input, base UIA, session state/capabilities, public actions, and
  asynchronous close. This plan consumes those public surfaces only.
- The authoritative master plan already assigns stable command ID
  `cmd/pane/openCommandShell` to the embedded terminal and requires preserving
  legacy insertion IDs while changing their behavior.
- `FolderWindow` now models distinct Folder/Preview/Terminal content and reuses
  a live opposite-pane Terminal. A multi-Terminal collection and the advanced
  hidden two-phase pending-insertion commit remain WIP.
- `ViewerPluginManager` and viewer associations remain unrelated. Terminal is
  created through `TerminalPluginManager`/`IID_ITerminal` only.
- Settings Store is schema v16 and already owns opaque
  `plugins.configurationByPluginId`; terminal fields must remain opaque to the
  host. Generic schema-v2 rendering/asynchronous administration is additive.
- The former bottom pseudo command line and `cmd.exe /C` path were removed
  after all three stable insertion commands gained Terminal coverage.

## Frozen user interaction contract

| User action | Exact behavior |
|---|---|
| `cmd/pane/openCommandShell` (`Ctrl+Alt+T` and `Alt+7`) | Snapshot the focused logical Folder source and select a live Starting/Running `builtin/terminal` in the physical pane opposite focus, publishing the source's committed location while retaining the same child/process tree. Create a new terminal only when no reusable session exists. |
| `cmd/pane/openCommandShellWithProfile` (`Open terminal with...`) | Same route with `ForceChooser`; no default shortcut. |
| Folder `Edit` | If the focused item is a directory, use the same opposite-pane terminal-tab route. File Edit retains its editor behavior. |
| Windows folder | Resolve only compatible cmd, Windows PowerShell 5.1, and safely discovered local PowerShell 7+ according to remembered/default/chooser policy. If unavailable, show a diagnostic linking to `Preferences > Plugins > Terminal > Shells`; never silently choose WSL or an arbitrary PATH executable. |
| WSL folder | Launch exact `<System32>\wsl.exe --distribution <exact distro> --cd <native Linux path>`. Never choose another/default distro or a Windows shell. WSL has no app history, follow, remembered shell, or trusted contextual leaf in the first implementation; explicit insertion uses the quoted full native Linux path only within the exact launch distro. |
| `Ctrl+Enter` (`cmd/pane/bringFilenameToCommandLine`) | Act on the focused item only. If its translated parent equals the selected opposite Terminal's latest authenticated cwd under the profile namespace rules, insert the shell-quoted native leaf; otherwise insert the shell-quoted full native path. Never send Enter. |
| `Ctrl+Shift+Enter` (`cmd/pane/bringFullPathToTerminal`) | Insert the shell-quoted full profile-native path; never send Enter. It must be a distinct command/binding. |
| `Ctrl+Space` and existing `Ctrl+Shift+Space` alias (`cmd/pane/bringCurrentDirToCommandLine`) | Insert the source's full profile-native current directory through the same safe opposite-terminal route; never send Enter. |
| Opposite selected tab is a safe Terminal | Insert into that exact selected tab only. Never search hidden/background tabs. Busy/incompatible/untrusted/Starting/Exited/closing targets show a localized reason and receive zero bytes. |
| Opposite selected tab is Folder or Preview | The plugin may create one hidden `PendingPathInsertion` terminal. Reveal it only after engine/profile/integration readiness, final source/item revalidation, atomic insertion admission, and allocation-free host commit. Every failure leaves no visible tab and sends zero bytes. |
| Navigation follow | Persisted `followPathWhenIdle=false` by default. Coalesce navigation from the terminal's original source and dispatch only to the authenticated root PowerShell idle prompt with no pending user input. One-way only: pane navigation may change shell cwd; shell `cd` never navigates the pane. Cmd and WSL remain unsupported. |

Once the three legacy insertion commands route through Terminal, remove the
bottom pseudo command-line control, `cmd.exe /C` execution path, layout/state/
debug seams, strings, and tests. Preserve the stable command IDs to migrate
existing shortcuts; do not leave two implementations.

## Commands you will need

| Purpose | Command | Expected on success |
|---|---|---|
| Host/plugin build | `.\build.ps1 -Configuration Debug -Platform x64` | exit 0 |
| Pane/profile native | `.\.build\x64\Debug\TerminalTests.exe --suite pane-profiles` | all pass |
| Shell state native | `.\.build\x64\Debug\TerminalTests.exe --suite shell-state` | all pass |
| PowerShell control proof | `.\Tools\Test-TerminalPowerShellControl.ps1 -PluginPath .\.build\x64\Debug\Plugins\Terminal.dll -RequireWindowsPowerShell51 -RequirePwsh` | exit 0; required shell families pass |
| Commands selftests | `.\.build\x64\Debug\RedSalamander.exe --run-self-tests --test-filter terminal_` | exit 0; nonzero terminal case count |
| Focused Pester | `Invoke-Pester -Path Tools/Tests/TerminalCommandsEvidence.Tests.ps1,Tools/Tests/TerminalPluginBoundarySourceContracts.Tests.ps1,Tools/Tests/TerminalStateStore.Tests.ps1 -Output Detailed` | all pass |
| Full gate | `.\Tools\Run-AllTests.ps1 -Suite Full` | exit 0 |

Use the live selftest syntax registered by `Tools/TestRunPlan.ps1` if the binary
switch differs; do not treat zero discovered cases as a pass.

## Scope

**In scope host/generic files**:

- `RedSalamander/TerminalHost/TerminalPaneHost.h/.cpp` (new)
- `RedSalamander/TerminalPluginManager.h/.cpp`
- `RedSalamander/FolderWindow.h`
- `RedSalamander/FolderWindow.cpp`
- `RedSalamander/FolderWindow.Layout.cpp`
- `RedSalamander/FolderWindow.FileSystem.Navigation.cpp`
- `RedSalamander/FolderWindow.Viewers.cpp`
- the existing `RedSalamander/FolderWindow.*.Commands.cpp` files owning Open
  Command Shell, Edit, and insertion dispatch
- `RedSalamander/CommandRegistry.h/.cpp`
- `RedSalamander/ShortcutDefaults.h/.cpp`
- `RedSalamander/ShortcutManager.h/.cpp`
- `RedSalamander/ShortcutText.h/.cpp`
- `RedSalamander/ShortcutsWindow.h/.cpp`
- `RedSalamander/CommandDispatch.Debug.h`
- `Common/DxUi/DxUi.h`
- `Common/DxUi/DxUi.Controls.cpp`
- `Common/WindowMessages.h`
- `Common/PlugInterfaces/Host.h` (new if absent)
- `Common/PlugInterfaces/PluginConfiguration.h` (new if absent)
- `Common/PluginConfiguration.h`
- `Common/Common/PluginConfiguration.cpp`
- `Common/SettingsStore.h`
- `Common/Common/SettingsStore.cpp`
- `RedSalamander/PluginConfigurationCoordinator.h/.cpp` (new)
- `RedSalamander/SettingsHotReload.h/.cpp`
- `RedSalamander/Preferences.h/.cpp`
- `RedSalamander/Preferences.Internal.h/.cpp`
- `RedSalamander/Preferences.Dialog.h/.cpp`
- `RedSalamander/Preferences.Plugins.h/.cpp`
- `RedSalamander/Preferences.Plugin.Configuration.h/.cpp`
- `RedSalamander/ManagePluginsDialog.h/.cpp`
- `RedSalamander/RedSalamander.rc`, `RedSalamander/resource.h`, and satellites
- `RedSalamander/AppTheme.h/.cpp` and shipped terminal theme keys

**In scope plugin files**:

- `Plugins/Terminal/TerminalPlugin.h/.cpp`
- `Plugins/Terminal/TerminalService.h/.cpp`
- `Plugins/Terminal/TerminalSession.h/.cpp`
- `Plugins/Terminal/TerminalProfileCatalog.h/.cpp` (new)
- `Plugins/Terminal/TerminalLocation.h/.cpp` (new)
- `Plugins/Terminal/TerminalPathInsertion.h/.cpp` (new)
- `Plugins/Terminal/TerminalInteractionBroker.h/.cpp` (new)
- `Plugins/Terminal/TerminalShellIntegration.h/.cpp` (new)
- `Plugins/Terminal/ShellAdapters/PowerShellReadLineAdapters.json` (new)
- generated private PowerShell adapter constants/resources
- `Plugins/Terminal/TerminalActivityState.h/.cpp` (new)
- `Plugins/Terminal/TerminalFollowCoordinator.h/.cpp` (new)
- `Plugins/Terminal/TerminalStateStore.h/.cpp` (new)
- `Plugins/Terminal/TerminalHistory.h/.cpp` (new)
- `Plugins/Terminal/TerminalConfiguration.h/.cpp` (new)
- `Plugins/Terminal/TerminalConfigurationMutationBackend.h` (new)
- `Plugins/Terminal/TerminalResources.rc` and satellites

**In scope tests/specs/tools**:

- `RedSalamander/SelfTest/Commands/Commands.SelfTest.Terminal.cpp` (new)
- matching existing Commands selftest registration files
- `Tests/TerminalTests/` pane/profile/shell/state/settings cases
- `Tests/PluginContractTests/PluginConfigurationTests.cpp`
- `Tools/Test-TerminalPowerShellControl.ps1` (new)
- `Tools/Test-TerminalPowerShellAdapters.ps1` (new)
- `Tools/TerminalCommandsEvidence.psm1` (new)
- `Tools/Run-TerminalCommandsEvidence.ps1` (new)
- `Tools/Tests/TerminalCommandsEvidence.Tests.ps1` (new)
- `Tools/Tests/TerminalStateStore.Tests.ps1` (new)
- `Tools/Tests/TerminalPluginBoundarySourceContracts.Tests.ps1`
- `Specs/SettingsStore.schema.json`
- `Specs/Core/Core_SettingsStore.md`
- `Specs/UI/UI_CommandMenuKeyboard.md`
- `Specs/UI/UI_KeyboardManagement.md`
- `Specs/UI/UI_PreferencesDialog.md`
- `Specs/UI/UI_DxUiWinUIDesign.md`
- `Specs/Plans/WIP/Terminal_PaneTabsProfilesAndLifecycle_2026-07-22.md`
- `Specs/Plans/WIP/Terminal_ShellIntegrationHistoryAndPaneFollow_2026-07-22.md`
- `Specs/Plans/WIP/Terminal_EmbeddedConPtyLibGhosttyVt_IntegrationPlan_2026-07-13.md`

**Out of scope**:

- any terminal shell/profile/quoting/history/follow/default interpretation in
  EXE/Common;
- WSL shell integration, WSL path insertion, WSL history/follow, SSH/Git Bash/
  MSYS/Cygwin/custom profiles, or automatic import of shell-native history;
- automatically executing inserted text or sending Enter;
- searching background tabs for an insertion target;
- making Terminal a viewer association or Preview implementation;
- adding an engine/runtime path, update button, or version override to settings.

## Steps

### Step 1: Generalize each pane to stable dynamic content tabs

1. Replace fixed Folder/Preview index/Boolean assumptions with stable IDs and
   `PaneContentKind = Folder|Preview|Terminal`. Preserve existing Folder and
   single-Preview behavior exactly.
2. Add one `ITerminal` reference and ownership-marked child HWND per terminal
   tab. The host may parent/layout/show/hide/focus it after validation; it never
   owns/destroys the HWND or calls plugin-private code.
3. Add optional tab glyph/fallback icon, cached layout, close target, DPI/theme/
   RTL/high-contrast behavior without changing title-only callers.
4. Project `TerminalViewState` generations into generic tab title/status/
   close/restart state. Reentrant/stale callbacks revalidate instance and host
   view generation; exactly one removal claim mutates the tab/MRU.
5. Implement hidden terminal prepare/finalize records. Prepare allocates all
   tab/UIA resources while hidden. Finalize is allocation-free, nonblocking,
   and cannot fail after plugin Commit succeeds.

**Verify**:

```powershell
.\.build\x64\Debug\TerminalTests.exe --suite pane-host
.\.build\x64\Debug\RedSalamander.exe --run-self-tests --test-filter terminal_pane_
```

Expected: Folder/Preview regressions and Terminal add/select/hide/close/MRU,
eight-tab overflow, DPI/theme, stale callback, destructive reentrancy, hidden
commit/abort, and HWND ownership cases pass with nonzero counts.

### Step 2: Implement source classification and profile discovery in the plugin

`TerminalProfileCatalog` owns this exact matrix:

- Windows local/UNC/plugin-backed filesystem locations: compatible cmd,
  Windows PowerShell 5.1, and safely discovered local pwsh 7+.
- `windowsDefaultProfileId=auto|pwsh|windows-powershell|cmd`; `auto` tries newest
  supported auto-eligible pwsh, Windows PowerShell, then cmd. Explicit family
  choices never silently fall through.
- Discover system executables from canonical Windows APIs/approved locations;
  do not trust PATH, App Paths, aliases, HKCU, or parent-process inheritance as
  automatic eligibility. An explicitly user-chosen validated pwsh may become a
  stable exact profile only after the frozen eligibility handshake.
- Recognize `\\wsl.localhost\<distro>\...`, `\\wsl$\<distro>\...`, and
  `wsl:<distro>:/...`; select only the exact normalized distro and Linux path.
  Resolve only `<GetSystemDirectoryW()>\wsl.exe`; use argv
  `[wsl.exe,--distribution,<exact-name>,--cd,<linux-path>]`.
- WSL catalog work is cancelable and serialized: 3-second default, 10-second
  hard maximum, completion must be strictly before deadline, timeout wins at/
  after deadline, cancellation wins equal timestamp, stale generations lose.
- Archive/cloud/MTP/remote/non-filesystem locations show an unsupported
  diagnostic and a Terminal settings deep link where useful.

Selection modes are exact: `ForceChooser` always chooses; `ExactProfile` never
substitutes; `ResolveDefault` uses valid compatible remembered choice, then the
compatible configured Windows default, else chooser. WSL never reads/writes
folder shell memory.

**Verify**:

```powershell
.\.build\x64\Debug\TerminalTests.exe --suite profiles-locations
```

Expected: Windows/UNC/plugin/WSL spelling/case/path traversal, discovery/dedupe,
default/remembered/missing/chooser, deadline/cancel/stale, unavailable settings
link, and no-cross-family fallback cases pass.

### Step 3: Reuse the shell command and route directory Edit

1. Keep stable `cmd/pane/openCommandShell`; give both `Ctrl+Alt+T` and `Alt+7`
   to the same `ApplicationGlobal` action.
2. Add `cmd/pane/openCommandShellWithProfile`, label `Open terminal with...`, no
   default shortcut, `ForceChooser` route.
3. Preserve a separately named `cmd/pane/openExternalCommandShell` only if the
   current external-shell behavior remains intentionally accessible; give it
   no conflicting default binding.
4. Snapshot the focused logical source and committed location, resolve the
   physical opposite pane, create through `IID_ITerminal`, and open a new tab.
   From a focused Terminal, use its immutable original source association rather
   than terminal cwd to avoid reverse-follow coupling.
5. Intercept Edit only when the focused item is a directory and route through
   the same path. A file continues through editor associations.
6. Commands are asynchronous after synchronous validation/child creation; no
   shell discovery/process work or plugin load wait blocks the UI turn.

**Verify**:

```powershell
.\.build\x64\Debug\RedSalamander.exe --run-self-tests --test-filter terminal_open_
```

Expected: opposite-pane, repeated independent tab, focus kinds, same/other
window, directory Edit, file Edit unchanged, disabled/missing plugin, unsupported
source, chooser, and no-UI-wait cases pass.

### Step 4: Migrate all path insertion and remove the pseudo prompt

1. Preserve stable `cmd/pane/bringFilenameToCommandLine` for `Ctrl+Enter` but
   rename its user label to `Insert focused item in terminal`.
2. Add distinct `cmd/pane/bringFullPathToTerminal` for `Ctrl+Shift+Enter`; remove
   the duplicate old binding to the contextual command.
3. Preserve `cmd/pane/bringCurrentDirToCommandLine` for Ctrl+Space and its
   existing Ctrl+Shift+Space alias.
4. Host marshals immutable initiating source/item generations and typed logical
   locations only. Plugin translates to target profile namespace, compares
   contextual parent with latest authenticated cwd, quotes for the target shell,
   revalidates activity/input generations, and admits one no-Enter descriptor.
5. If selected opposite content is Terminal, target only it. Unsafe state emits
   zero bytes and an exact localized reason.
6. If selected content is Folder/Preview, permit one hidden pending open per
   physical host. Plugin owns profile/start/integration/readiness. On
   `ReadyToCommit`, host revalidates source/item, prepares fallible UI resources,
   calls plugin Commit, then allocation-free Finalize. Every abort tears down
   hidden child/session and creates no visible/UIA tab.
7. Only PowerShell-family sessions advertising authenticated non-executing
   insertion are insertion-capable in v1. cmd and WSL remain truthful terminal-
   only targets and receive zero insertion bytes.
8. After all three commands pass Terminal tests, delete the pseudo command-line
   HWND/model/layout/debug state, `cmd.exe /C` execution, resources, and old
   tests. Source-contract tests reject resurrection.

**Verify**:

```powershell
.\.build\x64\Debug\TerminalTests.exe --suite path-insertion
.\.build\x64\Debug\RedSalamander.exe --run-self-tests --test-filter terminal_insert_
rg -n -i "pseudo.*command|command.*pseudo|cmd\.exe.*\/C" RedSalamander Common | ForEach-Object { throw "obsolete pseudo prompt remains: $_" }
```

Expected: contextual same-parent leaf, different-parent full, always-full,
current-dir, quoting, Unicode/UNC, focus-only, busy/incompatible, hidden commit/
every abort race, no-Enter, and source/item stale cases pass; final search emits
nothing outside historical specs/tests explicitly allowlisted.

### Step 5: Implement truthful PowerShell shell integration

#### Implemented first-implementation decision

The working first implementation deliberately uses a smaller contract than the
advanced adapter design below:

- one encoded script is carried by the root PowerShell/pwsh `-EncodedCommand`;
- one cryptographically random plugin-owned duplex named pipe is created with
  `FILE_FLAG_FIRST_PIPE_INSTANCE` and rejects any client PID other than the
  exact root process;
- the root publishes strict UTF-8 `Get-Location.ProviderPath` at most twice per
  second from `PowerShell.OnIdle`; the plugin accepts only bounded absolute
  Windows/UNC paths without NUL/CR/LF;
- the newline reply is either empty or the newest safe follow target. The
  event action applies it with module-qualified `Set-Location -LiteralPath`, so
  no bytes are injected into ConPTY and PSReadLine history is unchanged;
- typed input marks the prompt non-idle conservatively until the next
  authenticated report; cmd and WSL publish no trusted semantic capability.

This is sufficient for authenticated cwd, conservative prompt-idle state,
contextual leaf/full insertion, and history-free pane follow. It does **not**
claim exact accepted command text, child-process ownership, password/TUI state,
semantic close approval, or app command history. Those capabilities require the
allowlisted PSReadLine adapter/protocol below and remain release WIP. The larger
design below must not be partially layered onto the working channel; adopt it
only as one reviewed replacement when a remaining release gate requires it.

Do not infer prompt/cwd/history from terminal text or keystrokes. Implement an
authenticated, versioned PowerShell 5.1/pwsh adapter with independent
capabilities: trusted prompt, trusted cwd, exact accepted input, non-executing
insertion, integration-owned cwd change, and native-history suppression.

- Pin supported PowerShell/PSReadLine tuples in
  `PowerShellReadLineAdapters.json`; generate C++ and embedded script constants
  from one source and test byte identity.
- Put no semantic nonce in initial argv/environment. Establish process/rooting,
  then deliver a one-epoch secret over the private control channel and accept
  only ordered, authenticated messages for that incarnation.
- Interpose only the allowlisted root read/accept boundary, preserve `$?`,
  `$LASTEXITCODE`, prompt, key handlers, `Get-History`, and PSReadLine history.
- A protocol violation, nested shell/app ownership, unexpected version, nonce/
  sequence mismatch, or integration disable permanently revokes capabilities
  for that incarnation; Restart is required to rekey.
- cmd gets only one-use launch-root proof and no semantic capability. WSL gets
  exact launch and permanent `RequestedUnverified` cwd trust in v1.
- `shellIntegrationEnabled true -> false` installs a live generation fence,
  cancels pending history/follow/insertion, scrubs secrets, and ignores late
  messages before Apply reports success. False -> true is new/Restart only.

**Verify**:

```powershell
.\Tools\Test-TerminalPowerShellAdapters.ps1
.\Tools\Test-TerminalPowerShellControl.ps1 -PluginPath .\.build\x64\Debug\Plugins\Terminal.dll -RequireWindowsPowerShell51 -RequirePwsh
```

Expected: manifest/corpus and real PowerShell 5.1 plus at least one supported
pwsh complete authenticated accept/completion, cwd, follow, and non-executing
insertion proofs; prompt/status/native histories are unchanged; evidence is
content-free. A missing required family is failure, not skip.

### Step 6: Implement bounded private state and command history

Add a separate high-churn state file, never a main Settings rewrite per command:

- Debug/ASan: `%LOCALAPPDATA%\RedSalamander\Settings\<AppId>-debug.terminal-state.json`.
- Release: `%LOCALAPPDATA%\RedSalamander\Settings\<AppId>-<Major>.<Minor>.terminal-state.json`.
- Debug and Release never read/import each other. Release migration may inspect
  only compatible older Release state through the versioned rules.
- Use schema v1, atomic replace, cross-process family lock, revision/generation,
  corruption quarantine, bounded recovery siblings, exact ownership, and stale-
  writer fences. Never store runtime paths, nonce, raw terminal output, or
  unaccepted commands.
- Remember shell by canonical launch folder only after the profile/purpose
  eligibility point. Trusted later cwd never changes another folder's choice.
- Persist a command only after authenticated AcceptedInput plus matching later
  root-read completion. Attribute to acceptance cwd/profile. Never infer from
  keypresses or import PSReadLine/DOSKEY/bash history.
- Defaults: 100 commands/folder, 10,000 total, 256 folders, 90 days, 64 KiB
  command, 16 MiB file. Apply count/age/size/LRU at load, acceptance, settings
  change, and save. Consecutive dedupe follows exact text/timestamp/entry-ID
  ordering; duplicate IDs are idempotent only when byte-identical.
- History UI filters current trusted cwd/profile. Insert uses the same safe
  non-executing gate and never accepts the line; multiline is Copy-only unless
  proven safe. Do not hijack Up/Down/Ctrl+R.
- Disabling/clearing history or shell memory is a family-wide durable purge and
  stale-writer fence before success. Incompatible opaque recovery siblings are
  deleted wholly; failure remains disabled/read-only and offers confirmed Reset.

**Verify**:

```powershell
Invoke-Pester -Path Tools\Tests\TerminalStateStore.Tests.ps1 -Output Detailed
.\.build\x64\Debug\TerminalTests.exe --suite history-state
```

Expected: migration, corruption, crash/atomic replace, cross-process merge,
all bound/+1 cases, accepted/completed ordering, nested/revoked privacy, clear/
disable stale writer, Debug/Release isolation, chooser Insert/Copy/UIA, and no-
content diagnostics pass.

### Step 7: Implement navigation follow from the original source

Each terminal retains one application-unique original source key. The host sends
only committed typed source updates with monotonic generation. The working
slice already coalesces the newest safe Windows/UNC target, pauses on pending
input or absent trust, replies through the authenticated PowerShell idle pipe,
waits for the next exact cwd report, and creates no native history. The
following bullets describe the remaining advanced-adapter extension. The plugin:

- samples persisted `followPathWhenIdle` (off by default) at Open/Restart; per-tab override
  never persists and ends at Restart;
- coalesces to the newest location using `followPaneDebounceMs`;
- dispatches only when the adapter proves trusted prompt/cwd, integration-owned
  cwd change, native-history suppression, `IdleAtPrimaryPrompt=true`, and
  `HasRunningCommandOrChild=false`;
- serializes operation IDs, rechecks empty prompt immediately before write,
  permits no user-byte interleaving, waits for authenticated canonical cwd, and
  creates neither app nor native history;
- pauses on partial input, password prompt, nested shell, running child, alternate
  screen, unsupported location, ambiguity, or lost source; exposes status and
  Retry/Detach where frozen;
- never sends shell cwd back to navigate the pane.

Explicit close uses the same activity snapshot but is intentionally not the
inverse of follow: every Starting/Running session confirms when configured,
including a trusted idle prompt.

**Verify**:

```powershell
.\.build\x64\Debug\TerminalTests.exe --suite follow-navigation
```

Expected: default/override/Restart races, rapid coalescing, idle apply,
partial/running/password/TUI pause, source swap/close, unsupported-then-supported,
disable revocation, stale operation, exact cwd confirmation, no history, and
one-way behavior pass.

### Step 8: Expose every setting through generic Preferences

Keep Settings Store schema v16 and store one opaque plugin value at
`plugins.configurationByPluginId["builtin/terminal"]`. The host renders the
plugin's schema-v2 sections/fields/status/actions but owns no defaults or
validation. Add a `Terminal` entry under `Preferences > Plugins` and support
deep links to section/element.

The plugin schema must expose all v1 fields, grouped as follows:

- Shells/session: `windowsDefaultProfileId`, `shellIntegrationEnabled`,
  `maxTabsPerPane`, `confirmCloseRunningProcess`, `closeOnExit`.
- Pane sync: `followPathWhenIdle`, `followPaneDebounceMs`.
- Scrollback/input: `scrollbackMaxLines`, `scrollbackMaxMemoryMiB`,
  `inputQueueMaxMiB`.
- History/privacy: `rememberShellByFolder`, `rememberCommandHistory`,
  `historyIgnoreLeadingSpace`, `historyDeduplicateConsecutive`,
  `historyMaxCommandsPerFolder`, `historyMaxTotalCommands`, `historyMaxFolders`,
  `historyMaxAgeDays`, `historyMaxCommandBytes`, `historyMaxStateFileMiB`, clear
  folder/all history, forget folder/all shell choice, Reset incompatible state.
- Appearance: `fontFamily`, `fontSizePoints`, `fontWeight`, `ligatures`,
  `cellWidthScale`, `lineHeightScale`, `cursorStyle`, `cursorBlink`,
  `cursorBlinkIntervalMs`, `backgroundOpacity`, `boldIsBright`.
- Interaction: `copyOnSelect`, `trimTrailingWhitespaceOnCopy`,
  `scrollToBottomOnInput`, `showScrollbar`, `warnOnUnsafePaste`,
  `pasteWarningLineThreshold`, `pasteMaxBytes`, `bellStyle`.
- Security/images: `kittyGraphicsEnabled`, `kittyImageStorageMiB`,
  `kittyGlobalCpuStorageMiB`, `kittyGlobalGpuStorageMiB`, `kittyApcMaxMiB`,
  `kittyMaxSingleImagePixels`, `osc52Policy`, `hyperlinkPolicy`,
  `allowApplicationTitle`.
- Read-only status: engine pin/version/ABI health, profile/catalog/integration
  capability, state health, limits, and Restart-required notices. Show no user
  paths, command history, terminal content, or secret.

There is deliberately no engine path, Ghostty DLL selector, auto-update toggle,
or arbitrary command line setting.

`configurationVersion=1` is required. One pure parser serves startup, hot
reload, and Prepare. It preserves unknown members, migrates older versions
deterministically, quarantines future versions read-only, repairs individual
wrong fields to documented defaults/safest enum, clamps valid numeric out-of-
range values, never treats numeric/string truthiness as Boolean, and writes
canonical default-omitted JSON only through the revision-CAS transaction.

Apply is asynchronous two-phase across stable session IDs: Prepare each session
without holding service/other-session locks, abort all on failure, then publish
one generation and execute allocation-free/infallible Commit. Live-limit
lowering completes deterministic eviction/pruning before commit; privacy false
completes purge before commit; integration disable completes revocation before
commit. Failure retains the old setting.

**Verify**:

```powershell
.\.build\x64\Debug\TerminalTests.exe --suite settings
.\.build\x64\Debug\RedSalamander.exe --run-self-tests --test-filter terminal_preferences_
```

Expected: every field/default/bound/type/enum, unknown preservation, old/future
version, startup/hot reload/Apply equality, Apply/Cancel/compensation, Open/
close/root-exit races, privacy purge, live cap change, deep link, keyboard/UIA,
and no-host-default tests pass.

### Step 9: Complete command/settings accessibility and evidence

Extend the base UIA provider with accessible Terminal tabs, profile chooser,
history chooser, context actions, diagnostics/settings links, unsafe-paste,
OSC 52 ask, and exit/restart surfaces. Dismissal/teardown removes discoverability
and scrubs hidden payloads. Keyboard focus/order/accelerators and Narrator names
must match visible localized UI.

Run the `PaneProfiles` and `ShellState` native slices plus content-free Commands
evidence. Record exact archive paths and hashes; no test selects a “latest” run.

**Verify**:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Full
.\Tools\Run-TerminalCommandsEvidence.ps1 -Scenario PaneProfiles -Platform x64 -Configuration Debug -RepositoryCommit (git rev-parse HEAD) -EvidenceRoot .\Specs\TestRuns -PassThruRunPath
.\Tools\Run-TerminalCommandsEvidence.ps1 -Scenario ShellState -Platform x64 -Configuration Debug -RepositoryCommit (git rev-parse HEAD) -EvidenceRoot .\Specs\TestRuns -PassThruRunPath
```

Expected: full suite exits 0; both new archives validate with nonzero, passed,
unskipped required cases and privacy/content allowlists.

## Test plan

- Commands/panes: both Open shortcuts, chooser, repeated independent tabs,
  directory-versus-file Edit, opposite-pane focus/window/source mapping,
  dynamic Folder/Preview/Terminal tabs, hidden pending commit/abort, MRU,
  callbacks, DPI/theme/RTL, keyboard and UIA.
- Profiles/locations: Windows local/UNC/plugin paths, every WSL spelling and
  exact distro, unsupported providers, discovery/dedupe, remembered/default/
  chooser/explicit precedence, missing profiles, timeout/cancel/stale results,
  and no cross-family fallback.
- Insertion: focused-item only, authenticated same-parent leaf, otherwise/full
  path, current directory, shell quoting/Unicode/UNC, safe selected terminal,
  every unsafe state, every hidden-open abort race, no Enter, and no pseudo-
  prompt fallback.
- Shell integration: generated adapter byte identity, PowerShell 5.1 and pwsh
  live proof, prompt/status/native-history preservation, capabilities,
  sequencing/revocation/nested shell, content-free evidence; truthful cmd/WSL
  terminal-only behavior.
- State/history/follow: schema/migration/corruption/atomic merge, every bound +1,
  privacy purge/stale writer, Debug/Release isolation, accepted/completed
  capture, chooser Insert/Copy, navigation coalescing/safe-idle/one-way behavior.
- Settings: every field/default/range/type/enum, older/future/unknown values,
  startup/hot reload/Preferences equivalence, async Prepare/Commit/Abort,
  privacy/resource/integration live transitions, deep links, UIA, and no host
  interpretation.
- Verification is Step 9's full suite plus exact `PaneProfiles` and `ShellState`
  content-free evidence. Zero discovered or skipped required cases is failure.

## Done criteria

- [x] Open Command Shell (`Ctrl+Alt+T`/`Alt+7`) and directory Edit open or reuse
      the opposite-pane Terminal rooted/published at the focused source.
- [ ] Windows/WSL profile selection follows the exact matrix and exposes a
      settings link for unavailable choices; no silent fallback occurs.
- [x] Ctrl+Enter inserts leaf only on authenticated same-parent equality,
      otherwise full path; Ctrl+Shift+Enter always inserts full; neither sends
      Enter; focused-item and live-terminal reuse rules pass.
- [ ] The advanced hidden `PendingPathInsertion` commit/abort path guarantees
      zero visible tab and zero bytes on every failed new-session insertion.
- [x] The pseudo prompt and `cmd.exe /C` implementation are removed after all
      stable IDs migrate.
- [x] First-implementation PowerShell cwd/idle/follow capabilities are
      root-PID authenticated and independent; cmd/WSL remain truthful
      terminal-only. Advanced accepted-input/child/password/TUI capabilities
      remain open.
- [ ] History/state is bounded, private, atomic, versioned, purgeable, and
      Debug/Release isolated.
- [x] Follow defaults off, uses the original source, runs only at a trusted idle
      root PowerShell prompt with no pending input, creates no history, and is
      one-way.
- [ ] Preferences exposes every listed parameter/action/status through opaque
      plugin configuration and contains no runtime path.
- [x] All terminal-specific behavior remains in `Terminal.dll`; host changes
      pass the source-contract allowlist.
- [ ] Full suite and PaneProfiles/ShellState evidence pass.
- [x] Specs and `Specs/Plans/WIP/Terminal_FirstImplementationExecutionIndex_2026-08-01.md` status are updated for the working slice.

## STOP conditions

Stop and report if:

- an interaction requires terminal policy, shell quoting, profile selection,
  history, follow, or Settings defaults in the host;
- Folder/Preview regression cannot be avoided with stable dynamic tabs;
- hidden pending insertion cannot guarantee zero visible tab/zero bytes on every
  failure or cannot make host Finalize allocation-free;
- Ctrl+Enter equality would use display text, unauthenticated cwd, or physical
  alias guessing;
- PowerShell proof requires modifying the user's profile/history or exposing a
  nonce in argv/environment/logs;
- a required Windows PowerShell 5.1 or supported pwsh live proof cannot run;
- WSL needs semantic integration to satisfy v1 requirements;
- state purge cannot fence stale/cross-process writers before reporting success;
- Settings Apply can publish a mixed generation or block the UI;
- removing the pseudo prompt would break an unmigrated stable command; identify
  and migrate it before removal;
- a required deterministic/accessibility/privacy test fails twice after a
  reasonable correction.

## Maintenance notes

- Stable command IDs outlive labels. Future shortcut migrations must preserve
  case-insensitive identity and keep contextual/full-path actions distinct.
- Add new shell families only through an explicit profile/capability/security
  design; launchability never implies history/follow/insertion support.
- Future configuration versions must preserve opaque newer values on downgrade
  and use ordered idempotent migrations.
- History and shell-choice clears are privacy operations, not UI filters. Review
  all recovery files and older retained writers when state format changes.
