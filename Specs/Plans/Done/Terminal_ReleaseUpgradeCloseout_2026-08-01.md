# Plan 005: Prove security, performance, packaging, upgrades, and closeout

> **Status — Completed 2026-08-03.** Security boundaries, bounded queues/resources, instrumentation, deterministic
> tests, centralized package staging, five-notice legal closure, offline x64/ARM64 dependency reproduction, and the
> check/upgrade/rollback workflow are implemented. Durable authority is `../../Terminal/Terminal_EmbeddedPlugin.md`
> and `../../Installer/Installer_PortableZip.md`; later WIP/checkpoint language below is historical.

## Current checkpoint (2026-08-01)

**IN PROGRESS.** The exact source/Zig/overlay tuple, reproducible x64 and ARM64
engine hashes, centralized nested staging, runtime identity JSON, Ghostty
license, dependency rebuild command, deterministic terminal lifecycle test,
and first-slice performance instrumentation are implemented. The x64 Debug
product builds with zero warnings and zero errors.

This checkpoint is not a release sign-off. Loader hash/signature binding,
native ARM64 product compilation, clean portable extraction, notices/SBOM,
security/teardown stress, accessibility/advanced renderer gates, and full
performance budgets/evidence still require completion. Those open items keep
all related plans under `Specs/Plans/WIP/`.

> **Executor instructions**: This plan is mandatory release work, not optional
> cleanup. Execute after Plan 004 and only from a reviewed clean implementation
> commit. Run every verification on the named native architecture. Never claim
> a signed/installable or performance result from an unsigned, transferred,
> skipped, dirty, or mismatched lane. Update `Specs/Plans/WIP/Terminal_FirstImplementationExecutionIndex_2026-08-01.md` last.
>
> **Drift check (run first)**:
> `git diff --stat aac3ed260..HEAD -- Plugins/Terminal RedSalamander Common RuntimeDependencies.props Directory.Build.targets Tools Installer .github/workflows Specs/Terminal Specs/Testing Specs/Installer Specs/Plugins Specs/UI Specs/Core Specs/Plans/WIP Specs/Plans/Done External LICENSE.txt`
> Expected Plans 001-004 drift is allowed. STOP if any prior plan is still WIP,
> required evidence is missing, or the engine lock/runtime manifest/product ABI
> no longer match.

## Status

- **Execution state**: IN PROGRESS
- **Priority**: P1
- **Effort**: L (15-25 owner-days)
- **Risk**: HIGH
- **Depends on**: `Specs/Plans/WIP/Terminal_UserInteractionsAndSettings_2026-08-01.md`
- **Category**: security
- **Planned at**: commit `aac3ed260`, 2026-08-01

## Why this matters

A terminal parses attacker-controlled bytes, launches processes, touches the
clipboard, follows filesystem navigation, persists command history, loads a
private native runtime, and owns GPU/accessibility resources. Passing feature
tests is not sufficient. This plan establishes hard security/performance
budgets, proves every package topology and degraded mutation, freezes an easy
dependency-update/rollback workflow, and moves all durable requirements into
authoritative specifications before WIP closeout.

## Current state and release invariants

- `External/terminal-engine.lock.json` and generated runtime manifest must bind
  the same selected Ghostty commit, Zig/toolchain/dependencies, build flags,
  x64/ARM64 outputs, hashes, sizes, PE/import/export/ABI identity, and notices.
- Build, ZIP, MSI, and MSIX consume one validated runtime dependency manifest.
  `Terminal.dll` ships at `Plugins\Terminal.dll`; the complete non-system engine
  closure ships only below `Plugins\TerminalRuntime\`.
- Current repository release signing signs the MSIX envelope, not individual
  DLLs. The runtime lock hashes raw shipped DLL bytes. Any future individual PE
  signing must occur before manifest generation or use a separately
  authenticated manifest.
- Required product test lanes are x64 Debug, x64 Release, x64 ASan Debug,
  ARM64 Debug, ARM64 Release. Required package records are unsigned ZIP x64 and
  ARM64, unsigned MSI x64, unsigned MSIX x64 and ARM64, plus signed MSIX x64 and
  ARM64: seven exact records.
- MSI remains x64-only unless repository release policy is separately changed.
  Unsigned MSIX is inspection-only and is never labeled installable.
- Official evidence binds one clean 40-lowercase-hex repository commit, positive
  build number, three-part release version, source-input digest, engine lock,
  runtime manifest, notices, toolchain lock, test harness, and evidence schema.

## Commands you will need

| Purpose | Command | Expected on success |
|---|---|---|
| Full gate | `.\Tools\Run-AllTests.ps1 -Suite Full` | exit 0; no failed/skipped required Terminal cases |
| Release build | `.\build.ps1 -Configuration Release -Platform x64` | exit 0 |
| ZIP package | `.\build.ps1 -Configuration Release -Platform x64 -Zip` | exit 0; clean-extraction smoke passes |
| MSI package | `.\build.ps1 -Configuration Release -Platform x64 -Msi` | exit 0; administrative extraction smoke passes |
| MSIX package | `.\build.ps1 -Configuration Release -Platform x64 -Msix` | exit 0; unsigned inspection passes |
| Dependency exact check | `.\Tools\Restore-TerminalEngineInputs.ps1 -Offline` | all locked inputs verified offline |
| Engine lane verify | `.\Tools\Verify-TerminalEngine.ps1 -Platform x64 -Configuration Release -RunSmoke` | exit 0 |
| Perf evidence | `.\Tools\Run-TerminalCommandsEvidence.ps1 -Scenario FinalPerf -Platform x64 -Configuration Release -RepositoryCommit (git rev-parse HEAD) -EvidenceRoot .\Specs\TestRuns -PassThruRunPath` | returns one exact validated run path |
| Upgrade comparison | `.\Tools\Prepare-TerminalEngineUpgrade.ps1 -Help` | usage exits successfully; use locked production arguments from Step 4 |

## Scope

**In scope**:

- all terminal implementation/tests created by Plans 002-004 when a validation
  finding requires a fix
- `Plugins/Terminal/TerminalMetrics.h/.cpp`
- `External/terminal-engine.lock.json`
- `LICENSE.txt`
- `RuntimeDependencies.props`
- `Directory.Build.targets`
- `Tools/RuntimeDependencies.psm1`
- `Tools/PortablePackageSmoke.psm1`
- `Tools/Prepare-TerminalEngineUpgrade.ps1`
- `Tools/Restore-TerminalEngineInputs.ps1`
- `Tools/Build-TerminalEngine.ps1`
- `Tools/Verify-TerminalEngine.ps1`
- `Tools/Test-TerminalPackage.ps1` (new)
- `Tools/Publish-TerminalSignedMsix.ps1` (new)
- `Tools/New-TerminalPackageProvenance.ps1` (new)
- `Tools/Verify-TerminalPackageEvidence.ps1` (new)
- `Tools/Run-TerminalCommandsEvidence.ps1`
- `Tools/TerminalCommandsEvidence.psm1`
- matching `Tools/Tests/Terminal*.Tests.ps1`, package/runtime tests, and
  `Tools/TestRunPlan.ps1`
- `Installer/zip/build-zip.ps1`
- `Installer/msi/build-msi.ps1`
- `Installer/msix/RedSalamanderInstaller.wapproj`
- `.github/workflows/release.yml`
- `Tools/ReleaseArtifactPolicy.ps1`
- `Specs/Terminal/TerminalPackageToolchain.lock.json` (new)
- `Specs/Terminal/TerminalPackageEvidence.schema.json` and corpus (new)
- `Specs/Terminal/TerminalState.schema.json` and corpus (new if not completed)
- `Specs/Terminal/Terminal_EmbeddedPane.md` (authoritative; new/finalized)
- `Specs/Plugins/Plugins_Terminal.md` (authoritative; new)
- `Specs/Plugins/Plugins_PluginAPI.md`
- `Specs/Plugins/Plugins_ViewerPlugins.md`
- `Specs/Installer/Installer_PortableZip.md`
- `Specs/Testing/Testing_PerformanceValidation.md`
- `Specs/Testing/TerminalPerfBudgets.json5`
- `Specs/Core/Core_SettingsStore.md`
- `Specs/UI/UI_CommandMenuKeyboard.md`
- `Specs/UI/UI_PreferencesDialog.md`
- `Specs/UI/UI_DxUiWinUIDesign.md`
- every Terminal WIP/Done plan and `Specs/Plans/WIP/README.md`
- immutable evidence under `Specs/TestRuns/<MachineHash>/Commands/` and the
  versioned `TerminalPackage` area

**Out of scope**:

- adding new v1 features or relaxing gates to fit a failing result;
- automatic in-application engine updates or a runtime download path;
- installing unsigned MSIX or testing signed MSIX on a user's existing package
  family/machine state;
- changing Ghostty pin without the complete upgrade workflow;
- archiving terminal contents, paths, commands, history, clipboard, nonce, URLs,
  distro names, usernames, or image bytes in metrics/evidence.

## Steps

### Step 1: Complete threat model and source-contract enforcement

Update the authoritative security contract and tests for these boundaries:

- private runtime search/hijack, reparse/TOCTOU, PE/import/export/ABI mismatch,
  tampering, undeclared closure, and degraded behavior;
- VT/OSC/DCS/APC/Kitty malformed, oversized, fragmented, decompression-bomb,
  integer overflow, generation churn, and effect-queue attacks;
- OSC 52 default deny, ask/allow revocation, byte cap, one-use UI token, focus/
  generation revalidation, and secure payload scrub;
- hyperlink disabled/ctrl-click, scheme allowlist, canonical parse, policy/view/
  document recheck immediately before `ShellExecuteW`;
- unsafe paste warning versus non-bypassable hard byte/input queue caps;
- application-title length/policy and no title/path/history leakage;
- shell adapter nonce/channel/sequence/version/root-process authentication,
  permanent incarnation revocation, no profile/native-history mutation;
- state file path/ACL/reparse/lock/atomic-replace/corruption/recovery/stale writer,
  privacy purge, Debug/Release isolation, and no content diagnostics;
- cross-thread payload ownership, HWND reuse, callback drain, worker/module pin,
  ConPTY/process tree, UIA passive tail, GPU/device teardown.

Extend source-contract scans to reject raw `LoadLibrary` fallback, raw COM owner,
manual handle/window cleanup, `catch (...)`, detached thread, raw posted pointer,
UI wait, user-content logging, engine/header linkage outside plugin, runtime path
setting, plugin-local post-build copy, and terminal logic outside the reviewed
allowlist.

**Verify**:

```powershell
Invoke-Pester -Path Tools\Tests\TerminalPluginBoundarySourceContracts.Tests.ps1,Tools\Tests\TerminalEngineNetworkIsolation.Tests.ps1,Tools\Tests\TerminalStateStore.Tests.ps1 -Output Detailed
.\.build\x64\Debug\TerminalTests.exe --suite security-adversarial
```

Expected: all tests pass; every named attack has a deterministic reject/revoke/
degraded category; content canaries are absent from logs/evidence.

### Step 2: Instrument and freeze performance budgets from Release evidence

Emit aggregate, content-free metrics using existing PerfJsonl conventions:

- command-to-tab-visible, open-to-first-prompt by non-sensitive profile family,
  UI-close return, close-to-quiet;
- service slots/admission/quarantine and process private-byte delta;
- read batch bytes/time, parse, queue bytes/descriptors/reserve/rejections,
  presentation posts/coalescing;
- snapshot lock/build, paint/present/input-to-visible, dirty/full/hidden frames,
  glyph/fallback/color runs;
- Kitty copy/upload/current CPU/GPU/evictions/rejections;
- scrollback lines/bytes/compression and process text/snapshot/prune/reject;
- follow confirmed-cwd latency and pending/applied/coalesced/paused counts;
- history load/prune/flush/family-lock/file/family/folder/command/expiry/eviction.

Never emit per key/cell/row/chunk. Required deterministic scenarios:

`terminal_perf_open_command_tab_visible`,
`terminal_perf_tab_icons_overflow`,
`terminal_perf_local_shell_first_prompt`,
`terminal_perf_pane_follow_latest_wins`,
`terminal_perf_output_storm`,
`terminal_perf_resize_reflow_stress`,
`terminal_perf_scrollback_100k_bounded`,
`terminal_perf_multitab_background_idle`,
`terminal_perf_service_slots_and_process_text_limits`,
`terminal_perf_kitty_image_burst`,
`terminal_perf_rapid_close_quiet_point`,
`terminal_perf_history_bounded_flush`.

Use fixed generated inputs and WARP; no network/user profile/installed WSL for
hard budgets. First collect an ungated Release baseline. Require at least 200
samples for p95 and 1,000 for p99. Then add per-machine/build hard budgets in
`TerminalPerfBudgets.json5` for tab-visible, first-prompt, follow, input-visible,
paint, resize/reflow, scroll, history, and close-to-quiet. Immediate invariants:
zero UI ConPTY calls, zero unchanged hidden-tab product paints, bounded queues/
memory/history, and no teardown timeout.

**Verify**:

```powershell
$run = .\Tools\Run-TerminalCommandsEvidence.ps1 -Scenario FinalPerf -Platform x64 -Configuration Release -RepositoryCommit (git rev-parse HEAD) -EvidenceRoot .\Specs\TestRuns -PassThruRunPath
.\Tools\Test-TestRunArchive.ps1 -RunPath $run
pwsh .\Tools\Show-PerfRuns.ps1 -Area Commands -Run $run -Metric terminal.render.input_to_visible_us -FailOnQuality -ShowBuildFlavor
pwsh .\Tools\Show-PerfRuns.ps1 -Area Commands -Run $run -Metric terminal.render.paint_us -FailOnQuality -ShowBuildFlavor
```

Expected: exact Release archive validates; adequate samples/quality; budgets
pass; content privacy scan passes; no required scenario is skipped.

### Step 3: Prove clean build/package mutation behavior

Use the production manifest/loader for every topology:

1. Build Release x64 and ARM64 from clean, exact inputs. Verify the engine
   offline in native lanes.
2. Produce unsigned ZIP x64/ARM64, MSI x64, and MSIX x64/ARM64.
3. ZIP extraction uses portable smoke; MSI uses `msiexec /a`; MSIX uses the
   pinned Windows SDK `makeappx validate/unpack`. All commands run inside the
   existing kill-on-close containment with 120-second operation timeout and
   five-second post-timeout quiet proof.
4. Inspect exact file set, path spelling, PE architecture, notices/state schema,
   `Terminal.dll`, and lock-derived private closure. Reject extras/collisions/
   stale files/app-root engine copies.
5. For each extracted topology, run available plus mutations for missing,
   corrupt, wrong-architecture `Terminal.dll`; missing/corrupt/wrong-architecture
   Ghostty primary; and every declared private transitive leaf. The production
   loader must classify the exact failure with zero callbacks/handles while the
   application composition root starts and only Terminal is unavailable.
6. `TerminalTests.exe --package-root-smoke` and
   `RedSalamander.exe --terminal-package-smoke` call the production loader,
   create no shell/PTY/user state, and emit no path/digest/content.

**Verify**:

```powershell
.\build.ps1 -Configuration Release -Platform x64 -Zip
.\build.ps1 -Configuration Release -Platform x64 -Msi
.\build.ps1 -Configuration Release -Platform x64 -Msix
.\Tools\Test-TerminalPackage.ps1 -Help
```

Expected: package builds exit 0; the exact x64 ExtractAndSmoke records pass.
Repeat ZIP/MSIX and ExtractAndSmoke on native ARM64. No runtime mutation affects
non-Terminal startup/features.

### Step 4: Freeze the easy dependency observation, upgrade, and rollback lifecycle

This is the durable answer to “how do we check and upgrade dependencies?”

#### Normal checks

- On every engine-related change and release: validate lock schema, source/
  toolchain/cache/output hashes, runtime manifest, PE/import/export closure,
  notices, and all five exact build verification identities.
- After any restore/build: run offline exact-cache verification. Network access
  during offline build is a failure, not a warning.
- Monthly, before a release, and on relevant security advisory: observe the
  official Ghostty repository/channel, Zig security/release source, and actual
  lib-vt dependency advisories. Observation is read-only and records checked
  timestamp/ref/status. No automatic tracked mutation.
- Because libghostty has no independent stable version tag, discovery selects a
  reviewed exact official commit/ref. Missing owner, redirects, ambiguous ref,
  prerelease-only state, or incomplete advisory information blocks the
  observation decision.

#### Preparing an upgrade

1. Run `Tools/Prepare-TerminalEngineUpgrade.ps1` with an explicit candidate
   commit into a disposable `.build` comparison root. Never overwrite the
   tracked production lock, adapter, runtime, or notices during preparation.
2. Pin and restore every changed source/toolchain/dependency by URL/hash/size.
3. Build all five lanes and compute a structured diff against production:
   public C headers/layouts/enums/calling convention, required/extra exports,
   build info, output hash/size/machine/import/private closure, build commands/
   flags/toolchain, license/notices/SBOM, Gate 0 fixtures, product native tests,
   sanitizer results, and performance deltas.
4. If the ABI/export set changed, update only the plugin-private `GhosttyApi`
   table/adapter plus generated manifest, then rerun the ABI exercise probe.
   Host ABI must not change merely because Ghostty changed.
5. Run the full product, security, package mutation, and performance gates.
6. Review one atomic change containing source pin, toolchain/input hashes,
   adapter, lock, generated manifest, runtime outputs, notices, tests, and
   evidence. Reject any mixed production/candidate component.

#### Applying and rolling back

- Apply only after review by replacing the complete atomic set and rebuilding
  packages. Engine switch takes effect after application restart/complete quiet;
  never swap the DLL under live sessions.
- Rollback is the inverse atomic set: old adapter/lock/manifest/runtime/notices
  together. Rebuild/repackage and run smoke; do not copy an old DLL beside a new
  manifest.
- Preserve the prior exact package/evidence until the new release is verified.

**Verify**:

```powershell
.\Tools\Restore-TerminalEngineInputs.ps1
.\Tools\Restore-TerminalEngineInputs.ps1 -Offline
.\Tools\Verify-TerminalEngine.ps1 -Platform x64 -Configuration Release -RunSmoke
Invoke-Pester -Path Tools\Tests\TerminalDependencyLifecycle.Tests.ps1,Tools\Tests\TerminalEngineCandidatePatch.Tests.ps1,Tools\Tests\TerminalEngineNotices.Tests.ps1 -Output Detailed
```

Expected: production exact checks pass; disposable upgrade comparison produces
complete diff without tracked changes; mixed-set fixtures fail; atomic upgrade
and rollback fixture sets both pass their own smoke.

### Step 5: Harden MSIX signing and collect seven package records

Do not sign in place. Add a canonical helper that:

1. selects one exact unsigned MSIX and verifies its recorded hash/identity;
2. validates the supplied certificate identity/EKU/validity/Publisher without
   writing secrets to output/arguments;
3. imports only the exact certificate into an empty temporary CurrentUser slot,
   copies unsigned input to a unique staging file, invokes digest-pinned
   SignTool with `/fd SHA256 /tr <approved timestamp> /td SHA256`, verifies with
   pinned SignTool/MakeAppx, and removes temporary certificate/staging in
   `finally`;
4. publishes one canonical signed leaf atomically while leaving unsigned bytes
   unchanged;
5. makes release workflow upload unsigned and signed artifacts separately and
   attach only the verified signed MSIX.

On clean disposable matching-native x64 and ARM64 machines, SignedInstallSmoke
must reject a pre-existing package family, validate trusted signature/publisher,
install, derive AUMID, activate terminal package smoke, await exit 0, and remove
exactly that family in `finally`. Never install unsigned MSIX.

Compare signed versus unsigned extracted payload. Only these four envelope paths
may differ: `[Content_Types].xml`, `AppxBlockMap.xml`,
`AppxMetadata/CodeIntegrity.cat`, `AppxSignature.p7x`. Record their separate
identities; reject any ordinary payload difference or fifth exclusion.

Collect exactly seven ordered records bound to one approved commit/build/version:

1. ZIP x64 ExtractAndSmoke
2. ZIP ARM64 ExtractAndSmoke
3. MSI x64 ExtractAndSmoke
4. unsigned MSIX x64 ExtractAndSmoke
5. unsigned MSIX ARM64 ExtractAndSmoke
6. signed MSIX x64 SignedInstallSmoke
7. signed MSIX ARM64 SignedInstallSmoke

**Verify**:

```powershell
Invoke-Pester -Path Tools\Tests\TerminalMsixSigning.Tests.ps1,Tools\Tests\TerminalPackageEvidence.Tests.ps1,Tools\Tests\ReleaseWorkflowPolicy.Tests.ps1 -Output Detailed
.\Tools\Verify-TerminalPackageEvidence.ps1 -Help
```

Expected: tests pass; exactly seven schema-valid records and one evidence-set
pair are produced; signed records bind/reinspect their unsigned records; no
secret or content-bearing field exists.

### Step 6: Bind official evidence to exact clean source/tool identity

Add `TerminalPackageToolchain.lock.json` and a provenance helper. Before any
official build/package:

- require exact clean HEAD, approved positive build number, approved three-part
  release version, native platform, Release configuration, and official-release
  version context;
- hash exact committed Git tree/submodules plus compiler/MSBuild, PowerShell/
  compression, SDK/MakeAppx, WiX, SignTool, AppX deployment, and activation
  helper identities;
- recheck the same manifest immediately before the Release build and each
  package build; reject intervening tracked/submodule/version drift;
- provide offline read-only commit identity verification using locally present
  Git objects with lazy fetch/replacements disabled and zero filesystem output;
- permit only exact evidence/package transfer files in review/signed-smoke
  modes. Never select a latest run/artifact.

The package evidence-set verifier accepts exactly the seven canonical relative
manifest paths, reopens each unchanged two-file run directory, recomputes every
binding/digest, and emits one JCS/SHA-256 pair. Invalid/incomplete input emits no
set; a complete failure emits a failed set; only a complete recomputed pass is
success.

**Verify**:

```powershell
Invoke-Pester -Path Tools\Tests\TerminalPackageProvenance.Tests.ps1,Tools\Tests\TerminalPackageEvidence.Tests.ps1 -Output Detailed
.\Tools\Run-AllTests.ps1 -Suite Full
```

Expected: clean/mismatch/missing object/second checkout root/no-output/read-only
and evidence substitution/order/alias/digest cases pass; full suite exits 0.

### Step 7: Merge durable contracts and close all plans

Before moving any WIP plan, merge its normative behavior into authoritative
specs:

- `Specs/Terminal/Terminal_EmbeddedPane.md`: plugin-only/not-IViewer boundary,
  `ITerminal`, lifecycle, ConPTY/model/render/input/UIA, commands, opposite-pane
  tabs, directory Edit, contextual/full insertion/no-Enter, pseudo prompt
  retirement, profile/Windows/WSL matrix, shell capabilities, history/follow,
  settings/theme, resource/security/perf, packages, dependency upgrade/rollback,
  and the unsuffixed public API naming rule. In-process extensible records use
  `sizeBytes` only; concrete versions remain only in serialized schemas,
  evidence formats, and migrations.
- `Specs/Plugins/Plugins_Terminal.md`: factory/IID/configuration/child HWND,
  callbacks, module quiet, runtime loader/closure, passive UIA retention, host
  allowlist.
- `Specs/Plugins/Plugins_PluginAPI.md`: additive async configuration and generic
  `Busy|Unloaded|RetainedUntilProcessExit` lifecycle result.
- `Specs/Plugins/Plugins_ViewerPlugins.md`: Terminal is not `IViewer`, an
  association, or Preview allowlist entry.
- UI/Core/Installer/Testing specs named in Scope: durable command, tabs,
  Preferences, Settings Store, packaging, evidence, and perf contracts.

Authenticate every Gate 0/product child completion record and evidence digest.
Move completed child plans to `Specs/Plans/Done/`; move the master last. Resolve
every checkbox/WIP marker/unresolved handoff token before moving. Do not leave a
normative requirement only under `Specs/Plans/Done/`.

**Verify**:

```powershell
rg -n "^[*-] \[ \]|State:\*\* WIP|TERMINAL-HANDOFF-UNRESOLVED" Specs\Plans\WIP\Terminal_*.md
rg -n "ITerminal|Ctrl\+Enter|Ctrl\+Shift\+Enter|followPathWhenIdle|TerminalRuntime|dependency upgrade|rollback" Specs\Terminal\Terminal_EmbeddedPane.md Specs\Plugins\Plugins_Terminal.md Specs\UI\UI_CommandMenuKeyboard.md
.\Tools\Run-AllTests.ps1 -Suite Full
git status --short
```

Expected: no unfinished Terminal plan markers; authoritative specs contain all
durable contracts; all explicit Terminal plans exist only under Done; full
suite exits 0; status shows only intended implementation/spec/evidence changes.

## Test plan

- Security/adversarial: every protocol/resource limit at boundary/+1, malformed/
  fragmented/compressed input, runtime mutations, state attacks, shell adapter
  revocation, clipboard/hyperlink/title policy, content-canary privacy.
- Performance: all 12 deterministic scenarios, Release sample quality, fixed
  seeds, WARP, content-free metrics, hard invariants and measured budgets.
- Dependency lifecycle: current exact/offline, later candidate isolated diff,
  ABI/export/import/toolchain/license/perf changes, mixed-set failure, atomic
  apply and rollback.
- Packages: five unsigned extracted records plus two signed native install
  records; exact closure and terminal-only degraded mutations for each topology.
- Provenance/evidence: clean source/version/tool identities, no latest/aliases,
  second checkout, missing Git object/no lazy fetch, seven record order, JCS/
  SHA companion integrity, read-only re-verification.
- Closeout: WIP/Done placement, authoritative-spec coverage, protected Gate 0
  history, full suite, and clean/intended status.

## Done criteria

- [ ] Security/source-contract/adversarial suites pass with no content leak.
- [ ] Release performance evidence has adequate samples and all budgets/invariants
      pass on the approved baseline machine.
- [ ] Exact dependency restore/build/verify works offline for all five lanes.
- [ ] Monthly/release/advisory observation and isolated atomic upgrade/rollback
      workflow is documented and tested.
- [ ] ZIP x64/ARM64, MSI x64, unsigned MSIX x64/ARM64, signed MSIX x64/ARM64
      produce exactly seven verified records bound to one source/tool identity.
- [ ] Every package mutation leaves application composition healthy and only
      Terminal unavailable with the expected category.
- [ ] Signed packages install/activate/uninstall on clean native disposable
      machines; unsigned MSIX is never installed/labeled installable.
- [ ] Authoritative Terminal/plugin/UI/core/installer/testing specs contain all
      durable contracts, including user interactions and dependency lifecycle.
- [ ] Active public headers/specs contain no `*V1`/`*V2` struct, enum,
      interface, callback, or method names; extensible in-process records use
      `sizeBytes` only, and serialized schema/evidence/migration identifiers
      remain explicit.
- [ ] Every Terminal child and master plan is moved from WIP to Done in order.
- [ ] Protected Gate 0 history remains unchanged.
- [ ] `.\Tools\Run-AllTests.ps1 -Suite Full` exits 0.
- [ ] `Specs/Plans/WIP/Terminal_FirstImplementationExecutionIndex_2026-08-01.md` marks Plans 001-005 DONE.

## STOP conditions

Stop and report if:

- implementation/evidence commit is dirty, mismatched, unavailable locally, or
  lacks an exact source/toolchain identity;
- any required Terminal case is skipped or any native lane is substituted;
- a performance budget can pass only by reducing scenario work, samples, or
  instrumentation truth;
- any metric/evidence contains user content or a secret;
- build/ZIP/MSI/MSIX do not consume one runtime manifest;
- a runtime/package mutation makes the application or another plugin fail;
- private closure/license/provenance differs from the production lock;
- upgrade preparation overwrites tracked production state or needs mixed
  versions/a long-lived fork;
- signed/unsigned payload differs outside the exact four MSIX envelope paths;
- signed test machine has an existing package family or is not matching-native;
- required certificate/tool/signature evidence is unavailable; do not label the
  unsigned artifact installable;
- a normative requirement would remain only in a WIP/Done plan;
- full suite or an exact evidence verifier fails twice after a reasonable fix.

## Maintenance notes

- Dependency observation may be automated as a read-only report, but selection,
  lock mutation, evidence, and package publication remain reviewed operations.
- Never infer libghostty compatibility from Ghostty application version alone.
  The exact C header/export/layout/fixture diff is the upgrade authority.
- Keep the previous atomic engine/package set until the replacement is verified
  in production release topology; this makes rollback mechanical.
- Rebaseline performance only for an explained hardware/build/toolchain/product
  change and retain the old comparison. Do not silently widen budgets.
