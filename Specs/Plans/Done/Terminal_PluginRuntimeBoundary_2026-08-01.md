# Plan 002: Establish the Terminal plugin boundary, private loader, and package topology

> **Status — Completed 2026-08-03.** The lean `ITerminal` plugin boundary, private Ghostty loader, exact
> architecture/size/SHA-256 lock, canonical-path loading, held-file lifetime, centralized seven-file staging, and
> fail-closed degradation contracts are implemented. Durable authority is `../../Terminal/Terminal_EmbeddedPlugin.md`;
> later WIP/checkpoint language below is historical.

## First-implementation outcome (2026-08-01)

**FIRST SLICE COMPLETE.** `Common/PlugInterfaces/Terminal.h` provides the
unsuffixed `ITerminal` ABI; `Plugins/Terminal/Terminal.dll` privately resolves
the qualified Ghostty export table; and central runtime-dependency plumbing
stages `Plugins\TerminalRuntime\ghostty-vt.dll`, its identity JSON, and its
license. Terminal remains separate from `IViewer`. A later coordinated
source-tree ABI cleanup changed `ViewerOpenContext` and `ViewerTheme` to the
same leading-`sizeBytes` policy; no Terminal behavior uses those records.
Discovery entries are explicitly typed
Viewer or Terminal, so the existing manager supplies generic factory,
configuration, and module-lifecycle plumbing without becoming the terminal
ABI. Missing runtime state is isolated to a diagnostic Terminal child.

The remaining package-signature, loaded-file hash binding, clean-extraction,
and release-mutation gates stay open in Plan 005; therefore this plan remains
under WIP rather than being moved to Done.

The 2026-08-02 first-implementation closeout passed the full Debug x64 suite
(`20260802T190401Z-4616-084da5bd053c43418460a8bf7a4ee84a`, 1,180 passed,
0 failed, 53 justified skips), a 10-repeat real terminal lifecycle run
(`20260802T201146Z-15692-f697830e6f3c41a1926c6e7e1a15ea08`), the 10-case
build-reproducibility contract, and a full x64 Release solution build. Exact
evidence and the remaining release boundary are recorded in
`Specs/Terminal/Terminal_EmbeddedPlugin.md`.

> **Executor instructions**: Execute only after Plan 001 records Ghostty as a
> passing winner and creates the exact production engine lock. Follow every
> verification. Stop on any named condition; do not move terminal
> implementation into the EXE or stretch `IViewer`. Update the status row in
> `Specs/Plans/WIP/Terminal_FirstImplementationExecutionIndex_2026-08-01.md` when done.
>
> **Drift check (run first)**:
> `git diff --stat aac3ed260..HEAD -- Common/PlugInterfaces Common/PluginConfiguration.* RedSalamander/PluginModuleLifecycle.* RedSalamander/ViewerPluginManager.* RedSalamander/ApplicationContext.* RedSalamander/Preferences.Plugin.Configuration.* RuntimeDependencies.props Directory.Build.targets Tools/RuntimeDependencies.psm1 Tools/PortablePackageSmoke.psm1 Installer RedSalamander.sln Specs/Plans/WIP Specs/Terminal External LICENSE.txt`
> Expected drift from completed Plan 001 is allowed. STOP if the recorded
> winner is not Ghostty or its production lock/evidence is incomplete.

## Status

- **Execution state**: FIRST SLICE COMPLETE; release hardening open
- **ABI normalization**: COMPLETE on x64 Debug, 2026-08-03; ARM64 proof remains
  a release gate
- **Priority**: P1
- **Effort**: L (20-30 owner-days)
- **Risk**: HIGH
- **Depends on**: `Specs/Plans/WIP/Terminal_QualifyPrivateLibGhosttyWinner_2026-08-01.md`
- **Category**: migration
- **Planned at**: commit `aac3ed260`, 2026-08-01

## Why this matters

RedSalamander has a sound plugin factory and unload pattern, but its viewer
interface cannot represent terminal sessions and its runtime-dependency
pipeline currently assumes flat plugin output. This plan creates the additive
`ITerminal` boundary, a loadable-but-degradable `Terminal.dll`, an exact private
Ghostty loader, and one nested package manifest contract before session,
renderer, or UX complexity is added.

## Current state

- `Common/PlugInterfaces/Viewer.h` exposes only `Open`, `Close`, `SetTheme`, and
  `SetCallback`. Terminal does not extend or implement it. Its two extensible
  records share the repository's leading-`sizeBytes` policy.
- `Common/PlugInterfaces/Factory.h:141-152` already defines idempotent shutdown,
  unload-vote, and process-retention exports. Reuse those exports.
- `RedSalamander/PluginModuleLifecycle.cpp:29` invokes the shutdown export and
  supports deferred/process-retained unload. Extend the generic plumbing only
  if the new manager cannot reuse it unchanged.
- `RedSalamander/ViewerPluginManager.*` is viewer-specific. It is an exemplar,
  not the base class for a terminal.
- `RuntimeDependencies.props` records `Source` and `OutputName`; it has no
  validated nested destination.
- `Directory.Build.targets:105-137` copies dependencies to
  `$(OutDir)%(OutputName)` and removes stale leaves at the same flat path.
- `Tools/RuntimeDependencies.psm1:69-89` validates package outputs as
  `<build>\Plugins\<OutputName>`.
- `Specs/Installer/Installer_PortableZip.md:22-34` requires ZIP contents to use
  the same runtime-dependency manifest as build output. ZIP, MSI, and MSIX must
  not acquire separate terminal copy lists.
- The authoritative WIP master already freezes `builtin/terminal`, a new
  `ITerminal`/`ITerminalCallback`, plugin-owned child HWNDs, module quiet, and
  private `Plugins\TerminalRuntime\` packaging.

## Boundary contract to implement

```text
Host owns: plugin discovery, generic tab record, source/launch value marshaling,
           parent HWND, theme delivery, preferences routing, module lifetime.

Terminal.dll owns: runtime validation/loading, Ghostty API, ConPTY/session,
                   HWND/control, rendering/input/UIA, profiles/shell quoting,
                   history/follow/settings interpretation, all workers.
```

- Add `Common/PlugInterfaces/Terminal.h` with the unsuffixed size-based records
  and immutable interface IIDs. Do not
  add methods to `IViewer` or reuse `ViewerTheme` for Terminal.
- Factory plugin ID is exactly `builtin/terminal`; the same object supports
  `ITerminal` and `IInformations`. It must not QI to `IViewer`.
- All extensible ABI structs have one leading `sizeBytes`, fixed-width scalars, reserved-zero
  fields, copied UTF-16 views with explicit lengths, and no STL/Ghostty types.
- Public C++ identifiers are unsuffixed: use `TerminalOpenContext`,
  `TerminalPathInsertion`, `TerminalTheme`, `TerminalViewState`, `ITerminal`,
  `ITerminalCallback`, and equivalent unsuffixed PluginConfiguration/Host
  records. Do not declare a public struct, enum, interface, callback, or method
  whose identifier ends in `V1`/`V2`. Extensible in-process records use
  `sizeBytes` only; do not add `version`, `apiVersion`, `structSize`, public
  layout-size constants, or compile-time size/offset assertions. Compatible
  additions are optional tails. Incompatible changes use a new method or IID.
  Persisted schemas, evidence formats, and migration function names retain
  their concrete format versions.
- `SetCallback(nullptr, nullptr)` closes admission and synchronously drains
  in-flight callbacks. It is never invoked inside a callback.
- `Terminal.dll` must load and expose localized engine health even when the
  private runtime is missing, corrupt, wrong architecture, or ABI-incompatible.
  Other plugins and the application continue to work.

## Commands you will need

| Purpose | Command | Expected on success |
|---|---|---|
| Focused project build | `.\build.ps1 -ProjectName Terminal -Configuration Debug -Platform x64` | exit 0; `Plugins\Terminal.dll` staged |
| Host build | `.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64` | exit 0 |
| ABI/runtime Pester | `Invoke-Pester -Path Tools/Tests/TerminalPluginBoundarySourceContracts.Tests.ps1,Tools/Tests/BuildReproducibility.Tests.ps1,Tools/Tests/RuntimeDependencyManifest.Tests.ps1 -Output Detailed` | all pass |
| Plugin tests | `.\.build\x64\Debug\TerminalTests.exe --suite boundary-runtime` | exit 0; all cases pass |
| Package smoke | `.\build.ps1 -Configuration Release -Platform x64 -Zip` | exit 0; ZIP smoke passes |
| Full gate | `.\Tools\Run-AllTests.ps1 -Suite Full` | exit 0 |

If `RuntimeDependencyManifest.Tests.ps1` does not exist at execution time,
create it under that exact path and register it with the normal Tools Pester
suite; do not hide the cases in an unrelated test file.

## Scope

**In scope**:

- `Common/PlugInterfaces/Terminal.h` (new)
- `Common/PlugInterfaces/PluginConfiguration.h` (new)
- `Common/PlugInterfaces/Host.h` (new)
- `Common/PlugInterfaces/Factory.h` only if additive generic declarations are
  necessary; existing exports and IIDs remain unchanged
- `Common/PluginConfiguration.h`
- `Common/Common/PluginConfiguration.cpp`
- `RedSalamander/TerminalPluginManager.h` (new)
- `RedSalamander/TerminalPluginManager.cpp` (new)
- `RedSalamander/ApplicationContext.h`
- `RedSalamander/ApplicationContext.cpp`
- `RedSalamander/PluginModuleLifecycle.h`
- `RedSalamander/PluginModuleLifecycle.cpp`
- `RedSalamander/Preferences.Plugin.Configuration.h`
- `RedSalamander/Preferences.Plugin.Configuration.cpp`
- `Plugins/Terminal/Terminal.vcxproj` (new)
- `Plugins/Terminal/Terminal.vcxproj.filters` (new)
- `Plugins/Terminal/TerminalPlugin.h` (new)
- `Plugins/Terminal/TerminalPlugin.cpp` (new)
- `Plugins/Terminal/TerminalFactory.cpp` (new)
- `Plugins/Terminal/TerminalService.h` (new)
- `Plugins/Terminal/TerminalService.cpp` (new)
- `Tools/Generate-TerminalRuntimeManifest.ps1` (new); its
  `$(IntDir)\Generated\TerminalRuntimeManifest.h` output is untracked build
  output, never a source-tree file
- `Plugins/Terminal/TerminalRuntimeLoader.h` (new)
- `Plugins/Terminal/TerminalRuntimeLoader.cpp` (new)
- `Plugins/Terminal/GhosttyApi.h` (private; new)
- `Plugins/Terminal/GhosttyApi.cpp` (private; new)
- `Plugins/Terminal/TerminalResources.rc` and localized satellite resources
- `Tests/TerminalTests/TerminalTests.vcxproj` and boundary/runtime cases (new)
- `RedSalamander.sln`
- `RuntimeDependencies.props`
- `Directory.Build.targets`
- `Tools/RuntimeDependencies.psm1`
- `Tools/PortablePackageSmoke.psm1`
- `Tools/Tests/BuildReproducibility.Tests.ps1`
- `Tools/Tests/RuntimeDependencyManifest.Tests.ps1` (new)
- `Tools/Tests/TerminalPluginBoundarySourceContracts.Tests.ps1` (new)
- `Installer/zip/build-zip.ps1`
- `Installer/msi/build-msi.ps1`
- `Installer/msix/RedSalamanderInstaller.wapproj`
- `Specs/Installer/Installer_PortableZip.md`
- `Specs/Plans/WIP/Terminal_CoreConPtyRuntimeAndModel_2026-07-22.md`
- `Specs/Plans/WIP/Terminal_RendererControlAndAccessibility_2026-07-22.md`
- `Specs/Plans/WIP/Terminal_PaneTabsProfilesAndLifecycle_2026-07-22.md`
- `Specs/Plans/WIP/Terminal_ShellIntegrationHistoryAndPaneFollow_2026-07-22.md`
- `Specs/Plans/WIP/Terminal_SettingsSecurityPerfPackagingCloseout_2026-07-22.md`
- `Specs/Plans/WIP/Terminal_EmbeddedConPtyLibGhosttyVt_IntegrationPlan_2026-07-13.md`

**Out of scope**:

- ConPTY process creation, live Ghostty terminals, renderer, keyboard/IME, UIA,
  panes, commands, profiles, history, follow, and terminal settings UI;
- viewer methods, viewer associations, or Terminal use of viewer records; the
  coordinated `ViewerOpenContext`/`ViewerTheme` size-header normalization is
  the only viewer ABI change in this plan;
- a Ghostty import library reference from `RedSalamander.exe`, `Common`, the
  neutral host, or tests that claim to prove the public ABI;
- a user-configurable engine path or fallback DLL search;
- plugin-local `PostBuildEvent`, `xcopy`, or packager-only dependency lists.

## Git workflow

- Use a `codex/terminal-*` branch if starting a new branch. Do not push without
  operator instruction.
- Commit in four reviewable units: ABI/manager, plugin/loader, manifest/package,
  then tests/spec updates.
- Never commit generated build outputs or the Ghostty source checkout.

## Steps

### Step 1: Freeze and test the additive `ITerminal` ABI

First update the active Terminal WIP/master contracts to apply the public naming
rule consistently to every public type in `Terminal.h`,
`PluginConfiguration.h`, and `Host.h`. Examples include removing the suffix
from `TerminalOpenContext`, `TerminalPathInsertion`,
`TerminalViewState`, `TerminalTheme`, every Terminal enum/span/ID/location
record, every `PluginConfiguration*V1` record, `HostUtf16Span`, and
`HostPluginPreferencesRequest`. Do not rewrite immutable Done/evidence files;
their historical spelling remains historical.

Then add `Terminal.h` using the size-based contract in the updated master plan:

- `TerminalOpenContext` carries instance ID, physical host key, original
  source key, source generation/location, independent launch location, purpose,
  profile selection, parent HWND, and reserved fields.
- `TerminalPathInsertion` carries the prevalidated source item, requested
  insertion form, and correlation/generation needed for hidden-open commit.
- `ITerminal` provides configuration/theme, `Open`, idempotent asynchronous
  close/request-close, `GetViewState`, public actions, path insertion, callback
  registration, and diagnostic engine health. Use the precise method order and
  HRESULT rules from the master plan; do not redesign during implementation.
- `ITerminalCallback` method order is `TerminalStateChanged`,
  `TerminalPendingOpenResolved`, `TerminalOpenSiblingRequested`,
  `TerminalCloseApproved`, `TerminalViewRemovalRequested`, `TerminalQuiet`.
- The plugin creates exactly one direct child HWND under the supplied parent on
  successful `Open`; the plugin owns/destroys it. In this plan a diagnostic
  placeholder child is sufficient.

Add native ABI tests that create through `RedSalamanderCreate`, QI `ITerminal` and
`IInformations`, reject `IViewer`, attach/detach callback, and release/unload.
The tests reject undersized records and accept current and oversized records;
they do not freeze concrete record sizes or offsets.

**Verify**:

```powershell
.\build.ps1 -ProjectName TerminalTests -Configuration Debug -Platform x64
.\.build\x64\Debug\TerminalTests.exe --suite abi
$forbiddenPublicVersionedNames = rg -n "\b(struct|enum class|interface|using)\s+[A-Za-z_][A-Za-z0-9_]*V[0-9]+\b|\b[A-Za-z_][A-Za-z0-9_]*V[0-9]+\s*\(" Common\PlugInterfaces
if ($LASTEXITCODE -notin 0,1) { exit $LASTEXITCODE }
if ($forbiddenPublicVersionedNames) { $forbiddenPublicVersionedNames; throw 'Public ABI identifiers must be unsuffixed.' }
$stalePlanIdentifiers = rg -n -g 'Terminal_*.md' "\b(?:Terminal|PluginConfiguration|Host)[A-Za-z0-9_]*V[0-9]+\b|\bI[A-Za-z0-9_]*V[0-9]+\b" Specs\Plans\WIP
if ($LASTEXITCODE -notin 0,1) { exit $LASTEXITCODE }
if ($stalePlanIdentifiers) { $stalePlanIdentifiers; throw 'Active Terminal plans still use suffixed public API identifiers.' }
```

Expected: build and test exit 0; size-boundary cases pass; negative QI and
reserved-field cases return their frozen HRESULTs; callback drain completes;
both public-identifier scans emit no match. Serialized schema/evidence format
versions remain allowed and are not matched by these scans.

### Step 2: Add Terminal.dll and a generic terminal plugin manager

1. Add `Terminal.vcxproj` to every required solution mapping: x64 Debug,
   Release, ASan Debug; ARM64 Debug and Release. Add ARM64 ASan only if Plan 001
   proved support.
2. Export the standard factory/enumeration/configuration/lifecycle functions.
   The plugin ID is `builtin/terminal` and cannot collide case-insensitively.
3. Create one lazy module-owned `TerminalService`; inject it into every
   `ITerminal`. Do not expose the service to the EXE and do not use a process
   singleton outside the plugin module.
4. Add `TerminalPluginManager` following the loading/configuration/module
   ownership pattern of `ViewerPluginManager`, but returning only `ITerminal`.
   Factor a genuinely generic helper only where semantics match; update
   `Specs/Core/Core_SharedHelpers.md` if a new shared helper is introduced.
5. Reuse `PluginModuleLifecycle`: stop admission, call plugin shutdown, check
   `CanUnloadNow`, and retain until process exit only through the existing
   explicit vote. Follow the repository quiet-point order.
6. Add a localized diagnostic placeholder view so a missing engine is visible
   without making plugin discovery or application startup fail.

**Verify**:

```powershell
.\build.ps1 -ProjectName Terminal -Configuration Debug -Platform x64
.\.build\x64\Debug\TerminalTests.exe --suite factory-lifecycle
```

Expected: exact output `\.build\x64\Debug\Plugins\Terminal.dll` exists;
factory/lifecycle tests pass for normal unload, busy/deferred unload, retained
vote, configuration failure, and missing-runtime diagnostic mode.

### Step 3: Generate an immutable runtime manifest from the winner lock

Generate `$(IntDir)\Generated\TerminalRuntimeManifest.h` at build time from
`External/terminal-engine.lock.json` with
`Tools/Generate-TerminalRuntimeManifest.ps1`. Add that intermediate directory
to the Terminal project's include path. Never write or update a generated
header under the tracked source tree. It must contain only compile-time data:

- candidate ID and exact source pin/tree identity;
- exact package-relative path for primary and every private transitive leaf;
- raw final-file SHA-256 and byte length;
- PE machine, expected import/export sets, required Ghostty function names;
- header/layout/ABI fingerprint and expected build-info diagnostics;
- selected build flags/toolchain identity;
- manifest schema version.

Generation must fail if the lock is missing, schema-invalid, not the selected
winner, contains path tokens/wildcards/root/dot-dot/ADS, has case-insensitive
destination collisions, or does not describe the active platform/configuration.
Do not accept user configuration or an environment override.

**Verify**:

```powershell
Invoke-Pester -Path Tools\Tests\TerminalPluginBoundarySourceContracts.Tests.ps1 -Output Detailed
.\build.ps1 -ProjectName Terminal -Configuration Debug -Platform x64
```

Expected: tests/build pass; generated manifest matches the lock; searches find
no Ghostty header include or import-library reference outside `Plugins/Terminal`
and Gate 0 tooling.

### Step 4: Implement the hardened private runtime loader

`TerminalRuntimeLoader` must perform this exact sequence:

1. Derive the application package root from the loaded `Terminal.dll` path and
   manifest-relative path; never from current directory or PATH.
2. Canonicalize and require the result to remain under the exact
   `Plugins\TerminalRuntime` directory.
3. Open and hold every path with a WIL file handle denying write/delete for the
   module lifetime. Reject missing/non-regular files and any reparse point in
   the runtime root, ancestor chain, or leaf.
4. Verify file ID/path, length, raw SHA-256, PE machine, imports, exports, and
   complete declared private closure before loading any engine code.
5. Load only the exact primary absolute path with
   `LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR | LOAD_LIBRARY_SEARCH_SYSTEM32`. Never call
   `LoadLibrary*` with a basename and never add a global DLL directory.
6. Re-read the loaded module path and file identity and compare it with the held
   validated file. Reject mismatch.
7. Resolve every required export into one immutable `GhosttyApi` table. Validate
   build-info diagnostics and run the ABI exercise probe before installing
   callbacks or creating a terminal.
8. On any failure, release partial engine state safely, preserve a stable
   localized health category, and leave the rest of RedSalamander operational.

Do not log full user paths or runtime digests in ordinary user diagnostics.
Detailed digests may appear in explicit diagnostic evidence with privacy rules.

**Verify**:

```powershell
.\.build\x64\Debug\TerminalTests.exe --suite runtime-loader
```

Expected: available case passes; missing, corrupt, wrong-machine, hash, size,
export, import, build-info, reparse, outside-root, basename-collision, TOCTOU,
and undeclared-transitive mutations all return stable terminal-only unavailable
categories with zero Ghostty callbacks/handles created.

### Step 5: Extend the centralized runtime manifest for nested paths

Implement the master plan's `SourceRoot` plus package-relative `PackagePath`
model across all consumers:

1. Default legacy rows compatibly, but validate every path-bearing field as a
   literal: no MSBuild/environment/item/metadata expression, wildcard, empty,
   rooted, dot/dot-dot, ADS, or root escape. `OutputName` is exactly one filename
   and equals the `PackagePath` leaf ordinal-ignore-case.
2. Normalize existing ambiguous `Flavor=Any` AWS rows to explicit Debug/Release
   sources before enabling strict resolution. Merge duplicate zlib rows only
   where source/required semantics match and union project applicability.
3. Reject any applicable case-insensitive `PackagePath` collision after
   normalization; do not silently coalesce entries.
4. Create destination directories, copy required files, and remove stale files
   using one resolver shared by MSBuild and PowerShell package consumers.
5. Add `Terminal.dll` at `Plugins\Terminal.dll` and the exact lock-derived
   Ghostty closure only under `Plugins\TerminalRuntime\`. Release packages omit
   PDBs unless the existing symbol workflow includes them.
6. Make ZIP, MSI, MSIX, build-output validation, and clean-extraction smoke
   consume the same manifest. Missing required input fails the shared task.

**Verify**:

```powershell
Invoke-Pester -Path Tools\Tests\BuildReproducibility.Tests.ps1,Tools\Tests\RuntimeDependencyManifest.Tests.ps1 -Output Detailed
.\build.ps1 -Configuration Release -Platform x64 -Zip
```

Expected: all tests pass; legacy outputs are unchanged; invalid/colliding paths
fail; clean ZIP contains exactly `Plugins\Terminal.dll` and the locked private
closure; no engine DLL exists at app root or directly under `Plugins\`.

### Step 6: Prove lifecycle and degraded behavior before adding sessions

Exercise the production loader/factory through the public ABI, not a duplicate
test loader:

- app/plugin start with valid runtime;
- missing/corrupt/wrong-architecture primary and every conditional transitive;
- configuration before/after instance creation;
- 1,000 create placeholder/open/close/release cycles;
- callback posts racing detach, window destruction, shutdown, and module unload;
- first valid load, complete quiet/unload, second valid load in one process;
- plugin-busy and process-retained votes;
- application close with plugin unavailable.

Any cross-thread UI payload uses `PostMessagePayload`/`TakeMessagePayload`, calls
`InitPostedPayloadWindow` on creation and `DrainPostedPayloadsForWindow` in
`WM_NCDESTROY`, and ignores stale tokens.

**Verify**:

```powershell
.\.build\x64\Debug\TerminalTests.exe --suite boundary-runtime --repeat 1000
.\Tools\Run-AllTests.ps1 -Suite Full
```

Expected: tests and full suite exit 0; zero leaked windows, callbacks, threads,
payloads, module pins, file handles, or loaded runtime modules at quiet point.

## Test plan

- `Tests/TerminalTests`: ABI layout, QI, factory metadata/configuration,
  callback drain, placeholder HWND ownership, module votes, degraded health,
  hardened loader mutations, repeat load/unload.
- `Tools/Tests/TerminalPluginBoundarySourceContracts.Tests.ps1`: forbid
  Terminal use/implementation of `IViewer`, Ghostty headers/symbols/libraries in EXE/Common, terminal
  implementation outside plugin, direct EXE dependency on Terminal project, and
  plugin-local copy events.
- `Tools/Tests/BuildReproducibility.Tests.ps1` and
  `RuntimeDependencyManifest.Tests.ps1`: defaulting, literal validation,
  SourceRoot resolution, nested paths, case collisions, stale removal, missing
  required inputs, all five configurations, ZIP/MSI/MSIX parity.
- Package smoke must call the production loader and show only Terminal becomes
  unavailable for every runtime mutation.

## Done criteria

- [ ] Terminal neither implements nor extends `IViewer`; viewer associations
      are unchanged, and all in-tree viewers pass the coordinated size-header
      boundary tests.
- [ ] `ITerminal` IID/vtable is immutable, size-boundary-tested on x64/ARM64, and exposes no
      Ghostty/STL ownership.
- [ ] No ordinary public struct/enum/interface/callback/method identifier ends
      in `V1`/`V2`; in-process extensible records use `sizeBytes` only, while
      serialized schema/evidence/migration identifiers preserve concrete format identity.
- [ ] `Terminal.dll` is the only product owner of Ghostty and terminal-specific
      code; EXE/Common contain only allowed generic plumbing.
- [ ] Missing/invalid runtime produces localized terminal-only degraded health.
- [ ] Loader validates exact private path/hash/size/PE/import/export/ABI identity
      and never searches by basename/PATH.
- [ ] Central manifest stages nested runtime paths identically for build, ZIP,
      MSI, and MSIX and removes stale outputs declaratively.
- [ ] Normal, failure, repeat-load, callback-drain, and quiet unload tests pass.
- [ ] x64/ARM64 required build mappings exist; unsupported ARM64 ASan is
      recorded, not faked.
- [ ] `.\Tools\Run-AllTests.ps1 -Suite Full` exits 0.
- [ ] Specs and `Specs/Plans/WIP/Terminal_FirstImplementationExecutionIndex_2026-08-01.md` status are updated.

## STOP conditions

Stop and report if:

- Plan 001 did not record Ghostty as winner or lock identity is incomplete;
- implementation requires changing `IViewer` methods/IID, using it as the
  Terminal boundary, or exposing Ghostty in the EXE/Common;
- the plugin cannot load without the runtime present;
- loader security requires a global DLL directory, PATH, registry, or app-root
  engine copy;
- validated file identity cannot be bound to the loaded module;
- build/package consumers cannot use one literal nested-path manifest;
- an existing runtime row cannot be normalized without changing its shipped
  behavior and focused tests cannot explain the difference;
- a callback/thread/window/handle/module survives the quiet point;
- a step verification fails twice after a reasonable fix;
- a necessary change falls outside Scope. Update/review this plan before
  expanding it.

## Maintenance notes

- The dynamic boundary is intentional. Do not later replace it with an implicit
  import-library dependency merely to simplify calls.
- Runtime health categories and manifest schema are durable support contracts;
  version additions and preserve unknown-field behavior.
- If per-DLL Authenticode signing is added, sign the private runtime before
  manifest generation or adopt a signed detached manifest, then re-run every
  package mutation test.
- Keep `Terminal.dll`, engine runtime, adapter expectations, lock, and notices
  atomic in releases and rollbacks.
