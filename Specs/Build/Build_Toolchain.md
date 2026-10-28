# Official Build Toolchain Identity
## External RedMonitor build boundary

RedMonitor is an independent project. RedSalamander vendors only Client.h and its MIT notice under
ThirdParty/RedMonitor; Dependencies/RedMonitor.lock.json records the exact upstream commit, API and
file hashes. Ordinary builds run local-only verification and add the public include directory. The
explicit Tools/Update-RedMonitorClient.ps1 command fetches only those two files. The compatibility
Tools/Restore-RedMonitor.ps1 entrypoint now verifies that local client and rejects viewer requests.

Every first-party native project with ClCompile inputs verifies before PrepareForBuild, including a
direct plugin build that does not reference Common. Resource-only and external dependency projects
are excluded. Real MSBuild controls prove valid direct admission, tampered-header rejection before
admission, and both exclusions; offline verifier controls cover the immutable manifest and file bytes.

The full behavioral receipt closure includes the lock, header, license and local verification tools.
There are no viewer/test/satellite compatibility projects, helper build/staging operations, or helper
runtime artifacts in consumer receipts/packages. RedMonitor owns its own production Release ZIP and
installation. Debug may asynchronously launch an installed production helper from its two standard
locations; missing helpers are harmless. No independent process is terminated. See the consumer
contract in Specs/Core/Core_RedSalamanderMonitor.md. Standalone SDK qualification does not qualify
RedSalamander Preferences, runtime or packaging.


The MSIX project declares Debug, Release and ASan Debug for both x64 and ARM64. Solution mappings preserve the selected configuration; an ASan build must never package Release or ordinary Debug executables. Packaging remains an explicit operation rather than an automatic side effect of a sanitizer solution build, consistent across both architectures.

The reusable hosted build limits initial MSBuild fan-out to two workers following an ARM64 Release compiler heap-exhaustion failure. Native compilation remains a required gate; a worker bound does not establish that the previous failure is resolved until the complete profile builds successfully.

On an x64 Windows host, first-party ARM64 projects default to the x64 compiler and
linker host. MSVC otherwise defaults that cross-build to the x86 host, whose address
space can be exhausted during Release whole-program optimization. The early
`Directory.Build.props` default preserves an explicit `PreferredToolArchitecture`
selection and native ARM64 host selection. Consumer profile tests evaluate the real
MSBuild imports for all six profiles and verify that explicit overrides survive.
DxUi restore evaluates that first-party host policy before computing the library fingerprint.
The product's DxUi project reference explicitly forwards the selected host: a property evaluated
inside a consumer project is not inherited automatically by the independent library project.
Identity probing restores the caller's environment on both success and failure. The existing
consumer/archive host-mismatch rejection remains mandatory.
Synthetic restore tests exercise a checkout path whose Git hook paths exceed 260 characters,
matching the governed Full sandbox depth. Their local Git clones enable long paths without
changing global or developer-repository settings; focused runs must cover the same boundary.

Contained process environment overrides distinguish removal from an empty value.
`New-RSProcessStartInfo` and its contained launch callers remove a child key for an
explicit `$null` override, retain a present empty key for `''`, and install exact
nonempty values. Parameter binding must preserve that distinction before converting
non-null values to strings. None of these child overrides changes the caller's
environment. Constructor tests and an actual contained child cover all three states.

## Pinned DxUi consumption

`Dependencies/DxUi.lock.json` records the API revision the product has adopted (4) and selects the single `DxUi.lib`
target. The pinned source states its own revision in `capabilities.json`, and the identity probe records it. Restore
requires the two to match, before the library's own dependency restore: a pin that crosses a revision is adopted
deliberately, by adapting the product to the revision's change notes (`CHANGELOG.md` and any pending `Changes/` notes in the library) and then updating the
lock. `Tools/Update-DxUi.ps1` moves only the commit and leaves `apiRevision` alone.
All DxUi-using production projects consume that pin: RedSalamander, RedConfigure, Monitor, Terminal
and the viewer plugins. ProductUiTests, RedConfigureTests and PerformanceTests2 use the same imports.
The old library and DxUiTests projects are retired. Shared implementation and control tests live
in the canonical DxUi repository; product adapters and their regressions remain in this repository.
Removing legacy files does not complete the runtime/resource gates recorded in I19.
Consumers import `Build/RedSalamander.DxUi.props` and `.targets`; they never enumerate library sources.
The root build restores the exact source with `Tools/Modules/Build/DxUiDependency.psm1`, or developers
can use `Tools/Restore-DxUi.ps1` before an IDE build. Restore owns the repository/platform dependency
lease, contains its native child processes and leaves sibling repositories unchanged.
The clone enables Git long paths twice: a process-level `-c core.longpaths=true` covers `git init`, which copies its
template hooks before the new repository has a configuration, and `clone -c` stores the setting in that repository so every
later git command on the checkout (the clean-source check) works from a deep consumer root. User and global Git
configuration never change. The checkout is sparse (`/*`, `!/Measurements/`, `!/docs/gallery/`, `!/Specs/`): retained
measurements, specifications and gallery sheets are most of the library's tracked files and none of its consumer
interface. Omitted paths are not deletions, so the clean-source check still holds. After checkout the restore requires a
clean status (a checkout can name files it could not write and still exit 0) and the paths restore and build read
(`capabilities.json`, `vcpkg-install.ps1`, `Tools/ConsumerBuild.psm1`, `Build/DxUi.Consumer.props`/`.targets`,
`src/DxUi.vcxproj`, `include`). A restore that fails removes only the partial checkout it created, beneath the guarded
dependency root, so the next run clones again. A restore test from a 160-character root, where a kept library file's
restored path passes MAX_PATH, covers all of it. Long paths do not lift the 258-character Win32 limit on a current
directory, which `git -C` needs, so the checkout directory (the consumer root plus 72 characters) must stay within it: a
root of at most 186 characters.

Separate platform properties select an output fingerprint containing source/API, compiler host and
compiler/linker/MSBuild/SDK identities, CRT family and STL annotation policy. RedSalamander explicitly
sets `DxUiDisableStlAnnotations=true` to match its ordinary vcpkg C++ dependencies; ASAN heap/stack
instrumentation remains required. Debug, Release and ASan Debug on x64/ARM64 are the target matrix.
The output directory under `.build/dependencies/DxUi` is the fingerprint's first 16 hexadecimal
characters (64 bits), which keeps the nested vcpkg tool, build-tree and archive paths short. It is a
lookup key only: `DxUi.identity.<Platform>.json` and every module sidecar retain the full SHA-256
fingerprint, and provenance rejects a linker command that names any other archive path.

The tracked lock participates in existing source and project-graph identities. A modified managed
source fails snapshot acquisition; missing or dirty source prevents build-receipt reuse. Successful
builds verify the actual linker command names the selected archive, reject a legacy archive in that
module, and produce `<module>.DxUi.json` with source/header/archive/toolchain/module identities.
These sidecars belong to the module's existing attested runtime closure. This extends existing build
evidence without introducing a new qualification schema.

Builds request one advisory about newer main with green DxUi CI. Lookup failure does not fail a valid
fixed-pin build. The maintainer updates the lock and adapters on a product branch, runs product tests,
and submits the ordinary PR. Library defects are fixed upstream before repeating that consumer loop.
The build adapter shows a newer main revision without a completed successful validation in red and a
validated update in yellow; other advisory states use dark yellow. A validated update prints the normal
`Tools/Update-DxUi.ps1` command and its `-UpdateOnly` alternative on separate yellow lines. Coloring communicates
advisory severity only and never changes the selected pin or build result.
`Tools/Update-DxUi.ps1` implements the normal branch update: it accepts only a current `main` commit with successful
completed DxUi CI, atomically changes the lock, and runs the Full suite. `-UpdateOnly` skips that local suite only
when equivalent product validation completed in another environment; neither mode auto-commits, and a local validation
failure retains the changed lock for diagnosis.
Rollback reverts the complete adoption change and uses the retained previous product package.
Before merging or releasing a consumer upgrade, publish the tested library commit to
the public DxUi repository and verify restore without a local Git URL rewrite. Local
qualification of an unpublished commit does not establish that a clean remote consumer
can fetch it. Library publication and product merge/release remain explicit lifecycle actions.
The [I19 checklist](../Plans/Done/DxUi_SharedLibraryAdoptionAndReleasePlan_2026-09-09.md) records actual
qualification; a declared configuration or library test pass does not qualify a product or native ARM64.

This specification defines the compiler identity used by official RedSalamander builds
and the evidence every build artifact must retain. It is authoritative for CI and
release builds; a developer may use another compatible installed toolchain locally,
but local success does not replace official identity evidence.

## Pinned bootstrap packages

Official GitHub Actions jobs install reviewed Chocolatey package versions rather than
accepting the mutable latest package:

- `visualstudio2026buildtools` version `118.8.1`;
- `visualstudio2026-workload-vctools` version `1.0.0` for normal CI builds;
- `visualstudio2026-workload-universalbuildtools` version `1.0.0` for release builds
  that produce the complete x64 and ARM64 package matrix.

A version change is a reviewed dependency update. The build must fail before
compilation if installation, discovery, or identity collection fails; silently falling
back to another Visual Studio installation is forbidden in official jobs.

## Recorded identity

`BuildEvidence.psm1` resolves the effective toolchain through MSBuild before compilation
and constructs the one canonical
[`ToolchainIdentity.schema.json`](ToolchainIdentity.schema.json) record used locally and
in CI. It binds the selected platform and VC/Windows SDK versions plus file name, size,
version where available, and SHA-256 for the actual MSBuild, compiler, linker, and
resource compiler. It also binds representative SDK bytes from `Windows.h`,
`kernel32.lib`, and `ucrt.lib`. The portable record intentionally contains no absolute
checkout or installation paths.

The reusable workflow resolves this canonical digest before dependency-cache lookup and
includes it in the vcpkg and Terminal runtime cache keys. `build.ps1` resolves the same
record, publishes it as
`.build/<platform>/<configuration>/toolchain-identity.json` after suspending the mutable
receipt pointer, and includes that file in the successful receipt output set. The digest
is therefore identical authority for cache selection, local receipt identity, portable
attestation, reusable-workflow output, and the uploaded record. A missing tool/SDK input,
schema-invalid record, changed byte seam, or digest mismatch fails before product output
is accepted. Vcpkg source/tool pins and Terminal runtime inputs remain separately bound
receipt identities rather than being duplicated in this record.

RedSalamander does not claim bit-for-bit reproducibility solely from this record; it is
an auditable provenance boundary for diagnosing and approving drift.

## Update procedure

1. Pin the proposed package versions in every official workflow that installs them.
2. Review the vendor/package release and the resulting compiler, MSBuild, SDK, and
   vcpkg identity records.
3. Run the complete Debug and Release build/test matrix, including x64 and ARM64 release
   packaging where applicable.
4. Update this specification and workflow policy tests in the same change.
5. Retain the successful identity artifacts with the CI run used to approve the update.

Floating `latest` compiler bootstrap packages and hand-written identity summaries are
not accepted evidence.

## Build diagnostic presentation

`Tools/Modules/Build/MSBuildInvocation.psm1` uses one diagnostic-header classifier
for captured-log counts and replay colors. Follow the
[MSBuild diagnostic format](https://learn.microsoft.com/en-us/visualstudio/msbuild/msbuild-diagnostic-format-for-tasks):
origin may be blank, subcategory and code are optional, and category is the
case-insensitive, locale-neutral `warning` or `error` regardless of the process
culture or localized message text. A code is a non-space token, not necessarily
a letter-plus-number identifier. Accept normal Windows/UNC source paths and the
console's numeric project-node prefix.

Classify only the anchored header, never another diagnostic-looking phrase inside
its message. Ordinary compilation filenames, incomplete headers and numeric
`Warning(s)`/`Error(s)` summaries do not count. Counts represent captured diagnostic
lines, not deduplicated event identities. Warnings are yellow and errors red;
explicit stderr presentation retains its red fallback. An unavailable log is
reported as unavailable, not as zero diagnostics. This presentation does not change
MSBuild's exit status or turn a warning-bearing run into zero-warning qualification.

## Terminal-engine checkout-invariant text inputs

Terminal-engine governance and reproducible-runtime tooling use byte-exact SHA-256
identities for tracked text. `.gitattributes` therefore pins the complete frozen
governance surface, Terminal-engine tools/tests, the private-runtime patches, and
notice inputs to LF. Adding a raw-byte identity input requires adding it to that exact
LF surface in the same change. A Windows `core.autocrlf=true` checkout and an LF
checkout must produce identical identities and pass the same transition, fixture,
ABI, archive, and closeout tests.

The Ghostty private-runtime builder additionally converts its patch and notice inputs
to UTF-8 without BOM with LF line endings before hashing or copying them. This
normalization is part of the schema-5 build identity; worktree bytes or the current
locale are never accepted as an implicit dependency identity.

The reusable-build Terminal cache key MUST include the thin runtime wrapper, the
Ghostty lifecycle module and its complete direct/indirect helper closure, the
canonical Ghostty product lock and schema, the current private-runtime patch, the
complete notice-input set, and `RuntimeDependencies.props`.
`Get-RSGhosttyRuntimeIdentityInputPaths` owns that exact closure and
`ReleaseWorkflowPolicy.Tests.ps1` compares it with the workflow key; adding or
removing a behavioral input
requires changing both in the same review unit. `BuildEvidence.psm1` hashes that same
complete closure for `ghostty_runtime_identity`; the obsolete nonexistent
`TerminalEngine.lock.json` path is never an identity input.

## Version-state coordination

`.build\version\current-version.json` and `local-build-counter.txt` are shared
repository state. The semantic `Resolve-RSVersionContext` operation owns the
exclusive repository-local `.build\version\version-state.lock` file lease while it validates saved state, reuses or allocates
the build number, constructs the requested context, and atomically persists it.
`Read-RSVersionContext` provides a locked, validated diagnostic snapshot, and
`Get-RSVersionStatePath` exposes only the canonical state path. Raw unlocked
state, counter, lock, and save primitives remain private module implementation.

The counter and current context are published through a unique same-directory
temporary file followed by atomic replace or move. An interrupted writer must leave
the prior complete file or no file, never a partially written final path. Malformed,
wrongly typed, out-of-range, or incomplete saved state fails closed instead of
resetting or guessing a value. The exclusive file handle coordinates Windows
sessions and aliases of the same physical checkout; process termination releases it
automatically. The lease is short-lived and MUST NOT span MSBuild or another native child.

Saved context is local reuse state, not proof that a particular artifact profile
completed successfully. `New-RSVersionContext` is therefore a pure constructor for
an explicit positive build number and requested configuration/platform; it performs
no shared-state I/O. Standalone MSI, symbols MSI, and ZIP callers require that build
number, Winget requires either a complete three-part version or build number, and the
standalone MSIX wrapper requires a complete three-part version. Integrated
`build.ps1` packaging passes the version already resolved for its compiled profile.

### Build evidence authority (Operation Startrail implemented)

[`Operation_Startrail_TrustedIncrementalValidationAndResumableFullEvidence_2026-08-13.md`](../Plans/Done/Operation_Startrail_TrustedIncrementalValidationAndResumableFullEvidence_2026-08-13.md)
records the completed implementation history. The active receipt and validation contracts
live in this specification and `Testing_ValidationEvidence.md`; the completed
[`cross-domain remediation`](../Plans/Done/CodeReview_CrossDomain_Remediation_2026-08-10.md)
records the `L4-BUILD-01..08` and `L4-STAR-01..16` review repairs without
remaining an active authority. Existing profile locks, containment rules,
exact-path blocking diagnostics, and no-kill ownership policy remain authoritative.

`build.ps1` computes the canonical base/index/worktree source snapshot and
toolchain/project/runtime identities before MSBuild. It resolves the effective test
surface once, passes the exact `RSBuildEnableTests=true|false` value to every project,
and binds that value into receipt identity. Debug and ASan Debug default to test-enabled;
Release defaults to production/test-disabled, and an explicit caller override is honored
in either direction.

`RSBuildCompilerDebugInformation` is an invocation-wide, receipt-bound build input with
exact values `ProgramDatabase` (the default) or `Embedded`. `Embedded` maps first-party
C++ compilation to `/Z7`, keeping compiler debug records in object files while preserving
normal final linker PDB output. `ProgramDatabase` preserves the normal profile policy, including
Debug Edit and Continue; Release and ASan Debug use the compiler PDB format. Embedded mode updates
existing `ClCompile` items as well as defaults so explicit ASan metadata cannot silently retain `/Zi`.
Real MSBuild evaluation controls cover every compile item in representative Common, plugin, app and
test projects across all six profiles and both policies, including the exact shared-TU `/MP` exception.
The Full validation runner uses this mode only for the
test-enabled full-profile build, its nested artifact-writing deployment tests and post-writer re-attestation,
where a complete Rebuild
must coexist with disjoint worktree builds without sharing MSVC's process-global compiler-PDB service.
Nested writers reuse `Get-RSBuildEnvironmentOverrides -Suite Full`; they must not reconstruct only
the test-enable flag or silently change the selected compiler debug-information mode.
MSVC 14.51 has a separate `/Z7` + `/MP` internal-compiler failure in the shared,
template-heavy `Common/SearchTextHelpers.cpp` translation unit. Embedded-mode builds
therefore compile that one translation unit without `/MP`; solution, project, and all
other translation-unit parallelism remain enabled. This is a narrow compiler workaround,
not permission to serialize the profile or disjoint worktrees.
Unknown values fail before MSBuild, and the selected value participates in receipt
compatibility.

When `RSBuildEnableTests` is false, projects beneath `Tests/` are evaluated as
no-output utility nodes so a production Release solution build does not compile a test
executable against production libraries that intentionally omit test hooks. When the
value is true, those projects retain their declared application/library types. Individual
projects MUST NOT override `RSBuildEnableTests`; mixed test surfaces inside one build graph
invalidate both ABI assumptions and receipt truthfulness.

#### Distributable test surface

A test-enabled (`ENABLE_TESTS`) build MUST never be distributed, whatever its configuration.
Test-enabled Release remains valid for validation (Full keeps `RSBuildEnableTests=true`), but
its outputs are never package inputs. Every distributable path builds with
`RSBuildEnableTests=false` and fails closed rather than packaging a test surface:

- `build.ps1` refuses `-Zip`, `-Msi`, `-Msix`, `-GenerateWingetManifest`, and
  `-OfficialRelease` before any build when the effective tests setting would be true: the
  process `RSBuildEnableTests` is `true`, or it is unset and the configuration is not Release
  (whose default is tests on). It and the standalone `Installer/msix/build-msix.ps1` pass
  `/p:RSBuildEnableTests=false` to the MSIX packaging project so its project references cannot
  be rebuilt test-enabled.
- The reusable build workflow fails a call that requests `official_release` together with
  `run_full_tests` or a configuration other than Release; `release.yml` builds Release and never
  requests `run_full_tests`.
- `Common/VersionInfo.rc` stamps every test-enabled first-party executable and DLL with
  `VS_FF_SPECIALBUILD` and the version string `SpecialBuild` =
  `RedSalamander.TestSurface=ENABLE_TESTS`. A production build carries neither.
- Native test-only entry points use reserved, case-sensitive export names: a name containing
  `SelfTest` or `ForTest`, or a C name that starts with `RedSalamander` and continues with
  `Test` or `Debug` (`SelfTest|ForTest|^RedSalamander\w*(Test|Debug)`). Production exports
  MUST NOT use these forms. `Tools/Tests/ReleaseWorkflowPolicy.Tests.ps1` holds the reviewed
  production export list and fails on any other first-party C export (`__declspec(dllexport)` or
  `PLUGFACTORY_API`) outside the reserved forms. The version marker is the primary guard: it
  covers every first-party binary, including `Common.dll`, whose `COMMON_API` C++ exports the
  export-name scan does not review; the scan is defense in depth for C entry points.
- `Tools/Modules/Packaging/ReleaseArtifactPolicy.psm1` parses each packaged PE image without
  loading it and rejects the version marker, any reserved test export, and any malformed image.
  Release portable ZIP packaging (the staged archive), MSI packaging (the staged directory),
  MSIX packaging (the produced package), `build.ps1` package outputs, and the release workflow's
  final artifact gate all run that check. MSBuild publishes an MSIX before it can be inspected,
  so a rejected MSIX is deleted with its package folder beneath `.build\AppPackages` before the
  failure is reported, while the packaging lease is held and never through a reparse point
  (one refuses the removal and the package must not be distributed); checks made after the
  lease is released never delete. Debug portable ZIPs are diagnostic, are never
  published, and would fail the release gate.

For C++ projects, the evaluated `LibraryPath` used to materialize the build child's
`LIB` environment MUST omit nonexistent directories. Missing optional
ATL/MFC or .NET SDK directories cannot satisfy the linker and otherwise cause the
Roslyn inline-task compiler to repeat CS1668 for every project.

Each invocation uses a unique artifact-manifest root. Every built first-party application
or DLL declares its target path, effective test surface, current vcpkg write-tlog inputs,
and governed runtime dependencies there. Receipt outputs and runtime closures are derived
from those current declarations, not from a recursive scan of surviving output files. A
full-solution build rejects an undeclared executable or DLL in the selected output profile.
A targeted build may coexist with unrelated pre-existing files, but it MUST NOT attest
them. This prevents stale or removed binaries from bootstrapping into new authority.

Vcpkg write-tlog entries have two supported forms: legacy `applocal.ps1` destination
paths already beneath the selected output profile, and built-in `z-applocal` source paths
beneath the repository's canonical `.build/vcpkg_installed/<platform>/` root. A source
entry resolves only to the same-named deployed file beside that project's declared target.
Arbitrary external sources, reparse source files, and missing or reparse deployed outputs
fail closed; the dependency source itself is never admitted as a receipt artifact.

Shared output directories do not confer project-local cleanup ownership. For first-party
C++ projects, dependency DLLs beneath `OutDir` are shareable; only the exact current
`TargetPath` and files beneath that project's private intermediate directory retain
ordinary DLL cleanup ownership. `Directory.Build.targets` migrates old file-list caches
before `CoreClean` and filters both prior/current lists after
`_CleanGetCurrentAndPriorFileWrites`. This ordering covers incremental orphan cleanup,
write-recording, and late vcpkg link-skipped registrations without depending on vendor
hook import order. A consumer Clean MUST NOT delete another project's deployed runtime,
including when the DLL is in use. Do not suppress MSB3061, disable runtime copying, or
serialize the solution to conceal an ownership error.

Obsolete shared DLL removal remains explicit in `RuntimeDependencies.props`; a full
receipt still rejects undeclared surviving binaries. The guard neither attests stale
files nor deletes dependencies by guessing ownership from an old consumer history.
Real MSBuild cleanup fixtures in `Tools/Tests/BuildRuntimeClean.Tests.ps1` exercise old
and new histories, incremental and explicit cleanup, locked files, path normalization,
test-disabled transitions and a non-first-party control. They run in the
`RequiresBuildToolchain` lane with all fixture data beneath the marked test root.

The central runtime-dependency manifest also owns application-sibling runtimes. A
dependency whose `OutputRoot` is `BuildOutput` is copied beside the owning executable
after its build and is attested in that executable's receipt closure.
`RedSalamanderSearchService.exe` and `sqlite3.dll` are the current required pair. Staging
and cleanup ownership are separate obligations: copying after one project's Build cannot
protect against a different project's later Clean; the shared-history guard above does.

The same ownership rule applies to Visual C++ `CoreCppClean`. Its native cleanup task
MUST exclude other projects' DLLs in the shared output directory while retaining cleanup
of the exact current target and private intermediates, including when a test project
becomes a production Utility node. Common MSBuild `CoreClean`/`FileWrites` coverage alone
does not qualify this path; build-toolchain tests also execute the installed native C++ task.

For `ASan Debug`, the shared build target MUST copy the platform-specific
`clang_rt.asan_dynamic-*.dll` before writing application artifact manifests and MUST
declare that exact staged file as a runtime dependency. Full-profile receipt publication
therefore attests the sanitizer runtime and includes it in each ASan application closure;
an untracked or missing sanitizer DLL fails closed instead of being accepted from PATH.

A tests-disabled production profile still builds `PluginContractTests.exe` without
`ENABLE_TESTS` as the non-shipping portable-package contract host. The full-profile
receipt must attest that executable together with the shipping inputs. All other disabled
test projects are utility nodes. Every disabled test clean path, including the retained
`PluginContractTests` application, deletes only its own project-named outputs; it must not
replay app-local dependency entries retained from a prior test-enabled build because those
DLLs may be shared shipping dependencies.

The build removes the mutable success pointer before mutation, forces the selected target
through `Rebuild` when no compatible receipt exists, checks that source did not change
across every selected build/package operation, hashes the declared outputs, and publishes
a schema-valid receipt only after the complete command succeeds. A compatible receipt
permits normal incremental MSBuild, but every new success republishes actual post-build
bytes.

The canonical `toolchain-identity.json` is a governed receipt output, not workflow-added
metadata. Its canonical JSON digest must equal `toolchain_identity_digest`, and its file
bytes must also match the ordinary receipt artifact size and SHA-256.

Receipts are canonical, content-addressed records under
`.build\<platform>\<configuration>\receipts\` plus the atomically replaced
`build-receipt.json` pointer. Targeted and full records coexist. The shared verified loader
recomputes `receipt_id`, loads `receipts/<receipt_id>.json`, validates both records, and
requires the mutable pointer to be byte-for-byte and semantically identical to the
immutable record. Artifact IDs and paths are unique case-insensitively, every artifact ID
is derived from its governed path, and every non-empty runtime closure refers only to
attested outputs and includes its own artifact. Receipt identity excludes timestamps,
duration, original checkout spelling, and absolute diagnostic paths while binding the
source snapshot, requested target/arguments, effective test surface, toolchain, project
graph, runtime dependency identity, output sizes/digests, and declared runtime closures.
Any missing, malformed, incompatible, aliased, reparse-routed, semantically inconsistent,
or byte-mismatched artifact fails closed.

When an official workflow builds in one job and tests a downloaded artifact in another,
the producer first revalidates every receipt output's regular-file status, size, and
SHA-256, verifies the receipt-bound canonical toolchain record, then copies exactly those
verified outputs into a unique `.build\PortablePayload\<id>` staging root. The only
producer-added file is the schema-validated `portable-build-attestation.json`; arbitrary
files found by scanning the live build output are never admitted. The artifact upload
reads this staged root rather than the mutable build-output directory.

The producer emits the attestation and toolchain digests as reusable-workflow outputs.
The consumer accepts the downloaded payload only for the exact repository, workflow,
run, attempt, producer job, source commit/snapshot, platform, configuration,
control-plane attestation ID, and toolchain digest. It rejects absolute path authority,
reparse paths, missing/tampered files, extra unmanifested payload, and any post-download
staging not present byte-for-byte in the manifest. It also validates the downloaded
toolchain record against its schema and attested digest. Only then may it rebind
repository-relative paths and publish a local receipt. A local receipt containing
absolute checkout paths is never portable CI provenance.

`Run-AllTests.ps1 -SkipBuild` requires a verified, test-enabled, full-solution receipt for
the exact source snapshot/profile and rehashes every attested output before launching a
test. After constructing the selected CI or Full plan, the runner also requires every
file-backed executable or plugin beneath the build profile to be present in the receipt's
declared closure. The same complete-plan check runs after an artifact writer rebuilds and
re-attests the profile. Path existence alone is never build authority.

Full's `RequiresBuildToolchain` deployment writer is profile-scoped. The runner supplies
the selected platform/configuration explicitly, suspends that profile's receipt before
the writer can delete bytes, and permits the writer to lock, remove, and rebuild only
`.build/<platform>/<configuration>`. Re-attestation republishes a verified receipt for
that same profile before any artifact reader is admitted. Release and ARM64 validation
must never mutate x64 Debug as a hidden side effect.

Later Startrail affected/resume phases do not change build authority until their own
normative activation gates land. A plan checkbox alone does not activate those modes.
The runner may bind the receipt ID and its declared output/runtime closure into an entry
fingerprint, but that fingerprint cannot widen the receipt's target/profile authority.

## Local artifact-operation coordination

Builds and tests that can read or replace compiled artifacts MUST coordinate by the
normalized repository root, platform, and configuration. The canonical scope is
`<repo>|<platform>|<configuration>`; target, project, and suite names are diagnostic
metadata and MUST NOT subdivide that scope. Every production caller of
`ArtifactOperationLock.psm1` MUST supply a valid platform and configuration. An
unscoped repository lock is compatibility/test infrastructure only and is not an
acceptable lock for a new artifact-mutating caller.

Operations for the same scope are exclusive from acquisition through their final
artifact read, result archival, or disk audit. Different platforms, configurations,
and worktrees use independent artifact roots and MUST NOT be blocked merely because
another RedSalamander build or test exists. A narrower shared-resource lock is
required before a future build step writes state outside its profile-specific output
and intermediate roots; broadening the artifact lock to every profile is not the
default remedy.

The current reviewed exception is
`.build\vcpkg_installed\<platform>\<triplet>`, plus the mutable vcpkg metadata
beside that triplet beneath the same platform root. x64 and ARM64 MUST use separate
install roots because manifest installation can purge packages through the root's
shared `vcpkg\status`; triplet subdirectories alone are not an isolation boundary.
Every configuration for one platform shares its platform root. MSBuild phases and
direct vcpkg installation therefore hold a second repository/platform dependency
lock while they can read or write that root. This lock does not include configuration and
must detect residual build tools for the same platform regardless of their
configuration. It is not held by test execution, sandbox work, reporting, or a
build for the other platform. An install that targets both triplets acquires the
x64 and ARM64 dependency scopes in a stable order before any clean/delete/merge.
An explicit custom triplet is accepted only when its name begins with `x64-` or
`arm64-`; any other architecture has no reviewed platform lock mapping and MUST
fail before staging or canonical install state is changed. Residual-process
matching MUST apply that same prefix mapping to `vcpkg`, compiler, and linker
command lines that reference the platform-scoped canonical root or custom staging triplet paths; process
name alone, another repository, or the opposite architecture is not evidence.

CI stages its managed executable-tool checkout and download cache under
`.build/vcpkg-tool`, matching the generated-output boundary used by source
snapshots. The checkout must not appear as an untracked nested repository at
the product root. Its executable and manifest pins remain separately verified;
the source-integrity check must not gain an exclusion for a root `vcpkg/` checkout.

The `vcpkg-tool.json` executable revision and `vcpkg.json` `builtin-baseline`
remain independent identities. When a caller materializes them as separate
checkouts, `vcpkg-install.ps1 -VcpkgExe <tool> -VcpkgRoot <registry>` MUST preserve
the executable identity check and pass the registry checkout through vcpkg's
explicit `--vcpkg-root` option. Setting `VCPKG_ROOT` does not satisfy this contract
because vcpkg ignores that environment variable when its executable is already
inside a valid checkout. An explicit registry root must contain both `ports` and
`versions` before dependency mutation begins. Before bootstrapping or installing,
CI MUST read `scripts/vcpkg-tools.json` from both pinned commits and require the
executable commit's positive `schema-version` to be at least the registry
baseline's positive `schema-version`. A manifest-baseline update that raises this
schema therefore also updates `vcpkg-tool.json`; an older executable fails with a
pin-specific diagnostic before dependency mutation.

A caller that materializes only one checkout supplies the version database from
that working tree while `builtin-baseline` is still resolved through git. The two
pins must then name the same commit: a baseline newer than the checked-out tool
commit selects baseline default versions whose exact entries are absent from the
on-disk `versions` database, and vcpkg fails with `no version database entry`. A
lane that needs the pins to diverge MUST materialize the baseline as a separate
registry checkout and pass it through `--vcpkg-root`.

Every auxiliary Git checkout a lane materializes (the executable tool, a ports
registry, or a combined checkout) MUST live beneath an excluded path such as
`.build`. An untracked nested repository elsewhere in the workspace is reported by
Git as a single directory entry and fails repository identity before MSBuild runs.

The selected registry must also contain an exact version-database entry for every
object-form manifest dependency that declares `version>=` and every exact manifest
override. The governed local installer validates those constraints before it
acquires dependency locks, creates a staging directory, or invokes vcpkg. A
dependency-floor or exact-pin change and its compatible `builtin-baseline` update
are one review unit; CI and clean machines MUST fail with the missing port/version
diagnostic rather than inheriting a package from a populated local install root.
When source behavior is audited against one dependency version, `version>=` alone
is insufficient because the baseline may select a newer default; the manifest MUST
also carry an exact override until that behavior is re-audited.

In this repository `builtin-baseline` and the `vcpkg-tool.json` commit name the same vcpkg commit; no reviewed
divergence is recorded. The offline pin contract (`Tools/Tests/VcpkgInstallSafety.Tests.ps1`, "Vcpkg dependency pins")
fails when the two differ, unless that exact baseline and tool pair is added to its reviewed-divergence table and to this
section. It also fails when a direct `version>=` floor exceeds the exact override of the same port (vcpkg silently builds
the override, so the manifest would advertise a version never built), when the `aws-sdk-cpp` override is not paired with
its audited `aws-crt-cpp` override, when the Dependabot rules below are removed, and, whenever the executable checkout
`.build/vcpkg-tool` is present at the pinned commit, when a floor or override has no exact entry in that checkout's
version database.

Dependabot's vcpkg feed MUST NOT move these audit-gated pins. It ignores `github.com/microsoft/vcpkg` (the
`builtin-baseline`), because it proposes the newest release tag, which can be older than the commit both pins name: it
proposed exactly such a downgrade on 2026-09-14 (#98) and 2026-10-05 (#111). It ignores `aws-sdk-cpp` and `aws-crt-cpp`,
which move only together as the re-audited exact-override pair (see the AWS paragraph below and
[FileSystem_S3.md](../FileSystem/FileSystem_S3.md)). A dependency change merges only with a green CI build, including
when CI cannot start.

Dependabot also MUST NOT propose any other vcpkg port update. It cannot move the baseline together with a port floor, so
a floor it raises is absent from the pinned baseline and the PR is unmergeable on its own: #111-#113 did exactly that and
#117 (`sqlite3` `3.53.4#1` to `3.53.4#2`) repeated it after the ignores above landed. `.github/dependabot.yml`
therefore carries an unrestricted wildcard ignore (`dependency-name: "*"`, no `update-types` or `versions`) under the
`ignore:` list of its vcpkg block, and the offline pin contract fails without it. The wildcard ignore is the required
mechanism: the vcpkg block itself stays declared (the pin and core-stack group contracts check it), and
`open-pull-requests-limit: 0` is not a substitute because GitHub applies it to version updates only, while ignore rules
also cover security updates.

Updating vcpkg dependencies is therefore a manual, single reviewed change:

1. Choose the target vcpkg commit. Move `vcpkg-tool.json` (executable revision) and `vcpkg.json` `builtin-baseline`
   to that same commit together, since a newer baseline can require a newer tool schema.
2. Raise the affected `version>=` floors and exact overrides in the same change, keeping every floor at or below its
   override, and keep the `aws-sdk-cpp` and `aws-crt-cpp` overrides on an audited pair. Moving that pair needs the
   re-audit described in the AWS paragraph below and [FileSystem_S3.md](../FileSystem/FileSystem_S3.md).
3. Prove availability with the governed installer gate: `vcpkg-install.ps1` against the new tool and baseline on a clean
   restore, and the offline pin contract (`Tools/Tests/VcpkgInstallSafety.Tests.ps1`), which checks each floor and
   override against the pinned checkout's version database.
4. Merge only with a green CI build.

The governed installer pins every triplet it builds to the Visual Studio installation and the default MSVC toolset the
build discovers. vcpkg otherwise builds with the newest toolset of the installation it prefers. That toolset can lack a
compiler for a target (in the historical 14.51-default installation, the side-by-side 14.52 preview lacked an x64-hosted ARM64 compiler, so every fresh
ARM64 install failed), and for the compiled libraries linked into the product binaries (curl, 7-Zip, libjpeg-turbo) it is
newer than the toolset that links them, a known link hazard. `vcpkg-install.ps1` takes the installation as `build.ps1`
does (GitHub Actions: the workflow-selected MSBuild; otherwise the newest installation with MSBuild that vswhere reports)
and the toolset from its `VC\Auxiliary\Build\Microsoft.VCToolsVersion.default.txt`, before any lock, staging directory or
download; a missing or malformed file fails and names the file and the repair. For each selected triplet it writes
`.build\vcpkg-triplets\<platform>\<triplet>.cmake`: the selected vcpkg root's own triplet (`triplets\`, then
`triplets\community\`) copied unchanged in its own line endings, then `set(VCPKG_VISUAL_STUDIO_PATH ...)` and
`set(VCPKG_PLATFORM_TOOLSET_VERSION "<major.minor>")`, rewritten only when its bytes change. It passes
`--overlay-triplets=<that directory>` to `vcpkg install`. The custom-triplet rules above are unchanged (leaf names only,
x64-/arm64- prefix mapping). A different triplet is a different package ABI hash, so the first install after this
change, and after a change of installation or default major.minor, rebuilds the packages; vcpkg's binary cache then
serves any root whose triplet text is identical. The functions live in `Tools/Modules/Build/VcpkgInstallSafety.psm1` and
keep the contract of DxUi's `Tools/VisualStudio.psm1` and `Tools/VcpkgTriplet.psm1`, a documented consumer interface of
the pinned library. They are not imported, because this installer runs before the DxUi checkout is restored (the Ghostty
ARM64 workflow installs first); a parity test compares both implementations whenever the pinned checkout is restored.
MSBuild's own manifest-mode install is not limited to an empty tree: vcpkg's `vcpkg.targets` runs it whenever `vcpkg.json`
(or `vcpkg-configuration.json`) is newer than `.build\vcpkg_installed\<platform>\.msbuildstamp-<triplet>.<host>.stamp`, a
stamp the governed installer never writes. It uses the same overlay. For first-party projects `Directory.Build.props`
appends `"--overlay-triplets=.build\vcpkg-triplets\<platform>"` (absolute, quoted, no trailing backslash) to
`VcpkgAdditionalInstallOptions` only when that directory holds `<triplet>.cmake`, so vcpkg never receives a directory that
does not exist. Before each MSBuild run, inside the platform dependency lock, `build.ps1` refreshes that file through the
same functions, installation and default toolset, copying the stock triplet of the vcpkg root MSBuild's integration
installs with (an explicit `VcpkgRoot`, else the installation's bundled `VC\vcpkg`). The `x64-windows` and `arm64-windows`
stock triplets of the bundled (VS 18) and pinned roots are byte-identical, so both writers produce the same bytes and do not
alternate the package ABI. `build.ps1` skips the refresh when `VcpkgManifestInstall=false` or when no such root has the
triplet. MSBuild's install still runs its own vcpkg root, scripts and `vcpkg\status` database, so its package ABIs differ
from the installer's staged ones: when its stamp is stale it rebuilds what its own status lacks, with the pinned toolset.
The inline install of the reusable CI workflow does not call the installer and keeps vcpkg's default toolset choice; CI then
builds with `VcpkgManifestInstall=false`. The pin matters where the default toolset is older than the newest one installed
beside it (the historical VS 18 Insiders 14.51-default/14.52-preview installation), above all for an ARM64 target on an x64 host when only the default toolset has
that compiler.

AWS SDK `1.11.880` is paired with the exact `aws-crt-cpp` `0.43.4` override,
matching the SDK tag's `crt/aws-crt-cpp` submodule. CRT `0.43.0` lacks the
ByteCursor Base64 overloads and encoded/decoded-length functions used by that
SDK, so a successful dependency restore alone cannot qualify the pair: clean
SDK compilation and the consumer Release build are required. The `aws-sdk-cpp`
floor equals that override, and the pair moves only together after that
re-audit. The vcpkg tool and registry baseline pins name one commit whose
version database contains every current manifest floor and override, including
the SQLite (`3.53.4#1`), libjpeg-turbo, curl, AWS SDK and CRT entries.
See [the SDK's CRT source](https://github.com/aws/aws-sdk-cpp/tree/1.11.880/crt/aws-crt-cpp)
and [the paired CRT declarations](https://github.com/awslabs/aws-crt-cpp/blob/v0.43.4/include/aws/crt/Types.h).

Official clean-runner builds MUST NOT assume that the Visual Studio image includes
user-wide vcpkg MSBuild integration. After the governed install succeeds, the
caller imports `scripts/buildsystems/msbuild/vcpkg.props` and `vcpkg.targets` from
the exact executable-tool checkout for that process only, disables manifest
installation during MSBuild, and binds `VcpkgInstalledDir` to the canonical
platform-scoped `.build\vcpkg_installed\<platform>\` root. A direct compiler
include fallback may name that same installed triplet. The imported build-system
files, installed headers/libraries, executable-tool checkout, and ports/version
registry therefore remain separately pinned inputs; no user-wide integration or
runner-ambient package path is an allowed substitute. These official build paths
MUST set `VcpkgXUseBuiltInApplocalDeps=false` before MSBuild. The pinned
PowerShell app-local copier records copied destination paths in the project tlog,
which lets fail-closed build receipts authenticate only outputs beneath the
selected build directory; the built-in copier's source-path tlog format is not a
receipt declaration.

The Ghostty private-runtime builder has its own reviewed shared-resource scopes.
Its source, download, toolchain, cache, and product state beneath
`.build\TerminalEngine` is shared across configurations and platforms in one
repository, so a mutex derived only from the normalized repository root serializes
x64 and ARM64 runtime operations for that repository. Different worktrees derive
different mutexes. A validated cache hit MUST NOT take the session-wide canonical
drive lease. Only an actual rebuild takes the `R:` lease, and only while it maps
`R:`, applies the temporary overlay, runs the contained build, reverses the overlay,
and unmaps `R:`. An existing mapping is never removed automatically: the builder
reports the mapping and the exact manual inspection/removal commands. Every native
child is launched through the kill-on-close containment helper.

Packaging has a separate repository-wide coordination scope because x64 and
ARM64 packaging can both publish beneath `.build\AppPackages`. The lock is acquired
only after compilation succeeds and
is held through MSIX/MSI/ZIP/winget generation; ordinary builds and tests do not
hold it. Standalone packaging entrypoints that mutate those locations MUST enter
the same scope, so invoking a child from `build.ps1` is a balanced same-thread
reentrant lease rather than an independent lock. Abandoned packaging ownership
fails closed until a reviewer restores shared installer inputs/outputs and
explicitly removes the marker; a build-profile rebuild cannot silently clear it.

A standalone packager that reads `.build\<platform>\<configuration>` acquires
that artifact profile before packaging and releases packaging before the profile.
The standalone MSIX transaction additionally acquires its platform
`shared-dependency` scope after packaging and only around contained MSBuild,
then releases dependency, packaging, and profile in reverse order. Caller-selected
package destinations stay beneath this repository's `.build\AppPackages`, and the
tracked `Installer\msix\Package.appxmanifest` is always read-only. MSIX version and
architecture stamping writes a unique CreateNew manifest beneath
`.build\AppPackages\ManifestStaging\<id>` and passes that exact path to the WAP project
through `RSAppxManifestPath`. The integrated and standalone packagers remove only that
validated staging directory in `finally` before releasing packaging coordination. A
package failure or process interruption can leave at most a disposable staged copy and
an abandoned-operation diagnostic; it cannot partially change or require reconstructing
the tracked manifest. A repository lock does not authorize an arbitrary external path.
Every existing component from the repository
root through a caller-selected destination MUST be non-reparse before and after
directory creation. Final package-file publication MUST use a unique validated
same-directory staged file, revalidate the complete guarded path immediately before
publication, and atomically move or replace the destination. A failed publication
must leave the prior complete destination intact. Native packager children use the
kill-on-close process helper. Ordinary packaging never mutates machine-global
WiX extension state: required extensions are installed explicitly before entry.

The per-profile owner record MUST authenticate the owning PID instance with process
start time and executable path, as well as the exact lock path and random token.
Nested ownership is same-thread only, and distinct nested scopes release in strict
last-in/first-out order so process delegation markers cannot name the wrong lock.
A child may reuse ownership only through the
reviewed one-use delegation handshake, as an immediate child contained in the
owner's kill-on-close Job Object. Environment variables alone are not proof of lock
ownership. The admitted child may reenter that scope locally but MUST NOT delegate
the parent's authority to another process generation. Immediate-parent identity for
the current contained process is obtained through a local native process query; CIM is
only a fallback for other diagnostic process IDs and its availability is not required
for the delegation security boundary.

Before using a profile, residual `MSBuild.exe`, `cl.exe`, and `link.exe` checks MUST
require boundary-aware evidence for both this repository and the same output profile
(an exact output/intermediate path or matching MSBuild platform/configuration
properties). A process name by itself, a substring that only resembles the repository
path, or a process belonging to another worktree/profile MUST NOT block the operation.
Failure to enumerate process metadata is a blocking diagnostic, not permission to
guess or terminate.

Path evidence accepts exact quoted paths and reviewed attached compiler/linker
switches such as `/Fo<path>` and `/OUT:<path>`, while requiring token boundaries so
near-prefix or concatenated foreign paths do not match. MSBuild property evidence
parses separate switches and semicolon/comma combined property lists. The direct
`Invoke-SanitizedMsbuild.ps1` command rejects an invocation without explicit
Configuration and Platform properties rather than falling back to a broad scope.

`build.ps1` MUST never terminate an independently launched application merely to
replace its binary. If a process executable path exactly equals an output that the
selected build can replace, the build fails with PID, executable path, and available
command-line diagnostics and asks the operator to close it. Processes at other paths
are out of scope. Automatic termination is limited to descendants that the current
operation itself launched and contained in its kill-on-close Job Object.

Likewise, `Run-AllTests.ps1 -ExePath` is compatibility syntax only for the exact
normalized `RedSalamander.exe` beneath the selected repository, platform, and
configuration. A foreign or differently scoped path fails before artifact-lock
acquisition; the runner does not broaden coordination or terminate an unrelated
process to accommodate it.

An abandoned non-empty lock or stale authenticated owner record contaminates only
that platform/configuration scope. Tests and incremental/targeted builds for the scope
remain blocked until a successful full-solution `build.ps1 -Rebuild` for the same
platform and configuration clears the marker. A repair for one profile MUST NOT clear
or block another profile. Shared-dependency contamination is platform-scoped and
requires a successful full-solution rebuild for that platform; configuration does
not partition or independently clear the shared triplet marker.
Packaging contamination is repository-wide and is never cleared as a side effect
of an artifact-profile or dependency repair.


All first-party native project definitions expose Debug, Release and ASan Debug for x64 and ARM64,
including the manual Win32HelloCred and historical Gate-0 harness/probe projects. Additional build
lanes do not rewrite or requalify sealed historical terminal evidence. Reproducing that evidence
still uses its recorded source and toolchain identity. The current ASan Debug configuration forces
compiler instrumentation in every translation unit, `/MDd`, `/Od`, and compatible debug information;
a missing app-local sanitizer runtime fails the build. The existing PluginContractTests
`--asan-seed-heap-overflow` child remains the detection probe. Native ARM64 execution is required.

The seeded probe's two intentional volatile accesses retain the selected ARM64 volatile model.
Their C4746 annotation surrounds the isolated noinline probe definition; MSVC's ARM64 ASan
compiler does not honor a statement-only annotation for these accesses. It does not suppress diagnostics in ordinary tests
or replace the required expected-failure sanitizer detection check.
Toolchain receipts identify the MSBuild-selected compiler host and its matching SDK resource compiler,
rather than assuming Hostx64 on every machine.

DxUi is public: exact-pin HTTPS source restore requires no PAT or Actions secret. Anonymous Git and
public Actions API access were verified on 2026-09-09. CI can use its automatic read-only job token
for advisory API rate limits; unavailable/rate-limited advisory queries never change or reject the pin.

The native matrix is selected per event: pull requests build x64 Debug and ARM64 Debug,
pushes to `main` add both Release profiles, and ASan Debug runs from `asan.yml`. Each
job builds the test-enabled solution, verifies the ASAN defect probe where selected,
then uses the receipt-gated runner in that same job with the suite named by the
reusable workflow's `test_suite` input (`PR` on pull requests and pushes to `main`;
`CI` on the nightly `nightly-ci.yml` matrix and the weekly ASan lane, which never runs
on pull requests; `Full` remains the local closeout gate). Cross-job nightly/package handoffs still require portable attestation.
Runtime execution is rejected if host architecture differs from the selected target.
No custom DxUi secret is required; the public pin restores through HTTPS.

The reusable workflow restores the vcpkg and pinned Terminal runtime caches with
`actions/cache/restore` and saves them with explicit `actions/cache/save` steps placed
right after the dependency install and the solution build. The combined `actions/cache`
action saves only when the whole job succeeds, so a failing test step used to discard a
successful 25-minute dependency build on every red run.

The MSBuild executable must match the host: `MSBuild\Current\Bin\arm64\MSBuild.exe` on
ARM64 hosts and `Bin\amd64\MSBuild.exe` on x64 hosts, with `Bin\MSBuild.exe` (the 32-bit
x86 MSBuild) only as a last resort. `build.ps1` (every discovery strategy: vswhere, the
common installation roots, and `VSINSTALLDIR`), `Tools\Restore-DxUi.ps1`, the DxUi
consumer test, the CI "Select MSBuild" step, and the MSIX packaging job in `release.yml`
all follow that order. The VC toolset
(`Microsoft.Cpp.ToolsetLocation.props`) derives the compiler host from the MSBuild
process's own `PROCESSOR_ARCHITECTURE`: it downgrades a requested `arm64` host to x86
whenever that process is not ARM64 and an `x64` host to x86 whenever it is neither x64
nor ARM64, so an x86 MSBuild on the hosted ARM64 runner could only drive the 32-bit
`HostX86\arm64` tools, whose heap could not optimize the largest Release self-test
translation units (C1002; the 32-bit linker restarted itself as 64-bit). With the ARM64
MSBuild, `Directory.Build.props` selects `PreferredToolArchitecture=arm64` for ARM64
targets (mirroring the x64-host rule) and the toolset keeps it. Toolchain receipts and the
DxUi consumer identity record the selected host.

Windows CI enables Git `core.longpaths` before checkout: retained test evidence includes repository paths beyond
the default Windows Git path limit. This setup runs on the disposable runner before any project validation.

CI formatting checks changed owned native files with the hash-pinned clang-format 22.1.3 wheel in
`Build/clang-format.json`. `Tools/Install-ClangFormat.ps1` uses PowerShell 7 to download and checksum-verify the wheel and extract only its native executable. No Python or pip is required for formatting. Cached executable bytes are checked against the verified archive, and a failed hash, entry or version check preserves the previous executable. `-WheelPath` permits offline installation with the same pin; `-WhatIf` downloads and writes nothing. Ghostty upstream qualification and dependency-managed build tools retain their separate Python requirements. CI does not commit changes or format archived Specs/TestRuns source snapshots.

The PR build-policy test MUST verify the formatter job invokes this installer with `pwsh` and
does not bootstrap Python or install pip packages in that job. This job-scoped wiring check
does not constrain other dependencies' Python requirements. `Tools/Tests/ClangFormat.Tests.ps1`
owns the separate offline acquisition, integrity, publication and cache-lease behavior checks.

Formatter acquisition MUST reuse `Resolve-RSGuardedRepositoryPath` and
`Publish-RSGuardedRepositoryFile` for its `.build/format` cache, download staging and
executable publication. Every existing output component from the repository root through
the leaf MUST be admitted without reparses, including `.build` and executable parents;
paths are revalidated before writes, execution, publication and temporary cleanup. If the
chain becomes unsafe, cleanup MUST preserve redirected targets and retain the temporary
artifact for governed cleanup. These path checks do not establish an OS security boundary
against a process with permission to race filesystem changes.

The installer MUST hold an exclusive `.build/format/.install.lock` file lease while mutating
the shared formatter cache; contention fails immediately without blocking native artifact
profiles. The empty lock file remains after closing. `-WhatIf` creates no directories or
lease file. An explicit `-WheelPath` MUST resolve once through the PowerShell FileSystem
provider against the caller's PowerShell location. Hash verification and ZIP parsing MUST
consume the same open stream held against writes/deletion until extraction and publication
finish, regardless of the process current directory.

The reusable CI build step checks the native build exit status immediately, before
any matrix validator or other native command can overwrite it. A compiler failure
must fail the build step and prevent native test/provenance acceptance; a missing
receipt is not the primary diagnosis for an already failed compilation.

## MSVC 14.52 Release compilation and linkage

All native profiles require the installed default MSVC **14.52.x** family with the v145 platform toolset. Discovery remains installation-independent. Restore rejects another default family before dependency compilation; MSBuild validates the selected family before compilation. Record exact patch/compiler/linker identities; GL archives and their final links require exactly matching tools.

Release explicitly uses /Gy and /GL, /OPT:REF and full /LTCG, with ordinary incremental linking disabled. Debug/ASan diagnostic optimization settings are unchanged. Preserve existing project-specific size/speed, warning, ISA, CRT and subsystem choices. Evaluated-profile tests check all six native configurations, pinned archives, and negative compiler/linker controls; resource-only satellites contain no C++ bodies.

Project references explicitly forward WholeProgramOptimization and VCToolsVersion into the adopted DxUi archive. Dependency identity includes releaseOptimizationPolicy=Gy-REF-GL-LTCG-full-v1 and is rehashed before choosing writable outputs. This separates the new archive policy without changing the adopted API/source pin. Keep header/library/compiler/CRT identities matched.

LTO is scoped to each final EXE/DLL and does not guarantee reduced size or improved speed. Retain baseline/candidate image hashes, sizes, build/link durations, fixture and dependency identities, every runtime/resource sample and concurrent-activity observations. Different compiler policies are an intentional experiment, not identical-profile evidence reuse. Cross-built ARM64 and headless functional passes cannot close native ARM64, interactive, or existing product/performance acceptance gates.

### CI toolchain preparation and audit evidence

Active Windows compilation/tooling workflows use the local setup-msvc composite action. It reads RequiredMsvcToolsVersion, verifies the discovered default compiler, linker, ASan target libraries, SDK and x64 LLVM scanner tools, and provisions Visual Studio 2026 Insiders only on GitHub-hosted runners when necessary. The Microsoft-signed GUI bootstrapper runs under RUNNER_TEMP through the canonical contained streaming-process helper, with a 45-minute bound and retained stdout/stderr logs. The action waits for its actual exit, accepts only 0 or 3010, then re-verifies every required tool; a shell-global LASTEXITCODE is not installer-completion evidence. Existing installations are retained. A channel that no longer supplies 14.52 fails rather than accepting another compiler. Local -VerifyOnly is read-only; automatic local installation is refused. Hosted provisioning requires an actual successful CI receipt; local verification/refusal controls alone do not qualify it.

[Full evaluated settings and differences](../../Measurements/Toolchain/2026-10-23-msvc1452/Settings.md) include both platform/configuration matrices and exact per-file metadata. [Microsoft installation guidance](https://learn.microsoft.com/en-us/cpp/overview/acquire-msvc) documents channel-specific acquisition. The published [runner image inventory](https://github.com/actions/runner-images/blob/main/images/windows/Windows2025-VS2026-Readme.md) does not establish that the required toolset is installed; runtime versions alone are also not compiler proof.


### MSVC 14.52 ARM64 AWS backend compatibility

ARM64 dependency installation selects Build/vcpkg-ports/aws-c-common, preserving aws-c-common 1.0.0, its upstream SHA512 and existing CRT patch. A compiler/architecture-scoped CMake helper applies /Ob1 only to source/hash_table.c under the exact MSVC C compiler 19.52.37025.0, retaining /O2 elsewhere. The exact 14.52.37025 command crashes in lookup3.inl:640 with C1001; removing /MP, using /Zi, or /O1 does not repair it, while /Ob1 and /Ob0 compile. Canonical package/product builds remain required. This ARM64-only dependency change has no effect on the x64 size/performance experiment. Native ARM64/S3 performance remains unmeasured. The CMake regression checks affected and other compiler/architecture/file combinations; retained repros live under .build/msvc1452/compiler-repro-attempt2/3.

Workflow-selected MSBuild paths may use a custom or hosted temporary installation root. Discovery validates the executable's VS18-or-newer version, invariant MSBuild.exe binary identity and standard MSBuild/Current/Bin host layout; it must not require a Microsoft Visual Studio parent-directory name. Required MSVC14.52/compiler/SDK checks remain independent and mandatory.

The [native policy, console and dependency-path contract](Build_ConsoleAndPaths.md) owns /Wall without /WX, zero-warning acceptance, Debug/ASan symbols, Release-only ICF, common terminal results and verified short dependency folders.

The native build-policy candidate locks DxUi API4 at `1f385ee013189e967ad5298f2311801a066de2c7`. It preserves this consumer's existing API adoption while repairing native warning diagnostics and stable payload type identity under Release COMDAT folding. Restore and compile matching headers/archive from this exact lock. The indexed NativeBuildPolicyCoherence plan records outstanding product and desktop-focus qualification; this pin selection alone is not a green product receipt.
