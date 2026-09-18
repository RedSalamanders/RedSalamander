# Official Build Toolchain Identity

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

## Pinned DxUi consumption

`Dependencies/DxUi.lock.json` selects API revision 2 and the single `DxUi.lib` target.
All DxUi-using production projects consume that pin: RedSalamander, RedConfigure, Monitor, Terminal
and the viewer plugins. ProductUiTests, RedConfigureTests and PerformanceTests2 use the same imports.
The old library and DxUiTests projects are retired. Shared implementation and control tests live
in the canonical DxUi repository; product adapters and their regressions remain in this repository.
Removing legacy files does not complete the runtime/resource gates recorded in I19.
Consumers import `Build/RedSalamander.DxUi.props` and `.targets`; they never enumerate library sources.
The root build restores the exact source with `Tools/Modules/Build/DxUiDependency.psm1`, or developers
can use `Tools/Restore-DxUi.ps1` before an IDE build. Restore owns the repository/platform dependency
lease, contains its native child processes and leaves sibling repositories unchanged.

Separate platform properties select an output fingerprint containing source/API, compiler host and
compiler/linker/MSBuild/SDK identities, CRT family and STL annotation policy. RedSalamander explicitly
sets `DxUiDisableStlAnnotations=true` to match its ordinary vcpkg C++ dependencies; ASAN heap/stack
instrumentation remains required. Debug, Release and ASan Debug on x64/ARM64 are the target matrix.

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
normal final linker PDB output. The Full validation runner uses this mode only for the
test-enabled full-profile build and post-writer re-attestation, where a complete Rebuild
must coexist with disjoint worktree builds without sharing MSVC's process-global compiler-PDB service.
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
Toolchain receipts identify the MSBuild-selected compiler host and its matching SDK resource compiler,
rather than assuming Hostx64 on every machine.

DxUi is public: exact-pin HTTPS source restore requires no PAT or Actions secret. Anonymous Git and
public Actions API access were verified on 2026-09-09. CI can use its automatic read-only job token
for advisory API rate limits; unavailable/rate-limited advisory queries never change or reject the pin.

The native matrix is selected per event: pull requests build x64 Debug and ARM64 Debug,
pushes to `main` add both Release profiles, and ASan Debug runs from `asan.yml`. Each
job builds the test-enabled solution, verifies the ASAN defect probe where selected,
then uses the receipt-gated runner in that same job with the suite named by the
reusable workflow's `test_suite` input (`CI` by default; `Full` remains the local
closeout gate). Cross-job nightly/package handoffs still require portable attestation.
Runtime execution is rejected if host architecture differs from the selected target.
No custom DxUi secret is required; the public pin restores through HTTPS.

The reusable workflow restores the vcpkg and pinned Terminal runtime caches with
`actions/cache/restore` and saves them with explicit `actions/cache/save` steps placed
right after the dependency install and the solution build. The combined `actions/cache`
action saves only when the whole job succeeds, so a failing test step used to discard a
successful 25-minute dependency build on every red run.

On an ARM64 host, `Directory.Build.props` selects the native ARM64-hosted tools
(`PreferredToolArchitecture=arm64`) for ARM64 targets, mirroring the x64-host rule. Left
unset, the toolset fell back to the 32-bit x86-hosted cross tools (`HostX86\arm64`), whose
heap could not optimize the largest Release self-test translation units: the hosted ARM64
Release builds failed with C1002 and the 32-bit linker restarted as 64-bit. Toolchain
receipts and the DxUi consumer identity record the selected host.

Windows CI enables Git `core.longpaths` before checkout: retained test evidence includes repository paths beyond
the default Windows Git path limit. This setup runs on the disposable runner before any project validation.

CI formatting checks changed owned native files with the hash-pinned clang-format 22.1.3 wheel in
`Build/requirements-format.txt`. It does not commit changes or format archived Specs/TestRuns source snapshots.

The reusable CI build step checks the native build exit status immediately, before
any matrix validator or other native command can overwrite it. A compiler failure
must fail the build step and prevent native test/provenance acceptance; a missing
receipt is not the primary diagnosis for an already failed compilation.
