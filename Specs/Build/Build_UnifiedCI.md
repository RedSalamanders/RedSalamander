# Unified organization CI contract

Qualified SDK builds retain the actual compiler search-path guard after the
project's `LibraryPath` normalization. Freshly absent canonical optional ATL, selected VC
`Auxiliary/VS/lib/<architecture>` support and versioned NETFX library directories may be
omitted; core MSVC/Windows/CRT defaults,
existing optional entries, ordering, include/reference paths and immutable input
identity remain required. See the shared [SDK cache contract](https://github.com/RedSalamanders/.github/blob/main/Specs/SDK_Cache.md).
This compatibility rule grants no native or production acceptance; both main
regression branches must still pass before production can start.

The manual native suite defaults to Full and may explicitly select CommandsHandoff,
FileOpsVolumes or ViewerSpaceCapture, each mapped to one fixed diagnostic case. The adapter permits filtered runner
exit 2 only after validating the exact archived native cases, quarantine state, profile and
source-bound full-solution receipt. It labels that result filtered and NOT_EVALUATED for
repository coverage and creates no regression or production receipt. Other phases reject these
scopes before mutation; the pinned native matrix, desktop lease and build contracts still apply.

The manual `DesktopProbe` scope reads desktop and session facts on the same native hosted
matrix before dependency installation or builds. It records calling window-station and
thread-desktop names, input and foreground-desktop names, and individual Windows query
errors in source-bound `desktop.json` evidence. An unavailable name is unknown, not proof
of a separate desktop. This scope sends no input, activates no window and changes no
display or session settings. It remains NOT_EVALUATED and creates no test, regression or
production receipt. Other phases and local invocation reject it before engine/sandbox work.

Project GitHub-hosted native interactive entries use an owned input desktop under the
runner's per-entry interactive lease. The original native hotkey case and verified desktop
restoration passed on both native architectures before this fixture was enabled. Creation,
verified neutral foreground/focus ownership, native input assertions, contained child drain,
restoration, neutral-window thread retirement and handle retirement remain
required; successful fixture setup alone does not qualify a suite. Requested fixture identity
is bound to validation fingerprints. Local/self-hosted execution retains its current desktop.
R0ePaired also retains the original desktop on both sources so historical and current
implementations use the same fixture. The adapter clears unselected inherited requests and
restores its caller's value afterward. See the self-test ownership contract.

The separate manual `R0ePaired` scope is Release-only. It requires a nonzero full lowercase
baseline commit SHA in `baseline_ref`, an available exact commit object, and a clean candidate
whose HEAD equals the dispatch SHA and frozen descriptor. The candidate and an isolated,
retained baseline worktree use full-solution test-enabled Embedded Release builds with the
same positive frozen hosted build number and verified matching dependency/toolchain identity.
Run five fresh canonical `FileOps_ProviderCapabilityMatrix` processes per side on the same
native architecture and machine, at timeout scale 1 and repeat 1. Verify the reviewed source
projections, raw archive closure, profile, quarantine, counts and timing gates before reduction.
The owning [performance contract](../Testing/Testing_PerformanceValidation.md) retains the
two-second ceiling and matched p95 bound.

This paired study always leaves repository coverage NOT_EVALUATED and produces no regression
or production receipt. A baseline may be described as timing-only RED solely when all other
reviewed predicate clauses and complete raw evidence are verified. Its failed native result
remains failed; paired-command exit 2 makes the hosted job fail after preserving evidence.
Candidate failures, incomplete archives, other failed assertions or mismatched identities
also fail the job. A successful diagnostic pair never substitutes for complete native
qualification. Retain partial logs, run IDs and the owned baseline checkout on interruption.

This specification owns the project's CI routing. The versioned implementation and policy tests live in
[RedSalamanders/.github](https://github.com/RedSalamanders/.github). The exact reviewed full commit is recorded
in .ci/project.json and every reusable workflow caller. Update those pins together through a reviewed PR.
Older references to automatic native PR matrices or nightly native runs are superseded by this contract.

PRs run changed-source clang-format feedback and the reviewed compiler-free static subset. Formatting drift
fails and retains a patch. The explicit Commit PR formatting action accepts an open same-repository PR,
checks its current head, commits only selected formatted source paths by ordinary fast-forward push and
explicitly dispatches fresh compiler-free checks for that exact head. Forks retain a patch for their author.
PR checks never count as native acceptance. Compiler-dependent tooling runs in main qualification.

After main has remained unchanged for five continuous minutes, its cheap coordinator submits one native
admission. Cumulative native identity uses the complete committed Git input tree, file modes, dependencies,
project adapter and exact engine/policy pins. It is not a truncated files API response or last-merge-only diff.
A code merge followed by prose still requires qualification. Excluded-only edits retain the previous actual
success, failure, cancellation or pending status. Failed unchanged inputs require an explicit fresh retry.
Unknown paths are meaningful; unreadable inputs fail closed. Renames inspect both paths and deletions count.
There is no maximum-deferral override while merges continue.

.ci/project.json declares exclusions and runtime exceptions. Prose, plans, historical measurements and Changes
fragments normally do not require a native build. Shipped schemas, test fixtures, runtime examples, generated
build inputs, resources, dependencies and tooling remain meaningful, including runtime files with .md suffix.
Do not exclude all of Specs or docs without retaining their product/runtime exceptions.

The native workflow has a separate non-cancelling serialized queue. New pushes may cancel only obsolete
cheap coordinators. A stale queued automatic candidate exits before provisioning a compiler. Each admitted
pipeline freezes one commit, input digest, engine and version. Both native x64 and native ARM64 execute:

1. Release with test targets/hooks enabled; matching source/dependency/toolchain/build evidence is retained.
2. Actual normal regression execution through the canonical test harness and an always-on Release failure probe.
3. A global AND gate verifies both architectures before either production Release job may begin.
4. Separate production-only Release outputs, clean-extraction native smoke and ZIP artifact retention.

These are separate test and production roles. They intentionally compile production code twice; neither a PR
matrix nor a second push matrix repeats the same role. Test receipts and production artifacts never share
mutable cached object outputs. Dependency binary caches include the engine, native architecture, manifests,
locks, restore/toolchain scripts, props/targets and adapter closure. Cache reuse cannot waive a test gate.

Receipts bind repository, exact source/input identity, workflow run, architecture, role, count unit, positive
discovery/execution, active failure probe and artifact digests. Canonical per-case evidence remains authoritative
where the harness provides it; suite counts are explicitly suite counts, not invented native case counts.
Missing/zero/filtered execution, an unexpected skip, timeout, disabled failure checks or an identity mismatch
cannot qualify production. A focused pass is not Full repository or release qualification.

Debug, ASan, paired benchmarks, repeated stress, and other extended obligations run through explicit manual
qualification. Their canonical warnings, native-host rules, budgets, evidence preservation and desktop leases
remain unchanged. There are no scheduled native builds. Scheduled dependency, security and status automation
remains enabled; CodeQL uses explicit C/C++ build mode none, which does not invoke a native compiler.
Compiler-free CodeQL does not analyze generated code as completely as a build-based scan:
[GitHub build-mode contract](https://docs.github.com/en/code-security/how-tos/find-and-fix-code-vulnerabilities/manage-your-configuration/codeql-for-compiled-languages).

The manual Flake Convergence workflow uploads governed evidence and live selftest artifacts,
including hidden evidence metadata, only when its selected test root is available. It excludes
completed-run scratch, selftest work trees and VHD/VHDX/AVHDX payloads; it never uploads the whole
sandbox or uses a drive-root fallback. These upload exclusions preserve local retained evidence.

Publication is an explicit trusted-main action, defaulting to prepare-only. Select an original successful
qualify.yml run with complete matching receipts and retained ZIPs for both native architectures. Publication
promotes those exact bytes without building, restamping, signing afterwards or substituting an older green
run for a later failure. Missing/expired artifacts require a fresh explicit qualification. Existing conflicting
tags, assets or winget submissions are retained for review and never overwritten. Product releases publish
ZIP assets only. Only the five application products have hosted publication and winget support; DxUi SDK ZIPs remain internal artifacts and DxUi has no publication workflow. winget requires anonymous immutable
release URLs, matching ZIP hashes, the configured nested executable, an explicit product license and a working
publication credential. Private visibility or unavailable credentials are readiness gates, never permission
to change visibility, billing, licensing or access.

Implementation closeout does not qualify a product. Actual hosted acceptance remains source-bound and is tracked separately in the organization hosted-acceptance plan.
Source changes, tests written, static checks, and the shared engine's own behavior tests alone do not establish
that all six native product pipelines passed.

## Project adapter

Workflow controls execute the pinned manual platform selector for All, x64 and ARM64 and
verify both runner labels. Git fixture repositories that deliberately exceed MAX_PATH enable
Git long-path handling for initialization and every subsequent operation. Repeated test invocations
use distinct owned scratch cases; a prior run's output is evidence, not a destination to overwrite.

The adapter resolves a relative `EngineRoot` from the current PowerShell filesystem location before importing the engine; non-filesystem providers are rejected. Absolute engine paths retain their existing meaning.

The adapter uses RSBuildEnableTests=true for Release test builds and Suite CI through the hosted desktop wrapper. Production uses RSBuildEnableTests=false, -OfficialRelease and -ProductionOnly. The exact tested PluginContractTests.exe is downloaded as an immutable packaging input, checked against its regression digest and reused solely for portable smoke; it is not recompiled or shipped. Generated test scratch, logs and evidence stay in the governed marked RedSalamander.Perf sandbox. Downloaded frozen receipts and harness binaries are build inputs under .build/ci; native test generation does not write there. Production packaging retains the RuntimeDependencies.props and clean-extraction contracts. Full artifact-mutating qualification remains an explicit separate obligation.



The hosted adapter retains alternate-volume admission through the canonical sandbox/VHD contract; it does not discover or initialize unapproved volumes. Its manual action offers Normal (CI) and Full scope. ASan explicitly requires the seeded overflow detector proof before suite success. Canonical Pester execution is pinned to 3.4.0. The source-and-lock-bound online Ghostty observer remains a compiler-free publication preflight, and its failed report is retained. The owner confirmed on 2026-10-30 that the shared native x64/ARM64 Release regression gate replaces the former separate Ghostty Debug ARM64 publication gate. Debug, ASan and extended candidate investigations remain available manually; this publication decision does not close I50/I51's independently owned runtime findings. The frozen historical Gate0 workflow remains disabled through the control plane.

Native build phases restore the exact vcpkg-tool.json executable and vcpkg.json ports baseline through Initialize-RSVcpkgToolContext, then invoke vcpkg-install.ps1 for the selected platform. Process-local pinned props/targets, include/library paths and the canonical platform-private install tree are applied before MSBuild; hosted images and warm caches are not dependency prerequisites. ReleaseTest executes the already built profile without installing or rebuilding. The same prerequisite applies to production and manual native build roles.

## Emergency PR merge

An organization owner may explicitly bypass PR merge checks during a recorded emergency through the GitHub PR override or gh pr merge --admin --match-head-commit FULL_SHA. Record the reason, blocked checks, exact reviewed head, rollback and follow-up in the PR and merge commit. The default-branch ruleset uses OrganizationAdmin with pull_request bypass; automation receives no bypass permission and direct/force pushes stay protected. Failed checks remain their actual results, and normal main qualification still runs. Production ZIP/SDK and publication requirements remain unchanged.

Follow the common [emergency merge runbook](https://github.com/RedSalamanders/.github/blob/main/Specs/EmergencyMerge.md), including the reason file and exact-head CLI procedure. No emergency merge is an ordinary fallback after a failed merge attempt. Account restrictions must be resolved before enabling the ruleset; the pending RedPrism/RedTilde public visibility change remains the owner's operation.

## Dependency metadata coverage

The organization's shared engine runs a weekly and explicit manual metadata observer across all six default branches. It inventories the coordinated engine SHA, vcpkg baseline/tool/floors/overrides and custom dependency locks, and reports available producer updates as advisories. It never restores or compiles native dependencies, edits pins or treats an upstream main commit as qualified. Missing/private or undeployed repositories remain explicit incomplete coverage. Built-in dependency review and security alerts retain their supported scopes; custom lock inventory is not a native/transitive vulnerability scan. See the common [dependency status contract](https://github.com/RedSalamanders/.github/blob/main/Specs/CI.md).

Application publication uses WINGET_TOKEN. The owner's 2026-10-31 access policy selects the current public organization repositories except DxUi; reconcile that selection when repository visibility changes or a repository is added. The two private applications remain excluded until their separate readiness work and public transition. Repository secrets take precedence over the organization secret; remove an existing override only with the owner's explicit authorization. An organization owner manages the shared secret through GitHub settings; the publication token requires public_repo, not organization-secret administration. GitHub Free does not deliver organization secrets to private repositories, so repository selection alone does not establish their hosted readiness. See [GitHub's secret setup](https://docs.github.com/en/actions/how-tos/write-workflows/choose-what-workflows-do/use-secrets).

The explicitly dispatched trusted-main RedSalamander promote-publication-credential.yml path now performs read-only credential verification, retaining its historical filename for compatibility. It authenticates the effective WINGET_TOKEN through GET /user and requires the classic token's declared public_repo or repo scope. It never prints account data, exports a token, checks out code, compiles, changes secrets, creates a fork or publishes a product. A successful result proves authentication and declared scope only; it does not exercise write access or establish native installation, artifact or publication acceptance. Existing overrides must be accounted for before claiming the shared organization credential was checked. Missing permissions or failed native eligibility remain open acceptance gates.

## Evidence upload roots

The adapter's governed sandbox evidence is uploaded separately from workspace build logs and reports. Never combine
these roots in one upload path list: a Windows drive-root common ancestor can resolve as a drive-relative workspace
path and cause the artifact client to reject sandbox files. Keep test evidence in its marked sandbox, upload that
single explicit directory only when emitted, and retain workspace diagnostics independently after early failure.
## Dependency cache reuse

The pinned shared engine restores vcpkg's ABI-indexed package cache for normal and manual native qualification. Exact engine/input keys are preferred; fallback lookup stays within the same runner OS, host architecture and target platform, and the pinned vcpkg restore still selects each package by its computed ABI. A miss rebuilds the package. This supplies dependency inputs only: product objects, DxUi SDK outputs and build/test receipts are not restored by this cache. Dependency pins, native execution, failure detectors and production/publication gates remain required. The owning behavior is the shared engine's Specs/CI.md.

The pinned shared conformance rejects extension-wide Markdown exclusions. Pin maintenance migrates only the original global default to owned prose locations and known guides. Runtime/test include exceptions retain precedence; historical artifacts keep their original source identity.

Eligible Main ReleaseTestBuild and Production adapters may prepare the exact pinned qualified SDK input through the shared engine, then pass its independent descriptor to build.ps1. Static, ReleaseTest and manual Debug/ASan/Qualification do not prepare it. The build receipt and SDK reader still independently govern artifact reuse and production packaging; dependency availability never grants regression success or publication. No native/local obligation is removed.

Prepare SDK inputs before applying the returned vcpkg build environment. SDK
builds and their subsequent ReleaseTest reuse retain canonical vcpkg MSBuild
imports and installed paths but omit the helper's synthetic `CL` include override.
Source and manual qualification keep their existing environment. Never clear an
inherited caller compiler override: the SDK ABI reader must independently reject
unsupported overrides rather than accepting a different compiler environment.
## Filtered Ghostty product evidence

The manual Ghostty product qualification's Commands filters accept the canonical
runner's exit 2 only as directed evidence with repository coverage NOT_EVALUATED.
`Assert-RSHostedFilteredCaseRun` dispatches by the actual runner summary schema.
Durable PR/CI/Full runs write v2 with a v1 sibling and retain strict promoted-entry
validation. Single-suite Commands, FileOps and Compare runs write only v1 at the
stable summary path. A filtered Commands v1 result requires one archive-bound
prelaunch fingerprint/plan/decision proof (including native rotation to `previous_run`),
matching initial/post-entry/final source snapshots, the exact current byte-verified
receipt and filtered plan, one successful owned native PID/exit and matching timed
raw metrics. Missing, stale, ambiguous or malformed proof fails; v1 does not acquire
v2 promotion. Both paths require matching invocation window/filter/profile and
clean quarantine. The nonblocking sandbox audit remains visible without requiring
an unrelated globally clean root. The native Commands payload must match its completed archive
inventory hash and length, contain unique passing cases selected by the exact name
or trailing-underscore prefix, and agree with both aggregate counts. The current
source snapshot, executable path and byte-verified build receipt must match the durable v2
record or the validated v1 prelaunch proof, respectively. The exact application project-union receipt remains admissible under the
ordinary filtered-run contract; these checks grant no Full or release acceptance.
In-process hosted callers must capture information stream 6 alongside success output
because the runner-owned `Run id:` line is emitted by `Write-Host`; capturing that line
does not replace the exact source, profile, receipt and archived-case checks above.

Durable runners write v2 at `run-all-tests-results.json` and retain v1 at
`run-all-tests-results-v1.json`; single-suite runners keep v1 at the stable path.
Fixed diagnostic readers select the v1 sibling only when the stable summary is v2;
they continue verifying native result bytes against archive inventory. Aggregates
are generated after the native archive transaction and are not inventory members.
