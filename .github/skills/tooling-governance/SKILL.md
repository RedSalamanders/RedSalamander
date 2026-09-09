---
name: tooling-governance
description: Add, modify, move, deprecate, or review RedSalamander repository tools without creating inventory, help, caller, compatibility, CI, or specification drift. Use for any change under Tools/, for scripts consumed by build or workflows, for PowerShell modules and Pester support, and whenever introducing a new developer or automation command.
---

# Tooling Governance

Read `Specs/Testing/Testing_ToolingGovernance.md` and `Tools/README.md` before
changing tooling. Treat `Tools/tool-inventory.json` as the machine contract for
classification and ownership.

## Classify before editing

For every added or changed file, determine:

1. whether it is a public command, automation entrypoint, private module, test,
   data file, compatibility shim, generated artifact, or sealed historical tool;
2. all direct consumers using `rg` across scripts, MSBuild, workflows, specs,
   documentation, and tests;
3. its prerequisites, side effects, outputs, owning spec, and lifecycle;
4. whether its path or exact bytes participate in a cache key, frozen identity,
   build hook, release process, or external operator contract.

Update the inventory in the same change. A file is not complete merely because
it is reachable through a folder-wide inventory rule; public and root-level
files require an explicit, meaningful record.

## Choose the correct surface

- Keep `Tools/` root for stable human or automation entrypoints plus navigation
  and machine inventory files.
- Put reusable implementation in a domain module under `Tools/Modules/` and
  export only the reviewed public functions with `Export-ModuleMember`.
- Put Pester tests under `Tools/Tests/`; share common test helpers through the
  reviewed test-support module rather than copying them.
- Keep current Terminal runtime helpers separate from sealed Gate-0/Round-4
  selection tooling. Never make historical behavior part of default CI solely
  because it remains reproducible.
- Retain a historical file only when it is an explicit manual entrypoint or is
  required by the minimal transitive closure of one. Record exact
  `manualEntryPoints` in inventory. Delete orphan history; a completed plan can
  preserve its old path/digest without keeping the file live.
- Keep CI and Full on `CurrentToolingProfile.Tests.ps1`; obtain sealed membership
  only from `HistoricalTerminalProfiles.json`, preserve recorded paths and
  SHA-256 values, and use the explicit Gate0/Round4 archival launcher for
  historical reproduction.
- Use canonical `Verb-Noun.ps1` names for new public commands.

Do not add a function-only `.ps1` beside runnable commands. Do not introduce
top-level side effects in a `.psm1` import. Do not dot-source a broad helper when
an explicitly exported module can express the dependency.

## Coordinate builds and tests narrowly

When a command reads or replaces compiled artifacts, reuse
`Tools/Modules/Build/ArtifactOperationLock.psm1` and pass an explicit
`configuration` and `platform`. Same-repository operations for that output
profile serialize; project, target, and suite names are diagnostics rather than
additional lock partitions. Do not introduce a repository-wide lock that blocks
disjoint profiles or worktrees. If a new step writes genuinely shared state,
give that resource its own reviewed lock and tests.

For the current build graph, MSBuild and vcpkg installation use the separate
repository+platform `shared-dependency` scope because all configurations share
one vcpkg triplet. Keep that lock out of ordinary tests and reporting, match
same-platform residual build tools across configurations, and acquire both
platform scopes in stable order before a two-triplet install mutates anything.

Packaging is the other reviewed shared scope: platform builds share MSIX inputs
and `.build\AppPackages`, so build orchestration and standalone packaging
mutators enter repository-wide `coordination=packaging` only for generation.
Standalone packagers that read compiled output acquire the artifact profile
first and packaging second, then release in reverse. The standalone MSIX wrapper
also acquires its platform `shared-dependency` lease only around contained
MSBuild, after packaging, and releases it first. Keep ordinary compilation and
tests outside the packaging scope. An abandoned packaging marker requires
reviewed restoration and is not automatically cleared by a profile rebuild.

Version state is a separate short-lived repository transaction. Use
`Resolve-RSVersionContext` for locked validation, allocation, construction, and
atomic persistence; use pure `New-RSVersionContext` when the build number is
already explicit. `Read-RSVersionContext` is a locked, validated diagnostic read;
raw unlocked lock/read/save primitives stay private. The lock is an exclusive
repository-local file lease, so it coordinates Windows sessions and aliases of the
same physical checkout, is released after owner termination, and never spans a
child process. Standalone packaging requires an explicit
build number or complete version and never treats saved context as proof that
matching artifacts were produced. Winget additionally requires both exact-version
x64/ARM64 portable archives. Any caller-selected package destination must remain
beneath its reviewed repository root with no existing reparse component before or
after directory creation.

The selected Ghostty builder serializes `.build\TerminalEngine` through a
repository-derived mutex, contains native children, and takes the session-wide
`R:` mapping lease only during an actual rebuild. Its workflow cache key includes
the builder, every direct module import, the current patch, and all notice inputs.

Process ownership must be proven, not inferred from an executable name. Residual
build-tool checks require boundary-aware repository and profile evidence. A
build encountering an independently launched executable at the exact output path
fails with diagnostics and asks the operator to close it; it does not terminate
the process. Automatic termination is limited to descendants the current command
launched inside its kill-on-close Job Object.

Distinct nested coordination scopes are same-thread LIFO. A contained immediate
child may consume one delegation and reenter locally, but it cannot forward that
authority to a second process generation. Direct MSBuild wrappers require explicit
Configuration and Platform properties and parse combined property lists rather
than silently choosing a broad fallback.

Test-plan entries that touch foreground focus, UI Automation, pointer, clipboard,
prompts, or desktop-global ETW state set `RequiresInteractiveDesktop`. The runner
holds the session mutex only around that entry execution, including a matching
quarantine repair attempt. Noninteractive tests, builds, setup, archival, and
cleanup stay outside that mutex. Changes to this policy require focused guards for
same-profile exclusion, disjoint-profile concurrency, foreign-path immunity,
no-kill behavior, and plan-marker propagation.

## Trusted validation discipline

When changing trusted incremental validation or resumable evidence, read
`Specs/Testing/Testing_ValidationEvidence.md` and the Operation Startrail implementation
history. Fresh is the default and final-closeout authority. Explicit local exact Resume
and Affected are active but may not be widened without a new reviewed plan.

Classify every added schema, module, command, fixture, and behavioral data file in
`Tools/tool-inventory.json`. Audit all aggregate and runner consumers, including archive
validation, reporting, MTP live closeout, workflows, docs, and source-contract tests.
Treat artifact mutation as an explicit read/write/epoch contract. Any change to
`Affected`, `Resume`, or Commands-family behavior must update help, exit categories,
evidence roots, schemas, consumer compatibility, and focused tests together.

Validation runner entries use stable semantic IDs and schema-complete metadata. When a
tool test begins building, staging, or rewriting profile artifacts, move it behind the
`RequiresBuildToolchain` tag and the explicit `ToolsPesterBuildToolchain` plan entry;
never leave it hidden in the read-only Tools Pester entry. Update quarantine projection,
live inventory JSON, plan digest/schema tests, and the Testing specs whenever entry
behavior, coverage, outcome/skip policy, artifacts, environment, or resources change.

Phase 3 workspace and entry identity is active. Changes to the runner, plan modules,
fingerprint module, or their reviewed schemas MUST remain in the runner dependency
digest closure. New environment inputs require an explicit non-secret capability
projection; unknown inputs fail planning. Generated evidence must stay inside reviewed
excluded roots so it cannot invalidate the snapshot that authorized its own run.

Fresh impact planning remains shadow-only; Affected enforces only its filtered execution
set and Resume enforces only exact promoted reuse. Update `Tools/validation-impact.json`
and `ValidationImpact.Tests.ps1` when adding a project, import, runtime dependency,
runner/schema dependency, or new path class. Specific rules accumulate; unknown relevant
paths widen to Full. Use `Run-AllTests.ps1 -Suite Full -ExplainPlan` to review scope, but
never interpret an unchanged shadow entry as reusable evidence outside explicit Resume.

CI/Full checkpoints, aggregate v2, and exact Resume are active. When a
test writes build outputs, declare its read/write/produced and receipt-invalidation sets,
place it before readers or open a new modeled epoch, and add a re-attestation barrier.
The current `ToolsPesterBuildToolchain` lane runs first. After execution the runner must
produce and byte-verify a compatible full receipt, open a new artifact epoch, and
recompute the impact decision and downstream entry fingerprints before readers; do not
require rebuilt PE/PDB bytes to equal the pre-entry receipt. It uses `ExactArtifact`, so
an exact Resume may reuse its promoted post-entry evidence, while an executed writer
resets earlier reuse decisions conservatively. Evidence writers preserve terminal
artifacts before cleanup and atomically replace only `run-state.json`.

The post-writer build must run under the same suite-owned environment as the initial
build and restore the caller environment in `finally`. Full keeps
`RSBuildEnableTests=true` across this barrier, including `-SkipBuild` entry, so receipt
validity cannot mask artifacts that lack a declared test surface.

Resolve and schema-validate a Resume origin before current fingerprinting, then seed
current planning with the origin terminal artifact epoch. Otherwise promotions produced
after a writer's epoch transition can never match. Epoch seeding does not bypass receipt,
snapshot, plan, dependency, capability, coverage, or outcome identity checks.

When a runner invokes `build.ps1` or Pester in-process, treat that call as a PowerShell
module-scope boundary. Re-import the runner's validation/build-evidence modules before
the next snapshot, receipt, or evidence operation. Dependency modules must not use
`Import-Module -Force` merely for convenience when the caller owns the imported command
surface; forced replacement can remove the caller's commands during nested scope teardown.

Use:

```powershell
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Resume -ResumeFrom <run-id>
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Affected -ImpactBase <commit-ish>
.\Tools\Run-AllTests.ps1 -Suite Commands -CommandsFamily file-operations
```

Resume accepts exact-artifact promotions only. Affected and Commands-family runs report
repository `NOT_EVALUATED`. Commands families remain non-cacheable until a separately
fingerprintable test payload/runtime closure exists.

## Preserve compatibility deliberately

Before moving or renaming a command, search every tracked consumer and consider
untracked operator use when the path has been documented or shipped.

- Migrate in-repository callers atomically when safe.
- Retain a thin forwarding shim for stable or plausibly external paths.
- Give the shim complete help, name its canonical replacement, forward exit
  behavior and parameters, and cover it with a focused test.
- Record whether support is permanent or the release condition for removal.
- Never delete a protected historical path or edit frozen evidence bytes to make
  a cleanup convenient.

## Document public commands

Every public command must use advanced-function parameter binding where
appropriate and include comment-based help with:

- `.SYNOPSIS` and a behavior-focused `.DESCRIPTION`;
- every public `.PARAMETER`;
- `.OUTPUTS` and exit-code behavior;
- `.NOTES` covering prerequisites, side effects, and primary consumers;
- at least one safe `.EXAMPLE`.

Mutating commands must support `ShouldProcess` when preview semantics are
meaningful. Defaults must not silently perform a broader mutation than the
documented example.

## Keep dependency closure exact

When a workflow cache, release identity, or frozen manifest represents a tool:

1. enumerate every imported module, patch, schema, notice, and version input;
2. make the machine key or identity consume that complete set;
3. add a test that fails when implementation imports and declared inputs drift;
4. invalidate obsolete inputs instead of retaining them as misleading key
   material.

Historical manifests stay explicit and immutable. Do not derive frozen inputs
from a mutable current manifest.

For the Ghostty private runtime, compare the cache declaration with the
builder's actual direct `Import-Module` statements. The key must cover the
builder, every direct module dependency, the current private-runtime patch, and
the notices file. Maintain one current overlay: the Windows-wide shared-only
hunk in `Ghostty-WindowsPrivateRuntime.patch` supersedes the former ARM64-only
patch. An upgrade rebases or rederives that single overlay and proves any hunk
removal with clean x64 and ARM64 builds; do not restore or stack a superseded
patch as a second cache/build input.

## Test and reconcile

Add or update tests for inventory coverage, help, caller compatibility, module
exports, output/exit behavior, and cache or workflow policy as applicable.
Update the owning normative spec and user/developer documentation in the same
change. If the work closes an active plan, move it to `Specs/Plans/Done/` only
after those durable contracts are current.

Run at minimum:

```powershell
.\Tools\Get-ToolInventory.ps1 -Validate
Invoke-Pester .\Tools\Tests\ToolInventory.Tests.ps1 -PassThru
Invoke-Pester .\Tools\Tests\ToolingGovernance.Tests.ps1 -PassThru
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
```

Run the focused tests for every touched command/module and
`.\Tools\Run-AllTests.ps1 -Suite Full` before closeout.

## Reference patterns

Implementation patterns for repository PowerShell tooling:

- [references/windows-tool-discovery.md](references/windows-tool-discovery.md) — priority-chain
  auto-discovery of developer tools (explicit parameter → repo-local → env var → PATH → vswhere →
  Chocolatey/Scoop → common roots), as implemented by `vcpkg-install.ps1`.
- [references/powershell-multi-target-install.md](references/powershell-multi-target-install.md) —
  layered platform/variant parameters with `$PSBoundParameters` explicit-detection and a resolved
  target-array install loop.
