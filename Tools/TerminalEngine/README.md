# Terminal tooling profiles

The product runtime path is Ghostty-specific. `Tools/Build-GhosttyTerminalRuntime.ps1`
is the stable thin entrypoint. `GhosttyRuntimeLifecycle.psm1` validates the
selected lock against `Specs/Terminal/GhosttyRuntimeLock.schema.json`, owns
restore/build/inspection/notice/identity publication, and reuses the text,
archive, PE, import, export, contained-process, and canonical-JSON helpers.
Without candidate arguments it selects the canonical tracked Product lock and
may write `.build/TerminalEngine/Product` plus the generated pairing header.
Candidate mode requires both explicit paths below
`.build/TerminalEngineUpgrade`, has no canonical fallback, and cannot write
Product state or a pairing header.

The lifecycle consumes the lock's complete ordered dependency and notice
closures. Each dependency has independently verified source and canonical Zig
content archives plus a normalized license identity; each notice is resolved
from its repository, upstream, toolchain, or dependency source and staged at
its locked path. Both archive forms are cached so `-Offline` candidate builds
fail closed without hiding newly introduced inputs.

## Read-only update observation

`Tools/Get-GhosttyTerminalRuntimeStatus.ps1` is the only active Ghostty status
entrypoint. It validates the same canonical product lock through
`GhosttyRuntimeLifecycle.psm1` and publishes strict canonical evidence matching
`Specs/Terminal/GhosttyRuntimeStatus.schema.json` below `.build` only.

- `-Mode Locked` performs zero network access and validates the checked-in lock.
- `-Mode Online` observes the official default branch and OSV commit advisories.
- `-Mode Candidate -CandidateCommit <full-lowercase-sha>` resolves exactly that
  commit and refuses branch-head substitution.

The daily workflow performs the narrower advisory-only Online query; the monthly
workflow performs full Online observation; manual dispatch supports all three
modes. These jobs have read-only repository permissions and may warn or fail,
but never edit the lock, push a branch, create an issue, or open a pull request.
Release preflight reruns Online observation and binds the report to both the
checked-out source commit and the current lock SHA-256 before artifact builds.

Native ARM64 candidate qualification is dispatched through
`.github/workflows/ghostty-runtime-native-qualification.yml`. It authenticates
short-lived base64 lock/patch secrets against explicit dispatch hashes, runs
the exact isolated candidate build and three upstream lib-vt targets on
`windows-11-vs2026-arm` with fixed Zig seed `0x97d35bc7`, and uploads only the
schema-validated qualification record plus runtime identity. Because Zig 0.16.0
server-mode tests reproducibly exit with code 5 after returning every result on
native Windows ARM64, the workflow uses the pinned default Zig test-runner
source in simple mode, records its SHA-256, and proves the transport-only source
edit was fully removed. The unchanged upstream target graph runs with `-j1`
because a diagnostic native ARM64 run returned Windows code 5 for one concurrent
test executable while the other completed. A failing native
test also emits narrowly filtered `test.exe` Application Error event 1000
diagnostics before the job fails.

## Exact candidate freeze and review

`Tools/New-GhosttyTerminalRuntimeCandidateReview.ps1` freezes one explicit
Candidate observation before any overlay or product-lock change. It requires a
pristine official Ghostty checkout at the requested lowercase full commit, two
independently downloaded official archives with the same SHA-256, and the exact
canonical Candidate status report. It verifies the commit and tree against Git,
requires official-origin provenance and GitHub signature verification, proves
the candidate descends from the selected product pin, and never substitutes a
branch head.

The command evaluates the complete reviewed lib-vt path set (`build.zig`,
`build.zig.zon`, `include/ghostty/vt`, `src/build/GhosttyLibVt.zig`,
`src/lib_vt.zig`, and `src/terminal`). It records every changed path and line
count by build/toolchain, public ABI, parser/state, allocation/memory,
clipboard/security, Kitty graphics, or snapshot/serialization concern. It also
diffs public C declarations, explicit enum values, and every loader-required
export, then records bounded primary-source security queries and reviewed
commit findings without assigning unsupported vulnerability labels.

Both canonical outputs are untracked and confined to
`.build/TerminalEngineUpgrade/<candidate>/`: `candidate-review.json` validates
against `Specs/Terminal/GhosttyRuntimeCandidateReview.schema.json`, while
`candidate-descriptor.json` validates against
`Specs/Terminal/GhosttyRuntimeCandidateDescriptor.schema.json`. The descriptor
binds the product lock, exact observation, authenticated review, toolchain,
dependencies, current overlay, and a derived input digest so later candidate
build and admission phases can reject stale or substituted evidence.

`Tools/New-GhosttyTerminalRuntimeCandidateOverlayReview.ps1` advances that
descriptor only after a candidate overlay is rederived. It requires one clean
candidate checkout, one independently applied candidate checkout, the single
normalized patch, and a reviewer-authored disposition for every concern in the
current product lock. The candidate concern list is ordinal-sorted, unique, and
limited to five. The tool rejects more than 16 upstream files or 2,500 total
changed lines, a patch that does not apply cleanly, a regenerated diff whose
normalized SHA-256 differs, incomplete current-concern coverage, and any newly
added private C or Zig enum option whose name/value mapping collides or
disagrees. Every new private terminal option is recorded, not only the first.

The overlay review preserves the Phase-4 descriptor as
`candidate-descriptor.p4.json`, writes the schema-validated
`candidate-overlay-review.json`, and replaces only the untracked
`candidate-descriptor.json` with the candidate patch path, digest, ordered
concerns, and new input digest. The tracked Product patch and canonical lock do
not change until manual admission. This keeps normal product builds coherent
while isolated candidate builds consume the exact reviewed `.build` inputs.

## Current validation

`Tools/Modules/Testing/TestRunPlan.psm1` points CI and Full at
`CurrentToolingProfile.Tests.ps1`. That profile discovers every ordinary
`Tools/Tests/*.Tests.ps1` file except the sealed files listed in
`HistoricalTerminalProfiles.json`. It therefore picks up new current tooling
tests automatically while running only the narrow historical profile integrity
guard for the archived Gate0/Round4 material.

The Ghostty cache key must cover the wrapper, lifecycle module, its complete
direct and indirect helper closure, the canonical lock and schema,
`Patches/Ghostty-WindowsPrivateRuntime.patch`, the copied notices, and
`RuntimeDependencies.props`. BuildEvidence owns the exact file list, and the
release-workflow policy compares that same list with the workflow key.

For a Ghostty upgrade, rebase or rederive that one current overlay against the
new source/tool tuple, then force-build and attest both x64 and ARM64. The
deleted `Ghostty-WindowsARM64-SharedOnly.patch` is not an independent fallback:
its candidate-only static-companion optimization was deliberately removed from
the current overlay, and the unused companion is never staged or shipped. If
upstream changes make a current hunk unnecessary, prove its removal with both
architecture builds rather than stacking the superseded patch.

## Historical Gate0 and Round4 profiles

Gate0 and Round4 are sealed behavioral profiles, not ordinary product CI. Their
test paths and SHA-256 identities live in `HistoricalTerminalProfiles.json`.
Do not rename, move, edit, or silently add to those files. A new experiment gets
a new named profile and new evidence; it does not rewrite the sealed profile.

Run an archival profile explicitly:

```powershell
./Tools/TerminalEngine/Invoke-HistoricalTerminalTests.ps1 -Profile Gate0
./Tools/TerminalEngine/Invoke-HistoricalTerminalTests.ps1 -Profile Round4
```

The matching GitHub workflows retain their established filenames and frozen
source shape because the Round4 contract pins the Gate0 workflow's exact digest.
Round4 is manual-only. Gate0 is limited to manual dispatch or the sealed
`codex/terminal-engine-gate0` branch with an explicit `[run-gate0]` commit
marker; it is not a pull-request or ordinary product branch gate. Do not edit
either workflow to simplify the duplicate list: that source identity is part of
the historical proof.

## Generic lifecycle compatibility facade

`External/TerminalEngine/TerminalEngine.proj` and the public
`Build-TerminalEngine.ps1`, `Verify-TerminalEngine.ps1`, and
`Get-TerminalDependencyStatus.ps1` entrypoints are retained as an
`internal-historical-future-upgrade` compatibility facade. No tracked default
`External/terminal-engine.lock.json` currently exists, and current product builds
do not use this facade. Keep the paths stable for plausible external/manual
consumers. A future engine-upgrade effort may reactivate them only by supplying
an explicit reviewed lock and updating the manifest, tooling inventory, tests,
and authoritative Terminal/tooling specifications together.
