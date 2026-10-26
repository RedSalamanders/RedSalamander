# C6 shared runtime cleanup ownership

Starting commit: `96e037730254fe69c198ee82055bff226888f73b`.
This is an independent build-reliability repair, not C6 or whole-plan closeout.

## Defect and repair

Several first-party projects recorded the same deployed DLL in their private
`*.FileListAbsolute.txt` caches. A later consumer `CoreClean` could delete a
dependency another project had already staged, or emit MSB3061 if it was in use.
`IncrementalClean` could likewise delete it as an orphan. Copying the dependency
after one project's Build did not protect it from another project's later Clean.

The root targets now remove shared DLLs from prior/current cleanup ownership,
including migration before explicit Clean. Classification runs after common
MSBuild gathers writes, covering late vcpkg preservation hooks. The exact current
project target and private intermediates remain cleanable. Runtime copying,
manifest-owned obsolete-file removal, normal parallelism and fail-closed receipt
validation remain enabled. `Specs/Build/Build_Toolchain.md` owns the durable rule.

## Evidence

| Run | Result | Interpretation |
|---|---|---|
| `baseline` | 0/8 | Initial fixture setup errors, not a valid product baseline: missing OutputPath and Pester 3's regex-oriented Contain assertion. Retained unchanged. |
| `baseline-v2` | 1/8 | Valid pre-repair baseline: seven failures expose deletion, locked-DLL MSB3061, stale incremental ownership and the test-disabled transition. Non-first-party control passes. |
| `candidate` | 1/8 | Initial target-condition syntax error; retained, not counted as qualification. |
| `candidate-v2` | 8/8 | All original regression cases pass. |
| `policy` | 73/73, no skips | Expanded nine cleanup cases plus 64 build, inventory, governance and impact checks. Includes a first-build/no-history control. |

The fixtures import installed common MSBuild targets and the actual repository
targets. They model late link/link-skipped FileWrites registration; they do not
compile product binaries or contact remote servers. Locked-file and deliberate
stage-before-late-Clean ordering make the failure deterministic without a timing race.
Fixture files and logs remain under the five named `C:/RedSalamander.Perf/runs/`
roots. `run-provenance.json` records exact XML copies and hashes of those raw inputs.
The failed fixture-development attempts are not performance comparisons.

The full test-enabled x64 [Release](../2026-09-06_c6_shared_runtime_clean_release/README.md)
and [Debug](../2026-09-06_c6_shared_runtime_clean_debug/README.md) rebuilds both
pass with zero warnings/errors and matching source snapshot
`3754f54aea44516549094354eff5053a2588e68eb2fc6d939f6fec34518b7335`.
Durations: Release 257,128 ms; Debug 189,655 ms. The preceding Move-mode builds
were 249,285 / 184,614 ms respectively; these single-run timing differences are
not evidence of a throughput change. The earlier Release MSB3061 witness remains
in the Move-mode archive; no warnings were suppressed.

Both rebuilt profiles pass the three-case governed loopback FileOps CRUD smoke.
Their 28 cleanup histories per profile contain no entries for monitored shared
`z.dll`, `yyjson.dll`, `sqlite3.dll`, `blake3.dll` or `libcrypto-3-x64.dll`.
The complete captured history rows are retained alongside full build receipts.
`tested-sources.json` binds seven exact pre-build file copies; the build receipts
bind the complete dirty/untracked workspace snapshot. Subsequent plan/archive
documentation does not claim another runtime qualification.

An ordinary [incremental Release build](../2026-09-06_c6_shared_runtime_clean_incremental_release/README.md)
also passes with zero diagnostics (160,281 ms), the same source snapshot and
receipt `a6fece6675db82bf539822d2caf671a416ad4918b27e0c406522c637dff2576d`.
Its final governed CRUD smoke passes 3/0/0 and its 28 captured histories still
exclude the monitored shared DLLs. Several targets recompiled/relinked, so this
is not a no-op-build timing comparison or proof that every vendor preserve hook
ran. The deterministic late-registration fixture supplies that ordering guard.

## Remaining limits

No Fresh Full, production/test-disabled Release solution, ARM64, or live remote
mutation qualification is claimed here. The native smokes report 183/184/185 disk-audit
issues in its preserved runner results; none was removed or hidden. C0–C5 and
the remaining C6 gates, including warning-accounting and inherited I17, remain open.
