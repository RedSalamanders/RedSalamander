# Final cancellation fixes: x64 Release validation, 2026-10-19

The test-enabled Release solution build passed with zero warnings/errors. The twelve selected
native invocations finished with 93 passed, zero failed and 1 explicit configuration
skip: 67 passing scenarios plus 26 setup/cleanup checks. Three affected
ViewerSpace cases and 421 tooling checks also passed. These are selected x64 Release
results; they do not establish Fresh Full, ARM64, ASan or unrelated plugin/device qualification.

## Source, artifacts and interrupted continuation

Delivery parent: `880bad2b6fdfc37162aed8455fcd96a4a7537bc2`, comparison base
`265147ee62c296b1a47ed44efb4d8edb58ad60cf`. The independently running primary-checkout Release
app prevented output replacement, so compilation/testing used the existing managed checkout
`C:\Users\eric\.codex\worktrees\selection-shape-release\RedSalamander`.
Its Git HEAD remained `a48e7d7880cb37ea6898d462ae188a6213a5f3f8` with deliberately mirrored
dirty/untracked source bytes. No branch reset or independently launched process termination occurred.

Release receipt: `0e2153e829e31fdc13f72eb83d3e6be914b4681bd51f4eecd4fcc61424c87008`.
Source snapshot: `3725a0492d81ea4ec0796bf3741923e7858ead8f5aa0b25cfe54421c097792ea`.
[candidate-source.json](./candidate-source.json) records the exact isolated snapshot;
[tested-source-files.json](./tested-source-files.json) binds 81 affected files to this receipt.
[delivery-source-verification.json](./delivery-source-verification.json) additionally records
7,540 compared primary source/spec/tool files, zero byte mismatches, and the full raw manifest digest.
The complete original manifest is retained losslessly in [source-tree](./source-tree/manifest.json),
with its decompression hash verified before curation.
This is source-byte equivalence across checkouts, not a claim that their Git snapshots match.
[debug-source-comparison.json](./debug-source-comparison.json) verifies the final Debug code/test
bytes are unchanged; archive policy and the Release qualification specification changed before build.
The earlier Release packet predates the final stage-debt and observer-lease fixes and is not reused
as proof of those fixes. Their original failures and Debug repair evidence remain in the
[final Debug packet](../2026-10-18_final-cancellation-debug/README.md).

Four selections (58 native checks) completed before a host restart interrupted the first bridge
selection. Its last trace is 2026-10-19 05:32:35.674 UTC; System event 1074 records the restart
request at 05:32:35.845 UTC. Later OS-upgrade restart events and the partial original trace/metric
digest are retained under [host-restart-interruption](./host-restart-interruption/classification.json).
The interrupted invocation has no final result and is not counted as passed or failed. Its full
raw artifacts were preserved before cleanup. The profile lock had been cleared cleanly. The saved
source snapshot, receipt and every attested artifact hash were verified before running the eight
unfinished selections. The original four successful selections retain their original identities.
This selected continuation is not a formal Fresh Full or repository-wide Resume verdict.

The intermediate delivery `a7a78f31b301bdc5534ed7063866f1f99a94a234` recorded ten qualified
selections while the registered NAS was unavailable. Its original three-repeat Move attempt
retains three prerequisite skips (Windows error 53); those skips are separate from the final
qualified totals below. The user brought the NAS online, and its actual registered directory
and SMB connection were verified before the two missing selections ran. See
[network recovery](./network-recovery.json). The available-scope verification/metrics and earlier
network diagnostics remain historical snapshots; the final twelve-selection summary,
independent metrics and pre-staging barrier supersede their incomplete status. The original
unavailable attempt and interrupted invocation are preserved and never converted to passes.

## Verified final-fix behavior

- Six real registered-SMB rows covered folder and selected-file native Copy, verified bridge Copy
  and managed Move. Each reached four actual file transactions with nine files/32 MiB, exact
  destination content/source outcomes, totals closed before payload and zero growth after closure.
  The registered target was `\\rubygloom\RedSalamander-SelfTest\selftest`; no local-share substitute.
- The local matrix retained expected peaks 4,4,4,4,1,1. Selection shape does not override provider,
  session or storage constraints. The alternate-volume control used the marked D: sandbox.
- All five real-share parallel Move cleanup controls passed, plus the nested thirteen-file,
  six-directory Move with four active transfers. NAS unsupported-metadata observations remain
  explicit in the native trace; this packet does not claim preservation the provider cannot verify.
- Fresh metadata for 4,097 selected local files converged in 31 ms within its
  unchanged 2s budget before payload. Duplicate reports stayed stable; changed metadata withdrew
  certainty; source integrity and canceled-destination cleanup passed. Pane counts were not reused.
- Fifteen Move phase-cancellation trials passed across three repetitions, including 0
  naturally uncertain publication results and 0 retained stage names. The test requires
  typed Unknown publication/source/stage debt and guarded Possible artifacts whenever needed.
  Clean timing repeats do not prove that interrupted network mutation can never be uncertain.
- Six direct pre-rename controls (file/directory, repeated three times) proved known non-commit
  and bounded exact compensation. Six actually committed rename/reply-failure controls correctly
  remained Unknown with final objects preserved and no automatic replay. The primary cancellation
  watch and its existing grace were retained.
- Three deterministic stage-debt controls each covered twelve debt/sibling lifecycle combinations,
  two stronger-debt transitions and a same-binary counterfactual reproducing the previous Published
  result. Healthy siblings cannot overwrite another transaction's unresolved root debt.
- Three bounded observer-lease captures passed the actual active-finalization, known-size and
  full-fill assertions after at least 650ms. [Capture provenance](./captures/provenance.json) binds
  unchanged native no-activation screenshots to this receipt; all three were visually inspected.
- Four completed-card controls passed, including the height-retaining counterfactual, normal
  completion, genuine indeterminate status and reduced motion. The first default terminal frame
  was compact and idle timers stopped, with status labels preserved.
- The same-binary 128 MiB pipeline comparison measured 1.328s disabled / 0.86s
  enabled. Seeding stayed outside both copy clocks. The unchanged candidate <=80% gate passed;
  this artificial chunk-latency workload does not establish an SMB throughput speedup.

The exact raw-derived checks and observations are in
[independent-metric-checks.json](./independent-metric-checks.json). The
[pre-staging barrier](./pre-staging-verification.json) verifies source stability, every attested
artifact, each selection, ViewerSpace results and the tooling receipt after execution.

## Scope and configuration gate

| Selection | Repeats | Passed | Failed | Skipped | Receipt-bound run |
| --- | ---: | ---: | ---: | ---: | --- |
| FileOpsFamily_DiscoveryScope | 1 | 27 | 0 | 0 | [20261019T052413Z-45640-b1451ed522234ebcba445cace9a75f98](./20261019T052413Z-45640-b1451ed522234ebcba445cace9a75f98/qualification.json) |
| Commands: eleven selected cases | 1 | 11 | 0 | 0 | [20261019T052741Z-45640-9616abde18334543b41791a3d67fc156](./20261019T052741Z-45640-9616abde18334543b41791a3d67fc156/qualification.json) |
| FileOpsFamily_DecisionMirror | 1 | 9 | 0 | 0 | [20261019T052857Z-45640-1187fd1ed1d846ff82f9f5f7ba34c4a5](./20261019T052857Z-45640-1187fd1ed1d846ff82f9f5f7ba34c4a5/qualification.json) |
| FileOpsFamily_Phase05_Discovery | 1 | 11 | 0 | 0 | [20261019T053030Z-45640-9b628de0b2bb46cd837c969d57480290](./20261019T053030Z-45640-9b628de0b2bb46cd837c969d57480290/qualification.json) |
| FileOpsFamily_Phase11_BridgeAndConnections | 1 | 9 | 0 | 0 | [20261019T060751Z-58644-c339c2ddeea145d2ab364ec0accbcc6f](./20261019T060751Z-58644-c339c2ddeea145d2ab364ec0accbcc6f/qualification.json) |
| Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker | 1 | 2 | 0 | 1 | [20261019T061209Z-58644-7ab15f6e7ad246718a45f1ecbd22f391](./20261019T061209Z-58644-7ab15f6e7ad246718a45f1ecbd22f391/qualification.json) |
| FileOps_SmbNativeRouteConcurrencyBudget | 1 | 3 | 0 | 0 | [20261019T061314Z-58644-4eed764f8b794a8baf7e2293a7d295f6](./20261019T061314Z-58644-4eed764f8b794a8baf7e2293a7d295f6/qualification.json) |
| FileOps_SameProviderBridgeDestinationStorage | 1 | 3 | 0 | 0 | [20261019T061415Z-58644-e2c13c869b704f8595d3111a886af299](./20261019T061415Z-58644-e2c13c869b704f8595d3111a886af299/qualification.json) |
| FileOps_SelectionShapeLocal | 1 | 3 | 0 | 0 | [20261019T061943Z-55020-6e16558735d247b49dc5137d5335fa00](./20261019T061943Z-55020-6e16558735d247b49dc5137d5335fa00/qualification.json) |
| FileOps_ReviewRemediationContracts | 1 | 3 | 0 | 0 | [20261019T062116Z-55020-616dea4e418748329a8c25b278ccb6f5](./20261019T062116Z-55020-616dea4e418748329a8c25b278ccb6f5/qualification.json) |
| FileOps_SmbManagedMoveLeavesNoStage | 3 | 9 | 0 | 0 | [20261019T064015Z-65156-27b42f28ade9427fa85636a0d4c2b992](./20261019T064015Z-65156-27b42f28ade9427fa85636a0d4c2b992/qualification.json) |
| FileOps_SelectionShapeSmb | 1 | 3 | 0 | 0 | [20261019T064209Z-65156-102b36daba24484fb2bfd703da20614b](./20261019T064209Z-65156-102b36daba24484fb2bfd703da20614b/qualification.json) |

`Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker` explicitly skips in Release because
the underlying plugin export is Debug-only. Its native result/trace preserve the reason; it is
not a passing Release scenario. The prior final Debug packet owns that export's executed proof.
All other native cases executed without skips. The same three affected ViewerSpace contracts
passed in native no-activation mode. The nine selected Pester files passed with only the declared
RequiresBuildToolchain tag excluded. Both affected/PrePush explanations are retained as advisory
plans with repository NOT_EVALUATED, not executed wider gates or pending CI passes.

Formatting passed for all 36 affected whole C++ files (clang-format 22.1.3); spec inventory had
zero blocking findings; the tool inventory was complete. The sandbox audit retained its 41
preexisting unowned historical directories; this packet does not claim clean-sandbox acceptance.

## Reproduction and archive policy

Build with `RSBuildEnableTests=true`, `build.ps1 -Configuration Release -Platform x64 -MaxCpuCount 4`.
Use the exact selections/reasons/repeats in [candidate-scope.json](./candidate-scope.json), through
`Tools/Run-AllTests.ps1 -SkipBuild -Configuration Release -Platform x64 -TestRoot Z:/RedSalamander.Perf
-AllowAlternateVolumeTestRoot`; never join FileOps family names with commas. Commands uses the
eleven exact names recorded in its filter. All generated data stayed beneath the initialized
marked Z:/D: sandbox roots or the registered unique remote fixture children.

Full raw evidence and executable private controller/check scripts remain at
`Z:\RedSalamander.Perf\evidence\final-cancellation-release-2026-10-19`.
Each curated metric file contains complete original matching lines in original order, bound to
the raw path/SHA/bytes and selected SHA/count/regex. Full raw streams remain outside Git. Receipt
bytes are stored once at root; per-run SHA-bound relative references preserve provenance without
duplicating the receipt. The exact `.gitattributes` archive pattern disables text normalization.
The explicit archive and required inventory gates must pass within 2 MiB per file/5 MiB total.

No production code repair was needed for this Release qualification. This delivery updates the
authoritative testing qualification contract, archive policy and this new evidence packet.
Hosted CI for the delivery parent had failed before execution due billing/spending limits;
the final push's CI status is checked separately and is never inferred from local results.