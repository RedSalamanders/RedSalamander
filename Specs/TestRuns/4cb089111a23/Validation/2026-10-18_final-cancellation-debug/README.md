# Final cancellation repair: x64 Debug validation

The final test-enabled solution build and affected selection passed on machine 4cb089111a23.
Native results: 94 passed, zero failed/skipped across twelve selections: 68 scenario executions
and 26 setup/cleanup checks. Three affected ViewerSpace cases and 421 tooling checks also passed.
This is selected x64 Debug qualification, not Fresh Full, ARM64, ASan or a new Release qualification.

Final receipt: 0e8bae1ce209243a8e4c4bd75d7598e319ea1ccb8b8c0d9d914ac1cc2f179177.
Source snapshot: 295a8527742890ef7111979c26f25b87b1686b9f9d1e100591abf1ee06b46f35.
The build used parent commit a4b506d53b92d716a923199ba436463f95a6e375 plus the final nine changed
code/test/spec/archive-policy files. The exact 81-file source manifest and byte hashes are in
[tested-source-files.json](./tested-source-files.json); [pre-staging-verification.json](./pre-staging-verification.json)
binds source stability, native selections, viewer cases, tooling and all attested artifact bytes.

## Repairs verified

Parallel bridge transactions previously used last-worker-wins stage status: an unknown directory
publication could be overwritten by a healthy sibling's Published/Removed state. The root now
accumulates unresolved stage debt atomically. Unknown dominates retained-incomplete and retained debt;
sibling progress cannot resolve another object's debt. A fresh bridge attempt owns a fresh aggregate.
Three deterministic controls each cover twelve debt/sibling lifecycle combinations, two stronger-debt
transitions and a same-binary counterfactual that reproduces the old Published result.

The legacy FileOps_SmbManagedMoveLeavesNoStage case now checks exact truth. Known cancellation before
rename requires clean exact-stage compensation. An already attempted, unreconciled rename must
preserve Unknown publication, retained-source truth and completed-card stage debt; retained names
must remain inspectable Possible artifacts with guarded authority. A failed listing or unexplained
stage still fails. Actual surviving source/destination contents remain checked in every trial.
The existing committed-failure injection now covers the legacy FileRenameInfo fallback as well as
FileRenameInfoEx. Direct real-share file/directory controls exercise both pre-rename cancellation
and an actually committed rename whose reply cannot establish its outcome.

DiscoveryScope_FolderLiveProgress had a timing defect: the original overwrite frame sampled the
first file after its finalization ended, then captured a later file. Keep the 650 ms active-phase,
full-meter and known-size assertions. A test-only nonempty-commit observer lease holds the real phase
until capture completion; it honors operation cancellation and has a 30 s bound. Empty commits
remain unheld so the serial route can reach the sampled payload. The 1,500 ms slow-commit fixture
remains. Original and final actual harness captures, unmodified, are in [captures](./captures/provenance.json).
These are application captures through SaveWindowScreenshot with no activation, not proposal images.

## Final scenario evidence

- Six real registered-SMB rows (folder/selected files; native Copy, verified bridge Copy, managed Move)
  each reached four actual file transactions, with nine files/32 MiB, exact contents/source outcomes,
  totals closed before payload and zero growth after closure.
- The Local matrix retained expected transaction peaks 4,4,4,4,1,1 for its route/budget controls.
  Selection shape does not override provider, session or native-operation constraints.
- All five parallel Move cleanup controls passed on the actual SMB target, plus a thirteen-file,
  six-directory nested Move with four active transfers. The NAS unsupported-metadata note remains
  explicit; these results do not claim timestamp preservation where exact metadata rebind is unsupported.
- Fifteen phase-cancellation trials passed across three repetitions. Final trials retained zero stage
  names. Six forced committed-rename controls still reported Unknown correctly. These clean timing
  repeats do not prove that real interrupted network mutation can never be uncertain; the deterministic
  debt counterfactual supplies the regression witness.
- Six direct pre-rename checks (file/directory, repeated three times) proved known non-commit and
  exact bounded compensation. The primary cancellation watch and 500 ms grace remain active.
- Fresh metadata for 4,097 selected local files converged in 47 ms against a 2 s budget before payload.
  Duplicate reports stayed stable; changed metadata withdrew certainty; source integrity and clean
  canceled destinations passed. No pane counts/sizes were reused.
- Three finalization captures retained actual finalizing phase and known-size fill after 766/735/703 ms.
- Four completed-card controls passed, including the height-retaining counterfactual, default
  completion, genuine indeterminate status and reduced motion, with compact geometry on the first
  default terminal frame and no lingering idle timers.
- The same-binary 128 MiB bridge pipeline pair measured 1.406 s disabled / 0.953 s enabled.
  Its 187 ms seed was outside both copy timers; the unchanged candidate-at-most-80-percent gate passed.
  UI capture work and different profile/source receipts are not throughput comparison baselines.

The actual endpoint was the registered dedicated share
\\rubygloom\RedSalamander-SelfTest\selftest, using unique test-owned children.
Local scratch, logs and complete raw archives remain under marked Z:\RedSalamander.Perf; the
alternate-volume storage case used the existing marked D:\RedSalamander.Perf with explicit runner
admission. No share, mount or outside-root scratch was created by this validation.

| Runtime selection | Repetitions | Native passed | Governed run |
| --- | ---: | ---: | --- |
| FileOpsFamily_DiscoveryScope | 1 | 27 | [20261018T220758Z-75160-3ffec2aa623846168db33d5a158cf530](./20261018T220758Z-75160-3ffec2aa623846168db33d5a158cf530/qualification.json) |
| Commands: 11 selected cases | 1 | 11 | [20261018T221121Z-75160-8aad6869c06a4956829e29a166cd753f](./20261018T221121Z-75160-8aad6869c06a4956829e29a166cd753f/qualification.json) |
| FileOpsFamily_DecisionMirror | 1 | 9 | [20261018T221233Z-75160-4d50a4656321420eb800a15b2487fe36](./20261018T221233Z-75160-4d50a4656321420eb800a15b2487fe36/qualification.json) |
| FileOpsFamily_Phase05_Discovery | 1 | 11 | [20261018T221405Z-75160-c295d408fb8d460ab78f023542c1ca0d](./20261018T221405Z-75160-c295d408fb8d460ab78f023542c1ca0d/qualification.json) |
| FileOpsFamily_Phase11_BridgeAndConnections | 1 | 9 | [20261018T221514Z-75160-d481546fbff942be9258030415a22242](./20261018T221514Z-75160-d481546fbff942be9258030415a22242/qualification.json) |
| Riptide_SharedFileOpsSchedulerShutdownWaitsForBlockedWorker | 1 | 3 | [20261018T221924Z-75160-439e363e63f444819c2d0910b0281b53](./20261018T221924Z-75160-439e363e63f444819c2d0910b0281b53/qualification.json) |
| FileOps_SmbNativeRouteConcurrencyBudget | 1 | 3 | [20261018T222024Z-75160-00c1587ef5464a8080eed42203e44049](./20261018T222024Z-75160-00c1587ef5464a8080eed42203e44049/qualification.json) |
| FileOps_SameProviderBridgeDestinationStorage | 1 | 3 | [20261018T222118Z-75160-8c573eeb968d49318226620938fbfe09](./20261018T222118Z-75160-8c573eeb968d49318226620938fbfe09/qualification.json) |
| FileOps_SmbManagedMoveLeavesNoStage | 3 | 9 | [20261018T222221Z-75160-de1875e54d544765a0038f6a6bb18bc5](./20261018T222221Z-75160-de1875e54d544765a0038f6a6bb18bc5/qualification.json) |
| FileOps_SelectionShapeLocal | 1 | 3 | [20261018T222354Z-75160-d27ab81efd9941ffaedc8d25cf78986a](./20261018T222354Z-75160-d27ab81efd9941ffaedc8d25cf78986a/qualification.json) |
| FileOps_SelectionShapeSmb | 1 | 3 | [20261018T222538Z-75160-b75fad05cfff44ffa1eee693b5d0382a](./20261018T222538Z-75160-b75fad05cfff44ffa1eee693b5d0382a/qualification.json) |
| FileOps_ReviewRemediationContracts | 1 | 3 | [20261018T222742Z-75160-7b5c1f0d625d41c0b8be4c3627ddfaf7](./20261018T222742Z-75160-7b5c1f0d625d41c0b8be4c3627ddfaf7/qualification.json) |

The exact command list and selection rationale are in [candidate-scope.json](./candidate-scope.json).
Native JSON results, traces, aggregate runner results and selected original metric lines accompany
each run. Metric-selection manifests bind the complete retained raw JSONL by SHA-256 and disclose the
selection expression; no result predicates were rewritten. Readable independent checks are in
[independent-metric-checks.json](./independent-metric-checks.json).

## Preserved failures and qualification boundary

| Epoch | What it establishes |
| --- | --- |
| Original Debug, receipt 5a5f98f940e08ad7b0c6cef5e374b9f64332727ac6e7abacddbe6eb838ddcc15 | Three actual finalizing-1 failures: Win32 995 interrupted an in-flight directory rename on an identity-less share. No-stage-only acceptance contradicted conservative mutation truth. |
| Fallback control, receipt 9824a9476b20b0e07c2598ff8c0cdeb002e5dce1d1a0ec5e15aa240808b16f36 | Three new controls failed because successful legacy fallback bypassed failure injection. |
| Stage debt, receipt abe7c168d6a55ecf0c29a60747ba6f99138eb2d4fc7ee6e17bdc03793e0f2b97 | Three finalizing-3 regressions: publication Unknown but stage Published after a healthy sibling overwrote it. |
| Timing-dependent candidate, receipt 55ed5991c792071f0abd3bc887ccad5563f4e51d6348b8b8daa53297632ae9f3 | SMB repeats passed with no natural uncertain rename. Superseded to add deterministic debt coverage. Subsequent native discovery passed but source-stability validation rejected the deliberate source change; neither partial epoch qualifies final source. |
| Display timing, receipt 1203b6d40771b14fd35d0d1ac74d41d48ce290b1e50a5058e275731ed25cd25e | SMB repeat passed; discovery had 25 passed, one live-progress failure and one not-reached skip. Preserved failed overwrite capture shows the temporal mismatch. |
| Final, receipt 0e8bae1ce209243a8e4c4bd75d7598e319ea1ccb8b8c0d9d914ac1cc2f179177 | All twelve runtime selections, three viewer cases and 421 tooling checks passed on the same source/artifact profile. |

The tooling filename correction was an invocation error before Pester launched; its initial workspace
and reason remain under tooling-invocation-baseline. The final tooling selection uses the actual
ReleaseWorkflowPolicy.Tests.ps1 independently of the Debug build profile and excludes only the
RequiresBuildToolchain tag, matching the reviewed affected selection.

Formatting checked all 36 affected native files with clang-format 22.1.3. Spec inventory had zero
blocking findings and the tool inventory remained complete. Final affected/PrePush explanation
artifacts retain the broader advisory obligations and repository NOT_EVALUATED verdict.
These focused results do not replace a Fresh Full gate or establish complete repository/foreground
coverage. CI was billing/spending-limit blocked on the parent; no CI pass is claimed.

The prior [Release packet](../2026-10-18_affected-master-release/README.md) qualified the previous
cancellation checkpoint repair. It predates the additional stage-debt aggregation and capture-fixture
repairs here. New Release artifacts were not run in this Debug follow-up. The test-seam equivalence
proof concerns FileSystem.cpp's moved ENABLE_TESTS block only; it does not claim unchanged production
State.cpp or requalify the old Release artifacts.

Complete raw evidence: Z:\RedSalamander.Perf\evidence\final-cancellation-debug-2026-10-18.
Historical epochs are retained separately and are excluded from final pass totals.
Identical build receipts may be stored once with SHA-256-bound per-directory reference files;
receipt-pool-index.json discloses the exact original bytes and reference mapping.
The pre-staging barrier precedes Git metadata changes. Post-commit verification belongs in the
persistent raw root and must prove the 81 tested source files remained byte-identical and remote
master matches the published commit; it is not a new receipt or an exact Resume on the new Git identity.