# Terminal Engine Selection And Dependency Proof

> **Authority:** Execution plan 1 of 6 for `Terminal_EmbeddedConPtyLibGhosttyVt_IntegrationPlan_2026-07-13.md`. The master owns requirements and overrides this file. This plan performs evidence-only spikes and determines whether a terminal engine qualifies; an exact no-winner result is valid, and this plan does not authorize product integration.

## Status

- **State:** Done — Round 4 has no admissible engine. Its exact governance,
  decision, audit, zero-cohort, focused-test, dependency-lifecycle, and
  classified full-suite records are frozen below. Every later terminal product
  plan remains blocked.
- **Baseline:** RedSalamander `1a5d838d0c18298d9e560e243e67b821a40fcf6a`, 2026-07-30
- **Risk:** HIGH
- **Output:** a reviewed Round 4 no-winner decision and exact pre-admission
  governance records. Candidate-neutral evidence and synthetic
  dependency-lifecycle tooling/tests remain reusable for a later admissible
  round. A winner lock, engine artifacts, native winner lanes, runtime/SBOM
  closure, and product integration are not Round 4 outputs.

## Drift and prerequisites

Before work, compare HEAD with the baseline across the master plan, `RuntimeDependencies.props`, `Directory.Build.targets`, `Common/PlugInterfaces/{Viewer,Factory,Host,Informations}.h`, `RedSalamander/ViewerPluginManager.*`, `Common/MinimumOsVersion.*`, `Common/DxUi`, build scripts, installer scripts, testing guidance, and any existing `Plugins/Terminal`/`Specs/Terminal` files. Reconcile material drift and update the baseline before scoring. A material mismatch in plugin/module ownership, pane-tab ownership, per-plugin Settings behavior, centralized runtime staging, minimum-OS ConPTY policy, candidate/build topology, or the snapshot/ABI boundary is a STOP condition. Freeze the master's mandatory gates and weights before measurements.

The 2026-07-30 drift review found no product-boundary change in the listed host, plugin, runtime-staging, minimum-OS, build, or installer inputs. The master-plan expansion and its source-contract test are the only relevant baseline changes. Before freezing the requirements, Gate 0 also corrected one standards error against the authoritative Kitty graphics protocol: its wire format numbers are RGB, RGBA, and PNG, not separate gray/gray-alpha formats. Grayscale content remains supported through a protocol-defined wire format. This correction does not rescue any observed candidate: zlib, crop, deletion, generations, placements, three-layer state, and pre-allocation accounting remain mandatory.

The product boundary is frozen before candidate work: the winner must be linkable into, or privately loadable only by, `Plugins/Terminal/Terminal.dll`; all terminal-specific ConPTY/model/render/profile/history/follow/configuration implementation stays in that DLL. `IViewer` is not an allowed product boundary. Candidate cost/evidence includes DLL-safe initialization, callback/module quiet, x64/ARM64 PE/import/CRT closure, and private `Plugins\TerminalRuntime\` packaging where dynamic.

## Candidates

Evaluate exact current pins for:

1. Ghostty `libghostty`/VT C API.
2. Microsoft TerminalCore/text-buffer/parser source subset that exposes the same consumer-owned immutable snapshot boundary; the full/candidate-owned control is an unscored constraint control.
3. Contour `vtbackend` plus `vtrasterizer`.
4. WezTerm `wezterm-term` as a Rust-bridge contingency.
5. The current maintained libvterm source.
6. The current maintained libtsm source.
7. Microsoft/full candidate-owned control, WebView2/xterm.js, and a RedSalamander parser as unscored constraint controls whose product-constraint/ownership failures are documented.

Canonical discovery roots are [Ghostty](https://github.com/ghostty-org/ghostty), [Microsoft TerminalCore](https://github.com/microsoft/terminal/tree/main/src/cascadia/TerminalCore), [Contour vtbackend](https://github.com/contour-terminal/contour/tree/master/src/vtbackend), [WezTerm term](https://github.com/wezterm/wezterm/tree/main/term), [libvterm](https://www.leonerd.org.uk/code/libvterm/), and [libtsm](https://github.com/kmscon/libtsm). These moving branches/pages are discovery inputs only: before collection, resolve each to one exact commit or source-archive digest, record any redirect/fork/submodule identity, and make the frozen digest-verified pin—not a branch name, downloaded “latest,” README claim, or URL response—the evidence input. The frozen review and candidate-set records bind the exact source location, pin, and archive digest so a repository move cannot silently substitute different bytes; they do not claim a schema field that records retrieval time.

No candidate receives incumbent credit. Ghostty has a conditional pass to enter the spike only.

### Equal-adaptation universe and nomination close

`Specs/Terminal/TerminalEngineCandidateUniverse.v1.json` is the frozen base
seed universe. The current authority is the exact approved
`Specs/Terminal/TerminalEngineCandidateUniverse.v4.json`, which carries that
base through the Round-2/3 corrections and adds the Round-4 chronology and
immutable-consumption closure. Every final review, assessment, candidate-set,
and evidence digest binds the raw v4 file identity, not an effective-object
rehash or the v1 file alone. The universe is the master-plan seed list, not a
post-result shortlist. It grants one nomination to each seed whose exact pin
already contains the required subsystem and forbids authoring a missing
protocol family/parser/terminal model/render-data model merely to rescue a
candidate. Common plugin/UI/ConPTY product integration is outside the engine
nomination. The nomination window closes only when every allowed seed has one
exact admitted or measured-excluded outcome.

All nominations share the conjunctive caps and metric algorithms in that record: no more than 60 owner-days, 2,500 additions-plus-deletions, 24 touched files, five closed logical concerns, one product-new toolchain family, zero non-system runtime leaves, and no binary patch. All source/build/test/documentation/adapter/generated-glue paths count; no generated exclusion or rename discount exists. The exact patch must reproduce the result tree from the exact upstream tree. Two non-author reviewers approve the complete path-to-concern map and owner-day inputs, and the evaluator recomputes every cap. An incomplete lower-bound patch may exclude a nomination only when it already fails a cap under every applicable policy floor and omitted work can only increase the measure; it can never admit or score that nomination.

Round 3 measured and closed every nomination under the identical allowance.
Its downstream Review/Assessment/Set approvals were later found to predate the
approved Round 3 universe, so those records are invalid and quarantined. Round
4 changes no candidate, pin, cap, score, threshold, control, fixture, or
measurement recipe: it carries forward the three exact nomination identities
and re-closes them under strict universe-open, approval-window, common-freeze,
immutable-snapshot, and provider-consumption rules. Contour-RS v3
at exact upstream `0397f9cfbfcc94d46103d3cff70ebf59326dd695`
has a verified 3,041-line/15-path/six-concern patch and 76-day upper-bound owner
ledger, so it exceeds the changed-line, logical-concern, and owner-day caps.
An independent non-schema audit also found that its retained Kitty snapshot
pixels are charged as terminal-text CPU memory instead of image CPU memory.
That observation is non-dispositive and is not represented as an additional
approved failure code; the three schema-bound mechanical cap failures are
sufficient to reject Contour. Ghostty-RS v4 at
`70c498ac3273661aebf6cce9904c0d42b2e5d299` has an exact verified
2,213-line/21-file/five-concern lower-bound patch, but no frozen Gate-0 adapter
export. Its approved upper ledger derives 84 owner-days and every applicable
policy floor still derives 81 before that missing adapter is added. WezTerm-RS v3
at `46a166d6dc8188ede1a32a96375e308081ebafc5` has a verified
2,459-line/seven-file/six-concern patch and derives 87 owner-days. Ghostty fails
the owner-day cap; WezTerm fails the concern and owner-day caps. The three
absent-subsystem sources remain independently source-rejected. Therefore no
nomination enters native scored collection and Gate 0 ends with no current
winner; no successful proof build overrides admission. A candidate-changing
experiment requires a separately approved v5-or-later universe.

## Evidence contract

Create `.build\terminal-engine-spike` for disposable source/build output and archive only content-free review summaries/manifests under `Specs/TestRuns/<MachineHash>/TerminalEngine/<RunId>/`. Allowed archive data is schema/harness/requirements/anchor/fixture/measurement-protocol identities, candidate/pin/license/dependency identities, platform/configuration/native-attestation categories, artifact and reproducer digests/sizes, aggregate measurements, gate/result categories, call ordering, and cross-manifest digests. It must contain no source tree, binary/package, extracted file, local/absolute path, certificate/private key, secret, command, screen/clipboard/image payload, terminal bytes, user data, or user/machine name. Raw sources, artifacts, and minimal reproducers remain disposable under `.build`; the archive links only their non-sensitive identity/digest and result. The manifest schema and required digest fields are defined here before collection and later copied unchanged into `Specs/TestRuns/README.md`. For each candidate capture:

- canonical repository, exact commit/archive digest, maintenance/release state, license, notices, dependency/SBOM closure, toolchain and transitive digests;
- clean Windows x64 and ARM64 results, compiler/runtime/CRT/architecture, offline rebuild behavior, artifact/import/export/header sizes and hashes, package contribution;
- supported consumer boundary and its stability; all crossed type layouts/calling conventions; renderer/PTY/thread ownership;
- exact mutation/snapshot/borrowed-data/callback/teardown sequence from public source plus an executable proof;
- proof that the winner feeds the `builtin/terminal` plugin-owned native HWND/D2D/DWrite renderer through the fixed consumer-owned immutable cell/image/effect snapshot; a candidate-owned control/renderer cannot pass by promising a later plan rewrite;
- proof that a source/static winner links into `Terminal.dll` without a terminal implementation reference from `RedSalamander.exe`, or that a dynamic winner and its complete non-system closure load only from the locked private `Plugins\TerminalRuntime\` paths while the plugin module remains pinned to quiet;
- pre-allocation APC, image/decompression, scrollback, effect, and allocation-failure controls;
- common conformance, malformed-input, fuzz-seam, snapshot, Kitty, and performance results;
- known missing behavior and the amount of maintained glue/patching required for three years.

Freeze criterion-specific integer 0–5 anchors, three-year owner-day rules, crossed-ABI element counting, non-system dependency counting, and added-Release-package-byte measurement before testing. Fractions/interpolation are forbidden. Every score links to evidence. “Likely,” full-app behavior, README claims, and future upstream work are unknown, which disqualifies a scored candidate rather than becoming a discretionary value. Demonstrated absence scores 0. A mandatory failure rejects before scoring.

### Frozen C7 engine-effects score proof

The 2026-07-30 score-proof correction freezes `Specs/Terminal/TerminalEngineC7ProofContract.v1.json` under the closed `Specs/Terminal/TerminalEngineC7Proof.schema.json`. This is a candidate-neutral scoring extension; it does not add, remove, replace, or weaken any row in the mandatory 29-fixture/43-row Gate 0 corpus. `Tools/TerminalEngine/TerminalEngineC7Proof.psm1` is the sole owner of its authenticated-frame, nonce, epoch, sequence, replay, operation-correlation, canonical-JCS, size, and content-free evidence rules. Candidate patches may expose only the typed public engine events, `livePromptSpan`, and copied engine-owned cwd/title/hyperlink effects required by the frozen action; they may not implement the authenticator, accept a driver-supplied result, or synthesize a score object.

C7 rating 5 requires all of the following in each native x64 Release, x64 `ASan Debug`, and native ARM64 Release collection:

- the provider independently loads the exact adapter DLL already bound to that lane's candidate/pin/source/build/layout/header/capability proof, creates a fresh session, feeds the frozen 192-byte `c7-engine-effects-v1` stream under all-at-once, bytewise, and terminator-split schedules, invokes exact action `probe-c7-engine-effects`, and compares the returned canonical state with the frozen expected state;
- a separate fresh-session order run proves that provider-generated authenticated events are admitted only in exact A/B/C/D/A/B engine order, yield the required two available prompt spans and two unavailable states, copy the expected cwd/title/hyperlink effects, report exit code 23, and leave `hostParserCount=0`; the provider rejects any pre-created output or driver-supplied state;
- the neutral verifier passes every one of the 31 frozen authentication/composition/origin cases plus the three explicit boundary cases: an exact 2,097,152-byte canonical frame succeeds, a 2,097,153-byte frame invalidates the epoch, and the next sequence after `UINT64_MAX` rejects and invalidates instead of wrapping;
- the provider attaches exactly one bounded, content-free `score.c7-proof` evidence object with `resultCategory=verified`; its adapter identity and contract/engine/order proof digests are recomputed from the real run, and the proof digest is distinct across all three collection lanes so a replayed descriptor cannot score.

`Tools/TerminalEngineEvaluator.psm1` may match the C7 rating-5 anchor only after resolving and validating that provider evidence in every input manifest. Missing, non-verified, oversized, malformed, injected, replayed, or mismatched proof makes rating 5 unavailable; it is never rounded up or inferred from source review. `Tools/Tests/TerminalEngineC7Proof.Tests.ps1`, `Tools/Tests/TerminalEngineCollectionProvider.Tests.ps1`, and `Tools/Tests/TerminalEngineEvaluator.Tests.ps1` are required focused guards, and `.github/workflows/terminal-engine-gate0.yml` runs the proof suite.

This correction changes a frozen score input. All TerminalEngine collection and
performance evidence produced before the corrected
contract/provider/evaluator identities is stale and non-reusable. When the
frozen candidate set contains at least one `disposition="collect"` scored
candidate, restart Bootstrap/Clean/Offline collection for every admitted scored
candidate and declared control in all three required lanes, rerun the frozen
performance protocol, and finalize only from those new exact manifests. When
the set contains zero executable scored candidates, `Collect` must reject
before lane attestation, subprocess launch, or output mutation; close the round
from the exact universe, review, assessment, candidate-set, rejection proofs,
decision, schema-bound universe/source/adaptation approvals, and two named
non-schema final-decision cold audits, and create no collection or finalization
manifest. No earlier lane, metric, proof descriptor, or partial result may be
carried forward.

### TerminalEngine evidence manifest v1

Before the first collection, check in `Specs/Terminal/TerminalEngineEvidence.schema.json`, its positive/negative schema corpus, and tested `Tools/TerminalEvidence.psm1` as the sole JCS/SHA-256/archive-path writer later reused by Commands and package evidence. It exports exact `ConvertTo-RSJcsUtf8Bytes -InputObject`, `Get-RSJcsSha256 -InputObject`, `Get-RSJcsIdentitySetSha256 -Records -KeyProperty <string[]>`, `Write-RSEvidenceRecord -Manifest -Directory -JsonLeaf -CompanionLeaf`, and `Get-RSTestMachineHash`; every caller uses those functions rather than `ConvertTo-Json` or a private canonicalizer. The identity-set helper rejects missing/null sort keys and duplicate tuples, sorts key tuples with `StringComparer::Ordinal` field-by-field, preserves each record's exact closed shape, then hashes the RFC-8785 array. `Write-RSEvidenceRecord` is transactional/create-new and returns the one manifest path. Every TerminalEngine run directory contains exactly one canonical `terminal-engine-evidence.v1.json`, one `terminal-engine-evidence.v1.sha256`, and optional human summary files containing no additional evidence. The schema uses draft 2020-12, `additionalProperties: false` at every object, explicit bounds/enums, and a top-level `oneOf` for `collection` versus `finalization`.

Serialization and digest rules are fixed: JSON is RFC 8785 JCS, UTF-8 without BOM or trailing newline, with duplicate keys/non-finite numbers rejected. The companion file is exactly the lowercase SHA-256 of the manifest's exact bytes plus LF. Every `*Sha256` is 64 lowercase hex; an artifact/reproducer digest covers its exact raw bytes, while an identity-set digest covers the RFC-8785 canonical bytes of the named sorted identity array. Byte counts and raw timing/count samples are unsigned 64-bit decimal strings so JSON-number precision cannot vary. Candidate/evidence/metric IDs match `[a-z0-9][a-z0-9._-]{0,63}` and are unique within their arrays. Every array described as sorted uses ascending ordinal comparison of those case-sensitive ASCII IDs; tuple arrays compare fields left-to-right in the stated order. Locale/natural/case-folded ordering is forbidden. No signature is implied.

Both record kinds require exact `formatId="red-salamander-terminal-engine-evidence"`, `schemaVersion=1`, `recordKind`, `runId`, RFC-3339 UTC `createdUtc`, 40-lowercase-hex `repositoryCommit`, `repositoryClean=true`, `machineHash`, `schemaSha256`, `harness{version,sourceSha256}`, and `frozen{requirementsSha256,scoreAnchorsSha256,fixtureCorpusSha256,measurementProtocolSha256,candidateSetSha256,candidateReviewSha256}`. Before bootstrap, collection, or finalization, HEAD must equal `repositoryCommit` and `git status --porcelain=v1 --untracked-files=all` must be empty; ignored build/evidence roots are allowed only because Git classifies them ignored. `Tools/TerminalEvidence.psm1` owns one tested `Get-RSTestMachineHash` implementation byte-compatible with `SelfTest::GetComputerHashName`: read the `HKLM\SOFTWARE\Microsoft\Cryptography\MachineGuid` string, otherwise the Win32 computer-name string; encode the selected string as UTF-8 without BOM/terminator, SHA-256 those exact bytes, and use lowercase hex of the first six digest bytes. Empty/unreadable source or conversion/hash failure blocks evidence rather than emitting `unknown`. The 12-hex result must equal the parent directory. `runId` is UTC `yyyy-MM-dd_HHmmss`, must equal its directory, and is create-new: on collision wait for the next UTC second or fail, never overwrite/reuse. `candidateSetSha256` covers the RFC-8785 canonical bytes of the complete closed candidate-set record, including every pin digest, disposition, source-rejection reference, control constraint, and platform bootstrap-input identity. Collection additionally requires `platform` (`x64|ARM64`), `configuration` (`Release|ASan Debug`), `nativeAttestation{processMachine,nativeMachine,windowsBuild,isNative}`, `toolchain{compilerId,version,identitySha256}`, `collectionPolicy{bootstrap=true,clean=true,offline=true,bootstrapInputSetSha256,networkIsolation="runner-enforced"}`, and a candidate array sorted by `candidateId`. `bootstrapInputSetSha256` covers the frozen sorted candidate source/tool/dependency input identities after digest verification and before clean/offline build; the runner must attest outbound network denial for every candidate build/test subprocess after bootstrap. Raw `IsWow64Process2` values serialize `processMachine` as `UNKNOWN|AMD64|ARM64` and `nativeMachine` as `AMD64|ARM64`; `isNative=true` only when `processMachine=UNKNOWN` and `nativeMachine` matches the requested platform (`AMD64` for x64). Any non-`UNKNOWN` process machine is emulation and is rejected for native evidence. After the isolated build, the provider—not the candidate driver—must locate exactly one `CMakeFiles/**/CMakeCXXCompiler.cmake` and one `CMakeCache.txt` below the clean lane root, require MSVC and the requested target architecture, resolve absolute non-reparse `cl.exe`, `lib.exe`, and `link.exe` paths, and hash their exact bytes. `toolchain.identitySha256` is the JCS SHA-256 of the closed object `{compilerId="msvc",version,targetArchitecture,components[]}`, where components are ordered `compiler`, `librarian`, `linker` and each is exactly `{role,fileName,sha256,sizeBytes}`; the driver tuple must match this independently derived value. A duplicate compiler/cache file, relative/wrong-name path, architecture mismatch, or driver disagreement invalidates the lane. Hashing the frozen `toolchain-input` records alone is not toolchain-use evidence.

Each collection candidate is exactly one of:

- `complete`: `{candidateId,pin,pinSha256,outcome="complete",gates,evidenceObjects,measurements}`. `gates` contains G1 through G12 once in order, all `pass`. A gate record is exactly `{gateId,status,resultCategory,evidenceSha256}`: `gateId=G1..G12`; `status=pass|fail|not-run`; pass requires `resultCategory=passed` plus a digest; fail requires `resultCategory=mandatory-failure` plus a digest; not-run requires `resultCategory=failed-G<n>` naming the first failed gate and `evidenceSha256=null`. `evidenceObjects` is a sorted array of `{evidenceId,sha256,sizeBytes,resultCategory}`. `measurements` is a sorted array of `{metricId,unit,warmupCount,samples}`, with `unit` from `microseconds|bytes|count` and `samples` a nonempty array of unsigned decimal strings.
- `early-failure`: `{candidateId,pin,pinSha256,outcome="early-failure",gates,earlyFailure}` with no artifact, measurement, criterion, score, or tie fields. This is only a lane-executed failure: gates before `firstFailedGate` are proven `pass`, that G1–G12 record is `fail`, and every later gate is `not-run`. `earlyFailure{firstFailedGate,failureCode,reproducerSha256,reproducerSizeBytes}` is required; `failureCode` is one exact lower-case ID from the gate-specific allowlist frozen into `requirementsSha256`. The raw reproducer remains outside the archive.
- `source-rejected`: `{candidateId,pin,pinSha256,outcome="source-rejected",sourceRejection}` has no `gates`, measurements, scores, or implied earlier passes. `sourceRejection` is exactly `{gateId,failureCode,candidateReviewSha256,blockerId,blockerSha256,sourceArchiveSha256}` and must byte-resolve to the frozen candidate-review record. Only G6, G7, and G9 are source-review eligible. Each blocker binds the exact candidate/pin/pin digest, archive, relative file and file digest, inclusive line range and raw range digest; file hashing covers exact archive-member bytes, while the range begins after the `(lineStart-1)`th `0x0A` and ends after lineEnd's `0x0A` or at EOF, with no decoding/newline/Unicode normalization. Its direct requirement contradiction must be approved by two independent reviewers who did not author candidate glue before collection. Missing adapter/native/build/measurement evidence, a disputed claim, or an unapproved claim is incomplete evidence, never a rejection. If an executable collection cohort exists, the same source-rejection tuple appears unchanged in all three companion collections and no lane build is claimed for that candidate. If no executable cohort exists, the candidate-set references the rejection directly and no collection record is created.
- `constraint-control`: `{candidateId,pin,pinSha256,outcome="constraint-control",controlFailure}` with no G1–G12, measurements, or score fields. `controlFailure{constraintId,evidenceSha256}` uses exactly `candidate-owned-control|web-runtime-second-stack|new-parser-maintenance`. If an executable collection cohort exists, every declared control appears in all three companion collections with a reproducible failure reference, but is excluded from scoring and selection. Controls alone cannot attest a collection lane and do not authorize a static-only manifest.

Finalization exists only when the frozen candidate set contains at least one executable scored candidate. It then requires `inputManifests`, exactly the native x64 Release, x64 `ASan Debug`, and native ARM64 Release tuples sorted ordinally by `{platform,configuration}` with each exact manifest SHA-256; all shared top-level repository/schema/harness/frozen values and `repositoryClean=true` must be byte-identical. `candidateResults` has exactly one record per scored candidate sorted by `candidateId`: a mandatory-pass record `{candidateId,outcome="scored-complete",eligibility,criteria,total,tieTuple}`, or a failed record `{candidateId,outcome="mandatory-failed",eligibility="mandatory-failed",failureReferences}` with every criterion/score/tie property forbidden. A lane-executed failure reference is `{inputManifestSha256,firstFailedGate,failureCode}`. A source rejection reference is `{inputManifestSha256,gateId,failureCode,candidateReviewSha256,blockerId,blockerSha256,sourceArchiveSha256}`; exactly three identical rejection tuples, differing only by input-manifest digest, are required. References sort by their exact schema keys and resolve to that candidate's collection record. A scored-complete `eligibility` is `eligible|below-threshold`; `criteria` contains C1 through C11 exactly once in numeric criterion order. Each criterion is `{criterionId,weight,rating,matchedAnchorId,contribution,evidenceRefs}`: weight must equal the master table, rating is integer 0–5, `matchedAnchorId` is the exact frozen anchor ID, contribution is a two-decimal string, and the nonempty refs `{inputManifestSha256,evidenceKind,evidenceId}` are sorted by those fields in that order and resolve to that same candidate's collection `gate|object|metric` record. `total` is a two-decimal string; `tieTuple` is unsigned-decimal-string `{ownerDays,crossedAbiElements,nonSystemRuntimeDependencies,addedReleasePackageBytes}`. With zero executable scored candidates, creating either a collection or finalization manifest is invalid because no candidate build, native lane, or derived toolchain identity exists to attest.

`controlResults` contains every declared control exactly once, sorted by `candidateId`, and each exact record is `{candidateId,outcome="constraint-control",constraintId,failureReferences}` with all score/criterion/tie properties forbidden. Its `failureReferences` contains exactly three `{platform,configuration,inputManifestSha256,evidenceSha256}` records—native x64 Release, x64 `ASan Debug`, and native ARM64 Release—sorted ordinally by `{platform,configuration}` and resolving to that control's matching collection `controlFailure` digest; `constraintId` must be identical across all three. No other control-result shape is valid.

The top-level `selection{status,rawLeaderScore,cohort,winnerCandidateId}` is one of four exact shapes: `no-scored-candidate` has null score, empty cohort, and null winner; `below-threshold` has a two-decimal score string, empty cohort, and null winner; `winner` has a two-decimal score string, nonempty cohort, and required winner in that cohort; `tied-first-no-winner` has a two-decimal score string, at least two cohort IDs, and null winner. `cohort` always contains the inclusive cohort's unique candidate IDs in ascending ordinal `candidateId` order; its order never represents score or tie-break rank. Semantic validation resolves every evidence ref and recomputes criterion weights/contributions, totals, eligibility, raw leader, the sorted cohort, lexicographic order, controls, and outcome instead of trusting stored values. Any schema/digest/reference/semantic mismatch is incomplete/invalid evidence, never a candidate result.

## Common spike surface

Add `Tools/Evaluate-TerminalEngines.ps1` with `-Mode Collect|Finalize`, mandatory `-Candidate All`, collection-only `-Platform`/`-Configuration` plus mandatory `-Bootstrap -Clean -Offline`, `-EvidenceRoot`, `-PassThruManifest`, finalization-only `-EvidenceManifest`, and `-RequireAllEvidence`. A partial collection is neither archive-quality nor finalizable and has no implicit diagnostic mode. After resolving the frozen selection, reject source-only, controls-only, and mixed-static subsets before review/assessment, repository-state work, provider Bootstrap/Clean, subprocess launch, or output mutation. There are no implicit defaults for those three collection switches: omission or use in Finalize is a parameter error before mutation. In Collect, `-Bootstrap` is an explicit first phase that materializes only frozen digest-verified candidate/source/tool/dependency inputs. Every future executable round also freezes exactly one candidate-neutral `host-wil-headers` archive named `host-wil-headers.zip` per platform input set, shared by the exact executable cohort. Every consuming phase removes only the exact contained prior expansion, freshly safe-expands the digest-verified archive, requires `include/wil/resource.h`, rejects reparse points from the filesystem root through that member, and binds the complete include-tree identity. The provider verifies that identity immediately before and after the isolated candidate driver; Harness and Probe wrappers independently reverify immediately before and after MSBuild, put the explicit root first in include search, and disable `RSUseVcpkgManifest`, `VcpkgEnabled`, and `VcpkgEnableManifest`, so between-phase mutation, ambient `.build\vcpkg_installed`, user caches, or inherited manifest installs cannot supply evidence headers. Runtime neutrality treats only that exact structurally valid host tuple with ordinal-unique consumers equal to the complete nonempty executable cohort as neutral; toolchain file names retain their exact leaf without deriving a bare extension-stripped literal, while exact toolchain IDs, every source archive, and every other dependency remain candidate-bearing. Round 3 has no executable cohort and therefore creates no synthetic host-header input. Before exposing patched source to a driver, the provider reconstructs an isolated Git index from the exact source tar, restores every executable bit from the tar member modes, rejects links and unsafe members, proves the frozen upstream tree, applies the exact patch to the index, and proves result tree, no-renames numstat/name-status, binary count, canonical regenerated patch, and exact concern-mapping path closure. Trusting only the checked-in mapping or applying a patch to an unverified extracted directory is forbidden. After bootstrap exits, `-Clean` removes/verifies empty only the exact candidate build/output roots (never the verified input cache), and `-Offline` runs all build/test subprocesses under runner-enforced outbound-network denial with no global-tool or ambient-cache fallback. The manifest's collection policy and G3 evidence bind all three results. Both modes require `-EvidenceRoot`; both accept `-PassThruManifest`, which writes exactly the single resolved collection/finalization manifest path and no other success-stream value. `Collect` produces no scores/winner and rejects before attestation or mutation when the selection has no executable scored candidate. If one or more executable candidates are admitted, run those candidates plus every declared companion static record in native x64 Release, native x64 `ASan Debug`, and native ARM64 Release collection lanes before selection. `Finalize` then accepts only the explicitly named three manifests, performs no filesystem/latest-run discovery, archives its own v1 record, and alone computes scores/selection and complete-evidence/no-winner exits. It rejects a dirty/mismatched repository, duplicate candidate/platform/configuration tuples, absent scored-candidate evidence/early-failure records or control failure records, wrong/non-native architecture attestations, any false bootstrap/clean/offline/network-isolation flag, differing x64 Release/ASan bootstrap input-set digests, a platform bootstrap input set that does not equal the frozen lock, digest/order mismatches, and mixed repository, schema, harness, frozen requirements, anchors, fixtures, or measurement protocols. The harness must:

- in Round 4, recognize that the current authority also has no executable
  cohort and creates no synthetic host-header input; the retained Round-3
  sentence describes historical provenance, not the current authority;
- never modify product projects or use a developer-global compiler/tool silently;
- feed identical deterministic split UTF-8/VT, resize/reflow, alternate-screen, scrollback, color/style, grapheme/wide-cell, selection, key/mouse/focus/paste, synchronized-output, title/hyperlink, static-Kitty, malformed and allocation-failure fixtures;
- use the same dimensions, byte counts, clocks, warmup/sample counts, machine identity, and Release measurement method;
- run numbered mandatory gates first; after the first hard failure, preserve the reproducer and mark later work `not run — failed G<n>` instead of requiring a meaningless artifact;
- deterministically probe whether snapshot begin/update consumes dirty state and whether post-begin failure can force a later full snapshot; include mutation/notification barriers around every ownership-bit clear;
- emit machine-readable pass/fail gates, raw measurements, weighted scores, unknowns, artifact hashes, and distinct exits for incomplete evidence versus complete comparison/no winner. For a nonempty executable cohort, `-RequireAllEvidence` succeeds only when every admitted scored candidate has complete evidence or a reproducible early-failure record and every declared companion control has its reproducible constraint-failure record in all three collections. A zero-executable cohort closes only through the reviewed pre-admission record set and never through `Collect` or `Finalize`.

Do not create a compatibility layer broad enough to hide a candidate's weakness. Candidate glue is the smallest code needed to observe its real public/source-vendored contract and is counted in integration cost.

The final decision/lock must name `pluginLinkage=source-static|private-dynamic`, the exact `Terminal.dll` import/CRT consequences, every private dynamic runtime leaf, and the plugin module quiet-point/lifetime proof. A candidate that requires its terminal implementation, global runtime, renderer, or UI control to be owned by the EXE fails mandatory integration fit even if its full application works.

## Ghostty adversarial probes

- Verify the exact C API status and pinned header ABI-padding/layout concerns.
- Build/load clean x64 and ARM64 artifacts from pinned Zig/tool/dependency inputs; inspect PE imports/CRT/SIMD/system libraries.
- Generate and compare public header/export/type-layout/calling-convention fingerprints before any callback/handle.
- Prove begin/end snapshot and every borrowed cell/image lifetime; prove callbacks are bounded and quiesce before unload.
- Install a bounded PNG decoder callback if required. Reserve compressed, decoded, transient BGRA, retained-generation, execution-time, session, and process costs before work; file/temp/shared-memory media stay disabled.
- Prove static Kitty formats, chunking/compression, crop/placement/z-order/deletion/generations. Record animations/glyph/notification/Sixel/iTerm/clipboard gaps truthfully.
- Generate notices from actual linked dependencies, not only the top-level MIT license.
- Fuzz fragmented VT/APC, decompression bombs, allocator failure, resize/reset/cancel/teardown, and eight-session storms.

Any mandatory failure rejects that Ghostty pin for the frozen comparison. No post-result patch, glue exception, or reviewer approval can rescue it. A proposed patch must first become a new exact candidate pin, dependency/source digest set, owner-day/ABI estimate, and frozen candidate-set revision; discard the superseded comparison and restart every candidate/control across all three collection lanes before any new result is considered.

## Decision and lock

Apply the master's frozen 100-point matrix after mandatory rejection. Each
contribution is `weight × rating / 5`; retain full precision and publish totals
to two decimals. Compute `rawLeaderScore`. The inclusive tie cohort contains
every mandatory-pass candidate at or above 80.00 and at least
`rawLeaderScore - 5.00`. One member wins; multiple members are ordered
lexicographically by lower frozen three-year owner-days, fewer crossed ABI
functions/by-value types, fewer non-system runtime dependencies, then fewer
added Release package bytes, even if that displaces the raw leader. If two or
more cohort members remain tied for first after all frozen lexicographic keys,
there is no winner and Gate 0 stops. No candidate-name/reviewer discretion or
post-result anchor change is allowed. For a future executable cohort, two named
reviewers approve the decision record and exact evidence-manifest digests under
the then-frozen contract. Round 4 instead has exactly the schema-bound
universe-proposal, source-blocker, and adaptation-rejection approvals recorded
below; its two named reviewers cold-audit the final Decision proposal digest as
a non-schema closeout review and do not claim that the current governance
schemas approve final Markdown or final Review/Assessment/Set file hashes.

Write `Specs/Terminal/TerminalEngineDecision.md` with the frozen requirements, candidate pins, raw evidence links, scores, uncertainty, rejections, chosen boundary, threat model, exact snapshot algorithm, capability ledger, quotas, dependency/license closure, benchmarks, reviewers, and date.

For the winner add:

- `External/terminal-engine.lock.json` with source/tool/dependency/runtime/header/export/type-layout/capability/license digests per architecture/configuration plus stable per-dependency update-discovery metadata;
- `Specs/Terminal/TerminalEngineLock.schema.json` and `TerminalEngineLockCorpus.v1.json` as the closed lock contract. Every bootstrap input declares exactly one `source`: a durable `repository-file` path or a canonical credential-free HTTPS URI. HTTPS bytes live only at `.build\TerminalEngineInputCache\v1\sha256\<sha256>` and are never addressed by a mutable filename, tag, redirect, query, credential, or ambient package cache. The lock separately binds patched-source upstream/patch/result-tree identities and counts, derived extracted file-set identities, and installed toolchain identities;
- `Tools/Restore-TerminalEngineInputs.ps1` as the only input-cache population path. Online restore performs an exact no-redirect/no-credential HTTPS GET, enforces locked size and SHA-256 before an atomic same-directory publish, and never edits the lock. `-Offline` performs verification only. Build and verification always require the complete verified cache and never download;
- `Tools/Build-TerminalEngine.ps1` and `Tools/Verify-TerminalEngine.ps1` with clean/offline modes and explicit `-LockFile` support, materializing verified consumable outputs at `.build\TerminalEngine\<Platform>\<Configuration>\`; any ASan-to-Debug artifact reuse is explicit in the lock and materialized at the requested configuration path;
- `Tools/Get-TerminalDependencyStatus.ps1` for read-only offline lock/closure verification and explicitly online upstream/advisory status, plus `Tools/Prepare-TerminalEngineUpgrade.ps1` for isolated exact-pin Collect/Finalize comparison under `.build`; neither tool updates tracked files or selects “latest”;
- `Tools/Tests/TerminalDependencyLifecycle.Tests.ps1` using only local fake remotes/feeds and synthetic locks to prove status, candidate isolation/comparison, build/verifier lock routing, and previous-lock rollback;
- generated layout/assertion metadata under `.build/thirdparty`, never handwritten duplicate ABI constants;
- an exact app-local/source-static load/link policy and fail-closed validation order;
- the rewritten master and remaining child plans with the selected names/pin/ledger. No generic or nonselected implementation instructions remain before plan 2 starts.

For a dynamic winner, lock the complete non-system import closure. Open primary/dependencies by canonical path denying write/delete; validate identities/digests/PE/imports and retain handles; reject a preloaded wrong same-basename module; load with DLL-directory/System32 search only; enumerate actual non-system modules and compare the exact locked set before calling the identity export or installing callbacks. Test replacement between validation/load, modified transitive DLL, unexpected import, and preloaded-name substitution.

Before closing Gate 0, clean/offline build and verify the winner in Debug and Release on x64 and ARM64, run smoke/conformance natively on x64 and ARM64 Windows, and pass the x64 `ASan Debug` consumer boundary. ARM64-native runner absence blocks selection. Record whether ARM64 ASan is supported by the selected toolchain and whether x64 ASan reuses the Debug engine artifact. The decision also fixes `Plugins\TerminalRuntime\` package paths for a dynamic runtime/closure and app-root notice paths so Phase 1 can extend the centralized manifest consistently.

## Dependency maintenance workflow

For this zero-winner closeout, implement and test the master's pre-winner scaffold
without fabricating a production lock. A later winner-bearing Gate-0 round must
complete the production discovery/status and review-note gates below before it
can select a winner or add `External/terminal-engine.lock.json`:

1. `Get-TerminalDependencyStatus.ps1 -Mode Locked -EngineLock <path>` validates the closed lock and frozen declared identities offline without requiring disposable cache bytes and without writing any file. Add `-RequireCache` to verify every repository/HTTPS input byte; a missing or drifting cache entry then fails. `Restore-TerminalEngineInputs.ps1 -LockFile <path>` is the explicit online cache bootstrap, while `-Offline` is the mutation-free cache check. The pre-winner `-Mode Online` observes only the closed discovery record's kind, credential-free URI, optional JSON property, frozen value, and advisory bit and reports `current`, `update-observed`, `security-review`, `required-manual`, or `unreachable`; it is advisory and never downloads executable inputs or edits the lock. In this scaffold `-RequireOnline` rejects unreachable endpoints only. `-PassThruReport` writes only one content-free `.build\TerminalDependencyStatus\` report and returns exactly its path.
2. `Prepare-TerminalEngineUpgrade.ps1 -Mode Collect` requires an exact target pin, exact 64-lowercase-hex source tree/archive SHA-256 even for a content-addressed pin, expected current-lock SHA-256, and—after the mandatory production extension—an exact pre-reviewed `-CandidateLockFile` plus `-ExpectedCandidateLockSha256` and exact pre-reviewed `-UpgradeReviewFile` plus `-ExpectedUpgradeReviewSha256`, all below `.build`, plus native platform/configuration, clean/bootstrap/offline policy, `-ExclusiveDisposableRunner`, and an explicit candidate root. The isolation switch is mandatory because offline enforcement is host-wide; never run a collection lane on a shared or non-disposable machine. It creates only `.build\TerminalEngineUpgrade\<CurrentLockSha256>\<TargetPinDigest>\` outputs and reuses the canonical Gate-0/bootstrap/build/verify/evidence paths through the candidate lock.
3. `-Mode Finalize -EvidenceManifest <exact-five-paths>` also requires those same four explicit candidate-lock/review file-and-digest parameters. Transfer the five manifests, exact approved candidate-lock bytes, exact upgrade-review JSON and SHA companion, and their recorded digests to the review machine. Finalize reopens/recomputes both files, validates the companion, proves every native lane bound those identities, and copies the supplied bytes to the canonical candidate root before emitting a comparison of source/submodules, tools, full dependency/SBOM/import closure, licenses/notices, ABI/layout/exports, capabilities, security/conformance/fuzz, performance/memory/package size, changed comparison categories, and rollback compatibility. A digest/path inside a lane manifest is not a substitute for the transferred record bytes. Finalize rejects absent or substituted same-path bytes; it does not generate or repair the lock/review, derive exact repository edits, or make a tracked edit.
4. A human/agent applies an approved upgrade as one repository change containing the lock, adapter, runtime manifest, notices/licenses, decision/ledger, tests, and docs. Two reviewers approve the comparison, candidate-lock, and upgrade-review-record digests. Then the complete clean/offline native, Commands, performance, package, and signed-smoke matrix runs. Mixed old/new pieces, automatic latest selection, or an irreversible terminal-state migration are STOP conditions.
5. Rollback reverts that complete repository change and reconstructs artifacts/packages from the previous lock. An engine-only update changes no persisted user-state schema. If a state migration is necessary, split it into a separately reviewed backward-compatible migration with its own rollback tests.

`TerminalEngineDecision.md` owns the long-lived operator runbook: current dependency owner, periodic/manual check command, emergency advisory path, how to prepare/review/apply, exact verification matrix, release notes/license duties, and rollback. Plan 6 extends the same status view with `TerminalPackageToolchain.lock.json`; it does not introduce a second dependency checker.

The checked-in discovery schema/status implementation is deliberately a
pre-winner scaffold. Before a winner-bearing round may approve a production
lock, extend the closed metadata with the stable tracked channel/ref or release
feed, prerelease policy, separate advisory source, expected repository owner
and canonical final URI, disappeared-pin result, and ambiguity policy. Extend
the local-fake test matrix to prove redirects, fork/owner drift, prereleases,
advisories, unreachable endpoints, missing pins, multiple-result ambiguity, and
unreviewed dependencies; production `-RequireOnline` rejects unreachable or
ambiguous observations. Until those tests pass, Online output is only a
content-free observation and is not sufficient for a production engine lock.

The current v1 schema is likewise only a pre-winner scaffold: its mandatory
`upgrade.basedOnLockSha256` cannot represent the first selected engine. Before
a winner-bearing round, replace it with the master's closed v2 `provenance`
`oneOf`. `first-selection` binds the exact pre-lock universe/review/assessment/
candidate-set/finalization and lock-independent candidate-materialization
identity plus selected candidate ID/pin; `upgrade` alone binds a real
previous-lock digest, target source digest, and changed categories. No
sentinel, placeholder, decision/lock, or Verify-manifest hash cycle is valid.
The materialization record contains reviewed five-lane expected identities but
no final lock digest. Add and corpus-test
`TerminalEngineLockPayload.schema.json`/corpus,
`TerminalEngineMaterialization.schema.json`/corpus,
the closed candidate build-receipt, verification-receipt, and lane schemas,
`Verify-TerminalEngineCandidate.ps1`,
`New-TerminalEngineMaterialization.ps1`, and
`New-TerminalEngineLock.ps1 -Mode FirstSelection`, plus the closed post-adoption
verification-set schema/corpus and
`New-TerminalEngineVerificationSet.ps1`. The lock-independent
candidate verifier produces exactly five native lane records from the exact
winner/finalization/descriptor under clean offline isolation; each record binds
separate subordinate build and verification receipt digests, never its own
digest. The
materialization record contains the pre-lock selection identities, the shared
closed lock payload projection, the exact dedicated payload-schema digest, and
those five lane records, with no lock path/digest field; its writer reopens
every subject and rejects mixed or nonnative evidence. The lock record binds
the same payload-schema digest. The lock writer is the only non-overwriting
initial-lock writer. Only afterward run the five lock-bound Build/Verify lanes
and use the verification-set tool to bind their ten exact manifests and
identity-set digest in the decision/handoff, not in the lock. Keep
`Prepare-TerminalEngineUpgrade.ps1` exclusive to subsequent upgrades.

Before the first shipped-lock upgrade, implement the master's closed
`TerminalEngineUpgradeReview` v1 schema/corpus/module and add mandatory
candidate-lock/review file-and-expected-digest collection binding. Its JCS
record contains the two lock
digests, target identity, exactly 14 category snapshot reviews, and a sorted
unique repository-relative path/action list with conditional current/candidate
file digests. Collect binds the independent candidate-lock and review-record
paths/digests in every lane; Finalize requires both exact transferred files and
expected digests, reopens them, validates the review SHA companion, recomputes
the category and file digests, rejects any path/digest/ordering/closure
mismatch, and binds both digests into its comparison. The review binds the candidate-lock digest but the
lock does not bind the review, avoiding a hash cycle. Two non-author reviewers
approve the candidate-lock, review-record, and final comparison digests.

`Collect` does not invent a candidate lock. Before official evidence, the
Terminal plugin/dependency owner performs a non-archiving discovery build under
`.build`, authors one complete candidate lock under `.build` against
`TerminalEngineLock.schema.json`, and replaces every target-affected source,
tool, dependency, runtime, license, output, ABI, capability, qualification,
lane, and upgrade field with exact discovered identities. Copying an unchanged
field is allowed only after it is reverified. The lock binds
`upgrade.basedOnLockSha256`, `upgrade.targetSourceSha256`, the exact expected
outputs for all five lanes, and the complete derived change set; it contains no
placeholder, branch, or “latest” value. Restore its exact inputs, validate it
offline with `-RequireCache`, and have two non-author reviewers approve its
SHA-256 before any official `Collect`. Transfer the byte-identical lock,
review JSON/companion, and content-addressed cache to each exclusive disposable
native runner. Transfer the same exact lock/review bytes with the five returned
manifests to the review machine. `Finalize` then verifies that all five records
bind those identities, reopens the supplied bytes, copies them to its canonical
candidate root, and emits the comparison manifest; it does not synthesize or
repair either record.

This discovery step is deliberately a separately reviewed manual boundary, not
an automatic “upgrade to latest” command. Its required inputs are the reviewed
current-lock digest, one canonical credential-free target source URI, exact
target pin, source size/digest established by the first discovery transfer, and
the candidate-specific upstream build recipe/tool identities. Source
acquisition rejects redirects and verifies the downloaded bytes; non-archiving
exploratory builds run only on exclusive disposable matching-native machines.
Its required output is the complete candidate lock plus the closed
`TerminalEngineUpgradeReview` record binding the current/candidate lock digests
and target source digest. The record contains all 14 changed/retained canonical
category digests and, for each changed category, every normalized
repository-relative path, `add|modify|remove` action, role, and conditional
current/candidate file digest. It contains no absolute path, command,
credential, source payload, or terminal/user content. Two non-author reviewers
approve both the candidate-lock and review-record digests, and every subsequent
shipped-lock upgrade binds the record digest into evidence before `Collect`.
If the URI/pin cannot be authenticated, the build needs ambient state, the
record is unbound, or reviewers cannot establish every lock field without a
placeholder, STOP before `Collect`. The easy routine operation is read-only
status/cache verification; an upgrade is intentionally not one-click.

Long-lived operator commands after the first lock is shipped:

```powershell
.\Tools\Get-TerminalDependencyStatus.ps1 -Mode Locked -EngineLock .\External\terminal-engine.lock.json
.\Tools\Restore-TerminalEngineInputs.ps1 -LockFile .\External\terminal-engine.lock.json
.\Tools\Restore-TerminalEngineInputs.ps1 -LockFile .\External\terminal-engine.lock.json -Offline
.\Tools\Get-TerminalDependencyStatus.ps1 -Mode Locked -EngineLock .\External\terminal-engine.lock.json -RequireCache
$onlineReport = .\Tools\Get-TerminalDependencyStatus.ps1 -Mode Online -EngineLock .\External\terminal-engine.lock.json -RequireOnline -PassThruReport
if (-not (Test-Path -LiteralPath $onlineReport -PathType Leaf)) { throw 'Online dependency status did not return its exact report path.' }
```

After the target discovery spike has produced the complete pre-reviewed
candidate lock, acquire and verify its exact inputs:

```powershell
$candidateLockFile = '<exact candidate lock below .build>'
$expectedCandidateLockSha256 = '<two-reviewer-approved 64-lowercase-hex digest>'
$upgradeReviewFile = '<exact TerminalEngineUpgradeReview v1 record below .build>'
$expectedUpgradeReviewSha256 = '<two-reviewer-approved 64-lowercase-hex digest>'
if ((Get-FileHash -LiteralPath $candidateLockFile -Algorithm SHA256).Hash.ToLowerInvariant() -cne $expectedCandidateLockSha256) { throw 'Candidate lock differs from the approved bytes.' }
if ((Get-FileHash -LiteralPath $upgradeReviewFile -Algorithm SHA256).Hash.ToLowerInvariant() -cne $expectedUpgradeReviewSha256) { throw 'Upgrade review differs from the approved bytes.' }
.\Tools\Restore-TerminalEngineInputs.ps1 -LockFile $candidateLockFile
.\Tools\Restore-TerminalEngineInputs.ps1 -LockFile $candidateLockFile -Offline
.\Tools\Get-TerminalDependencyStatus.ps1 -Mode Locked -EngineLock $candidateLockFile -RequireCache
```

Prepare an exact candidate on a native x64 machine:

```powershell
$currentLockSha256 = (Get-FileHash -LiteralPath .\External\terminal-engine.lock.json -Algorithm SHA256).Hash.ToLowerInvariant()
$targetPin = '<exact target commit or release identity>'
$targetSourceSha256 = '<64-lowercase-hex source tree or archive digest>'
$candidateLockFile = '<exact pre-reviewed candidate lock below .build>'
$expectedCandidateLockSha256 = '<same approved candidate-lock digest>'
$upgradeReviewFile = '<exact pre-reviewed TerminalEngineUpgradeReview v1 record below .build>'
$expectedUpgradeReviewSha256 = '<same approved upgrade-review digest>'
$upgradeRoot = '.\.build\TerminalEngineUpgrade'
if ((Get-FileHash -LiteralPath $candidateLockFile -Algorithm SHA256).Hash.ToLowerInvariant() -cne $expectedCandidateLockSha256) { throw 'Candidate lock differs from the approved bytes.' }
if ((Get-FileHash -LiteralPath $upgradeReviewFile -Algorithm SHA256).Hash.ToLowerInvariant() -cne $expectedUpgradeReviewSha256) { throw 'Upgrade review differs from the approved bytes.' }
.\Tools\Restore-TerminalEngineInputs.ps1 -LockFile $candidateLockFile -Offline
.\Tools\Get-TerminalDependencyStatus.ps1 -Mode Locked -EngineLock $candidateLockFile -RequireCache
$x64DebugUpgradeManifest = .\Tools\Prepare-TerminalEngineUpgrade.ps1 -Mode Collect -TargetPin $targetPin -TargetSourceSha256 $targetSourceSha256 -ExpectedCurrentLockSha256 $currentLockSha256 -CandidateLockFile $candidateLockFile -ExpectedCandidateLockSha256 $expectedCandidateLockSha256 -UpgradeReviewFile $upgradeReviewFile -ExpectedUpgradeReviewSha256 $expectedUpgradeReviewSha256 -Platform x64 -Configuration Debug -Bootstrap -Clean -Offline -ExclusiveDisposableRunner -UpgradeRoot $upgradeRoot -PassThruManifest
$x64ReleaseUpgradeManifest = .\Tools\Prepare-TerminalEngineUpgrade.ps1 -Mode Collect -TargetPin $targetPin -TargetSourceSha256 $targetSourceSha256 -ExpectedCurrentLockSha256 $currentLockSha256 -CandidateLockFile $candidateLockFile -ExpectedCandidateLockSha256 $expectedCandidateLockSha256 -UpgradeReviewFile $upgradeReviewFile -ExpectedUpgradeReviewSha256 $expectedUpgradeReviewSha256 -Platform x64 -Configuration Release -Bootstrap -Clean -Offline -ExclusiveDisposableRunner -UpgradeRoot $upgradeRoot -PassThruManifest
$x64AsanUpgradeManifest = .\Tools\Prepare-TerminalEngineUpgrade.ps1 -Mode Collect -TargetPin $targetPin -TargetSourceSha256 $targetSourceSha256 -ExpectedCurrentLockSha256 $currentLockSha256 -CandidateLockFile $candidateLockFile -ExpectedCandidateLockSha256 $expectedCandidateLockSha256 -UpgradeReviewFile $upgradeReviewFile -ExpectedUpgradeReviewSha256 $expectedUpgradeReviewSha256 -Platform x64 -Configuration 'ASan Debug' -Bootstrap -Clean -Offline -ExclusiveDisposableRunner -UpgradeRoot $upgradeRoot -PassThruManifest
if (-not (Test-Path -LiteralPath $x64DebugUpgradeManifest -PathType Leaf) -or -not (Test-Path -LiteralPath $x64ReleaseUpgradeManifest -PathType Leaf) -or -not (Test-Path -LiteralPath $x64AsanUpgradeManifest -PathType Leaf)) { throw 'x64 upgrade collection did not return all three exact manifests.' }
```

Prepare the same candidate on a native ARM64 machine:

```powershell
$currentLockSha256 = (Get-FileHash -LiteralPath .\External\terminal-engine.lock.json -Algorithm SHA256).Hash.ToLowerInvariant()
$targetPin = '<same exact target commit or release identity>'
$targetSourceSha256 = '<same 64-lowercase-hex source tree or archive digest>'
$candidateLockFile = '<same exact pre-reviewed candidate lock below .build>'
$expectedCandidateLockSha256 = '<same approved candidate-lock digest>'
$upgradeReviewFile = '<same exact pre-reviewed TerminalEngineUpgradeReview v1 record below .build>'
$expectedUpgradeReviewSha256 = '<same approved upgrade-review digest>'
$upgradeRoot = '.\.build\TerminalEngineUpgrade'
if ((Get-FileHash -LiteralPath $candidateLockFile -Algorithm SHA256).Hash.ToLowerInvariant() -cne $expectedCandidateLockSha256) { throw 'Candidate lock differs from the approved bytes.' }
if ((Get-FileHash -LiteralPath $upgradeReviewFile -Algorithm SHA256).Hash.ToLowerInvariant() -cne $expectedUpgradeReviewSha256) { throw 'Upgrade review differs from the approved bytes.' }
.\Tools\Restore-TerminalEngineInputs.ps1 -LockFile $candidateLockFile -Offline
.\Tools\Get-TerminalDependencyStatus.ps1 -Mode Locked -EngineLock $candidateLockFile -RequireCache
$arm64DebugUpgradeManifest = .\Tools\Prepare-TerminalEngineUpgrade.ps1 -Mode Collect -TargetPin $targetPin -TargetSourceSha256 $targetSourceSha256 -ExpectedCurrentLockSha256 $currentLockSha256 -CandidateLockFile $candidateLockFile -ExpectedCandidateLockSha256 $expectedCandidateLockSha256 -UpgradeReviewFile $upgradeReviewFile -ExpectedUpgradeReviewSha256 $expectedUpgradeReviewSha256 -Platform ARM64 -Configuration Debug -Bootstrap -Clean -Offline -ExclusiveDisposableRunner -UpgradeRoot $upgradeRoot -PassThruManifest
$arm64ReleaseUpgradeManifest = .\Tools\Prepare-TerminalEngineUpgrade.ps1 -Mode Collect -TargetPin $targetPin -TargetSourceSha256 $targetSourceSha256 -ExpectedCurrentLockSha256 $currentLockSha256 -CandidateLockFile $candidateLockFile -ExpectedCandidateLockSha256 $expectedCandidateLockSha256 -UpgradeReviewFile $upgradeReviewFile -ExpectedUpgradeReviewSha256 $expectedUpgradeReviewSha256 -Platform ARM64 -Configuration Release -Bootstrap -Clean -Offline -ExclusiveDisposableRunner -UpgradeRoot $upgradeRoot -PassThruManifest
if (-not (Test-Path -LiteralPath $arm64DebugUpgradeManifest -PathType Leaf) -or -not (Test-Path -LiteralPath $arm64ReleaseUpgradeManifest -PathType Leaf)) { throw 'ARM64 upgrade collection did not return both exact manifests.' }
```

Transfer those exact manifests, the approved candidate lock, and the approved
upgrade-review JSON plus SHA companion to the review machine and finalize
without discovery:

```powershell
$currentLockSha256 = '<same 64-lowercase-hex current lock digest>'
$targetPin = '<same exact target commit or release identity>'
$targetSourceSha256 = '<same 64-lowercase-hex source tree or archive digest>'
$x64DebugUpgradeManifest = '<exact transferred x64 Debug upgrade manifest>'
$x64ReleaseUpgradeManifest = '<exact transferred x64 Release upgrade manifest>'
$x64AsanUpgradeManifest = '<exact transferred x64 ASan Debug upgrade manifest>'
$arm64DebugUpgradeManifest = '<exact transferred ARM64 Debug upgrade manifest>'
$arm64ReleaseUpgradeManifest = '<exact transferred ARM64 Release upgrade manifest>'
$candidateLockFile = '<exact transferred approved candidate lock below .build>'
$expectedCandidateLockSha256 = '<same approved candidate-lock digest>'
$upgradeReviewFile = '<exact transferred approved upgrade-review JSON below .build>'
$expectedUpgradeReviewSha256 = '<same approved upgrade-review digest>'
if ((Get-FileHash -LiteralPath $candidateLockFile -Algorithm SHA256).Hash.ToLowerInvariant() -cne $expectedCandidateLockSha256) { throw 'Transferred candidate lock differs from approved bytes.' }
if ((Get-FileHash -LiteralPath $upgradeReviewFile -Algorithm SHA256).Hash.ToLowerInvariant() -cne $expectedUpgradeReviewSha256) { throw 'Transferred upgrade review differs from approved bytes.' }
$upgradeComparisonManifest = .\Tools\Prepare-TerminalEngineUpgrade.ps1 -Mode Finalize -TargetPin $targetPin -TargetSourceSha256 $targetSourceSha256 -ExpectedCurrentLockSha256 $currentLockSha256 -CandidateLockFile $candidateLockFile -ExpectedCandidateLockSha256 $expectedCandidateLockSha256 -UpgradeReviewFile $upgradeReviewFile -ExpectedUpgradeReviewSha256 $expectedUpgradeReviewSha256 -EvidenceManifest @($x64DebugUpgradeManifest, $x64ReleaseUpgradeManifest, $x64AsanUpgradeManifest, $arm64DebugUpgradeManifest, $arm64ReleaseUpgradeManifest) -UpgradeRoot .\.build\TerminalEngineUpgrade -RequireAllEvidence -PassThruManifest
if (-not (Test-Path -LiteralPath $upgradeComparisonManifest -PathType Leaf)) { throw 'Upgrade finalization did not return its exact comparison manifest.' }
```

`Finalize` returns only the exact comparison-manifest path. It reopens the
explicit candidate-lock and review bytes, validates the review companion,
matches both expected digests against all five lane bindings, and rejects an
absent or substituted same-path file. The manifest binds the five input
digests, current/candidate lock digests, upgrade-review digest, target
pin/source digest, every compared dependency/license/ABI/capability/performance
field, changed comparison categories, rollback result, and overall pass/fail.
It does not claim to emit the exact repository patch; reviewers reconcile those
categories with the approved repository-relative action record. A complete
failed comparison may be retained under `.build` but cannot be applied. The
runbook never tells an operator to overwrite the current lock from an
unverified path.

## Verification

### Round 4 zero-cohort closeout (active)

Round 4 has zero `disposition="collect"` nominations. Its valid verification
path is therefore the pre-admission path: validate the exact independently
approved universe, candidate review, empty admitted-candidate assessment,
candidate set, patch/concern-map/rejection proofs, immutable snapshot boundary,
and decision; then prove both the evaluator `Collect` entry point and the direct
provider entry point reject before lane attestation, bootstrap, clean,
subprocess launch, or output-root mutation. Run the focused Gate-0 and
dependency-lifecycle Pester contracts and the repository's final required test
suite. Record their exact final results in the completion record below.

Round 4 MUST NOT create collection or finalization manifests merely to fill a
handoff field. It also MUST NOT run or claim a native winner lane. The commands
in the next subsection are retained as candidate-neutral reusable instructions
for a later v5-or-later universe that actually admits an executable cohort;
they are not Round 4 closeout steps.

### Executable-cohort verification (future v5+ only)

Native x64 candidate collection; retain both exact printed paths:

```powershell
$x64ReleaseManifest = .\Tools\Evaluate-TerminalEngines.ps1 -Mode Collect -Candidate All -Platform x64 -Configuration Release -Bootstrap -Clean -Offline -EvidenceRoot .\Specs\TestRuns -PassThruManifest
$x64AsanManifest = .\Tools\Evaluate-TerminalEngines.ps1 -Mode Collect -Candidate All -Platform x64 -Configuration "ASan Debug" -Bootstrap -Clean -Offline -EvidenceRoot .\Specs\TestRuns -PassThruManifest
if (-not (Test-Path -LiteralPath $x64ReleaseManifest -PathType Leaf) -or -not (Test-Path -LiteralPath $x64AsanManifest -PathType Leaf)) { throw 'x64 collection did not return both exact manifest paths.' }
```

Native ARM64 candidate collection; retain the exact printed path:

```powershell
$arm64ReleaseManifest = .\Tools\Evaluate-TerminalEngines.ps1 -Mode Collect -Candidate All -Platform ARM64 -Configuration Release -Bootstrap -Clean -Offline -EvidenceRoot .\Specs\TestRuns -PassThruManifest
if (-not (Test-Path -LiteralPath $arm64ReleaseManifest -PathType Leaf)) { throw 'ARM64 collection did not return its exact manifest path.' }
```

Review-machine finalization after transferring those exact manifests; there is no “latest” fallback:

```powershell
$x64ReleaseManifest = '<exact x64 Release collection-manifest path>'
$x64AsanManifest = '<exact x64 ASan Debug collection-manifest path>'
$arm64ReleaseManifest = '<exact ARM64 Release collection-manifest path>'
$finalizationManifest = .\Tools\Evaluate-TerminalEngines.ps1 -Mode Finalize -Candidate All -EvidenceManifest @($x64ReleaseManifest, $x64AsanManifest, $arm64ReleaseManifest) -EvidenceRoot .\Specs\TestRuns -RequireAllEvidence -PassThruManifest
if (-not (Test-Path -LiteralPath $finalizationManifest -PathType Leaf)) { throw 'Finalize did not return its exact manifest path.' }
```

After a future Finalize names one threshold-passing winner, use this exact
first-selection sequence. These tools/schemas are future winner-bearing-round
deliverables and are intentionally absent from the Round-4 zero-winner
scaffold.

On a native x64 exclusive disposable runner, create three lock-independent
candidate lanes:

```powershell
$finalizationManifest = '<exact approved Gate-0 finalization manifest>'
$candidateSet = '.\Specs\Terminal\TerminalEngineCandidateSet.v1.json'
$candidateDescriptor = '<exact frozen winner descriptor>'
$candidateInputRoot = '<exact restored frozen winner input root>'
$candidateEvidenceRoot = '.\.build\TerminalEngineFirstSelection'
$x64DebugCandidateLane = .\Tools\Verify-TerminalEngineCandidate.ps1 -FinalizationManifest $finalizationManifest -CandidateSet $candidateSet -CandidateDescriptor $candidateDescriptor -InputRoot $candidateInputRoot -Platform x64 -Configuration Debug -Bootstrap -Clean -Offline -ExclusiveDisposableRunner -EvidenceRoot $candidateEvidenceRoot -PassThruLaneRecord
$x64ReleaseCandidateLane = .\Tools\Verify-TerminalEngineCandidate.ps1 -FinalizationManifest $finalizationManifest -CandidateSet $candidateSet -CandidateDescriptor $candidateDescriptor -InputRoot $candidateInputRoot -Platform x64 -Configuration Release -Bootstrap -Clean -Offline -ExclusiveDisposableRunner -EvidenceRoot $candidateEvidenceRoot -PassThruLaneRecord
$x64AsanCandidateLane = .\Tools\Verify-TerminalEngineCandidate.ps1 -FinalizationManifest $finalizationManifest -CandidateSet $candidateSet -CandidateDescriptor $candidateDescriptor -InputRoot $candidateInputRoot -Platform x64 -Configuration 'ASan Debug' -Bootstrap -Clean -Offline -ExclusiveDisposableRunner -EvidenceRoot $candidateEvidenceRoot -PassThruLaneRecord
@($x64DebugCandidateLane,$x64ReleaseCandidateLane,$x64AsanCandidateLane) | ForEach-Object { if (-not (Test-Path -LiteralPath $_ -PathType Leaf)) { throw 'x64 candidate lane record is missing.' } }
```

On a native ARM64 exclusive disposable runner, using byte-identical
finalization/set/descriptor/inputs, create the other two:

```powershell
$finalizationManifest = '<same exact approved Gate-0 finalization manifest>'
$candidateSet = '.\Specs\Terminal\TerminalEngineCandidateSet.v1.json'
$candidateDescriptor = '<same exact frozen winner descriptor>'
$candidateInputRoot = '<same exact restored frozen winner input root>'
$candidateEvidenceRoot = '.\.build\TerminalEngineFirstSelection'
$arm64DebugCandidateLane = .\Tools\Verify-TerminalEngineCandidate.ps1 -FinalizationManifest $finalizationManifest -CandidateSet $candidateSet -CandidateDescriptor $candidateDescriptor -InputRoot $candidateInputRoot -Platform ARM64 -Configuration Debug -Bootstrap -Clean -Offline -ExclusiveDisposableRunner -EvidenceRoot $candidateEvidenceRoot -PassThruLaneRecord
$arm64ReleaseCandidateLane = .\Tools\Verify-TerminalEngineCandidate.ps1 -FinalizationManifest $finalizationManifest -CandidateSet $candidateSet -CandidateDescriptor $candidateDescriptor -InputRoot $candidateInputRoot -Platform ARM64 -Configuration Release -Bootstrap -Clean -Offline -ExclusiveDisposableRunner -EvidenceRoot $candidateEvidenceRoot -PassThruLaneRecord
@($arm64DebugCandidateLane,$arm64ReleaseCandidateLane) | ForEach-Object { if (-not (Test-Path -LiteralPath $_ -PathType Leaf)) { throw 'ARM64 candidate lane record is missing.' } }
```

Transfer each complete lane directory—not only its top record—to the review
machine. Each directory must contain the exact candidate build receipt,
verification receipt, lane record, and all three SHA companions. Also transfer
the exact finalization, candidate-set, descriptor, restored-input identities,
and schema bytes. On the review machine:

```powershell
$finalizationManifest = '<exact transferred finalization manifest>'
$candidateSet = '<exact transferred candidate set>'
$candidateDescriptor = '<exact transferred winner descriptor>'
$candidateLaneRecords = @(
    '<exact transferred x64 Debug candidate-lane record>',
    '<exact transferred x64 Release candidate-lane record>',
    '<exact transferred x64 ASan Debug candidate-lane record>',
    '<exact transferred ARM64 Debug candidate-lane record>',
    '<exact transferred ARM64 Release candidate-lane record>')
$materialization = .\Tools\New-TerminalEngineMaterialization.ps1 -FinalizationManifest $finalizationManifest -CandidateSet $candidateSet -CandidateDescriptor $candidateDescriptor -LaneRecord $candidateLaneRecords -OutputRoot .\.build\TerminalEngineFirstSelection\Review -FailIfExists -PassThruManifest
$materializationSha256 = (Get-FileHash -LiteralPath $materialization -Algorithm SHA256).Hash.ToLowerInvariant()
if ($materializationSha256 -notmatch '^[0-9a-f]{64}$') { throw 'Materialization digest is invalid.' }
```

Two distinct non-author reviewers re-open those exact bytes and approve
`$materializationSha256`; record reviewer IDs, UTC-second times, and the subject
digest. Only then:

```powershell
$approvedMaterializationSha256 = '<exact digest approved by both reviewers>'
if ((Get-FileHash -LiteralPath $materialization -Algorithm SHA256).Hash.ToLowerInvariant() -cne $approvedMaterializationSha256) { throw 'Materialization differs from the approved bytes.' }
$candidateLock = .\Tools\New-TerminalEngineLock.ps1 -Mode FirstSelection -FinalizationManifest $finalizationManifest -CandidateSet $candidateSet -MaterializationFile $materialization -ExpectedMaterializationSha256 $approvedMaterializationSha256 -OutputFile .\.build\TerminalEngineFirstSelection\Review\terminal-engine.candidate.lock.json -FailIfExists -PassThruLock
$candidateLockSha256 = (Get-FileHash -LiteralPath $candidateLock -Algorithm SHA256).Hash.ToLowerInvariant()
if ($candidateLockSha256 -notmatch '^[0-9a-f]{64}$') { throw 'Candidate lock digest is invalid.' }
```

Two distinct non-author reviewers re-open and approve that exact candidate-lock
digest. In one reviewed adoption change, add the approved schemas/corpora/tools,
selected adapter/runtime declarations, notices/licenses, decision/ledger/tests/
documentation, and copy—not regenerate—the candidate lock:

```powershell
$approvedCandidateLockSha256 = '<exact digest approved by both reviewers>'
if ((Get-FileHash -LiteralPath $candidateLock -Algorithm SHA256).Hash.ToLowerInvariant() -cne $approvedCandidateLockSha256) { throw 'Candidate lock differs from the approved bytes.' }
if (Test-Path -LiteralPath .\External\terminal-engine.lock.json) { throw 'First selection refuses to overwrite a production lock.' }
Copy-Item -LiteralPath $candidateLock -Destination .\External\terminal-engine.lock.json
if ((Get-FileHash -LiteralPath .\External\terminal-engine.lock.json -Algorithm SHA256).Hash.ToLowerInvariant() -cne $approvedCandidateLockSha256) { throw 'Adopted lock differs from the approved candidate.' }
```

After reviewers approve that complete tracked diff, capture its exact commit.
On native x64 at that clean adoption commit, produce six lock-bound manifests:

```powershell
$lockFile = '.\External\terminal-engine.lock.json'
$x64DebugBuild = .\Tools\Build-TerminalEngine.ps1 -LockFile $lockFile -Platform x64 -Configuration Debug -Clean -Offline -ExclusiveDisposableRunner -PassThruManifest
$x64DebugVerify = .\Tools\Verify-TerminalEngine.ps1 -LockFile $lockFile -Platform x64 -Configuration Debug -RunSmoke -Offline -ExclusiveDisposableRunner -PassThruManifest
$x64ReleaseBuild = .\Tools\Build-TerminalEngine.ps1 -LockFile $lockFile -Platform x64 -Configuration Release -Clean -Offline -ExclusiveDisposableRunner -PassThruManifest
$x64ReleaseVerify = .\Tools\Verify-TerminalEngine.ps1 -LockFile $lockFile -Platform x64 -Configuration Release -RunSmoke -Offline -ExclusiveDisposableRunner -PassThruManifest
$x64AsanBuild = .\Tools\Build-TerminalEngine.ps1 -LockFile $lockFile -Platform x64 -Configuration 'ASan Debug' -Clean -Offline -ExclusiveDisposableRunner -PassThruManifest
$x64AsanVerify = .\Tools\Verify-TerminalEngine.ps1 -LockFile $lockFile -Platform x64 -Configuration 'ASan Debug' -RunSmoke -Offline -ExclusiveDisposableRunner -PassThruManifest
```

On native ARM64 at the same clean adoption commit, produce four:

```powershell
$lockFile = '.\External\terminal-engine.lock.json'
$arm64DebugBuild = .\Tools\Build-TerminalEngine.ps1 -LockFile $lockFile -Platform ARM64 -Configuration Debug -Clean -Offline -ExclusiveDisposableRunner -PassThruManifest
$arm64DebugVerify = .\Tools\Verify-TerminalEngine.ps1 -LockFile $lockFile -Platform ARM64 -Configuration Debug -RunSmoke -Offline -ExclusiveDisposableRunner -PassThruManifest
$arm64ReleaseBuild = .\Tools\Build-TerminalEngine.ps1 -LockFile $lockFile -Platform ARM64 -Configuration Release -Clean -Offline -ExclusiveDisposableRunner -PassThruManifest
$arm64ReleaseVerify = .\Tools\Verify-TerminalEngine.ps1 -LockFile $lockFile -Platform ARM64 -Configuration Release -RunSmoke -Offline -ExclusiveDisposableRunner -PassThruManifest
```

Transfer all ten exact manifests and companions to the review machine at the
same adoption commit, then close the set without path discovery:

```powershell
$buildManifests = @('<x64 Debug build>','<x64 Release build>','<x64 ASan Debug build>','<ARM64 Debug build>','<ARM64 Release build>')
$verificationManifests = @('<x64 Debug verify>','<x64 Release verify>','<x64 ASan Debug verify>','<ARM64 Debug verify>','<ARM64 Release verify>')
$verificationSet = .\Tools\New-TerminalEngineVerificationSet.ps1 -LockFile .\External\terminal-engine.lock.json -BuildManifest $buildManifests -VerificationManifest $verificationManifests -OutputRoot .\.build\TerminalEngineFirstSelection\PostAdoption -FailIfExists -PassThruManifest
$verificationSetSha256 = (Get-FileHash -LiteralPath $verificationSet -Algorithm SHA256).Hash.ToLowerInvariant()
$verificationIdentitySetSha256 = [string]((Get-Content -LiteralPath $verificationSet -Raw | ConvertFrom-Json -Depth 100).identitySetSha256)
if ($verificationSetSha256 -notmatch '^[0-9a-f]{64}$' -or $verificationIdentitySetSha256 -notmatch '^[0-9a-f]{64}$') { throw 'Post-adoption verification-set identity is invalid.' }
```

The handoff records exact finalization, materialization, candidate-lock,
adoption-commit, ten manifest, verification-set file, and
`identitySetSha256` identities plus both reviewer pairs/times. Any failed
lock-independent lane creates no materialization/lock/adoption and never falls
through to a runner-up; fix via a new pin/design in a freshly frozen round. Any
failed lock-bound lane bars adoption closeout and requires reverting the whole
adoption change before retry; an unadopted candidate lock stays below `.build`.

Expected only for a future executable cohort: all three explicitly named
manifests validate as one frozen evidence set; every scored candidate has
complete comparable x64 Release/x64 ASan/ARM64 Release evidence or a
reproducible first-gate failure, and every control has three reproducible
constraint-failure records. Finalize returns one exact finalization-manifest
path whose companion digest is approved with the decision. Constraint controls
are unscored; winner artifacts reproduce offline with exact
architecture/build/type/capability/license identity and run natively on
x64/ARM64. Incomplete/invalid evidence and complete/no-winner have distinct
nonzero results.

### Reusable dependency-lifecycle scaffold

Run the focused dependency-lifecycle contract in Round 4:

```powershell
. .\Tools\TestRunPlan.ps1
$pesterParameters = New-RSPesterInvokeParameters -Path (Resolve-Path .\Tools\Tests\TerminalDependencyLifecycle.Tests.ps1).Path
$pesterResult = Invoke-Pester @pesterParameters
$failedProperty = $pesterResult.PSObject.Properties['FailedCount']
$failed = if ($failedProperty) { $failedProperty.Value } else { $pesterResult.PSObject.Properties['Failed'].Value }
if ([int]$failed -ne 0) { exit 1 }
```

Expected for Round 4: the tests pass against only local fake remotes, synthetic
locks, and disposable build roots. They prove that
`Get-TerminalDependencyStatus.ps1`,
`Restore-TerminalEngineInputs.ps1`,
`Prepare-TerminalEngineUpgrade.ps1`, the lock schema/corpus, rollback, and
build/verifier routing remain reusable. They do not prove a selected engine or
authorize creation of `External/terminal-engine.lock.json`.

After a future winner has produced a reviewed real lock, the operator also runs
the mutation-free `Get-TerminalDependencyStatus.ps1 -Mode Locked` and
`Restore-TerminalEngineInputs.ps1 -Offline` checks from the dependency
maintenance workflow. Those real-lock checks are N/A in Round 4 because no
engine lock exists.

## Child completion and move

- **Completion record:** `TERMINAL-GATE0-ROUND4-NO-WINNER`
- **Frozen-review commit:**
  `3e84e87b717841e28e11e1fe7e5bc04b4e3a3e0d`
- **Universe:** `Specs/Terminal/TerminalEngineCandidateUniverse.v4.json`,
  SHA-256
  `7621ab8562cb4b1d530d24355c67d90c905781a5d4f6bbc3364699ac059cb740`,
  proposal SHA-256
  `22ba55a19bfa631093af06729b54f4eba46c5e23198f8f421accad634d106129`,
  frozen UTC `2026-07-31T01:47:10Z`, approvals `codex-dalton` at
  `2026-07-31T01:51:32Z` and `codex-contour-independent` at
  `2026-07-31T01:53:33Z`
- **Candidate review:** `Specs/Terminal/TerminalEngineCandidateReview.v1.json`,
  SHA-256
  `ba2942c7abe8f94b6f50be22ddb78570eaade5f60a152a8b0cca48dcea91542d`,
  source approvals `codex-root` at `2026-07-31T02:27:20Z` and
  `codex-contour-independent` at `2026-07-31T02:28:37Z`
- **Candidate assessment:**
  `Specs/Terminal/TerminalEngineCandidateAssessment.v1.json`, SHA-256
  `adf5ae89c3c90fd6018652f9ebd88aa809048a849ba29f49b04a4fde6d20343e`;
  admitted candidate count `0`; admitted-candidate approvals **N/A**
- **Candidate set:** `Specs/Terminal/TerminalEngineCandidateSet.v1.json`,
  SHA-256
  `fb240d3ef838daea0740417fafd2d7d44b9536295f16c607e45fe57e64d11f4d`,
  common freeze UTC `2026-07-31T02:29:00Z`, adaptation-rejection approvals
  `codex-root` at `2026-07-31T02:27:20Z` and
  `codex-contour-independent` at `2026-07-31T02:28:37Z`
- **Decision:** `Specs/Terminal/TerminalEngineDecision.md`, SHA-256
  `e10e6c5b885fc40496fb3eb1a7db9feb82084a74a368567305374bc0787e8260`,
  decision UTC `2026-07-31T17:42:43Z`, cold audits
  `codex-round4-cold-audit-one` at `2026-07-31T17:48:06Z` and
  `codex-round4-cold-audit-two` at `2026-07-31T17:47:47Z`; both independently
  recomputed prefix SHA-256
  `c9cf3f18ebe2fb53736c9589bdcbad5d441d1e5cfcd98d8ab07ba7d2e362e3b3`
  and returned clean with no P1/P2 findings
- **Verification:** the exact 20-file focused Gate-0 Pester set passed `239/239`
  twice before the frozen review, including the post-runtime-edit run
  immediately before that commit. After quarantining nine disclosed historical
  TestSandbox-root fixture directories, accepted isolated run
  `r4clean-284e0752` passed the same `239/239` in 639.63 Pester seconds and its
  same-process disk audit returned clean with zero issues; the dependency-
  lifecycle synthetic contract passed `12/12` in 7.005 wall seconds; the
  required post-full zero-cohort proof
  completed `2026-07-31T16:59:44Z`–`2026-07-31T17:05:00Z` with
  `Status=valid-no-winner-product-blocked`, six candidate records, zero
  `collect` candidates, evaluator exit `21`, zero success-stream bytes, absent
  evidence roots, and no lock/product/runtime changes. The canonical full suite
  completed all remaining lanes and exited `1` only for the reproduced,
  pre-existing, non-Terminal Commands Preferences/DxUi non-return, ViewerText
  clipboard-warning assertion, and three Tools Pester baseline-contract
  exceptions. Its aggregate run
  `20260731T153831Z-35744-9dceccf8fcac42a8b91d71f30c3a489e`
  recorded 17 suites, 776 total entries, 722 passed, 3 classified failed
  suite/case entries, and 51 skipped; the aggregate SHA-256 is
  `e9a9e7bf6b5e2882753d5c5ff566121f3b1c914c089cf1c129e7c9032f7998f4`.
  Exact current/prior traces, owner classifications, zero-impact diff, and
  wrapper streams are archived under
  `Specs/TestRuns/7d3a1247382a/Continuation/2026-07-31_161745_terminal_gate0_commands_plugins_uia_hang/`.

All Round-4 closeout values above are exact frozen identities. Angle-bracketed
metavariables in the future-v5 command examples elsewhere remain illustrative
and are not Round-4 Done blockers.

Round 4's exact no-winner disposition is:

| Potential handoff output | Round 4 disposition |
|---|---|
| Scored collection manifests and run IDs | **N/A** — zero nominations were admitted, so `Collect` must reject before effects. |
| Finalization manifest | **N/A** — zero admitted candidates make finalization invalid. |
| Winner score, threshold cohort, and tie-break tuple | **N/A** — no candidate reached scored collection. |
| Winner identity, adapter identity, and engine lock | **N/A** — no `External/terminal-engine.lock.json` may be fabricated. |
| `offlineLockedStatus` and `offlineInputCacheStatus` for a real engine | **N/A** — there is no real engine lock or cache; synthetic lifecycle tests remain reusable evidence about the tooling only. |
| x64 Debug/Release/ASan Debug and ARM64 Debug/Release winner artifacts | **N/A** — native winner lanes cannot be claimed for an unadmitted candidate. |
| Selected-engine runtime/import/SBOM/license/notice/package closure | **N/A** — no selected engine is staged or shipped. |
| Product plans 2–6 | **BLOCKED** — no product terminal implementation is authorized. |

Plan 1 is complete when the placeholders are replaced, the exact Round 4
universe/review/empty-assessment/set chronology and independent approvals
validate, the immutable governance-consumption and zero-effect rejection tests
pass, every focused Gate-0/dependency/lifecycle test is green, every regression
introduced by this work is repaired, and the canonical full-suite result meets
the master's failure-classification policy. A genuine unrelated pre-existing
environmental failure may be recorded only with its exact command/test/log,
reproduction and owner classification, prior matching baseline/archive, proof
of zero Gate-0 impact, and the wrapper's completed remaining-lane results; it is
not permission to skip a focused or required Gate-0 lane. Finally,
`Specs/Terminal/TerminalEngineDecision.md` records the exact no-current-winner
outcome and durable dependency workflow. The candidate-neutral schemas,
harnesses, evaluator/provider, and synthetic dependency-lifecycle tools/tests
remain checked-in reusable infrastructure; their existence is not a winner
artifact.

The master plan remains `WIP`. Plans 2–6 remain blocked and MUST NOT be rewritten
with a least-bad candidate, guessed adapter, or placeholder lock. Round 4 is the
root of the append-only Gate-0 chain, has no predecessor, and contains no
`successfulGate0Handoff` success marker. The user has authorized the proposed
synchronous caller-owned WezTerm redesign only as the separate evidence-only
Round-5 plan
`Specs/Plans/WIP/Terminal_EngineSelectionAndDependencyProof_Round5_WezTermSynchronousCallerOwned_2026-07-31.md`.
That successor stays blocked until this Round-4 plan is in Done and its exact
closeout commit, Done blob, and commit-bound decision digest have been copied
into and verified by the Round-5 predecessor fields. This does not authorize
product implementation.

Archive only schema-allowed content-free TerminalEngine evidence when an
executable cohort exists. Round 4 produces no collection/finalization archive.
Before moving, ensure the durable no-winner, dependency-lifecycle, immutable
governance-consumption, centralized runtime-dependency, and product-block
contracts are present in `Specs/Terminal/TerminalEngineDecision.md` and the
affected authoritative terminal/testing guidance.

Then set this file's `State` to `Done`, ensure it has no unchecked checkbox, and execute exactly:

```powershell
git mv Specs\Plans\WIP\Terminal_EngineSelectionAndDependencyProof_2026-07-22.md Specs\Plans\Done\Terminal_EngineSelectionAndDependencyProof_2026-07-22.md
```

After the move commit exists, make one reviewed handoff-publication change that
does not edit this Done file: record its exact Done path, closeout commit, Done
blob, commit-bound decision path/digest, and `outcome=no-winner` in the master's
exact `gate0Round[0].*` machine-readable tokens; copy the same commit/blob into
the `gate0Round[1].predecessor*` tokens and Round-5 predecessor fields; update
`Specs/Plans/WIP/README.md`; and route N3 to that exact Round-5 WIP plan. The master stays
WIP, its successful-handoff pointer remains `none`, and plans 2–6 remain blocked.
Do not point N3 at plan 2 until a later round publishes the sole authenticated
winner marker.

## Product-integration STOP conditions

- Frozen requirements or scoring must change after seeing candidate results.
- No candidate passes every mandatory gate and is the resulting deterministic
  inclusive-cohort/lexicographic-tie-break winner at or above 80.00. Round 4
  reached this condition: it blocks product integration but is a valid completed
  no-winner outcome for this Gate-0 child.
- A score or capability cannot be tied to reproducible evidence.
- The winner needs a second VT/Kitty parser, private unstable types, floating source, normal network access, or unbounded pre-allocation behavior.
- The winner cannot be wholly linked into or privately loaded by `Terminal.dll`, requires an EXE-owned terminal runtime/control/renderer, or cannot prove plugin-module callback/worker quiet before unload.
- Legal/dependency closure, x64/ARM64 build, snapshot lifetime, ABI identity, or two-reviewer approval is incomplete.

On product STOP, publish the decision findings and leave all product plans
blocked. Complete this evidence-only child with the exact no-winner record set;
do not select the least-bad failure.
