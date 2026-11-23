# I65 10-day audit remediation and evidence ledger

Status: ACTIVE
Priority: User-authorized multi-domain correctness/governance remediation; ordering follows current-source validation.
Planned-at commit: `78a8a75c46b784c52c5aeb82c0a3bf0773e61f6e` (`origin/main` at plan start).
Audit scope: user-supplied report `2ffe09a70..78a8a75c4`, covering 354 files over 10 days.
Starting worktree identity: 25 tracked dirty files; captured normalized `git diff --binary` SHA-256 `C8E670ED82A34C151B322E5FD9F86F6D670D2EE7552D6065B7259D185D061E29`.
Drift check: inspect `git status --short --branch`, resolve the current base/merge base, and compare each report finding against current source, callers, tests and contracts. Keep the starting worktree identity distinct from subsequent worker changes.

## Ownership and boundaries

Runtime and tooling workers own implementation in their assigned source/test surfaces. The integration owner maintains this ledger, its WIP index entry, and one dated `Changes/` fragment. I60 and I62 remain their existing roadmap owners; this plan may reconcile their status/anchors but cannot claim broader implementation complete. I64 remains the File Operations collected-collisions roadmap; mapped audit findings that overlap its code remain in scope without taking over its product decisions.

## Finding ledger

Status terminology: **validated** means current source review supports the defect; **implemented** describes candidate changes only; **executed/pass** is limited to the exact tests listed below. A Pester pass does not qualify native behavior.

| Report ID | Disposition / candidate status | Current evidence and validation state |
|---|---|---|
| BUG-01 | Rejected | Current File Operations contract/source around `FileSystem_FileOperations.md` section on queued execution holds admitted tasks in Preparing order; queued tasks wait behind earlier Parallel work. No defect reproduced. |
| BUG-02 | Validated; fix reported | Shared `WithdrawUnstartedTask` wake/remove behavior repaired. Deterministic runtime coverage is pending execution. |
| BUG-03 | Validated; bounded Release control passed | Bounded file replacement keeps mismatch errors deferred and non-mismatch failures separate. The Release R3 dummy control passed; other routes remain outside that selected-control result. |
| BUG-04 | Validated; implemented | Missing solution satellite dependencies under the `CurrentSolutionConfigurationContents` guard repaired; locale/project graph parity test reported **12 passed** in `ResourceLocalizationContracts`. |
| BUG-05 | Validated; actual v1 Commands diagnostic accepted; broader qualification pending | The verifier preserves strict durable v2 Fresh/promoted checks and accepts the actual single-suite Commands producer's v1 stable-path record through a validated prelaunch proof. The no-desktop `cmd_directory_cache_invalidation_during_enumeration_is_retained` probe passed through the real verifier at source 1cea (native exit 0; diagnostic Change PASSED, repository NOT_EVALUATED). This accepts the filtered diagnostic path only; Full/ARM64/native suite qualification remains pending. Historical 88-pass/2-fail tooling campaign remains retained. |
| BUG-06 | Partly validated | Overlay/navigation focus and drag paths repaired. Right-click-after-menu allegation rejected: current handler focuses the target first. Runtime execution remains pending. |
| BUG-07 | Validated; corrected Debug control passed | Quarter-wheel input clamp corrected. The corrected Debug Popup selected case passed and exercised quarter-detent controls; broader UI acceptance remains pending. |
| BUG-08 | Validated; implemented | Japanese Move wording now reflects the destructive move-only discard-source action; placeholders remain unchanged. The historical localization run predates the newly added runtime strings, so final current-resource validation remains pending. |
| BUG-09 | Validated cross-root namespace defects; focused controls executed | Local directory `FileIdInfo` includes RSLFID2, volume serial, and file ID, but does not prove IDs are comparable across redirector/export/WSL namespaces; UNC root IDs encode server/share while Local roots encode volume GUID. Ready traversed Local directory transfers across qualified roots run typed `LocalDirectoryAncestryProbeResult` marker proof after positive identity checks. The generic identity-domain helper is unchanged; the overbroad allegation that every differing `rootId` proves incomparability remains rejected. Same-path unsupported/bound Copy remains a duplicate and Move remains `ERROR_ALREADY_EXISTS`. A later accepted safety extension withholds `DiscardSource` when distinctness across qualified Local roots is unproved; identityless overwrite may publish while retaining the source. The directed Review run passed its selected case plus setup/cleanup (3 passed/0 failed/0 skipped, native exit 0); it exercised 21 ancestry, 6 power and 3 cross-root controls, including namespace-different same-entry Cancel, distinct-bound Cancel, and identityless Replace requiring Published + Completed + Retained with exact source/destination bytes. The canonical wrapper reports `NOT_EVALUATED` because this was a selected-case run. Paired performance evidence remains pending. |
| BUG-10 | Validated; fix reported | Extended distinct opens corrected. Runtime execution remains pending. |
| BUG-11 | Partly validated | Microsoft Drive rollback retry no longer mutates the fresh rollback leaf after collision; fake-Graph collision/exhaustion coverage is implemented, not executed. Explicit max-component length 255 allegation rejected because the descriptor default already provides 255; the boundary contract remains covered. |
| BUG-12 | Validated; implemented | Artifact-registry ASCII case-insensitive marker/extension checks use the canonical helper and fixed-width MTP suffix location. `FileOps_IdentityLessGuardPartition` coverage is implemented, not executed. |
| BUG-13 | Validated; fix reported | Shared reserved-device-name policy now trims before extension handling and covers COM/LPT superscript aliases; duplicate bridge policy removed. Runtime execution remains pending. |
| BUG-14 | Validated; corrected Debug controls passed | DPI pixel state, footer restoration and exact-ceiling growth corrected. The corrected Debug Popup selected case exercised same-physical-DPI state and one-pixel resize controls; broader UI acceptance remains pending. |
| BUG-15 | Partly validated by corrected Debug control | Completed-group reveal was exercised for both targets in the corrected Debug Popup selected case. Issues-pane copied-facts regression coverage is implemented in `RunFileOpsIssuesPaneUxChecks` (actual displayed fact strings plus exact paths); current-source execution and broader UI acceptance remain pending. |
| BUG-16 | Validated; fix reported | Queue reorder under lock corrected. Runtime execution remains pending. |
| BUG-17 | Validated; implemented | Historical-tree hashing now batches through the 64 MiB caller; meaningful batch-byte tests reported passed. No native builds/tests are claimed. |
| BUG-18 | Validated; implemented, subclaims 18.1–18.4 | Shared workflow/PATH MSBuild selection and vcpkg alignment; sanitized/contained SDK invocation; current-process PowerShell and no-optional-locks Git reads with deeper JSON identity; fail-closed single Clean before Rebuild. Initial PATH-only VS selection issue also fixed. Focused tooling passes are listed below; native acceptance is pending. |
| BUG-19 | Validated; fix reported | Managed ABI/thread exception boundaries retain justified broad catches, propagate errors after retirement, and aggregate action/display restoration failures. Focus-taking behavior unchanged. Native/runtime execution remains pending. |
| BUG-20 | Validated; Debug/Release private controls passed | Power request reason/writability are cached before the leaf mutex; exact last-error is preserved and failed-clear handles are retired before reacquisition; conflict waits use scoped cancellation/decision/progress events and a monotonic quiet deadline instead of 50 ms polling. Six private scenarios cover pre-answer, pre-cancel, quiet cancellation, resume/answer, stop-only, and spurious wakes, with Debug and Release selected controls passed. Real system sleep/CPU behavior is not claimed. |
| BUG-21 | Validated; implementation reported, subclaims 21.1–21.6 | Expected-occupant `NOT_SUPPORTED` prompts from fresh observed facts; S3 Table advertises no writer proof; retention reasons localized in five locales; dynamic NT buffer; debug delete-pending behavior; unused constant/stale “More” documentation fixed. Native/runtime execution remains pending. |
| GOV-01 | Live merge verified; W5 qualification unresolved | GitHub and local history confirm PR #201 merged 2026-11-08 as `a9e366240d2083484a12aa6177c5864d8e617164`. I60 Rule 3 required W5.1–W5.4 implementation and local + physical NAS verification before merge. The archive `Specs/TestRuns/4cb089111a23/Validation/2026-11-07_055500_managed-controls/README.md` records a preceding PR201 head `73d562a6` scale-3 Full pass, but its `Delete_IdentityLessPinnedTreeScale` took 398,235 ms, and the README states default NAS deadline failures and current-head qualification remained pending. The live merged-PR check rollup shows tooling/format/CodeQL and `ci-gate`, not a native W5 NAS gate. This evidence does not meet W5.1's <60 s target or prove all W5 gates on the final merged head. The merge does not prove or waive those gates; keep I60 ACTIVE pending exact per-wave evidence. |

### Additional report categories

| ID | Disposition / candidate status | Current evidence and validation state |
|---|---|---|
| PERF-01 | Partly implemented; one proposal deferred; queue-component v2 selected-case evidence executed | Temporary queue vector/sort removed. The reviewed `queue_head_selector_v1` capsule carries frozen-baseline provenance/hash; the directed Identity run passed its selected case plus setup/cleanup (3 passed/0 failed/0 skipped, native exit 0), producing 48 same-binary selector pairs, 96 timing rows and 48 capacity rows against the same current-harness baseline algorithm. These results are in the published native component packet below. This is not a whole-application historical baseline; Debug cohort timing regressions are retained and no speedup claim is made. Popup formatting cache optimization remains unproven and deferred. |
| PERF-02 | Implemented; bounded component evidence executed | Historical tree cache, one-scan archive path, and governance reuse by identical snapshot are implemented. `Specs/TestRuns/4cb089111a23/Validation/2026-11-13_150237_ToolingComponents/` contains the archive-validated six-file packet with 32 counterbalanced component observations, bound to baseline `78a8a75c4` and the exact candidate target hashes. It records tree queries 2→1; index queries 2→2 with staged addition observed (256→257 rows); fresh inventory calls 5→4; recovery 2→2; same-snapshot governance 1→0 and changed-snapshot governance 1→1 over 126 graph inputs. This is bounded component evidence only; no whole-planner or native performance claim is made. |
| ARCH-01.1 | Implemented | Windows helper unified. Runtime validation remains pending where applicable. |
| ARCH-01.2 | Rejected | No current duplicate prefix helper found; canonical helper remains authoritative. |
| ARCH-01.3 | Implemented | Dead scoped receipt/source/artifact/run/environment functions and stale exports/metadata removed; test migrated to obsolete-export absence and canonical runner capability guard. `ScopedTesting` and tool inventory checks passed. |
| UX-01 | Intentional restriction; selected Debug help controls passed | Broker capability restriction is by design; tooltip/UIA help now explains limited Delete Source repeat. Corrected Debug Popup controls passed for Move-positive, default-false and Copy-negative help; broader UI/interaction acceptance remains pending. |
| TEST-01 | Coverage map expanded; selected controls passed; Full/Affected and Full/Fresh runs failed | The no-activation `FileOps_ReviewRemediationContracts` case runs ancestry and cross-view matrices after its activation blocker and before prerequisite skips. Existing coverage includes Ready plus two bound handles for one physical directory with a test-only reported-namespace difference, same-object stability/revalidation, refusal of a nested alias before parent creation, exact bytes for ordinary distinct cross-root Copy, and the new `DiscardSource`/identityless-overwrite retention controls. Directed Review passed selected case plus setup/cleanup (3/0/0, native exit 0) but its wrapper is `NOT_EVALUATED`; the larger Suite Full / ValidationMode Affected run at e83 FAILED. The corrected Popup selected case also passed at source 1cea (3/0/0, native exit 0; wrapper `NOT_EVALUATED`). Source-726 focused UI controls are recorded in the 2026-11-14 follow-up (Europe/Paris); four broad desktop entries, hosted Full qualification, and ARM64 remain pending. |

### Qualification follow-up findings (2026-11-14)

| ID | Disposition / candidate status | Evidence and remaining limits |
|---|---|---|
| FOLLOWUP-01 | Validated validation-report root violation; report-free mode implemented | `Get-SpecInventory.ps1 -NoReports` retains specification/test-inventory reconciliation and failure policy without initializing or writing reports. The Static adapter and canonical runner both select it. The focused 148-pass campaign and real-checkout zero-BLOCKING result are recorded below; that run did not capture the nonblocking count. |
| FOLLOWUP-02 | Validated ViewerSpace capacity-fixture premise; test-only repair, native unexecuted | Holding every disposal callback prevented the fresh layout publication required by the old setup. Preparation now phases model-only retention, fresh publication and two charged retained owners, with no-expiry/completion witnesses. The earlier 9.9 ms Close observation followed invalid capacity setup and is not qualifying timing. |
| FOLLOWUP-03 | Validated live-drive-count assumption; test-only repair, native unexecuted | The shell case no longer assumes two machine drives. It checks copied live provider identity/order/value-presence parity; separate exactly-two-row display data retains actual renderer alignment coverage without creating drives or navigating. |
| FOLLOWUP-04 | Clipboard/foreground causes unresolved; metadata-only diagnostics implemented | Historical Boolean clipboard failure did not identify its transport stage, and later foreground reacquisition did not identify the earlier owner. Commands and hosted refusal diagnostics now retain the captured identities and setup/read metadata without contents/titles. Refusal, input, restoration, deadlines and one product-copy attempt are unchanged; diagnostic coverage is not a runtime repair or qualification claim. |

### Tooling evidence (not native qualification)

All paths below are beneath the canonically marker-verified `Z:\RedSalamander.Perf`. The tooling worker reports:

- `ResourceLocalizationContracts`: 12 passed with the added runtime resource strings in `Z:\RedSalamander.Perf\runs\20261113T150840Z-63848-c94ed3ab9d3c414cada0faa67490e1a6`.
- `ValidationFingerprint`: 29 passed; `ScopedTesting`: 11 passed; `MSBuildInvocation`: final 37 passed (earlier 33-pass result superseded); run folder `Z:\RedSalamander.Perf\runs\20261113T143855Z-75772-6701fa32993a4d98bd9440f874c68375` for the first two and `Z:\RedSalamander.Perf\runs\20261113T144432Z-60240-a35ed74c094645bcabb310a66a063518` for the final invocation result.
- `TestRunEvidenceArchive`: 26 passed; `ToolInventory`: 17 passed; `ToolingGovernance`: 12 passed in `Z:\RedSalamander.Perf\runs\20261113T144032Z-46488-f3b6a1584b434d6b96c66b4f692a4c01`. This run's combined Pester log records 103 passed/1 failed, including `HostedTestDesktop` 48 passed/1 failed; retain as a failed run.
- Directed `HostedTestDesktop` aggregate-exception/filtered-case check: 2 passed/0 failed in `Z:\RedSalamander.Perf\runs\20261113T144200Z-50748-6a44f065b919437b850ae16030753622`; this is not a full suite. The later complete `HostedTestDesktop` run passed 49/0/0 in `Z:\RedSalamander.Perf\runs\20261113T144531Z-2400-30bda3c35b234ddd8464a5ec3e1112f0`.
- vcpkg default-toolset/hosted parity: 9 passed, 0 failed, 1 skipped because the existing pinned-DxUi mirror prerequisite was absent, in the same `144531` run. The initial broad campaign's 88 passed/2 failed remain preserved in `Z:\RedSalamander.Perf\runs\20261113T143534Z-82584-3dbf6733d6cd4639ac3794d4b1ad496b\tooling-pester.log`; these distinct run results are not an overall green.
- The final compiler-free Pester campaign reports 263 passed, 1 failed, 0 skipped in `Z:\RedSalamander.Perf\runs\20261113T150840Z-63848-c94ed3ab9d3c414cada0faa67490e1a6`; the failure was a stale BUG-06 source-guard assertion. Its repaired affected describe reports 190 passed, 0 failed, 0 skipped in `Z:\RedSalamander.Perf\runs\20261113T151140Z-81084-91f0fb47fbb8418da6179ae8c7fc634b`; this rerun covers the targeted describe, not all 223 original source-contract cases. The 150840 run also records `ToolInventory` 17 passed and `ToolingGovernance` 12 passed. The preserved spec-inventory result from run `20261113T150840Z-63848-c94ed3ab9d3c414cada0faa67490e1a6` records zero blocking and 60 nonblocking findings; `Test-Changelog.ps1 -Base origin/main` reports 156 pending fragments. Tool inventory: 230/230, zero findings; modified PowerShell parser and scoped `git diff --check` reported pass. The newer affected run and focused repair outcomes are recorded below; Full/Release, focus-taking acceptance, and broader native/full-planner performance qualification remain pending.

## Contracts and validation

Confirmed/likely normative owners to reconcile per report row include `Specs/Core/Core_SharedHelpers.md`, `Specs/Core/Core_FileSystemBridge.md`, `Specs/FileSystem/FileSystem_FileOperations.md`, `Specs/FileSystem/FileSystem_MicrosoftDrive.md`, `Specs/FileSystem/FileSystem_S3.md`, `Specs/Plugins/Plugins_VirtualFileSystem.md`, `Specs/Testing/Testing_ValidationEvidence.md`, `Testing_SelfTests.md`, `Testing_ToolingGovernance.md`, `Specs/Build/Build_Toolchain.md`, `Build_UnifiedCI.md`, and relevant UI contracts. Runtime/tooling owners report these candidate updates; this ledger does not claim every spec closeout is complete.

Use `Test-Changes.ps1 -Explain`, then the derived affected scope and `Test-Changes.ps1 -Mode PrePush` under `Specs/Testing/Testing_ValidationEvidence.md`; run Full only where the owning contract requires it. Keep compile, focused execution, broader qualification, performance evidence, and hosted CI as separate states. The first x64 Debug attempt failed during build; the earlier Suite PR / Fresh build passed compilation but failed overall in Pester; directed Review/Identity cases passed with their wrappers still `NOT_EVALUATED`; and the latest Suite Full / Affected run failed as recorded below. Native qualification remains pending because that broader run failed and desktop, Release, and hosted Gates acceptance remain incomplete. BUG-09 paired evidence and broader native/full-planner performance qualification remain pending; the bounded PERF-02 component packet is not a whole-planner speedup claim.

### First native build attempt and reviewed repair

- At source `301c0dc56c5e56e987505055f433c59c96339a44`, `Tools/Run-AllTests.ps1 -Suite PR -ValidationMode Fresh -Platform x64 -Configuration Debug -TestRoot Z:\RedSalamander.Perf -NonInteractive` failed during the full-solution build after 9 minutes 12 seconds. Run `20261113T151944Z-81180-56a94db279ba42fc863448abc18e033d` returned 1; no native cases executed and no compatible build receipt was published. Its interrupted archive is preserved under `Z:\RedSalamander.Perf\evidence\completed-runs\20261113T151944Z-81180-56a94db279ba42fc863448abc18e033d`.
- The exact authoritative MSBuild log was copied and byte-verified at `Z:\RedSalamander.Perf\evidence\native-build-failures\20261113T151944Z-81180-56a94db279ba42fc863448abc18e033d\msbuild.txt`, with `failure.json` beside it. SHA-256: `56a8523415370c6cbecb8aed2b1bf7e3094af29629df2270a1545bb3ae7e094c`. This companion does not modify the original completed-run archive.
- Two distinct `C4365` locations in `FolderWindow.FileOperations.Popup.cpp` at original lines 9132/9134 produced ten warning/context rows. The reviewed repair retains the Move-only Delete-source tooltip and exact action-count clamp, using a bounded span/ranges lookup. `Commands.Tests.ViewCommands.cpp` at original line 30013 produced `C3688` for `2000ms` without a literal namespace, followed by `C2660`; the repair uses explicit `std::chrono::milliseconds{2000}` without weakening the focus assertion.
- Parent-reviewed hosted formatter feedback is applied to S3 Core/header and the FileOps runner. Pinned formatter dry-run/Werror and `git diff --check` pass for the repair. The later x64 Debug run built this source cleanly; its overall result is recorded below and was not a pass.
- The sole new power-control invocation now runs beside the ancestry controls under `FileOps_ReviewRemediationContracts`' existing activation blocker, before prerequisite-based skips. Later native evidence must contain all six `FileOps.Power.ConflictWaitCalls` controls, not merely a partial case result. The no-desktop admission lists are reconciled without changing classification; the existing focused plan-helper describe reports 66 passed, zero failed/skipped.

### Latest validation attempts and follow-up repairs

- The final committed scope check for `301c0dc56c5e56e987505055f433c59c96339a44` against `origin/main` returned the expected exit 2 / `NOT_EVALUATED`: 80 changed paths, full-solution advisory build, 14 affected noninteractive obligations, and four pending desktop entries (`Commands`, `FileOperations`, `ViewerPETests.Interactive`, `ViewerSqliteTests.Interactive`). Run `20261113T151710Z-67596-5b2a8dee90d04fbeaec3235b2ba1937c`; this scope snapshot predates the later source changes and is not a pass.
- At source `bfc468f0741597a9c2557b9e9015961e0ae10526`, x64 Debug full-solution build completed in 5 minutes 54 seconds with zero warnings/errors and published receipt `422cd6cb5b2c268c498ffdc0f76c7ce0b41a136509aa7de71190cf615908ef78`. The sealed runner packet is `Z:\RedSalamander.Perf\evidence\completed-runs\20261113T154000Z-73744-a59b7e0162084a63b7497591b087dbde`; the archived `run-all-tests-results.json` SHA-256 is `e301069ef452568221163c7617218159a18f52880ff2429f4862f831e20e105a`. The overall result is FAILED: 12 obligations executed, 11 promoted, and 1 incomplete. All 11 non-Pester entries passed; the broad Pester campaign reported 1,291 passed, 6 failed, and 2 unallowed skips. Do not count this as a green Fresh run or complete qualification.
- The tooling repair worker reports a focused repair suite of 105 passed/0 failed/0 skipped in `Z:\RedSalamander.Perf\runs\20261113T161815Z-76584-59ee9be7d76a4a1b9024dc0f25d11636\scratch\tooling-terminal-repair`; it covers the six failed rows, both formerly skipped baseline cases, BuildEvidence, and archived-proof tamper/freshness checks. A final inventory/governance guard suite reports 29 passed/0 failed/0 skipped in `Z:\RedSalamander.Perf\runs\20261113T161903Z-73984-cf4b6d3da6f7423d982cad27e4d081b6\scratch\tooling-repair-governance`. These focused reruns do not rerun the full PR Fresh campaign. Both baseline cases now resolve exact objects from the verified existing DxUi object store; nine exact pinned Git objects were hydrated while HEAD, refs, and `FETCH_HEAD` remained unchanged. No Full/Release/native qualification is implied.
- The directed Review and Identity runs at e83 passed their selected cases with setup/cleanup (3/0/0 each, native exit 0; canonical wrappers exit 2 / `NOT_EVALUATED` v1). Review exercised 21 ancestry, 6 power and 3 cross-root controls. Identity exercised the 48 selector pairs/96 timing rows/48 capacity rows; no whole-application speedup is claimed. Nonblocking selected-root disk audits were not clean: Review 100 and Identity 99 findings, including 19 unexpected root children, 8 stale entries and 72 manually created runs; ownership is unproven and nothing was cleaned.
- The x64 Debug build at source commit `e83dd755f44cef4f720d3f439f5fbd06383a569a` completed in 7 minutes with zero warnings/errors for source snapshot `bf6aed7cb9a8479f998dbde44717845600a8735933c4387f518d531e8100b04d`; receipt `2a421a6b4d303b142f68a56faf7b070fdeff32953a4803884fba1ecd0e105135`.
- The subsequent Suite Full / ValidationMode Affected run `20261113T165543Z-60300-679c095ca4fe474cbcaa5a8d5f233164`, launched through `Test-Changes`, took 42 minutes 16.6 seconds and ended `FAILED/FAILED`: 14 executed, 12 promoted, 2 incomplete, 0 reused/invalidated; totals 1,598 passed, 3 failed, 27 allowed capability skips out of 1,628. Pester reported 1,297 passed, 2 failed, 0 skipped; Compare reported 257 passed, 0 failed and 27 capability skips; its entry promoted. Standalone Curl passed; PerformanceTests2 passed all 15 cases. The sealed v2 result is `Z:\RedSalamander.Perf\evidence\completed-runs\20261113T165543Z-60300-679c095ca4fe474cbcaa5a8d5f233164\artifacts\selftest\last_run\run-all-tests-results.json`, SHA-256 `afb5ba3e5f1ed32c03fb3a70a04146c1b2b3dbc9ba42bf9dd3c110928cbcbe56`. Writer reported 34 passed; a 6-minute-44-second reattestation build had zero warnings/errors and receipt `ca202ebf467d9ed8cc16aed9d19c9672cc3bd425dfabe747306667716dadf835`, same source and artifact epoch 1. The earlier bfc `20261113T154000Z-73744-...` packet remains the distinct Suite PR / Fresh result.
- The failed checks were one invalid Curl PID-substring assertion (the cryptographic generator is correct; the concrete leaf/PID was not logged) and two stale constructor-positional source checks at original lines 4106 and 4613. The Curl/S3 debug-selftest repair now checks the existing hexadecimal-token grammar and valid fictional-PID digits, with malformed controls and an S3 basename-digit control; production naming, RNG and retries are unchanged. Native validation remains pending. Tooling separately identified a Ghostty Stream 6 capture gap: `Write-Host` run IDs were omitted. A deliberate red control against the unchanged workflow reported 26 passed/1 failed/0 skipped, confirming the missing run ID; evidence is `Z:\RedSalamander.Perf\runs\20261113T174610Z-86824-52500cba748a401eba43959e50666136\scratch\ghostty-stream-red`. The 6>&1 workflow capture and stricter source guards are implemented; the focused 281 passed/0 failed/0 skipped campaign at `Z:\RedSalamander.Perf\runs\20261113T174901Z-59144-9cc763e9ae184fff9a70c8931fefe13c\scratch\stream-source-contract-final` passed those fixtures but did not model the actual legacy Commands writer. The later source review identified and repaired the verifier's missing v1 stable-path proof support; the real no-desktop Commands acceptance probe remains pending.
- The final compiler-free tooling campaign reports 405 passed/0 failed/0 skipped in `Z:\RedSalamander.Perf\runs\20261113T180944Z-71464-6c5c63757a5148adb5691c62236739f2\scratch\final-hosted-v1-tooling`. It includes whole source-contract, BuildEvidence, MSBuildInvocation, HostedTestDesktop, inventory/governance and writer-faithful v1/proof negative cases. The 321 measured target hashes, including the parent-reviewed Curl/S3 files, were unchanged before/after; I65 reporting was excluded. The earlier 404 passed/1 failed fixture run remains retained. This compiler-free result did not replace the actual admitted Commands probe recorded below and does not establish Full, Release, or ARM64 qualification.
- Exact-head automatic CI/security/format jobs passed at e83 (CI `43465350387`: tooling, format and `ci-gate`; security `43465350077`: dependency review and CodeQL; separate CodeQL success). The aggregate Gates workflow remains in its expected draft/approval wait; hosted checks do not turn the failed Suite Full / Affected packet into native acceptance.
- At source `bb48b49ce582a8373c443095f185ddb4999a2cd3`, the full Debug build exited 1 after 5 minutes 33 seconds with 4 warnings and 2 errors; no native tests ran and no new receipt was published. `ItemResultReducer` controls referenced unavailable `PerfNowUs`/`PerfElapsedUs` (`C3861`). `SelectorSample`'s atomic array triggered implicit copy/move warnings (`C4625`, `C5026`, `C4626`, `C5027`). The frozen log and `build-failure.json` are preserved under `Z:\RedSalamander.Perf\evidence\completed-runs\20261113T163353Z-86432-a19c88b4caab442cbc05e75a6d947691\scratch\audit-native\full-debug-build`; log SHA-256 `ea258579d386ee8a1328078d242b107d0cf47d0b2ad904bade51fd95e0bdbe79`. The timer/lifecycle repair was recompiled and exercised at e83 as documented above.
- Exact-head automatic hosted CI checks passed on `bb48b49ce582a8373c443095f185ddb4999a2cd3`: CI run `43462989848` passed tooling, format and `ci-gate`; security run `43462989907` passed dependency review and CodeQL, with the separate CodeQL check also successful. The aggregate Gates workflow remains draft/approval-wait. These hosted results do not satisfy native acceptance.
- At source `3bf94cb022fda7bac66ae6bf9231e68058f3df99`, the first full Release build exited 1 after 7 minutes 32 seconds with exactly 3 `C5258` warnings and 0 errors. The required zero-warning gate rejected the build; no Release receipt was issued and no tests or controls ran. All three warnings came from a redundant `constexpr` prefix capture in the Curl grammar lambda (original line 891, columns 31/57/103). Its sealed log and failure metadata are `Z:\RedSalamander.Perf\evidence\completed-runs\20261113T181716Z-75840-2ef620a0307f4910be6203ea9d021753\scratch\audit-native\full-release-build\msbuild.txt` and the adjacent `transcript`/`build-failure.json`; log SHA-256 `cb669c4107b0649c0b46848887bf80bc881c1760d44f7102b51dc622333fadc7`. The narrow `[prefix]` to `[]` repair is applied with unchanged semantics; recompilation later passed at `535ffa40` as recorded below. Exact-head CI/security/format checks passed on this commit (CI `43480962482`, Security `43480961812`), but do not satisfy native acceptance.
- At source `535ffa40bd87d8b070e22cb2aa15fc70d038326a`, full x64 Release built in 8 minutes 3 seconds with zero warnings/errors and receipt `2354fa6e57b4050d001c295d11a1e6f305eb11c6cc8ca21853006cfe3f5bf1ea`; source snapshot `9e97ca50d1f2f6f142156fafd97fa20cc5337de2589cbf7f3eea60536151cbf2`. Directed `ReleaseReview` run `20261113T183701Z` and `Identity` run `20261113T184042Z` each passed selected case plus setup/cleanup (3/0/0, native exit 0; wrapper exit 2 / `NOT_EVALUATED`). Review exercised 21 ancestry, 6 power and 3 cross-root controls; Identity verified 48 pairs, 96 timing rows and 48 capacity rows. Selector results retained regressions: 10/12 cohort medians were lower, while two 64-task mixed-order cases changed 61→62 and 50.5→52.5 µs per 16 calls; six slower pairs, three ties and one zero remain in the evidence. No uniform or whole-application speedup claim is made. Three leased no-activation cases also passed selected case/setup/cleanup (Safety `20261113T184730Z`, IdentityLessGuardPartition `20261113T184952Z`, R3.1 IdentityLessReplace Dummy `20261113T185252Z`; each 3/0/0, native exit 0, wrapper `NOT_EVALUATED`). Focus-taking acceptance remains pending.
- The following `PopupStableHostedTree` run `20261113T185408Z-12252-992d0740671b43ed96d607fa07ef79ae` failed (2 passed/1 failed/0 skipped; native exit 1, quarantine 0). Sealed summary SHA-256 `a38caef000a5720bf2db9eec73a1376a041f05f9355b6ac45e8923a95e2f81c0`; trace SHA-256 `428a770aa5dfe83eebd0c71e9b7e69ef5c386d52965a16041a0ed24f7fef8c66`. Five assertion messages trace to two invalid fixture premises: only one finished record was created despite the production group requiring two, and the Move visual fixture never enabled `DiscardSource`. These results do not establish product regressions or a timing baseline. The test-only fixture/API repair now creates two finished tasks, explicitly publishes `DiscardSource`, covers fail-prerequisite ordering, and checks negative help. Its compiler-free follow-up passed 291/0/0 in `Z:\RedSalamander.Perf\runs\20261113T190714Z-73696-f098bd60dd0d4b70b025e3c346e37adc\scratch\popup-fixture-static-validation` (SourceContracts 223, desktop-inventory admission 1, directory-cache admission 1, plan-helper 66). The corrected Popup native retry is recorded below. Debug native PGT repair, PrePush, focus acceptance and ARM64 qualification remain pending.
- At source `1cea954f88a00514e8655016e1926665462ce44d`, the x64 Debug full-solution build passed with zero warnings/errors in 5 minutes 27 seconds; receipt `9697cf3b4c93f801718f5e807f3180aff59c6aae9b102a34776fa8ffe31086b4`, source snapshot `3c98ae018a904a80c987a4e92ac06363eccd1ed577f38e56f4bcbe305a11381d`.
- The actual Commands v1 probe `cmd_directory_cache_invalidation_during_enumeration_is_retained` passed (1/0/0, native PID 90132 exit 0; wrapper exit 2 / `NOT_EVALUATED`). `Assert-RSHostedFilteredCaseRun` accepted the actual archived-v1 proof, including source/receipt/PID/time/native bytes; diagnostic Change is PASSED and repository coverage remains NOT_EVALUATED. Sealed run `20261113T191928Z-73296-6588d04ac0b143d9ae2ecddcb75269d8`; archive-intent SHA-256 `d2271bf673f4ddd80af1f75f7fe98662f93c0a32e2111208278bde86d53c8522`. Outer proof evidence `Z:\RedSalamander.Perf\evidence\completed-runs\20261113T191928Z-73296-15ad1cae3b3046d3838d8968ef396d8a\scratch\audit-native\debug-commands-real-v1-proof`: captured-output SHA-256 `040627245d3d9581c498315aa6153a1cd9caf3d6684fb1fcc32ae53fdff5ec10`; proof SHA-256 `ae61d55af848938a9a90dada6042a6709fd4c98c245d289dce3ae89a5f2ee205`.
- Corrected `PopupStableHostedTree` run `20261113T192051Z-77112-d5e4b285e7be4af3b0bbccee802d9d46` passed selected case plus setup/cleanup (3/0/0, native PID 13408 exit 0; wrapper exit 2 / `NOT_EVALUATED`, lease/no-activation, quarantine 0). Retained audit count 101. Quarter-wheel metric recorded 4 messages, result scroll 18, hr=0; trace SHA-256 `16864c954543721c13c90da23a60927106e2a374f6633a360477c776c7707297` confirms both group-focus targets, quarter-detent, DPI-state and one-pixel-resize assertions plus Move-positive/default-false/Copy-negative help assertions completed. No separate DPI/scope metric exists. The earlier 535 Popup failure remains failed history and is not a timing baseline.
- Exact-head automatic CI/security/format checks passed at 1cea (CI `43490007220`, Security `43490006949`; separate CodeQL success). The draft Gates workflow is not native acceptance. Still pending: Release Popup refresh, PluginContract entropy checks through governed PrePush, current PrePush accounting, genuine focus tests and ARM64 qualification.

### Published native component evidence

- `Specs/TestRuns/4cb089111a23/FileOps/2026-11-13_184357_AuditNativeComponents/` contains 108 files totaling 2,092,660 bytes. Explicit `RunPath` validation passed with 108 files; required whole inventory validation passed with 18,647 files. All 108 copied/staged file hashes and lengths matched with zero mismatches; 216 text/whitespace attributes are unset. The 105 original reviewed evidence files are unchanged; the publication adds an authored README, provenance and manifest.
- Publication gate logs/JSON are in `Z:\RedSalamander.Perf\runs\20261113T192527Z-63900-d2d119904a52492e810b397e3b8d8a30\scratch\NativeComponentPublication`. Copied manifest SHA-256 `9ed56edb331a9dac20b13eda07617a5aad6bf303fd7941909f859fad3c28a1ea`; publication manifest SHA-256 `0d0dece6ee34470c82ea93ec8193f99868615fd203bc6222f082816920aa4f4e`.
- The archive contains source-specific e83 Debug and 535 Release evidence, all 48 Identity selector pairs with timing/capacity results and retained regressions, plus the directed controls. It does not establish current Full/native qualification; the e83 Suite Full / Affected run remains FAILED. Five raw SID-bearing intent records remain private. Public labeled selected-member projections have independent recipe/peer semantic review, but the generic gate did not provide dedicated schema validation for those projections.

### 2026-11-14 follow-up (Europe/Paris): source-726 Full run and native guard refinements

- At source `7266562634f1466117be55b52aed911e9482de13`, the x64 lane of hosted [Full Release run 43499066444](https://github.com/RedSalamanders/RedSalamander/actions/runs/43499066444) built the full solution with zero warnings/errors. Its Full/Fresh v2 result `20261113T204736Z-2376-b5d742961bb24e559371a761f67cf3d0` ended `FAILED/FAILED`: 18 executed, 16 promoted, 2 incomplete; 2,854 passed, 2 failed, 69 skipped. The sealed summary SHA-256 is `131733442a7642322941236db2acc54fc5d556c20ec9984f9b44f1ff967c019d`. The two reported failures were reviewed as fixture/guard drift, not established product defects; the test-only repairs described below were unbuilt and unexecuted at this record.
- The exact Managed SMB new-child control was skipped because `path.smb.primary` was unconfigured. A separate rename-merge control passed, but it does not exercise the skipped scenario. The follow-on assertion probe did not run after the suite exited 1. The upload did not include manual-descriptor or full-build-receipt bodies; no offline, portable, main-branch, or production-attestation claim is made. The later ARM64 terminal failure is recorded below; subsequent source-bound results are tracked in PR #225.
- The focused UI window on 2026-11-13 UTC ran each of three Commands cases once in x64 Debug and Release (six launches): `cmd_fileops_artifact_name_shape_projection_scaling`, `cmd_pane_navigationView_unfocused_pane_click_focuses_target_pane`, and `cmd_pane_folderView_activation_press_keeps_selection`. All six passed (1/0/0; native exit 0; filtered wrapper exit 2 / `NOT_EVALUATED`; quarantine 0) in 7m21.14s. Outer evidence is `20261113T221009Z-89580-47e61cb7ff89426fabb75f00151efe64`, terminal SHA-256 `562bef19f901c079006d22e58d4a83e1163c997528bbc87eff23d217cdd4a199`. This evidence predates the test-only refinements recorded below. Four broad local desktop entries (`Commands`, `FileOperations`, `ViewerPETests.Interactive`, `ViewerSqliteTests.Interactive`) remain pending; no final focus/cursor-restoration metric was captured.
- At this record, two test-only refinements were unbuilt and unexecuted; subsequent source-bound results are tracked in PR #225. `FolderWindow.FileOperations.Tests.CrossProviderRoutes.cpp` requires zero errors after recovered-mismatch Skip and reads the destination occupant before asserting its preserved bytes. `Commands.Tests.PluginConfig.cpp` replaces a stale English-literal needle with a function/branch-scoped localized resource-ID guard covering Managed cleanup's no-prompt retained-source path, sequential/parallel propagation, and typed Retained disposition. Neither change alters product or resource behavior.

### 2026-11-14 coordinated report and fixture closeout

- Historical source `7266562634f1466117be55b52aed911e9482de13` hosted run `43499066444` finished FAILED on both architectures. ARM64 built with zero warnings/errors. Its actual native Commands payload reported 960 passed, 6 failed and 5 skipped; the hosted Default-boundary retirement exception then aborted outer aggregation. Canonical run `20261113T210134Z-7668-7d2bd0dac0bd4a83b4d0215e8e9b0f60` retains 18 registered entries, 12 promoted, Commands incomplete/running and 5 pending. There is no final Full v2/v1 summary, ARM64 FileOps execution or post-suite assertion probe. The exception refused to replace unexpected foreground ownership; the precise owner at that instant is not proved by the later `hosted-compute-agent` snapshot. The six native rows include the known stale source guard, Batch Rename clipboard copy, two pre-body foreground-settling refusals, the drive-row assumption and ViewerSpace disposal setup. This is not six validated application defects.
- Both architecture artifacts and raw logs are preserved in monitor context `20261113T221137Z-59744-b7651372f6f249b39ac4690ead7fad02`, `scratch/hosted-full-release-72665626-continuation`. Packet-index SHA-256 is `30aa9a752a9e7390a1569808aaa3f32604bdd8ee12adc73dc5241aef48b7f50e`; ARM file-inventory SHA-256 is `02c8ff445ce54d4e503b834ae024691d6fb7178e0df8326c4942ae338ba9aafa`. The original dispatch context is preserved through its canonical completed-run archive. Neither upload includes the actual frozen descriptor or full build-receipt bodies. Partial canonical bindings, synthetic Pester scratch summaries and successful uploads cannot establish Full or offline portable acceptance.
- Historical source `11c2efba764d1bce3292fbcd389991890155367b` subsequently built x64 Debug and Release with zero warnings/errors, source snapshot `f4ab0adf3e778dea714af13dd31665d39de1cc3c932bd65376922f1fb9c7b05b`; receipts are `abe8ba330a320465e7c9e17356dc8a661c2e6d2f95d64c41e7aa14aa9d1711e0` and `ac81068fac451bc73addd0b0c13841b6d94155418affac966c52c4447326824b`. Repaired MTP selected runs `20261113T232313Z-67580-bee223056347486ea4d757a57cb9c618` and `20261113T233053Z-14512-fc0d05546fe546c08bb88929c9559e5d` each passed selected case/setup/cleanup (3/0/0, native exit 0, filtered wrapper exit 2, quarantine 0). Combined proof SHA-256 is `3e9e62030964b8f9d9686399ce3549bcaa77d2775711ac975ebdf9c1bc2506a6`; nonclean audit 101 remains retained. These builds and selected passes predate the combined changes below and do not qualify that later source.
- Report-free specification validation adds `-NoReports`; both actual validation callers preserve the same validation/result handling while producing no reports. The source-11c2-based focused campaign at `Z:\RedSalamander.Perf\runs\20261113T235715Z-78672-672779876e72418c9ca41bfc6980ca35\scratch\SpecInventoryNoReports\final` passed 148/0/0 (142 complete governance/spec suites plus 6 selected actual caller controls). Twelve measured target hashes were stable. The real checkout reconciled 505 specification files and captured **zero BLOCKING findings**; it did not capture the nonblocking count. Existing real `.build/SpecInventory` existence/bytes were unchanged, reports written 0; tool inventory also passed. Summary SHA-256 is `cbe9113e742034258f3fc24de04bc4a25cfc7a4b65e75047fd499896fc197c4f`. The initial 8-pass/4-fail fixture attempt and repaired 12-pass retry remain preserved; fixture controls still report their real protected-history findings.
- Hosted refusal enrichment passed the whole existing compiler-free suite, 55/0/0 in 16.835 seconds, at `Z:\RedSalamander.Perf\runs\20261114T000853Z-46952-42aa029da09a42f78e5dd469621f2fe8\scratch\hosted-refusal-observability`; five measured module/test/support hashes were stable. These are fake/managed controls, including distinct identity roles, unavailable diagnostics and fatal diagnostic exceptions that are not suppressed and follow existing owner-thread propagation. No live desktop call, native execution, ownership/input/restore policy or deadline change is claimed.
- The ViewerSpace test-only follow-up phases model-only retention, fresh display/layout/presented publication and one completed resize with two held owners. Appended `refreshDisposeBytes` observes submitted/awaiting-submission charge separately from pending canceled-builder charge; all later barriers require exact bytes/generations and no completion/expiry. Scoped static review and pinned formatting passed, but this repaired native scenario is unbuilt/unexecuted at this record. The old source-726 9.9 ms observation is invalid as capacity/Close evidence.
- The drive-menu test separates live provider parity from exactly two synthetic display rows and preserves actual renderer, ownership, dismissal, bitmap, shell and focus checks. Commands window/clipboard observations are metadata-only, use the same captured HWND for comparison/logging and expose unavailable fields; Batch Rename requires clear/quiet before one copy attempt and omits contents. Pinned formatting and bounded static invariants passed; native execution of these changes remains pending. The durable `Commands boundary observation and clipboard source contracts` describe in `Tools/Tests/TestHarnessSourceContracts.Tests.ps1` passed its four compiler-free controls (4/0/0, 883 ms), with three native-source and guard hashes stable. Evidence is `Z:\RedSalamander.Perf\runs\20261114T003209Z-69308-20d38860958e4de78d341ff59a6250e5\scratch\tooling\commands-boundary-global-call-guard`, summary SHA-256 `7279adce67e86e3e414f2f32272c47f94463c06fbe76fc2d75bd94fd593be133`; the earlier two four-pass results remain preserved. These controls enforce API/privacy allowlists, same-HWND use and the single-copy prerequisite/return contract; they do not establish native behavior or update the earlier source-726 UI passes.
- I65 remains ACTIVE. The six selected focus-taking source-726 UI passes remain historical; the user's 15-minute foreground window expired at `2026-11-13T22:24:07.7201508Z`. The four broad local desktop entries (`Commands`, `FileOperations`, `ViewerPETests.Interactive`, `ViewerSqliteTests.Interactive`), current combined native/platform qualification and unresolved SMB prerequisites remain pending. No further foreground time, Full pass or ARM64 acceptance follows from these component/static results. Subsequent source-bound results belong in PR #225.

### 2026-11-14 source-084 qualification and bounded attribution follow-up

- Source `084085b367f7905e74d98b3252f83ccf68bd329d` has three local full-solution builds with zero warnings/errors. Its Debug PrePush run `20261114T005548Z-76952-ff85bbb845c342c0943817a26e2257d7` is Full/Fresh **filtered**, typed exit 2, Change PASSED and repository NOT_EVALUATED: 14 executed/promoted, 1,619 passed, zero failed and 27 declared Compare capability skips. Writer coverage includes the actual native family-union control; broad Pester passed 1,317. Exact repaired MTP Debug and Release runs each passed selected case/setup/cleanup (3/0/0, native exit 0, filtered outer exit 2), with final metric state 2/0/hr0. Combined native proof SHA-256 is `1e7af7a874ce1022c9b8de5daf9d78a3d56e4ea1d444784d26ec011b292705e6`, canonical context `20261114T020320Z-94392-864afc7f1d9541b180162b7f9895353e`. Combined compiler-free controls passed 570/0; the actual specification scan covered 505 files with zero BLOCKING findings and 60 informational findings. These source-specific component results do not establish Full or later-source acceptance.
- Exact-head automatic CI/security/format passed at source 084. The single approved test-only [hosted Full Release run 43544978035](https://github.com/RedSalamanders/RedSalamander/actions/runs/43544978035) then **FAILED on both native architectures**, each after a zero-warning/zero-error build. x64 canonical run `20261114T030412Z-7736-9f40edcaa6f245ba85842c51deb2a3ba` has complete Full/Fresh v2 FAILED/FAILED: 18 executed, 17 promoted, one incomplete; 2,875 passed, one failed, 65 skipped. Its sole native failure is `FileOps_PopupIdleCadence` (paused Copy: 2 popup timer ticks, 13 renders and 2 owner ticks); the separate `validation_acceptance` detail is not a second native failure. Commands passed 966/0/5 and the interactive standalone entries passed. The later assertion probe did not execute. Summary SHA-256 is `262069c8ee6b5a67f6ac47b6226b346a1e479ebb0ace356a6cdd096bfb08f878`; raw FileOps perf SHA-256 is `e3f84de700179050840dfbaa1147a950c764e485c55b1950dd1ff6687148e4ae`. The render-burst cause remains under investigation; no threshold, timer or product-policy change is justified by this count alone.
- ARM64 canonical run `20261114T030036Z-10376-933320090bf647dcb0ff10718d18fd3c` has actual native Commands 963 passed/3 failed/5 skipped. Failures are `cmd_preferences_dialog_keyboard_reset_live_dx_interaction`, `cmd_pane_fileops_popup_throughput_disclosure`, and `cmd_app_menuBar_persistent_view_to_plugins_hover_switches_popup`. Commands PID 9152 exited 1; subsequent hosted retirement refused another captured foreground and aborted outer aggregation. State remains 18 registered, 12 promoted, Commands incomplete/running and five pending, with **no final Full aggregate, ARM64 FileOps execution or assertion probe**. The Preferences refusal sampled a distinct CASCADIA HWND/PID 8760; retirement later sampled PID 2660, while the original Progman PID 7584 remained unchanged. Those observations do not establish a common actor, parent/job ownership or a product defect. Native results SHA-256 is `82b13d338ae7da381248790639eb248373cbe162e1d2b2a4d5fe7f8547c9b24f`; trace SHA-256 is `0754dfa6d3644ff2a0408fb3e2eaabd8caf43ee28c4d86b41c17423efd7f2a93`.
- Both architecture ZIP digests match GitHub artifact metadata. Frozen per-file inventories cover 77,879 x64 and 97,146 ARM64 extracted files; hashes are `bbf89c635f14004b951c68ef1b360e27c0111ce9bc24fd9d97a99bef917d1af9` and `cc12d20a1fccfdb15e152e3541a1e39ceb506ede48ce31f8e7f9c3498325d463`. The relative-member packet index is SHA-256 `3606a0dde2cd1fdeddc91a12018c4b2273d9cc7a3110b1ab985dff5f9a7dbc11` in monitor context `20261114T021220Z-65620-b7ff68eee6b74657913361ead3ae66eb`, `scratch/hosted-full-release-084085b3`; resolve its final location through the canonical archive record. Existing upload paths omit the actual workflow descriptor and full build-receipt bodies, so receipt IDs and source-bound payloads do not establish offline portable/main-production acceptance. Historical source-726 failure and local/UI evidence remain separate.
- The uncommitted follow-up adds diagnostic-only process enrichment with query-only held identities, per-API errors and exact UTF-16LE SHA256 image-path digests; minimal original creation identity is captured at hosted entry. Fixed image categories, unavailable exact-job identity, bounded parent creation compatibility and HWND ABA limits remain explicit. No raw path, arbitrary basename, command line or title is logged, and no metadata grants close/input/restore authority. Native fake controls run before unchanged assertions in the existing `red_salamander_help_lists_diagnostics_options` case. Managed/fake focused controls passed 26/0/0; native compilation and execution of this new metadata remain pending.
- Popup and persistent-menu fixtures now fail a lost positive-focus prerequisite before Enter, require a live owned/readable menu session before first hover, and retain bounded audited menu metadata with prior TLS sink/target restoration and capture-loss accounting. Existing input/focus assertions, deadlines and background-negative semantics remain intact. These are unqualified local test/diagnostic changes; they do not fix an established application defect or explain the historical clipboard failure. The exact `FileOps_PopupIdleCadence` case alone joins the reviewed no-desktop list; broad/family/prefix/multi-case/casing-near-miss and proof-injected selections remain leased, with no new local foreground authorization.
- The fixed `CommandsUiAttribution` route runs the observer-help prerequisite then the three named UI cases once each with independent canonical evidence, unchanged receipt bytes and strict v1 native proof. The review-discovered extra-repair risk is closed by shared `Assert-RSTestQuarantineClean` and runner `-RequireCleanQuarantine`, enforced before native work and at actual repair dispatch; no alternate ledger or default behavior is substituted. Canonical missing-date validation now reaches its intended empty-to-invalid-state path. The seven-target focused correction passed 31/0/0 with stable hashes in `20261114T062226Z-75672-58423897dc1b4628a1b182816a7a83c6`, `scratch/tooling/commands-ui-attribution-controls`; earlier 22/1, 0/1 and 30/1 development failures remain preserved. Independent source review accepted the bounded route correction.
- The cadence follow-up is diagnostic/test-only: default-off bounded scalar observations sample actual pinned graph/host debug state, retain loss and reentrant-query availability, and emit only after measurement. A synthetic nonzero-running-to-paused transition invokes the actual reducer while preserving graph identity/history. Existing 600 ms/2,200 ms spans, ceilings and resume assertions are unchanged. The three-path patch passed independent source/format review; it remains unbuilt/unexecuted and establishes no cause or animation fix. The monitor packet is now canonically sealed: complete-record SHA-256 `7cc192c955e832ca766556c5294d1af7ff27a92f76d9f34a5a50c592702edbf0` binds its unchanged packet index and raw evidence. Current combined native/platform acceptance, the four broad local desktop gates, SMB and applicable I60 NAS obligations remain pending; I65 stays ACTIVE and no additional native run or foreground window is authorized by these controls.

- Subsequent dirty-28 Debug build completed with zero warnings/errors and receipt prefix `47dac8`, source snapshot prefix `89fef7`. The exact no-activation cadence Explain passed, but canonical runner `20261114T070511Z-4236-770e331975cf4d97a751ec9cc0a92bf3` failed admission because the `TestRunPlan.psm1` facade omitted the owning module's `Assert-RSTestQuarantineClean` export. **No native cadence case or callback launched.** The interrupted archive and outer observation proof SHA-256 `1e1cbb3ca2b6553d3ba1f8be8f52ffa7c03008d0fe3f8d11e788283c96b19507` remain preserved. This genuine facade defect is distinct from the earlier private-driver setup failures. The narrow correction re-exports the canonical assertion and adds a fresh-process actual-facade reader/assertion control; the older 28-path build receipt does not qualify this changed source.

- Subsequent dirty29 x64 Debug and Release cadence observations passed setup/selected/cleanup (3/0/0 each, native 0 / filtered runner 2 NOT_EVALUATED) after zero-warning/error builds. Snapshot 8ecc88a82986aa64ba3f2e9fd6fb27641367f704a25f3f0ca0f70ea81fbd9018 and manifest ce5789c5b8cfe5cbd531f00da35e9b7a654136c5aaed2ec4a4aac3d58897346f remain measured identities; receipts ce7bec/e76fc belong to that uncommitted source, not later publication. Real paused 2201/2203 ms samples each had 2 ticks / 2 renders / 2 owner ticks; all 147 rows per profile had loss 0 / reducedMotion false. [The component packet](../../TestRuns/4cb089111a23/FileOps/2026-11-14_074108_PopupCadenceDiagnostics/README.md) preserves exact selected derivatives and explicit custody projections. Graph/marker settling is observed; old source-084 13-render cause is neither proved nor fixed. Blank raw commit, observer overhead, nonportable receipts, prior failures and pending broad/platform/provider gates remain explicit. I65 stays ACTIVE. Publication changes documentation/workspace identity without rewriting measured proof.

- At committed source `8539231efca0259577fa80147bf48f1ec9a906ed`, canonical noninteractive PrePush `20261114T083406Z-84472-74cf2ba142584f9b87ed14c1edc1aba1` ended **FAILED/FAILED**, typed exit 1: 14 entries executed, 13 promoted and one incomplete; 1,652 passed, one failed and 27 declared Compare skips. The sole failure was the stale closed nightly-scope assertion in `Tools/Tests/BuildReproducibility.Tests.ps1:754`, which omitted the already approved `CommandsUiAttribution` choice; broad Pester reported 1,350/1/0. Initial and post-writer full x64 Debug builds had zero warnings/errors, the writer controls passed 34/0/0, and all other executed entries passed. Final sealed proof SHA-256 is `59bef291852c6ab8db40eb57ff19cab89cba494019820532ac3cd5888ee99961`, canonical custody context `20261114T091847Z-24908-939b7143e41f4d66b798230d05b24601`, `scratch/audit-native/853-prepush-final-proof/final-prepush-proof.json`. The failed row and source/workflow hashes remain preserved in the completed outer `20261114T083356Z-88968-e293301800f4483baddcf777a2ef2927` packet. The bounded correction updates the strict case-sensitive scope/platform/default/forwarding guard with negative controls and aligns three active usage/spec passages with UnifiedCI; it changes no workflow, pin, native or product behavior. The published cadence packet, measured dirty29 manifest and private proofs remain immutable; its at-publication source closure is not a claim that later usage/test edits match that snapshot. A corrected candidate requires fresh exact-source validation; I65 stays ACTIVE.

### 469 ARM64 stock refusal and captured DxUi identity follow-up (2026-11-15)

- Exact-head CI/security/format passed at `46939438b3a9101a73faec539bb4eaec93176bb7`; noninteractive PrePush remained source-specific (1,672 passed, zero failed, 27 declared Compare skips; 14 promoted, Change PASSED / Repository NOT_EVALUATED). Test-only ARM64 Release attribution run [43641993393](https://github.com/RedSalamanders/RedSalamander/actions/runs/43641993393) remains **FAILED** after a zero-warning/error build. Help passed once (native 1/0/0, PID 10868 exit 0), and strict v1 proof accepted trusted timeout 2. Preferences was refused during stock preparation before native launch; the final two selections were unexecuted.
- The new held-instance observation records terminal-host creation at 22:33:17.2274403Z, between the first child's return and the next selection's launch, after logged build completion. The exact captured HWND/PID/thread relation remained stable, creation matched the borrowed lease, and parent creation was compatible; parent image/session/job and creator/payload identity were not captured. Any-job membership is not exact owned-job membership. These facts establish neither creator nor dismissal authority. Raw official logs, three ZIP digests, 76 extracted-file hashes, 34 original member joins and separate late-control custody are immutable in completed context `20261114T211621Z-64560-f7df089a12164daba555830b88e4c1a1`; terminal proof SHA-256 `d3f03eebfb2cb2f78e602ae639e1b9864366d3e7935498cdde165567a3ffa0bd`. Missing receipt/descriptor bodies retain the offline artifact-byte-verification limit.
- A separate source-policy review found four bare read-only Git calls in `DxUiDependency.psm1`: HEAD, clean status, public-header tree and post-checkout status. The hosted image manifest lists PowerShell 7.6.10; its tagged native-command implementation can attempt hidden console allocation and foreground restoration for bare CUI invocation. This is a conditional launch-policy gap, **not proof that these calls created the observed terminal**. The candidate routes all four through the already imported canonical captured no-window launcher, with first actual Git Application resolution, literal argv/explicit CWD, separate silent streams, actual exit and no artifact-write delegation. Pin/dirty/cleanup/environment rules and hosted refusal authority stay unchanged. The build/restore-only consumer compiler-host MSBuild query is outside this four-Git change; no whole-module native-launch coverage is claimed.
- Local Full Debug at source 469 was deliberately retired before all four interactive entries: zero foreground entries executed, 11 entries promoted, writer 34 and broad Pester 1,370 passed, build zero warnings/errors; Compare retained partial 256/0/27 without a final native exit or promotion. Release Full did not launch. Sealed retirement proof SHA-256 `4adc9f4659f17ab1d84a4237544616c32fc4b84a3116259e32d3250ff124ce2f` remains in completed context `20261114T215807Z-111204-52be9170f3e64f3f8a572f6184cdcdef`; no Full acceptance is claimed. Selected source-469 cadence/SMB prerequisites and source-c55 six focused UI passes remain historical source-bound results.
- The new candidate-only Git controls preserve development parse/mock/recording failures before correction. Actual safe children require no-window/redirected PSI before launch, exact contained ownership, literal argv/CWD, actual exit, Unicode/separate-stream integrity and inherited environment; no unsafe bare-Git baseline is executed. The compact [DxUi Git capture component](../../TestRuns/4cb089111a23/Testing/2026-11-14_225602_DxUiGitCapture/README.md) records bounded work and source/custody rather than a speedup or native/platform verdict. Current-source PrePush, hosted ARM64 and broad Debug/Release desktop qualification remain pending. I65 stays **ACTIVE**; exact future Done-path admission neither moves nor completes the plan. Historical actor/clipboard/13-render causes and applicable broader provider/NAS obligations remain separate.

- Final frozen candidate compiler-free controls passed **325/0/0**, with all nine non-reporting target hashes unchanged: selected DxUi behavior 21, whole source contracts 237, build/cache identity 8, inventory 17, governance 12 and documentation drift 30. Canonical context `20261114T231409Z-87496-d44e7ff3500b41a68b7319f3a7407328`, `scratch/tooling/dxui-git-final-cf/summary.json`, SHA-256 `12090612704bc6122cfd4cad42b51fda7eef013dc7038ba2970c07396cb862eb`, retains the three exact mock queries and five guarded real child records. Seven source/result/work members have verified unique original inventory membership and completion-to-intent binding. Earlier development 5/3 and 7/1 results remain preserved; the final recording correction stopped a test-local re-import from replacing its interception after the first child. No production helper behavior was changed to accommodate that fixture. SDK/MSBuild profile evaluations, product builds, local UI and hosted execution were excluded; this is component evidence for the dirty candidate, not later-source acceptance.

### C receipt republication failure and idempotency correction (2026-11-15)

- Source `32ab3db627dcddb4afb844de9bf269d00e937146` compiled with zero warnings/errors, but its canonical PrePush `20261114T234241Z-107648-fd1dc4768feb4ed0838fea0316c607ea` failed during initial receipt publication before the later test entries. It attempted the same receipt ID `58dbae623e6cd78007883ebc11fe60a63546fc0caf7876910b33ef0ce755b272` as the accepted recovery build, while the publisher required full-record equality. Governing identity intentionally excludes timing and absolute diagnostic fields. The preserved disjoint build intervals prove fresh timing metadata must differ; the incoming receipt body was not persisted, so a complete field-by-field difference is unavailable. The original immutable receipt and all 188 current output bytes remain verified. Forensic proof SHA-256 `c35cf93d3f169e749dadd73e0faf793b736fdb4d75939be81e21de605150e8e3` is retained in completed context `20261114T234618Z-110228-346259254dff41a38ebcbc529f558e4d`, `scratch/audit-native/C-publication-forensics/proof.json`. C is not a passed PrePush or Full qualification.
- The narrow candidate validates the complete incoming canonical object/schema and receipt semantics, validates the stored canonical record, compares full governing projection text, and republishes the first stored bytes without replacing their original metadata. Initial immutable writes use the existing create-only primitive. Pointer full-record/byte equality, source/profile/toolchain/output/closure checks, canonical path/bounds and portable authority remain unchanged. No schema/API, native, workflow, pin or desktop-admission change is involved.
- Regressions were added before implementation. The genuine same-identity/new-interval failure and six schema-specific candidate controls retained a 10-pass/7-fail red run in completed context `20261114T235010Z-44908-f1e121526e1c48ba95c7d522a6ff4613`, summary SHA-256 `92f40941b2bfb719e835ba6485fc77fa5696ef95f0310622e989dfc7df175f94`. The old full-record comparison already refused those six malformed repeat objects; they were not six prior unsafe acceptances. The earlier 0/17 attempt also remains preserved; generic Pester-3 throw assertions required the existing actual-action helper and are not 17 product defects. The corrected implementation passed all 17 focused controls, including retained first bytes, malformed candidate fields, governing mutations, a simulated collision boundary and unsupported/corrupt stored bytes. The controls do not claim an injected concurrent-publication race. These are dirty follow-up compiler-free component results; fresh committed-source PrePush, ARM64 and broad desktop qualification remain pending. Existing public packets and all prior failures remain immutable; I65 stays **ACTIVE**.

### D broad strict-scope fixture failure and E module-ownership correction (2026-11-15)

- Exact D `0d71dd66c2ca590307f52defdf60c65ce1ea536c` completed its admitted canonical noninteractive run `20261115T001034Z-103392-0dac7c6fba474215899405fb3aaeee6e` as **FAILED/FAILED**: 1,697 passed, one failed, 27 declared Compare skips; 14 entries executed, 13 promoted and one incomplete. Broad Pester reported 1,395/1/0. The sole failure was `DxUiDependency.Tests.ps1:115`, before sandbox or real Git fixture launch: PSModuleInfo has no `InstanceId` property. The exact reader's 188 outputs were verified; the first real repeated receipt publication independently passed. The v2 and v1 terminal summaries are late control records retained through the separately sealed final-proof hash custody, not members of the original native archive intent. Final proof SHA-256 `0ef5638a937ee65ea0c3df7ffd5ba213af4d43b48880da87ee6011ec2fae4d69` is retained in completed context `20261115T011245Z-75856-aa6dc39fa2044ca6919368d4b5817f08`, `scratch/audit-native/prepush-D-final-proof/proof.json`.
- The earlier isolated module-metadata assertion silently compared null with null outside strict scope and was vacuous. The five real no-window/captured/contained child records remain bounded actual observations, but that missing-property comparison provided no independent module-instance assurance. Existing public C/D component bytes are not rewritten or relabelled. The broad profile's strict entry exposed the test API defect; no production Git-capture or native UI failure is established by this row.
- A strict-entry-only reproduction retained 7 passed/1 failed with the same missing property before any fixture child in completed context `20261115T011353Z-63636-9122d8d3485541e199e33d0c7a247b2d`, summary SHA-256 `9918fefb2e498cbd92f2cb2ace5594bf193fe17978a27785971d68e300431230`. The test-only correction resolves the canonical Script module in the consumer session, checks type/normalized path and real export-owner object reference, and refuses missing/lookalike/wrong-path/foreign-owner/fake-export/ambiguous states. It never treats caller Pester-proxy ownership or module GUID as instance identity and does not reimport over the hook. Initial strict green passed 8/0/0 with all five actual safe starts; final filtered-profile/source-bound controls remain separate from committed native acceptance. Production helpers, workflow/pins, desktop authority and earlier packets are unchanged. I65 remains **ACTIVE**; fresh corrected-source PrePush, ARM64 and broad desktop/provider/NAS/W5 obligations remain pending.

### F local metadata launch-policy correction (2026-11-15)

- Source E `8a48bb7a621a0033245f2670276b0248109dc32e` contains the strict module-ownership correction; no committed-source native qualification is claimed for E. A separate bounded source review found seven local bare VSwhere/MSBuild fixture probes and the production consumer compiler-host query. Parent no-window flags do not govern PowerShell's nested bare native-command route. The follow-up factors one private captured backend already using ProcessStreaming and shares it with the thin four-query Git wrapper and these local metadata readers. Literal arguments, filesystem CWD, actual exits, supported host selection, JSON/definition parsing, separate silent streams, parent environment, locks and no artifact-write delegation are preserved. No native, workflow, pin, interactive-build, public API or admission policy changes are involved.
- The pinned upstream `ConsumerBuild.psm1` identity helper has separate bare Git/MSBuild probes and no capture callback. It is not executed by the local restore fixture, which supplies its own benign synthetic helper and mocks actual restore work. Complete build-launch closure remains pending that separate disposition; neither this correction nor age/class/parent compatibility identifies the historical terminal creator.
- Source-only pre-fix control retained 0/1 in completed context `20261115T013120Z-68400-5922bfa7a9ca4020bd12bf019140e045`, summary SHA-256 `364c25dffe300407428193638ed14426fd2ecdd417bda8c7504c46d8918ac8e2`. Initial mocked development retained 13/1 because a diagnostic-message regex did not span a newline; the actual nonzero rejection was correct. Corrected strict Git/metadata controls passed 15/0/0 with all five guarded actual Git starts and unsafe metadata-PSI refusal, summary SHA-256 `f845c8452585a1e455556372259848ac26e3916cd25d4c812363c9f83ed9e858`, completed context `20261115T013858Z-89252-c9326136cb09442db4d5a5726ffa0c5a`.
- The first guarded metadata campaign retained 2/1 in completed context `20261115T013910Z-72700-a7070c5bc76e4087bf28d73948abf0a0`, summary SHA-256 `28f96a8f1bc9c7517760abee633844d0f1fb60fe8ae1f87bbf7fd42fa42e4ddc`: header-policy and six-profile metadata evaluations passed, but the restore fixture's Git clone returned nonzero before its metadata probes. Its original stderr was discarded, so the cause remains unproved. One separately authorized instrumented control retained exact Git args/CWD/exit/stdout/stderr and passed the unchanged benign restore fixture with 15 guarded MSBuild and one VSwhere child in completed context `20261115T014130Z-94316-4a79c6d34bc74c1f93d7a56f7bfe2408`, summary SHA-256 `09fafe05db641f277f8a1d56638cdf7bc764d103808d7f0c09dfc74fff1dd95f`. The ordinary context prefix was shorter; no semantic path workaround or historical-cause claim is made. No unsafe legacy baseline, product compilation, pinned upstream helper, desktop input or native acceptance is part of these filtered component controls. Current-source final compiler-free gates, PrePush, ARM64 and broad desktop/provider/NAS/W5 qualification remain separate; I65 stays **ACTIVE**.

- Independent review found a real fixture-only output-authority defect in the newly added diagnostic recorder: an inherited nonempty record environment variable alone could authorize an append. The correction binds both recorder paths to active script-owned audit state and canonical marked-root containment, captures the state in the launch interception, and restores prior state/environment in finally. An owned-sentinel mocked negative proves inherited paths grant no write authority; no external path is probed or written. Corrected strict Git/metadata controls passed 16/0/0 in completed context `20261115T014717Z-59000-d3b9f0a761c345dab88e3f8e75ee6590`, summary SHA-256 `7c2e046132a5ae025b408430c8198ad72f4caa54366762b9518dcc67ae2527e4`. The prior source/governance result 297/0/0 and filtered 68-file import-order parity 8/0/0 remain retained against their earlier fixture hashes, not relabelled as the corrected source.

- The length-matched instrumented campaign retained 31/1/0 in completed context `20261115T014810Z-104216-15ae79d18aa34727bbb87c5b64123ac0`, summary SHA-256 `a360327c7e843e473eaca63d644e1013879efd6763890f5f26cbacff0df815de`. The prelaunch witness proves a 266-character clone destination; actual captured Git exit 128 reported `this operation must be run in a work tree` and `Clone succeeded, but checkout failed`. This identifies the failing checkout boundary, without proving OS internals or a production capture defect. The bounded fixture correction adopts the existing canonical two-phase no-checkout/explicit detached-checkout pattern and process/repository long-path settings, copies objects with `--no-hardlinks`, and retains the same pinned-commit assertion and length witness. The original failures remain immutable; no shortened-path workaround or unsafe baseline is used.

- The two-phase length-matched setup retained another 31/1/0 in completed context `20261115T015149Z-107636-63c1de277b074893b1ec79c39b1708b2`, summary SHA-256 `a475a53030df7e4c8a9742bbd6b5d3635effc415d7cc2ffb49da6695f91d79c8`: clone passed, but explicit Git `-C` checkout refused the same 266-character directory with `Filename too long`. The actual owning fixture asserts metadata paths >260, not Git `-C` directories >260. Its bounded correction therefore acquires a shallower canonical scratch, retains long metadata and the complete host/environment matrix, and refuses source directories above 240 before launch. Path-budget records distinguish ordinary source placement from the failed >=265-source discriminator; no production long-source guarantee or original-failure relabelling is made.

- The ordinary-source/long-metadata corrected control passed 19/0/0 in completed context `20261115T015445Z-93988-bdd8de54307a46d2bac97a8d7fd44ed2`, summary SHA-256 `5957e68324adbbf6e39f1ed4de98ff0c8504b809d6090f4be8d1dddd05e04539`: all three audited metadata Its, strict Git/owner controls and five actual Git starts, ambient-record and unsafe-PSI negatives. Its source-directory budget and >260 metadata witness are retained with exact capture records. These are filtered compiler-free component observations; no product compilation, pinned upstream identity-helper execution, Git `-C`>=265 support, creator attribution, ARM64 or Full acceptance is claimed. Remaining source/governance/parity and actual specification/inventory/changelog gates are reported separately against the frozen follow-up.

## Done criteria

- Every report ID has an explicit current-source disposition and validated defects have a repair, meaningful regression coverage, and owning-spec update.
- I60/I62 status and anchors are reconciled without claiming unrelated roadmap items complete or waiving outstanding release gates.
- Required affected, PrePush, build/runtime/platform/performance, focus-taking, and hosted gates are completed or remain explicitly pending/blocked with evidence.
- One valid dated change fragment remains in the single authorized feature PR; `CHANGELOG.md` remains unmodified and no separate changelog PR is opened.

## Stop conditions

- Stop only the dependent slice when report and current contracts cannot resolve a product/compatibility decision, the exact sandbox marker cannot be verified, output ownership changes unexpectedly, or focus-taking acceptance lacks time-specific authorization.
- Keep independent fixes moving and record failed attempts, skips, and unexecuted gates plainly.


### 1c3a ARM64 attribution execution and trusted-timeout repair (2026-11-14)

- Exact-head automatic CI, formatting, dependency review and CodeQL passed at `1c3a5a12cff150161a8b5f429a934cf5b3e441a1` (CI `43584463256`, security `43584463350`). The separately dispatched test-only ARM64 Release `CommandsUiAttribution` run [43584996256](https://github.com/RedSalamanders/RedSalamander/actions/runs/43584996256) **FAILED**; the native build passed with zero warnings/errors in 36m39s.
- The first observer/help prerequisite actually passed once (native **1/0/0**, PID 1420 exit 0; canonical child `20261114T111214Z-7832-f70fd2cf30714996919b11b9e575410c`, runner exit 2 / repository `NOT_EVALUATED`, zero active/invalid quarantine, retries and repairs). Its archived plan used the trusted adapter's timeout multiplier 2, while strict verification reconstructed default 1. The first campaign row therefore retained `unsafe-or-invalid` with `Archived Commands plan, profile, filter, source or receipt proof does not match.` All three UI selections remained explicitly **unexecuted**; this was not a demonstrated UI-case or desktop-retirement failure.
- Official raw logs, both artifact ZIPs/digests and every extracted file were preserved in `Z:\RedSalamander.Perf\evidence\completed-runs\20261114T102338Z-74528-0790cc2688394e938758dfd3da448ba5\scratch\tooling\hosted-attribution-1c3-monitor-recovered\terminal-preservation-raw`. Inspection proof SHA-256 `8e7219f9d123aeb8bae7a7e13281dbc133ec6d91118448b19bb4df08ed69220f` binds all 23 child and four campaign original inventory members, source snapshots, native exit and 243 PID/time-matched metric rows. The late v1 aggregate is separately hash-observed through the official artifact ZIP, not labelled an original native-intent member. Full receipt/descriptor bodies were not uploaded, so offline artifact byte-verification remains unavailable. The initial cosmetic queued-title monitor refusal is separately preserved; it caused no replacement dispatch.
- The candidate correction propagates an optional trusted multiplier (default 1) through the strict v1 proof and both campaign outcome paths; the adapter supplies one value 2 to actual invocation and expected proof. It also correlates runner/native timeout observations without accepting expectations from either. The complete real archived plan matches the owning canonical planner exactly at 2 and differs at 1 or 3. All other native planner controls remain their fixed defaults. Focused compiler-free validation of the dirty nine-path correction passed 106/0/0 with unchanged target hashes: strict verifier 3 (default-1 and trusted-2 profiles plus mismatch/tamper controls), ordinary-failure/campaign/clean-quarantine 15, adapter/workflow 29, inventory 17, governance 12, documentation drift 30. Evidence is retained in Z:\RedSalamander.Perf\evidence\completed-runs\20261114T112236Z-93880-5593f07809724ed9a0818d4706dd6623\scratch\tooling\trusted-timeout-propagation; no native rerun or new hosted execution is claimed. Run 43584996256 stays FAILED, historical source-specific passes remain historical, I65 stays ACTIVE, and the existing broad desktop/SMB/applicable NAS gates remain pending.

### c55 hosted admission evidence and stock-instance observation follow-up (2026-11-14)

- Corrected test-only ARM64 Release attribution run [43595987751](https://github.com/RedSalamanders/RedSalamander/actions/runs/43595987751), exact source `c55d6e75af53d02928d1a7f5a1f4ac0ddedcddb8`, remains **FAILED** after a zero-warning/zero-error build. Its first help/observer selection passed once (native 1/0/0, PID 7112 exit 0) and the real strict v1 verifier accepted the trusted timeout 2. The next Preferences child was refused by stock preparation before boundary creation or native launch: `An unapproved foreground surface prevents hosted preparation.` The missing runner-owned ID is downstream of this refusal. The final two UI selections are unexecuted. Raw official logs, artifact ZIP/service digests, all 34 original member hashes and late-control custody are retained in canonical context `20261114T124033Z-60248-5f5805e31adc4e82b668a6c8552b66dd`; terminal-inspection proof SHA-256 `481a69407cdc610247873f2a7d1c43d4e31c194f8f0d5bda0ca5c545915a1369`. Receipt/descriptor bodies were absent from uploads, preserving the offline byte-verification limit.
- The observed Unknown terminal-host tuple lacks retained process creation/parent/job provenance at this pre-boundary refusal. Read-only review found no validated creator. The actual help, direct CI and contained native launch paths already request no console window. Separate background dependency-install and sanitized SDK/MSBuild callers used the uncaptured console-capable route; the follow-up moves those two callers to the existing captured no-window streaming helper while preserving delegation, locks, arguments and exit status. This fixes that reachable launch-policy gap without proving that it created the historical terminal, which could also have pre-existed in the background. Class, title and timing do not prove ownership or payload. The observation portion only closes this evidence gap: minimal first/last captured-instance baselines plus failure-only canonical enrichment of the exact held process lease, creation matching and final relation checks. Classification, approved closes, input, polling/deadlines, containment, strict proof and quarantine rules are unchanged. No foreign-process cleanup, alternate-desktop bypass, actor attribution or cause fix is claimed. Current hosted query outcomes remain unexecuted for this follow-up.
- Source-c55 local focused UI execution remains independently accepted: Debug and Release each passed the three exact UI selections once (3/0/0, native exit 0, selected runner exit 2; no retry/repair), using current byte-verified receipts. Final proof SHA-256 `65d4392552984ae194d3a77a0f85e074f884abd34fbf889b280c6cde9ee8ea3d`, canonical context `20261114T142356Z-84788-046ff59e0abf4ad595cedee7c82728c7`, retains 88 original members and both profile joins. It contains no dedicated external final global cursor/foreground equality metric. These passes and the green c55 noninteractive PrePush remain source-specific; they do not qualify this later observation patch, Full ARM64, the four broad desktop entries or remaining SMB/applicable NAS gates. I65 remains **ACTIVE**; historical 853/084 failures and published dirty-29 component/projection limits remain unchanged.

- The initial observer campaign passed 369/0/0 in canonical context 20261114T171326Z-104004-b585a173556f410d9542832fbee50626; its superseded three-file packet/source manifest is preserved byte-for-byte in completed context 20261114T172900Z-72004-0f966b5292db4e5988832cb6cdc70231. The first joint campaign (174429/PID106616) retained 436 passed/4 failed/0 skipped from one missing derived-consumer inventory declaration, then repaired that declaration. The next campaign (174732/PID109328) retained 439/1/0: an existing managed GUI-bootstrapper fixture returned -1 before its marker; the exact failed iteration/cause remains unknown after automatic fixture cleanup. The fixture now acquires bounded short marked-sandbox paths, checks executable/config/log budgets before launch, and retains actual 0/3010/17 exits, arguments, logs and original errors before cleanup. Its targeted 11/0/0 and final pure primary/secondary/fatal/restore controls 9/0/0 are retained; no production wait, timeout or assertion was weakened, and MAX_PATH is not promoted to the historical cause.
- Final frozen joint-source compiler-free validation passed **440/0/0** in 140.311 seconds: whole HostedTestDesktop, source contracts, inventory, governance, documentation drift, ProcessStreaming, Vcpkg safety and compiler-free BuildReproducibility (RequiresBuildToolchain excluded). All 18 source hashes were stable; the [compact stock-observation component packet](../../TestRuns/4cb089111a23/Testing/2026-11-14_171326_HostedStockObservation/README.md) pins the 17 non-reporting targets and all 40 exact fake-work rows from completed context 20261114T181221Z-106388-bcb8bdefe1754f749767db431e92cd3a. The actual background launch control requires no-window/redirected PSI, window zero, both pipes, exact owned job/PID-parent and exit 17; observed console-server association count 1 is diagnostic, not a zero-count requirement. These component results do not prove the old terminal creator, native API cost, animation causality or a performance improvement. New-source native/ARM64 and broad desktop qualification remain pending; I65 stays ACTIVE.

### VSwhere metadata caller follow-up prepared privately (2026-11-15)

A private unapplied patch lifts the existing silent DxUi capture projection into
the canonical ProcessStreaming module and routes the standalone restorer, root
build and local vcpkg VSwhere metadata queries through it. The original bare
probes were read-only discovery, not intentional interactive installers. The
reachable launch-policy gap is validated; historical console authorship remains
unproved. Queries, workflow fast paths, native-host ordering and build fallback
remain intact; nonzero standalone/vcpkg discovery fails closed. No installer or
real interactive MSBuild inheritance is changed. This preparation changes no actual
RS source, dependency pin, SDK policy or engine. Focused native-query controls are
separate from native product/Full/ARM64/desktop qualification. I65 stays ACTIVE.

Private final controls cover eight actual contained no-window CUI starts through
the shared projection and all three caller ASTs, with exit0/17, separate streams,
literal arguments, owned console visibility and cleanup; twenty existing focused
source/metadata/workflow controls also passed. Earlier private fixture failures
remain retained. These are metadata-query controls, not product native acceptance.
The final composed source requires a reviewed PR execution-digest refresh or
conservative local PrePush accounting; the old digest must not delegate new work.


### BUG-09 paired marker-proof coverage prepared privately (2026-11-15)

An unapplied exact-source candidate extends the ordinary FileOps runner with
`FileOps_LocalAncestryMarkerProofCost`, included in unfiltered Full and the existing
safety family. It compares observation-only current identity guards with the same
actual guards plus marker proof on identical disjoint owned fixtures: 16 cohorts,
six alternating ABBA/BAAB blocks and four calls/window (1,536 timed guards/768 proofs).
The existing scoped identity-less injection and actual counter must agree; 16 untimed
real Copy witnesses retain source and exact payload and must traverse the production
proof. Primary/alternate roots remain existing marked resources with explicit opt-in;
there is no scan fallback, new link/share/mount/root or no-lease exemption.

Only private source/format checks are intended here. No build, native run, timing
measurement, pin or Full acceptance has occurred; BUG-09 paired evidence remains
pending until the exact final consumer source's admitted Release Full executes and
its raw data/receipt/profile/fixture bindings are reviewed. The bounded component
protocol does not qualify transfer throughput, NAS, I60/W5 or other roadmap work.
I65 remains ACTIVE. Compose this checkpoint and owning-spec additions with the
separate VSwhere and accepted producer-pin changes; keep the existing RS fragment
and do not fold CHANGELOG.

### BUG-09 long-path cleanup observation repair prepared (2026-11-23)

Exact M `1895ee443289e3275cf61402a971d55b884879a2` x64 Debug Full preserved a
FileOperations failure in `FileOps_LocalAncestryMarkerProofCost`: the first deep
existing-parent cohort completed 98 guard calls/guard-status rows, the successful proof
checks, and the actual Copy/payload checks, then the plain-path
`std::filesystem::directory_iterator` failed during the owned control-parent cleanup
observation. The run recorded primary ReFS/alternate NTFS; source/target fixture strings
were 749/747 characters and the enumerated parent exceeded `MAX_PATH`. The API error
value was not retained, so no specific Win32 cause is asserted. The exact failed run is
immutable at
`Z:/RedSalamander.Perf/evidence/completed-runs/20261123T122737Z-84824-3896c883c2c34a50ac58d81e877d7375`.

The R-based correction uses `Common::Paths::ToExtendedWin32Path` for that enumeration,
retains fail-closed marker/error checks, and reports cohort, error value and path length.
No timing, cohort, copy-witness, skip or acceptance rule changed. Compiler/native
validation has not yet been run on this correction; the matching focused Release case,
paired ReFS/NTFS evidence review and required broader Release/Full gates remain pending.
I65 remains ACTIVE.

### Composed H8 API4 candidate prepared privately (2026-11-15)

This unapplied private candidate combines the reviewed three-VSwhere metadata
capture fix and BUG-09 marker-proof component case with proposed DxUi
[`7fb2ab71852087d93b3a3e5b3ba504c37f56267c`](https://github.com/RedSalamanders/DxUi/tree/7fb2ab71852087d93b3a3e5b3ba504c37f56267c),
retaining API revision 4. H8 is **UNQUALIFIED/PENDING**: fresh producer gates
remain required. Historical H7 passed six native profiles and eight consumers but
its normal paired measurement timed out at 150 minutes after 135/144 receipts; no
complete summary or verdict exists. This is a failed partial A/B study, not A/A,
and not adoption evidence. H8 changes only the outer workflow deadlines and owning
source contracts; its twenty-file measured harness remains 7249D6FCCF538E722FAC8E615CB1656BD516F9224828BBB7CDFBB64FB82C0E56.
Fresh H8 [validation](https://github.com/RedSalamanders/DxUi/actions/runs/43769081609)
and [format](https://github.com/RedSalamanders/DxUi/actions/runs/43769081636) remain pending.
The earlier H6 and H7 candidates are superseded; neither failed cohort is adoption evidence. Actual RS source/index remain at 83fd/B pin;
apply only after the exact producer source qualifies, or supersede this candidate
if the accepted source changes.

The lock and sole hardcoded test expectation change together. Use canonical Source
restore/build when SDK NotDistributed, never B-SDK reuse, a descriptor/allowlist
exception, engine update or API5 migration. Rebuild with matching public headers/archive
and retain actual consumer/toolchain/CRT bytes. Earlier metadata-capture focused
passes and BUG-09 pure source/format checks remain component observations, not
execution of this composition. No build, native timing, PrePush, ABI, hosted or Full
acceptance is claimed.

After focused branch proof and normal required PR checks, merge the normal follow-up
PR with broad qualification explicitly pending. After the normal follow-up PR merges,
reconcile the exact merged main commit/tree, then produce fresh test-enabled full-solution
x64 Debug/Release Embedded builds and byte-verified receipts for that merged source.
The unused authorized desktop window is capped at six hours including the existing
three-minute teardown buffer and starts only after those corrected merged-source
preparations pass. Native ARM64 Release Full and the fixed published-head
CommandsUiAttribution diagnostic may run in parallel; neither ARM result is a prerequisite
for starting local x64 desktop coverage. Broad qualification remains pending until its
actual receipts/results are reviewed.

Verify actual hosted GITHUB_SHA against the reconciled merged identity. If main advances,
retain its distinct identity and pending state; do not relabel results. Branch-focused
packets/receipts remain historical, even if a new HEAD has the same tree. Existing Full
covers local regressions once; filtered follow-up is for a demonstrated gap/investigation.
SMB/resource and I60 NAS/W5 work retain their contractual scope. I65 remains **ACTIVE**;
the desktop clock has not started.

### Qualified H8 composition applied locally; consumer execution pending (2026-11-16)

The reviewed nineteen-path composition is now **APPLIED LOCAL / CONSUMER
QUALIFICATION PENDING** on RS `83fd6a3faeddd3228c5f462a857db3ba00889fd4`. It selects
qualified API4 DxUi `7fb2ab71852087d93b3a3e5b3ba504c37f56267c`: exact producer
[validation43769081609](https://github.com/RedSalamanders/DxUi/actions/runs/43769081609)
and [format43769081636](https://github.com/RedSalamanders/DxUi/actions/runs/43769081636)
passed, including the complete 144-report normal P/T8 study, trusted-policy Gate
and all 24 exact-budget slots. Independent offline replay and retained-binary rehash
were reviewed before source-adoption authorization. H7's 135/144 timeout/no verdict
remains failed history; historical private checkpoints above remain point-in-time
preparation records.

The shared VSwhere capture, BUG09v2 component case and existing BUG15 coverage are
integrated without changing production C++ or their assertions. BUG09 actual timing,
copy witnesses and cleanup assertions remain unexecuted. Current owning Build norm
now specifies qualified Source consumption; SDK NotDistributed grants no SDK reuse,
allowlist/descriptor/engine change, API5 migration or cooking port. Matching consumer
headers/archive, zero-warning full builds, byte-verified receipts, ABI, PrePush,
ARM attribution and fresh x64 Debug/Release Full plus separate ARM Full remain pending.
The authoritative Tests/test-scopes.json PR workflow digest is refreshed only after
these composed source bytes settle; a digest refresh grants no completed qualification.
RS CHANGELOG remains untouched and the existing dated fragment is retained.

Only contained compiler-free tooling/source/inventory/governance/changelog/spec checks
are authorized at this stage; their outcomes accompany the staged review. No local
native build, product test, focus, workflow dispatch, commit or push is claimed here.
I65 remains **ACTIVE**. The authorized 4–6 hour desktop clock has not started; its
conductor announcement follows corrected consumer builds and the ARM diagnostic,
with the existing three-minute teardown buffer. Applicable SMB/resource prerequisites
and separately owned I60 NAS/W5 scope retain their actual contracts.

### Remaining background Git launch closure implemented (2026-11-16)

A bounded nine-file follow-up closes actual scoped-planning, tooling/spec inventory, snapshot/archive, runner graph, hosted HEAD and vcpkg provenance/bootstrap launch gaps. Canonical contained capture serves graph/HEAD/probe reads; provisioning writes retain delegated streaming and the existing lock. Reviewed engine-independent private transports preserve raw/NUL, line/stdin, inherited environment and exit/error semantics; existing raw readers are unchanged. No temporary caller shim is the production fix. Exact source guards, controlled owned-CUI observations and focused semantic tests are reported with the final staged bytes. The previous cc848 digest was intermediate; the canonical scope-manifest field is refreshed after all production/test/inventory bytes settle. This is no historical Terminal attribution, UI fix or native qualification. I65 remains ACTIVE; consumer builds/ABI/Full and desktop clock remain pending/unstarted.

The CP437 investigation also validated real Git stdout Unicode corruption, rather than only a fixture mismatch. Specialized Git text decoders now select UTF-8; the canonical capture/streaming boundary adds one optional per-child stdout override for known Git callers, including the DxUi Git shim. Defaults, arbitrary stderr, non-Git decoding, ordered stdin and raw readers remain unchanged. No production global console encoding is changed. Existing UTF-8-only passes are retained with their limits; new actual guarded Git path controls exercise both CP437 and UTF-8. First assertion-pipeline fixture failures and the exact Pester3 Boolean-output diagnostic are preserved; equivalent terminating field checks prevent those doubles from corrupting command return values.

### Compiler-free integration checkpoint; real dependency preparation pending (2026-11-16)

Ten actual whole compiler-free suites pass **512/0/0**: ScopedTesting 13,
ToolInventory 21, SpecInformationArchitecture 85, ValidationFingerprint 31,
TestRunArchive 28, ToolingGovernance 12, DocumentationDriftContracts 30,
TestHarnessSourceContracts 243, ProcessStreaming 17 and DxUiDependency 32.
The scheduled fixture scope uses 136 migrated native Git calls and twelve explicit
protocol-control calls through the prepared no-window helper. Real controls preserve
scalar/multiline/blank versus empty output, literal arguments, redirected diagnostics,
Unicode under CP437/UTF-8 and actual native exits. The first scalar-list wrapper,
stale longpaths source-pattern and missing README map failures remain preserved;
only the equivalent guarded source-pattern assertion was intentionally retargeted.
The inventory classified all 231 files with zero blockers; fragment validation and
specification inventory also passed with zero blocking findings.

Whole Vcpkg safety is **PENDING**, not passed: its pre-build attempt reported
31 passed, one Describe-setup failure and one declared missing-pinned-module skip.
The existing strict object-store locator check correctly rejects the previous DxUi
identity after the H8 lock change. No locator is edited and no assertion is weakened.
Canonical Source preparation/build must establish current locators and restored
modules before rerunning those dependency-sensitive checks. The derived workflow
digest binds settled source only and does not promote this prerequisite or any
consumer native/ABI/Full result. I65 remains ACTIVE; desktop clock is unstarted.
### Final noninteractive launch-safety source awaiting receipt refresh (2026-11-16)

Committed source c842f29686e2dbebab5b36c456b4d4e866273180 completed test-enabled
full-solution x64 Debug and Release Source builds with four workers, saved build 1 and
zero warnings/errors. Each actual receipt bound 188 outputs, 110 uniformly test-enabled
manifests, 25 public headers and 13 exact H8/API4 module/archive/link joins; 188 remains
an observation. Current-pin Vcpkg parity passed 32/0 with two declared optional immutable
baseline-data skips. Explain D/R selected 14 noninteractive and four PENDING_INTERACTIVE
entries, with no PR delegation. Private wrapper exit/child-null cleanup limitations are
retained separately; canonical receipts and independent bytes establish those c842
builds, rather than their outer wrapper exit alone.

The finite selected-profile trace found 20 additional fixture Git calls, 22 native
fixture calls and three asynchronous Versioning visibility sites. These use explicit
native/Git fixture capture or retain asynchronous Hidden routing and ownership.
Formatter/Ghostty metadata use their reviewed engine-free transports; the actual batch
keeps inherited production routing and is tested through its private launcher seam.
Native-caller controls passed 20/20, including the tiny C++ fixture's actual rebuilds
and expected-warning gates. Formatter/Ghostty/prefix controls, root formatter
preview/payload, real batch marker/argument/exit 19 and shared guards passed within
their stated component scope. Original assertions and nonzero outcomes remain binding;
there is no unsafe baseline, fabricated bootstrap success, retag or waiver.

These follow-up bytes require a new commit and canonical receipt refresh before
SkipBuild PrePush. Immutable c842 receipts do not qualify the future HEAD. No actual
PrePush execution, foreground work or publication is claimed here. I65 remains ACTIVE;
the desktop clock is unstarted and Full/ARM/resource obligations retain their owners.

### Actual 3ebf qualification failure and diagnostic-only follow-up (2026-11-16)

Committed source 3ebf04d83abd705799fba44eae286c0c4a7ec047 completed test-enabled
full-solution x64 Debug and Release Source refreshes with actual `/t:Rebuild`, four
workers, build 1 and zero warnings/errors. Fresh receipts and independent byte checks
bound the exact H8/API4 headers, archive, toolchain and CRT; these results do not
qualify later source. Both corrected wrappers propagated actual exits and restored
an absent build-test environment key as absent.

Actual Debug noninteractive PrePush failed: 14 entries executed, 13 promoted and
one incomplete; 1733 cases passed, one failed and 30 skipped. The archive transport
test observed Git staging exit 128. Two immutable-baseline data skips were also
rejected by aggregate acceptance. Original raw results, receipt bindings and the
111 reported-only disk-audit inventory are retained unchanged. Release PrePush did
not start; the four desktop obligations remain pending.

The required immutable data was subsequently hydrated by one bounded fetch of nine
missing blobs; the two unchanged data cases passed without skips. This is prerequisite
preparation, not an acceptance-policy waiver. The historical Git 128 cause remains
unproved: isolated exact-case execution passed, and path-length hypotheses were not
validated. The only tracked test follow-up adds bounded failure-only diagnostics
before the original staging assertion, with two pure diagnostic controls. It does
not repair or reclassify that historical failure, alter workload/limits/assertions,
or identify a failed process from a later command-resolution probe.

Fresh current-source validation in the complete 68-file order remains pending after
normal commit and compatible Debug/Release receipt refresh. I65 remains ACTIVE;
there is no publication or Full/ARM acceptance, and the desktop clock is unstarted.

### Actual 70fad local gates and failed ARM boundary attribution (2026-11-16)

Source 70fad2b84101f35552f4615e0d03217b8b9a9f59 completed test-enabled
full-solution x64 Debug and Release Source Rebuilds with four workers, build 1 and
zero warnings/errors. Fresh current-source noninteractive PrePush promoted all
14/14 applicable local entries in each profile: Debug 1738 passed/0 failed/28
skipped; Release 1736 passed/0 failed/30 skipped. Canonical change acceptance
PASSED with repository NOT_EVALUATED and four desktop obligations pending; neither
outer exit 2 nor these local passes establish Full acceptance. Exact consumed
writer receipts and profile-specific byte/header/archive/toolchain/CRT bindings
remain retained separately. The original failed 3ebf run and its unproved Git 128
cause remain unchanged.

[ARM64 attribution run 43826250371](https://github.com/RedSalamanders/RedSalamander/actions/runs/43826250371)
at the same source built ARM64 Release with zero warnings/errors. Help passed
1/0/0 with native exit 0, filtered runner exit 2 and a strict source-bound proof.
Preferences then failed before native invocation while establishing the hosted
foreground boundary; the final two selections did not run. Setup and retirement
wrappers can carry the same worker failure, so their labels do not prove two
independent faults. Underlying exception types, HResult and native failure codes
were not retained; the historical boundary cause remains UNKNOWN. Successful
stock preparation and Help's boundary proof do not establish Preferences admission
or complete restoration.

The diagnostic-only follow-up retains the original captured setup ErrorRecord
and visited exception branches, existing HResult/Win32 codes and bounded stack
text before the unchanged primary throw. Aggregate composite messages are omitted
rather than recursively materialized; depth/node bounds cover traversal and text
bounds cover emitted UTF-16 characters, not arbitrary property-getter allocations.
Its corrected V2 pure affected lifetime controls passed 26/0/0. Deep and wide
getter-counter controls reject the exact prior formatter at excluded descendant
reads and pass with zero excluded reads after correction. Original 76 cases,
foreground authority, warnings, readiness/cleanup, timeouts and native lifecycle
remain unchanged. This supplies observability, not a root-cause repair or new ARM
qualification.

Corrected V2 evidence is retained under
Z:\RedSalamander.Perf\evidence\completed-runs\20261116T095215Z-89544-26e912feed6042c99278cf0de76d3af3\scratch\implementation\RS-hosted-setup-diagnostics-v2.
Fresh current-head build/local obligations and ARM attribution remain pending
after the diagnostic change. I65 remains ACTIVE; no tracked TestRuns closeout or
Done move is claimed. The approved first local 4–6 hour foreground window has
never started, and fresh unfiltered x64 Debug/Release Full plus the separate
applicable ARM Full qualification remain pending.
### Actual 22a local failure and bounded archive fixture housing (2026-11-16)

Source 22a1fedf3aef1dc3e8a19fa70f2ba6c6cb0d6c44 completed test-enabled
full-solution x64 Debug and Release Source Rebuilds with zero warnings/errors and
current-head byte/profile/header/archive/toolchain/CRT verification. Actual fresh
Debug PrePush then FAILED: 14 executed, 13 promoted, one incomplete; 1746 passed,
one failed and 27 skipped. Tooling recorded 1444/1/0, with all eight new hosted-setup
diagnostic controls passing. Quarantine and repair counts were zero. Release
PrePush did not start; no failed result was retried or reclassified.

The retained archive staging diagnostic now identifies this occurrence's native
failure: Git add --all exited 128 with Filename too long, at a 152-character
fixture root plus separator and 107-character generated leaf (260 total).
The original 3ebf occurrence remains unproved. The command-resolution probe is
not failed-process image identity. The failed source, raw stderr/argv, actual
consumed writer receipt 62e1cb8b4244d62bdd490be6de6760df5721341e86c37f9125353a75f7b9c4d6
and original archived run 20261116T102505Z-108548-f57ef8e9e04a4cb3a88363db19c11427
remain unchanged.

The test-only correction uses canonical ps/ar-fullGUID housing and checks the
maximum normalized full native path is below 260 before workload writes. Every
800 relative filename, payload, UTF-8/ordered-stdin/worktree-map assertion, original
exit/error expectation and greater-than-65536 stderr requirement remains intact.
No production helper, Git policy, environment default, user/global config, alias,
junction, mount or root changes. Native inherited/sanitized controls preserved
four real 260/268-path failures and two bounded-housing successes at maximum path
236, including distinct ordered hash/worktree witnesses and Git-emitted argv.
The exact actual 800-file case passed 1/0/0 under an owned child-only global-false
configuration. Earlier local/command-policy controls remain diagnostic attempts,
not the chosen fix or a claim about all Git versions.

Fresh integrated current-head build/local qualification remains pending after
this fixture change. Separate ARM boundary failure evidence and investigation
retain their own source/cause owner. I65 remains ACTIVE; the first approved
foreground window is unstarted, four desktop obligations and Full remain pending,
and there is no new ARM dispatch, tracked TestRuns closeout or Done move here.
### Actual 22a ARM focus-query failure and next context diagnostics (2026-11-16)

[ARM attribution run 43846481982](https://github.com/RedSalamanders/RedSalamander/actions/runs/43846481982)
at source 22a1fedf3aef1dc3e8a19fa70f2ba6c6cb0d6c44 retained Help 1/0/0,
then a Preferences pre-native hosted-boundary failure; the final two selections
did not run. The new logger retained Win32Exception code 87/HResult 0x80004005
from HostedForegroundBoundary.ReadFocus in the worker's Run path. The actual
requested thread and its liveness/relation were not retained, so the underlying
cause remains unknown. This is separate from the local Debug 1746/1/27 failure
above; Release PrePush did not start, and neither failed result is reclassified.

The accepted next diagnostics save that original query error before a fixed
failure-only read-only sample of the requested window relation and foreground.
Requested HWND/TID/PID, input size, closed stage, fresh observed facts and explicit
non-atomic/unspecified-output markers precede the potentially long saved identity
under the unchanged formatter limits. Observer closure construction occurs only
after query failure. A private normalized source-template identity refuses stale
inline native types before initialization. No retry, additional focus/input,
attachment, handles, unknown-focus acceptance, wait or restoration-policy change.

The changed-source pure controls passed 43/0/0: nine focus-context controls plus
26 existing lifetime and eight return-policy controls. Compiled/source controls
reject the old closure shape, and the real bounded formatter retains fresh facts
with a long saved identity while rejecting the old ordering. All 84 prior cases
remain exact; no native window sampler or UI ran. These are observability/cache
controls, not native ARM qualification or a behavioral root-cause repair. Earlier
control provenance failures and superseded diagnostic packets remain preserved.
The accepted packet is retained under
Z:\RedSalamander.Perf\evidence\completed-runs\20261116T120253Z-68180-35e6b75e365f4ea9a703c4a474e6bc90\scratch\validation\RS-focus-error-context-v2.

The bounded archive fixture component and these focus diagnostics are integrated
for one future reviewed source cohort. Fresh current-head builds, actual local
PrePush and hosted ARM gates remain pending after commit. Existing producer pin,
engine, C++ and test assertions retain their accepted identities. I65 stays ACTIVE;
no foreground lease has started, and no Full, merge, TestRuns closeout or Done move
is established by these component/static checks.

### Actual b985 local qualification and DesktopProbe queue evidence (2026-11-16)

Source b985bbade0402eeacf2d4c01c29fa72aa3bcda7e completed test-enabled x64
Debug and Release full-solution Rebuilds with zero warnings/errors and current
H8 Source/header/archive/profile/toolchain byte joins. Fresh noninteractive
PrePush promoted all 14 local entries in each profile: Debug 1756/0/27 and
Release 1753/0/30. The archive 800-file transport case and all 43 selected
Hosted lifetime/focus-context/return pure controls passed first attempt. Actual
consumed writer receipts were fcc18173462892c7179abec380e3ef75ae62ecd15bad8444d9fa4b8e7ef3e0be
and b7ef1191b465a2f0b33ea562ef85a7a18245ef821e8b8313c0373826c0088a39;
neither is substituted with the initial build receipt. Both change verdicts were
PASSED, repository verdicts NOT_EVALUATED, with zero quarantine/repair/retry and
four desktop obligations still pending. Release's three additional skips declare
unavailable Search service test hooks; unchanged canonical policy accepted them.
The closed packet is retained under
Z:\RedSalamander.Perf\evidence\completed-runs\20261116T121652Z-75132-755ae4f9e9634f01a7c40da5a6d16598\scratch\qualification\RS225-fixture-focus-current-prepush.

[ARM attribution run 43866426276](https://github.com/RedSalamanders/RedSalamander/actions/runs/43866426276)
at the same b985 source retained a zero-warning/error build and Help 1/0/0,
then Preferences failed before native execution at ReadFocus with saved error
87/HResult 0x80004005; the final two selections did not execute. The requested
HWND 0x201C6/TID 8720/PID 8732, input size 72 and pointer width 8 were retained.
Fresh sequential observations found the window valid, matching TID/PID and
current foreground, but are explicitly non-atomic and prove no queue liveness
or underlying cause. Original b985, 22a and earlier failed results are preserved.

The next accepted DesktopProbe extension is evidence only: four fixed foreground
HWND/TID/PID samples bracket one explicit-thread query and one foreground-zero
query, with one bounded class observation and actual process/OS architecture.
Null or changed identities prohibit interpretation; saved errors and successful
outputs grant no activation, fallback, retry or qualification authority. Local
x64 observations where both queries succeeded are not native ARM evidence.
The accepted 26/0/0 pure controls include six new orchestration/source-negative
cases; no native sampler or UI ran. All 93 prior cases and existing ReadFocus,
acquisition and restoration behavior remain unchanged. The new source's first
remote diagnostic is intended to be the read-only ARM DesktopProbe, retaining
NOT_EVALUATED rather than full-build attribution or acceptance.
The accepted component packet is retained under
Z:\RedSalamander.Perf\evidence\completed-runs\20261116T142907Z-87404-e52e6b798c59442c88cb07c5773d4667\scratch\review\RS-DesktopProbe-queue-evidence-v2.

I65 remains ACTIVE. The b985 local receipts qualify only their own source;
current-source validation remains pending after this diagnostic extension.
ARM correctness, four desktop obligations, fresh unfiltered Full and normative
closeout remain pending. The first approved desktop lease is still unstarted;
no tracked TestRuns result, Done move or merge is established here.

### Primary original-foreground queue contract and prewrite authority (2026-11-16)

Published diagnostics source `b1a618aea416fe957fbb5d65b035fe6aa28299d3` passed
its three required compiler-free GitHub contexts. The one explicit ARM64 Release
[DesktopProbe run 43889225762](https://github.com/RedSalamanders/RedSalamander/actions/runs/43889225762)
completed successfully as `NOT_EVALUATED`, without dependency installation or
product builds (existing MSVC/cache setup still occurred). Its source/platform
and actual process/OS architecture were joined as native Arm64, pointer width 8,
input structure size 72. Four foreground identities were stable; explicit thread
9004 and documented foreground-zero queries both succeeded with identical GUI
results. The observed class was `Windows.UI.Core.CoreWindow`, not a console;
the foreground-desktop name remained unavailable with error 5. This establishes
neither a desktop mismatch nor the cause of historical error 87. The original
b985 failure and local x64 characterization remain unchanged and source-bound.

The applied local successor now selects the documented foreground-zero query as
the primary original-focus sampler, with strict before/after original identity,
exact returned active/root and independent related-focus identity/root checks.
Successful null focus is distinct from unavailable GUI state. Saved focus is
immutable; related cross-owner focus may be observed already restored but cannot
grant write admission. The fixture refuses before acquisition if the saved state
cannot satisfy that conservative restoration predicate. Restore revalidates the
existing Default input-desktop predicate at both prewrite validations, keeps the
full inherited/input context check after waits, and retains the final foreground
recheck, one saved-original request, owned readiness and both 1,500 ms limits.
No original-focus query runs while owned/null foreground settles; no returned
focus is written, no error-87 fallback or retry is added, and DesktopProbe is
unchanged. This is an API-semantics/authority correction, not historical-cause
proof or a console-compatibility claim.

Independent V2 review accepted the five-file component and its canonical managed
capture: 75 actual affected-source pure controls passed, zero failed/skipped,
true contained NoWindow child 44464 exit 0, streams drained and parent environment
restored. The added actual-orchestration refusal permits the first validation,
then rejects the second/prewrite desktop check with the exact primary error,
two validations, zero requests and zero GUI queries. All 108 V1 case extents
remain unchanged, with one case added; the original 99 names remain, including
the four explicitly justified query/context/refusal-binding contract updates.
The first 73/1 development capture, final V1 74/0 and original immutable packets
remain preserved. Final V2 proof SHA-256 is
`E9CD2F349503609ACA3A28AD4D25A5B0DBFF04A5E52046189B749EDD08BC4A52`,
seal `9FD00C27EF0075430E6351428C41F80B5D2CE63FC41147654792157FAB17CB1E`,
packet run `20261116T152356Z-4928-510aba118a51458aaadcf7abb3d272b8`.
These controls compile managed interop and exercise injected readers; they invoke
no native window sampler or activation and grant no native acceptance.

Shared inventory metadata, one ordinary dated fragment and the canonical scoped
workflow digest `19cea470a9f85f2c7999a56b662dda1811ba5cd05deeff2ecb32b1b86858bd86`
are composed around those exact reviewed source bytes. The H8 API 4 source pin and a094 engine remain
unchanged. At this integration freeze, no new source commit, build, PrePush, ARM qualification, foreground
lease, workflow dispatch or merge has occurred for this correction. The b985
Debug/Release receipts are historical and are not rebound; new-head builds,
scoped local gates, required GitHub contexts, actual ARM interactive qualification
and restoration, applicable unfiltered Full and the four broad desktop entries
remain pending. Console compatibility remains unproved; this creates no separate
bespoke console-fixture gate. I65 stays **ACTIVE**, its first approved desktop clock
is unstarted, and no Done move, generated changelog fold or qualification waiver
is established.

### Merged-audit qualification restart and preserved failures (2026-11-20)

Qualification resumes in an isolated candidate from main `7ff4e90d95e71fb09c41fca5095ea0e867c57828`;
the original audit checkout and shared main checkout remain preserved. The candidate
uses exact CI engine `f965c7945aff98dbab352623a0cfbe333e7f282f` and unchanged H8
API 4 source `7fb2ab71852087d93b3a3e5b3ba504c37f56267c`. Pinned x64 dependency
provisioning and canonical Source restore passed; this is not a product build.

The original merge `53d0a9d58257f53efe817f52a3c309a25454ca1b`
[native CI run 43917127850](https://github.com/RedSalamanders/RedSalamander/actions/runs/43917127850)
remains **FAILED**. Both native Release build steps passed, but native tests failed
and production Release/ZIP gates did not execute. The x64 job's paused-Copy case
reported 2 timer ticks, 10 renders and 2 owner ticks. Its isolated rerun passed;
shuffle seed 1015184194 did not reach the cadence case, while seeds 587162809 and
2023453774 passed that case despite other failures. These shuffles do not establish
a reproduced shuffled cadence failure or prove its historical cause.

The original ARM64 Commands child (PID 2932, exit 1) reported 964 passed, two failed
and five declared opt-in skips: `cmd_pane_fileops_completed_group_and_navigation`
lost its expected primary Dismiss action, and `modeless_window_ownership` refused
foreground settling at Terminal PID 4408/HWND `0x25206BC`. The classifier's isolated
first-case rerun passed (PID 296, exit 0); only afterward did boundary retirement
refuse a different Terminal PID 9812/HWND `0x30604A8`. Neither observation identifies
a Terminal creator. The original ZIP and source-bound raw results remain preserved
under completed context `20261120T091126Z-30428-a0bb893801f94675bcfd20f87a859b82`,
`scratch/qualification/RS53d0-ARM64-connector-original`; classifier success does not
replace either original failure.

The separate current-main source `7ff4e90d95e71fb09c41fca5095ea0e867c57828`
[run 44665794670](https://github.com/RedSalamanders/RedSalamander/actions/runs/44665794670)
also remains **FAILED** after both native Release builds passed. Its original x64
Commands result was 965/1/5 (PID 7780, raw exit 1), failing the traversal-policy
source guard; isolated PID 3188 also exited 1. Original FileOperations was 288/1/35
(PID 6768, raw exit 1), failing `FileOps_ReviewRemediationContracts`; isolated
PID 9172 passed, while three separate shuffle children exited 1. The last shuffle's
47/26/251 is not the original result and does not establish a product cause.
Both architectures' original ToolsPester result was 1449/7/0 and failed again on
entry classification: formatter-sensitive Terminal/Floodgate guards and three
stale project-skill paths were deterministic source-contract failures.

Current-main ARM64 Compare exited 0. Its original Commands child 2568 exited 1
with 962/3/6: `file_operations_phase0_typed_contract_source_guard`,
`cmd_pane_fileops_popup_throughput_disclosure`, and
`cmd_pane_fileops_conflict_close_hides_escape_cancels_snapshot`. Boundary retirement
then refused foreign Terminal PID 1312/HWND `0x3403EE`, before any Commands classifier
or ARM FileOperations child ran. The throughput trace separately observed Terminal
PID 6008/HWND `0x1203AA` during cycle 21; the conflict-return trace observed foreign
PID 3092 before View. These identities do not identify a creator or a product cause.
Bounded reports remain in completed context
`20261120T101925Z-69372-a67e9e20b9534e5ea73d4b46c73b01b5`,
`scratch/review/current-main-44665794670-reports`, with exact entry joins in
`20261120T102113Z-22260-eaa84ae763954edf8a57f5910a53d5bb`.
Neither this run nor its classifier results replace the distinct original merge failures.

The reviewed candidate repairs paused-rate projection, isolates destructive Phase14
shutdown and strict-empty ClipboardReady work in private real owners, preserves
legitimate retained rows and exact retained payloads, and routes completions using
their actual owner rather than colliding engine-local IDs. Direct launch/retirement
diagnostics preserve primary errors and mandatory cleanup without claiming foreign
process ownership. These product/native fixture changes are implemented but remain
unbuilt and unqualified. The failed diagnostic development campaigns remain retained,
including V3's 42 passed/two failed and V4's first 37 passed/one failed. V3's later
directory-metadata drift is separately recorded by observation SHA-256
`F4A524D677389178DAE2571D4BBD532CFE1013F7A9E32800A16255B549329783`;
no whole-inventory equality or acceptance/reuse is claimed for that failed archive.

Final V5 diagnostic source passed the full affected pure campaign once: Pester 3.4.0,
contained NoWindow child 60768, true exit 0, 51 passed/zero failed/zero skipped,
unchanged measured code/test bytes and all 36 historical selected cases passed.
Evidence is closed in completed context `20261120T101616Z-49784-d423f37ad606434390af55ba2c8de1b1`,
`scratch/pure-review/failure-launch-v5-affected`; result SHA-256 is
`1B960FFA095A6E49FD3BAF115787D12D34C1EBC6F2A8C9990D2435F714B4D1B4`.
These managed/mock controls invoke no product executable, native sampler or desktop
activation and do not establish full Tooling, native, platform or repository acceptance.

The subsequent complete direct compiler-free Tools profile remains **FAILED**:
1469 passed, two failed and one skipped, contained child 67280 raw exit 1, Pester
3.4.0 and identical pre/post inventories of 28,472 repository files. Its closed
context is `20261120T103141Z-3676-2c6082783ddb4a4ea3c4f2eb91106e3c`,
`scratch/pure-tooling/current-profile-complete`. The audit failure exposed missing
classifications for the two new hidden private fixture parents; exact call-site
classifications and actual visible/wrong-path rejection controls passed the owning
audit Describe 2/0/0 (child 63408, exit 0). The workspace fixture clone failed at a
286-character path because the external driver used a deeper TEMP than the canonical
runner's `scratch/process-temp`; the unchanged workspace Describe passed 9/0/0
(child 10824, exit 0) using that canonical layout. No fixture or global Git setting
was changed. The external driver's later access to an unavailable Pester 3
`InconclusiveCount` property is a separate reporting error; raw XML/CLIXML and both
source inventories were already saved. The derived failure report is separately
sealed with SHA-256 `84D7E1A1430DF3CAEE1019A4D37E705950DBE6F8A57FEB10C5E5CE4EF9E0B144`.
The explicit missing-built-ProductUiTests failure-report probe remains skipped and
pending a build. These focused corrections do not turn the failed full-profile run
green. Direct Pester evidence has no canonical validation promotion or build-receipt
join and cannot replace required post-commit/profile-build PrePush tooling execution.

I65 remains **ACTIVE**. The bounded Release historical-behavior/candidate observation
pair, current-source Debug/Release builds and native scoped gates, actual ARM64,
the four broad desktop entries, applicable unfiltered Full, provider/SMB/NAS obligations
and normative closeout remain pending under their existing owners and contracts.
The previously authorized first six-hour desktop window is still unstarted. No
source commit, native build/test, workflow dispatch, publication, Done move or
generated changelog fold is claimed by this integration update.

### Public build integration repair after the failed A attempt (2026-11-20)

The first ordinary local fix commit is d89e51f9dc52e7f9ac044b3eef1644e2f2865c8a.
The serialized historical-behavior A attempt applied the reviewed product-only inverse,
then its canonical selected x64 Release build FAILED: contained build child 63344 exited
1 with 3,056 vcpkg overwrite warnings and one Terminal MSB3073 error after the locked
aro download timed out at Codeberg (curl exit 28, zero received bytes). No compatible
build receipt was published, no cadence case or desktop lease started, and no A negative
premise or A/B measurement was established. The finally path restored the B source bytes
and clean d89e identity; it did not restore or accept build outputs or receipts. The raw
attempt remains in completed context
20261120T112303Z-59380-b298168aad9b475f9b92075856a90be3,
scratch/qualification/paused-rate-A; the selected failed-output preservation supplement is
20261120T113719Z-42056-49a7a86a4e9f47799b9a0708cb405b54,
scratch/qualification/A-failed-build-preservation.

The dependency warning cause is a validated public build integration defect, not merely
an external driver omission. The documented installer promoted staged payload without
its status database; plain build.ps1 then ran the bundled Visual Studio root's manifest
install into that populated canonical destination. The executable bytes matched, but
four helper-script identities and six package ABIs differed, with actual DLL differences.
All 3,056 warning rows were overwrite warnings, not an exemption from the zero-warning
gate. The root/helper/status/ABI joins are preserved in completed context
20261120T114239Z-69716-9c4c4aed99ae4509873350114c1fa19b,
scratch/audit-forensics/A-vcpkg-warning-cause (report SHA-256
BA035035D71ACD2AFDD364D6E507D2BBF1559625450184B9C0F2A4C076D58770).

The accepted public repair validates prepared inputs, verifies canonical managed pins
and applies child-only pinned integration to direct/streaming MSBuild and receipt
metadata queries. It retains SDK raw-override refusal before composition, removes
ambient CL/LINK concatenation and preserves existing locks and the zero-warning gate.
It adds no package installer under a dependency lock, readiness cache, fabricated
status/stamp or bundled-root fallback. Its 14-file source packet is
20261120T120242Z-71584-eef48f84b7fe467d8f0eb292419c18e5,
scratch/audit-repair/public-vcpkg-source-freeze-v1, patch SHA-256
638AC3151EED41C89B98AF02E5715AC6AAC133803938FE2200F2CA1454B5FBF5.
Root and independent source review accepted the exact frozen bytes. Source-bound pure
Pester controls passed 16/0/0 with stable code/test hashes in completed context
20261120T115749Z-65112-51c65ec1c93140a2ac7b22117f85e151,
scratch/pure-review/public-vcpkg-focused-v2. The prior 12/2 fixture-development run is
retained. This implementation/pure state grants no native or broader tooling acceptance.

The person actually ran the manual prerequisite V3 and reported: "The property
'PSIsContainer' cannot be found on this object". The deterministic first-root walker
failure was reproduced before five-input verification or copying; no transferred cache
or proof is claimed. The corrected V4 uses plain .NET ancestor traversal and the first
canonical tar application. Its exact script SHA-256 is
2CDF039FF300319526C7CD0B190A828377BAC8B6380C562B496BFDC3E04F8947.
Safe actual path/resolver/negative controls passed 14/0 in completed context
20261120T120300Z-44780-2bee4676522a4dd48510b2a4bf73c11e,
scratch/manual-prerequisites/walker-v4; the unchanged earlier nine footer controls remain
applicable. Root and independent source review accepted V4. The person subsequently
executed the exact V4 script under PowerShell 7.6.10 (operator PID 8388). Its canonical
archived report is in completed context
20261120T121258Z-8388-a49315a07a2440bf9f6298856c675abd,
scratch/manual-prerequisites/verified-Ghostty-inputs/manual-verified-inputs-report.json,
SHA-256 0264C95A3873E613099FA81AE262F21C9402C3D1C785B2C2E740BF0AB4E705BB.
CopyRequested and AllFiveVerified are true: all five layout/license readers exited zero,
five copy rows were recorded, and Error is null. Independent destination hashes match all
five locked inputs (2,311,712 bytes); canonical archive complete/intent records join the
operator report and the original live run is absent. The derived reconciliation is in
completed context 20261120T121413Z-10800-166f276ea7ad4d11b116ea6361a3c694,
scratch/manual-prerequisites/V4-operator-reconciliation/operator-reconciliation.json,
SHA-256 E818F1E0A9B9055C5DCA0BB5FE2983B8A9E6B6E65117E13E420CBF5AD6417D5A.
This verifies the prerequisite cache transfer only; it establishes no build, A/B or native
acceptance. Agents independently read the report and destination hashes without rerunning
the archive/layout/license operations or copying the inputs.
A prior agent archive/layout/license command was rejected before execution by automatic
review with only "blocked by policy"; agents have not retried that action.

This final integration metadata refresh records the canonical workflow digest and
required structural checks in the integration freeze packet. Public modules/tests remain
at their accepted frozen hashes. No second commit, package-payload repair, further
source inverse, native build/test, workflow dispatch, publication or changelog fold is
claimed here. Required canonical post-build PrePush, the behavior pair, current-source
Debug/Release/native ARM64, broad desktop/Full/provider obligations and I65 closeout
remain pending. The first desktop window remains unstarted; I65 remains ACTIVE.

### Build-console lifetime follow-up prepared; merged-source qualification pending (2026-11-20)

The B3 lifecycle helper preserves every original console/path assertion, validates
borrowed canonical active ownership before creation, and archives only its standalone
run with primary/fatal archive failures retained. Canonical Pester 3.4 owning coverage
passed **17/0/0**, actual contained child 79528 exit 0, with source hashes unchanged
through execution (run `20261120T153151Z-79528-df61e42c82e7426da5f68ea3ba0541bf`).
The original end-to-end assertion body and representative current canonical module-closure
fixture ran. Boundary mocks prove borrowing/refusal/error behavior; they are not real
archive proof. Earlier outside-candidate 13/0 plus getter-delta5/0 remain separate
preparation evidence. The first owning attempts 4/13 and 16/1 remain failed history;
new mock assertions now explicitly use the current It scope, and the dead-owner control
inspects the actual exception instead of Pester 3.4's empty-message Throw adapter.

The reviewed [qualitative Source pair](../../TestRuns/Local-x64/FileOps/2026-11-20_142303_PausedRateSource/README.md)
is supporting state-correctness and bounded-work evidence only. A source `c88af9df`
(the reviewed five-hunk inverse in three product files) built with zero warnings/errors
under its own `eddbd46d` receipt, then remained **FAILED**, native/runner exits 1/1,
2 passed/1 failed/0 skipped. Its fresh nonzero paused rate stayed stale with no telemetry
timer and no collection. The deliberate ZERO repeated-rate ordering control missed
collection/reset. B source `736eff83` on `ebebc14b` built with zero warnings/errors under
its own `870be6db` receipt, and the exact filtered case passed **3/0/0**, native exit 0,
runner/operator exit 2: change verdict PASSED, repository NOT_EVALUATED. B cleared its
fresh nonzero paused rate before another telemetry timer with one collection and reset
the repeated ZERO-rate window with one collection. Both halves retained the original
600 ms settle, 2200 ms measurement and 3/4/4 caps; actual paused counts were 2/2/2.
Neither reproduced the earlier ten-render symptom. Single CPU observations (A 0,
B 15,625 microseconds) are not benchmark samples or a speedup claim.

All five common-input stages per half retained `8329dd08` (35,584 files, 199 command
tlogs, 465 retained producer read-witness paths and 3,056 prepared payload files).
The V5-to-V6 operator change was reporting selection only, with 15 safe controls;
functional build/case/control code stayed unchanged. Each generated SDK/module identity
matches its own receipt; no output swap, receipt relabel or executable-code-equivalence
claim is made. Native metric commit fields remain empty; derived sidecars bind actual
source and receipts. The original A parent selector failure and later read-only exact-ID
reconciliation remain separate; Capture.StartedUtc denotes only a held invocation/launch
window, not observed kernel ProcessStartTime. The tool-host rejection before B launch,
curator-only preparation failures and their unchanged raw outcomes are retained privately.
Final source/pin/output checks and owned actor/lease retirement completed; full B2 source
was restored. B3 changes to tooling/docs do not promote these B2 branch receipts.
The combined public subset supersedes only the A-only preparation's incorrect
PauseProjectionUs.value1 description: it is the rateCleared && noTimerTick success flag.

The ebeb selected Release build (zero warnings/errors) and Review native3/0/0/native0,
filtered outer2/NOT_EVALUATED retain their original branch identities. Historical whole
Tools1469/2/1 stays FAILED; SandboxAudit119 stays unclean. Debug/Release Full, exact merged-M
preparation, ARM64/desktop, PrePush and closeout remain pending. PR225 source45fbc76214d24c9869f129fef382b19bb326eb19
merged as53d0a9d58257f53efe817f52a3c309a25454ca1b; H8/API4 remains unchanged. No version6
port, producer repin, output swap or receipt rebinding is part of this follow-up.

### Pester copied-module fixture isolation (2026-11-23)

The hosted Q ToolsPesterTests failure came from the ordered pair: canonical root module
imports, the real bp28 copied-module BuildConsole fixture, then the global
ArtifactOperationLock import and `InModuleScope ArtifactOperationLock`. Before the
fixture the lookup succeeds; after it, all six affected modules also exist under the
copied `Tools/Modules` path, and Pester 3.4.0 reports multiple named script modules.
The unpatched reproduction is archive run
`20261123T113224Z-43112-135e0a49391547579eef1dd03570d593`;
`scratch/I65/pester-module-ambiguity/attempt-08/stdout.txt` SHA-256 is
`7F72A8214817FDB269D43CE06E56F546426F17A6707A9A0283553BFFD9158C3F`; module inventory
`scratch/I65/pester-module-ambiguity/attempt-08/after.json` has SHA-256
`73AC05D76A91B1A14158D23AFEB0F0A0F0C659BD397BBFA96908150C14CD3F92`. That hidden
PowerShell run was not proven Job-contained.

The fixture now uses `Invoke-RSTestFixtureNative` to start a contained no-window child
using the same PowerShell executable as the runner, with the marked run and TEMP
inherited. The owning test asserts the helper's zero exit and success output, then
verifies no module path under the copied fixture remains in the caller. The canonical
ordered replay passed 18/0/0 with native exit 0 in run
`20261123T114622Z-84816-db3b981db67f445f92681108f8d4c7e0`; its
`scratch/I65/pester-isolated-ordered-final/ordered-test-output.txt` SHA-256 is
`8BFCF13B532AC77D5797A4215DC71B95138EF4AA336F16BE6417F3D4C3716757`. A prior 17/1
replay imported the target module from shared main instead of this isolated worktree;
it is retained as a mixed-root driver error, not a product result. This focused result
verifies the fixture isolation fix; native CI remains pending. I65 remains ACTIVE;
merged-source Debug/Release Full, PrePush, ARM64, desktop and provider gates remain open.

### Scripted-vcpkg query compatibility (2026-11-23)

The merged-M Full run `20261123T122737Z-84824-3896c883c2c34a50ac58d81e877d7375`
exposed two stale tooling contracts after PR #240: the residual-process classifier
still required removed `VcpkgManifestRoot`, and the SDK adapter test still required
the removed CL-deletion guard instead of PATH-only `-RuntimeOnly` binding. The narrow
follow-up checks current `VcpkgEnabled=false` / `VcpkgEnableManifest=false` policy and
the actual RuntimeOnly helper and environment application. The two exact Pester 3.4.0
cases passed 2/0/0 with exit 0 in archived run
`20261123T133037Z-91472-a6869b3e9595449fb5bd27baba6b2110`; output SHA-256 is
`5dc4bd84b075f2a7087ef0c5652d70e12fc390692cfffe9f7f6ffdadae131288`. The PR
execution digest is refreshed. Broader Full and hosted qualification remain pending;
this does not qualify a product build or the separate runtime-copy failure. I65 remains ACTIVE.
