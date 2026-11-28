# I66 File Operations Workflow and Plugin Contract

**Status: ACTIVE.** Final target specification and delivery plan.
**Priority: P1.** Restore ordinary Rename first. Land the small ingress fixes and the progress correction independently.
**Planned at:** `ac619f860` (2026-11-23). Run the [drift check](#drift-check) before each delivery.

## Implementation checklist

### Current qualification

- [ ] Finish hosted native qualification on integrated main. The exact Pack merge and its two queued descendants were superseded at admission before native tests. Main `ae26292ae` carries exact Pack cleanup and both hosted x64 follow-up repairs. Run `46085203828` admitted that source and started both native regression jobs at 00:56 CET. Its x64 job finished with 3,045 passes, one failure and 72 skips. The sole failure was the old Pack source guard. Commands, File Operations and Viewer PE passed. ARM64 reached the five-hour regression timeout after recording one original Commands failure and one File Operations failure. [x64 result](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#hosted-x64-result-on-the-viewerweb-follow-up-tree) Follow-up run `46089888885` on `73c70e330` includes the Pack source-contract repair. It admitted and both native jobs have started. [Follow-up qualification](https://github.com/RedSalamanders/RedSalamander/actions/runs/46089888885) The preceding native run on `651a700fc` passed x64 Commands and File Operations. Its x64 qualification failed the HWND and ViewerWeb checks. ARM64 also failed the HWND check, then Commands retirement refused a foreign WSL console. [Pack merge admission](https://github.com/RedSalamanders/RedSalamander/actions/runs/46076994462) [Current x64 and ARM64 qualification](https://github.com/RedSalamanders/RedSalamander/actions/runs/46085203828) [Follow-up evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#hosted-x64-follow-up-after-the-commands-trace-fix)
- [x] Merge the HWND source-contract and ViewerWeb origin-observation follow-ups. PR #389 corrected the Pester assertion to match the existing COM helper. PR #390 separated controller startup from exact browser navigation, built x64 Release with zero warnings and errors, and passed the non-activating Viewer PE entry. The original ViewerWeb symptom was not reproduced locally. [ViewerWeb evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#viewerweb-private-origin-follow-up)
- [x] Build both hosted test-enabled Release profiles for the corrected Pack tree. In run `46089888885`, x64 and ARM64 finished building and started native regression and the Release assertion probe. Their suite verdicts are pending. [Hosted qualification](https://github.com/RedSalamanders/RedSalamander/actions/runs/46089888885)
- [ ] Follow the timed-out ARM64 regression. Its original Commands failure captured the re-expanded Completed card before it published a paint. The isolated case passed. Shuffle attempts reported focus and Terminal creation failures. File Operations recorded a foreground stability failure before the five-hour timeout. The corrected hosted run has admitted and started both native jobs. [ARM64 timeout evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#hosted-arm64-timeout-on-the-viewerweb-follow-up-tree)
- [x] Replay the three timed-out ARM64 run cases on the corrected x64 Release tree. Completed group and navigation, popup hosted-tree stability, and Terminal settings each passed all three repetitions with zero failures or skips. The desktop cleared at 06:50:07 CET. [Timed-out cases replay](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#focused-replay-after-the-arm64-timeout)
- [x] Merge the Commands Dismiss check after the expanded card publishes its body. PR #398 retained the action and first-frame accessibility assertions and merged at `02755d02b` after three exact x64 Release repetitions passed. [PR #398](https://github.com/RedSalamanders/RedSalamander/pull/398) [Focused evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#completed-card-paint-and-queue-desktop-checks)
- [x] Recheck Start now and Pause all on the same x64 Release build. The selected task released its queue gate while its neighbor kept waiting. The popup paused and resumed the active transfer. Both focused selections passed with setup and cleanup. The desktop cleared at 07:18:31 CET. [Queue and pause evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#completed-card-paint-and-queue-desktop-checks)
- [ ] Verify the Start option from an idle Queue default. The pane commands currently derive the dialog's initial choice from whether another task is active. The x64 Release test-only build adds a witness for the initial choice and its scope across consecutive Copy and Move submissions.
- [ ] Verify the virtual-pane context Move picker from the actual menu command through Local publication. The existing Full run passed Local picker admission and same-folder Ctrl+drag. The added controlled case covers the virtual source, Cancel and the single Copy-only question.
- [ ] Follow the latest ARM64 Commands cases. Original results recorded late conflict-inspection teardown and ViewerSpace hover failures. The teardown replay passed, then the boundary refused restoration over a new WSL console. The hover replay and subsequent interactive suites remained pending. Both cases passed in the local x64 Release Full Commands entry. [Hosted ARM64 evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#hosted-arm64-follow-up-after-the-commands-trace-fix)
- [ ] Follow the two earlier ARM64 Commands failures. The recovered original results identify fallback Cancel all focus and the mouse-opened menu keyboard case. The separate retirement observation identifies the WSL image and leaves its launch ownership open. Both cases passed three local x64 Release repetitions. Native ARM64 remains open. [Recovered ARM64 evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#earlier-arm64-commands-and-desktop-trace)
- [x] Verify the identity-less Delete and Pack interlock, deferred exit, Copy-only source retention and the provider capability matrix on the Pack build. Four non-activating x64 Release selections passed setup, their cases and cleanup with zero failures or skips. [Follow-up evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#follow-up-checks-on-the-pack-build)
- [x] Run the approved focused desktop checks for clipboard Move admission, deferred Recycle consent, and the File Operations footer and mixed-result controls. All four x64 Release selections passed. The runners drained after 185 seconds inside the approved 12-minute window. [Desktop follow-up evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#approved-clipboard-consent-and-file-operations-desktop-pass)
- [x] Recheck controlled mixed Recycle and sensitive metadata loss on the ViewerWeb follow-up tree. The no-activation File Operations remediation run passed setup, its case and cleanup. It covered exact Recycle and permanent approval, cancellation, source retention after publication and a staged Move Skip. [Focused evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#controlled-mixed-recycle-and-metadata-recheck)
- [x] Identify the upstream issue for recurring hosted ARM64 WSL prompts. The runner-image tracker reports recurring WSL prompts and has an open image fix. RedSalamander's archived foreground observation confirms the WSL executable. It does not identify the process that launched it. [Runner image issue](https://github.com/actions/runner-images/issues/14264) [Proposed upstream image fix](https://github.com/actions/runner-images/pull/15362)
- [x] Reconcile the Stopping action list with fixed destinations. The popup specification and user guide agree that source and destination path inspection stays available while the task stops.
- [x] Verify the Pack source-contract correction in x64 Release. The corrected HWND assertion passed in the selected x64 Release Tooling entry. That entry otherwise found one stale Pack assertion which assumed all cleanup returned to the pane. PR #395 merged at `73c70e330` after checking the regular-file receipt and the directory worker separately. All 244 focused checks passed. The complete exact-build Tooling Pester selection passed and promoted all 1,519 checks, including the HWND and Pack guards. [Repair PR](https://github.com/RedSalamanders/RedSalamander/pull/395) [Tooling evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#pack-cleanup-source-contract-follow-up)
- [x] Recheck the Pack and Unpack round trip and the Viewer PE origin witness on the merged Pack repair tree. The archive case passed three repetitions and the viewer entry promoted without activation. [Exact recheck evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#non-activating-rechecks-on-the-merged-pack-repair-tree)
- [x] Finish and record the fresh x64 Release Full run on `05252ae38`. Full finished with 3,091 passes, 58 skips and two failures in the older HWND and Pack Tooling assertions. Commands, File Operations and Viewer PE Interactive passed. Release SQLite reported zero compiled interactive cases. The contained runner drained and restored the C manifest byte for byte. [Full terminal evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#local-pack-full-terminal-result)
- [x] Replay the two latest ARM64 Commands cases on the corrected x64 Release tree. Conflict-inspection teardown and ViewerSpace hover each passed all three repetitions. The contained desktop replay ended at 02:28:51 CET. [Focused replay evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#focused-replay-on-the-merged-pack-repair-tree)
- [x] Pass corrected x64 Release Full. The tree merged by PR #395 passed 3,093 cases with zero failures and 58 skips. All 18 entries promoted. Commands passed 965 cases and File Operations passed 306 cases, including the independent C and Z volume case. The runner drained at 04:29:37 CET and restored the C manifest byte for byte. [Corrected Full evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#corrected-x64-release-full-terminal-result)
- [x] Run the selected Retry, read-only replacement and metadata preservation desktop checks. All three x64 Release cases passed with their setup and cleanup, for nine passes and zero failures or skips. The launcher drained in 122 seconds within the approved eight-minute window. [Desktop replay evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#approved-conflict-and-metadata-desktop-pass)
- [x] Merge the regular-file archive read protection. Both Pack formats refused an active writer after exact cleanup consent and preserved the source and prior output. The versioned reader denied a new writer and refused a stale same-size revision. The x64 Release build had zero warnings and errors. All three non-activating archive repetitions and the Local File Operations remediation case passed. PR #381 merged at `e72f5f02f`. [Archive reader PR #381](https://github.com/RedSalamanders/RedSalamander/pull/381) [Archive read evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#local-archive-source-revision-protection)
- [x] Merge exact Pack directory cleanup. Three exact x64 Release repetitions passed ZIP and 7z cleanup, Cancel, source retention on an unproved tree, and zero removals after a late addition, a same-size rewrite or an ancestor identity change. The focused File Operations remediation check also passed. PR #384 merged at `b632fbf29` after the required checks passed. [Pack directory PR #384](https://github.com/RedSalamanders/RedSalamander/pull/384) [Archive evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#exact-pack-directory-cleanup)
- [x] Recheck the registered RubyGloom Samba and SINON Windows SMB fixtures. Samba corroborated Rename, Windows SMB exact Rename, and Local Native plus SMB Managed Move on both shares passed setup, the live case, and cleanup without activation. The Windows SMB selection restored the original resource registry byte for byte. The isolated SFTP and OneDrive Personal attempts skipped because their connection profiles were unavailable in the run. [Live recheck evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#follow-up-live-provider-checks)
- [ ] Complete the live provider gates. Finish SFTP Batch Rename and Change Case qualification, FTP and SCP sign-in and Rename, and the live Google Drive, OneDrive Business and SharePoint checks. The automatic approval review rejected disabling Windows Hello for the isolated FTP and SCP profiles. Google Drive, OneDrive Business and SharePoint test profiles are currently unavailable. The historical OneDrive Personal fixture was recovered separately. Finish the remaining mixed Move endpoint cases.
- [x] Merge the Commands diagnostic trace freshness correction. The runner compares the guarded file state before launch and after containment closes. A disk timestamp earlier than the launch clock no longer suppresses a newly written trace. The first canonical attempt found an incomplete inventory fixture. Its corrected x64 Release run promoted 1,519 checks with zero failures or skips. PR #375 merged at `320f1dc04`. [Merged trace repair PR #375](https://github.com/RedSalamanders/RedSalamander/pull/375) [Trace evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#commands-trace-freshness-correction)
- [x] Preserve both hosted verdicts on `42691d6bc`. x64 Release reported 3,030 passes, two flaky failures and 71 skips. Both retry checks passed. ARM64 reported two original Commands failures and two different failures in the shuffled replay. Its isolated retry children passed. Desktop retirement refused a new WSL instance after the shuffle, after the runner had already retired an earlier confirmed startup console. [Hosted evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#hosted-qualification-after-the-commands-diagnostic-merge)
- [x] Merge the Pack source preparation and retained-originals alert. Exact x64 Release built with zero warnings and errors. Three non-activating repetitions verified both ZIP and 7z alerts, publication and source retention. All 12 localization checks passed. PR #376 merged at `ab4523996`. [Merged PR #376](https://github.com/RedSalamanders/RedSalamander/pull/376) [Pack evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#pack-shared-preparation-and-retained-originals-alert)
- [x] Merge one decision for known pre-content EFS and sparse allocation losses. The exact x64 Release candidate passed three non-activating repetitions with one combined question each and preserved both payloads after Copy and keep source. The repeat decision matrix passed on the same receipt. PR #377 merged at `b7bea3b68`. The full metadata and conflict walkthrough remains pending desktop admission. [Merged combined metadata PR #377](https://github.com/RedSalamanders/RedSalamander/pull/377) [Metadata evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#combined-efs-and-sparse-allocation-consent)

### Restart qualification (2026-11-27)

- [x] Verify the saved x64 Release artifacts. The canonical runner refused the interrupted profile before taking focus. The independently sealed solution build on PR #348 head `96e495135` passed byte verification and became the desktop basis.
- [x] Preserve the first local x64 Release attempt. The older build's Tooling Pester entry reported 1,491 passes and five failures. The approved window closed at 08:55 CET during classification, before any of the six desktop entries started. The contained processes drained.
- [x] Record the approved x64 Release desktop window on merged `3f2b01251`. The build passed with zero warnings and errors and all 188 artifacts matched receipt `946bd916ef3e93a02ca5abcf9868baa1f45f8eeac6462611d1e8982047a1b47c`. The follow-up CI run `20261127T080428Z-67632-151d4572a667447aa54cc951dd320513` started at 09:04 CET. At 10:54 CET the contained launcher closed the interrupted retry and the process tree drained.
- [x] Complete local Commands on merged x64 Release. The canonical terminal record reports 965 passes, zero failures and five skips.
- [x] Preserve the first local File Operations failure. The original process exited with `0xC0000374` while the nested-roots delete test was allocating a fixture path. The exact trace, dump and symbolized stack are archived. The location of the original corruption is undetermined. [Crash evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#local-file-operations-native-failure)
- [x] Correct the confirmed Local Delete context alias. The crash dump shows 17 scheduler workers. The old Delete path assigned worker 16 to worker zero's mutable context. The candidate uses the full scheduler range and checks completions for each selected root. [Merged PR #360](https://github.com/RedSalamanders/RedSalamander/pull/360)
- [x] Verify the Local Delete repair. Five exact x64 Release runs passed without activation. They completed 100 nested-delete rounds with one successful result per selected root. The repair merged at `3a82609a0`. [Local Delete evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#focused-local-delete-repair)
- [x] Execute the broader local File Operations replay on the repaired x64 Release build. The nested-delete stage passed in the full suite. The runner exited normally with 300 passes, 29 skips and one failure. The independent-volume witness required a flag that the launch omitted.
- [x] Verify the independent-volume witness with the approved marked C and Z roots. The explicit alternate-volume run passed setup, R4A19 and cleanup. The launcher restored the resource manifest byte for byte. The broad replay retains its original failure. [Independent-volume evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#independent-volume-recheck)
- [x] Run the four selected viewer entries in the approved x64 Release window. Viewer Text Find and Go To passed. Viewer PE stopped after refusing to restore an unowned foreground in both attempts. Viewer SQLite exited successfully with zero compiled interactive cases in Release. [Viewer evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#dedicated-viewer-desktop-selection)
- [x] Merge the exact viewer selector and verify both test-enabled x64 profiles. Release and Debug builds completed with zero warnings and errors. All 188 listed outputs in each profile verified. [Merged PR #362](https://github.com/RedSalamanders/RedSalamander/pull/362)
- [x] Confirm Viewer PE Interactive in x64 Release and the SQLite Tab witness in x64 Debug. Each selected entry passed inside the approved 12-minute window, and the owned processes drained. The SQLite console confirms forward and reverse Tab traversal and Escape focus return. Both selections retain repository verdict `NOT_EVALUATED`. [Approved viewer recheck](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#approved-viewer-recheck)
- [x] Recheck the previously failing S3 multipart witness locally. The promoted Plug-in Contract entry on merged x64 Release reported 50 multipart passes and zero failures.
- [x] Recheck the previously failing ViewerSpace blocked-provider close witness locally. The non-interactive Viewer PE entry on merged x64 Release passed the 500 ms close assertion and verified provider exit after release. The failed historical Full verdict remains recorded.
- [x] Pass the complete local Tooling Pester entry on the merged x64 Release build. The follow-up reported 1,500 passes, zero failures and zero skips before starting Compare Directories.
- [x] Read both hosted job verdicts and check the archived run states. On `09c663878`, x64 completed all six interactive entries and its tooling suite failed six checks. ARM64 recorded two Commands failures after foreign-console interference. The ViewerSpace hover case also failed when repeated alone. Retirement then refused to restore over the foreground console. The archived state records failed tooling and pending later desktop entries. [Hosted x64 and ARM64 qualification](https://github.com/RedSalamanders/RedSalamander/actions/runs/45907308619)
- [x] Run the read-only ARM64 desktop probe. The host reported the Default input desktop, with a separate foreground process present at the sample. The repository verdict is `NOT_EVALUATED`. [ARM64 desktop probe](https://github.com/RedSalamanders/RedSalamander/actions/runs/45958485341)
- [x] Repair the five RedSalamander tooling failures. The focused Pester 3.4.0 rerun passed 312 checks with zero failures or skips. [Restart and hosted evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md)
- [x] Reproduce and repair the shared PrePush fixture failure. The corrected fixture passed all 40 cases on the product. [Shared-engine PR #130](https://github.com/RedSalamanders/.github/pull/130)
- [x] Integrate shared engine `317dfd0f` and rerun the affected tooling suites. Pester passed 315 checks with zero failures or skips. The archive inventory passed for 18,765 files.
- [ ] Complete hosted native qualification. On `8c3c56801`, x64 Release passed 3,028 required tests and the Release assertion probe, with 72 expected skips. ARM64 Commands exited 1 with two failures during foreign-console interference. Desktop retirement refused restoration, leaving File Operations and the four interactive viewer entries pending. [Hosted replay](https://github.com/RedSalamanders/RedSalamander/actions/runs/45993909744) [x64 results](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#hosted-x64-success-after-live-provider-qualification) [ARM64 results](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#hosted-arm64-foreground-failure-after-live-provider-qualification) [Earlier ARM64 callback failure](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#hosted-arm64-commands-callback-failure-on-the-smb-rename-merge)
- [x] Qualify and merge the Commands crash diagnostic. Twelve focused checks passed. The exact test-enabled x64 Release solution built successfully, and the canonical selected Tooling Pester entry archived 1,504 passes with zero failures or skips. PR #370 merged at `a9ca4a647`. The diagnostic preserves the last observed case boundary before desktop retirement. [Merged PR #370](https://github.com/RedSalamanders/RedSalamander/pull/370) [Tooling evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#commands-native-exit-diagnostic)
- [x] Verify the existing Pack directory-retention behavior on hosted x64 Release. The archive command case passed in 2,284 ms, covering complete ZIP and 7z output and retaining directory originals without cleanup consent or Delete when the tree receipt is unavailable. [Hosted x64 evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#hosted-x64-success-on-the-smb-rename-merge)
- [x] Finish integrated non-interactive x64 Release PrePush on merged `fc3fe2499`. All 14 selected entries promoted with 1,801 passes, zero failures and 30 skips. Both the original build and re-attestation finished with zero warnings and errors. The repository remains unevaluated, with four desktop entries pending. [Integrated PrePush evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#integrated-non-interactive-x64-release-prepush)
- [x] Merge the qualification tooling repairs. All required checks passed and main advanced to `3f2b01251`. The plan retains its desktop, native, live provider and exact archive cleanup gates. [PR #357](https://github.com/RedSalamanders/RedSalamander/pull/357)

**Live qualification checkpoint (2026-11-26).**

- [x] Verify one Local Native item and one Windows SMB Managed item in the same Move. Exact x64 Release passed 3/0/0 without activation, verified both destination payloads and source removals, and retired the SMB fixture. The Samba exact-identity control skipped before creating its fixture. [PR #348](https://github.com/RedSalamanders/RedSalamander/pull/348) [Mixed Move and control evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-26_221027_live_mixed_smb/README.md)
- [x] Recheck SFTP F2 Rename, both SMB Rename routes, and OneDrive Personal Rename and transfers. All five scoped runs passed 3/0/0 without activation. The isolated SFTP and OneDrive test profiles retired with their installed settings digests verified. [Live recheck results](../../TestRuns/4cb089111a23/FileOps/2026-11-26_221027_live_mixed_smb/README.md#existing-live-route-rechecks)
- [x] Verify the live SMB parent-replacement and same-name decoy checks. The x64 Release witness passed 3/0/0 on natural Samba and controlled Windows SMB without activation. Both protected the original and decoy bytes and retired the fixture. [Merged PR #359](https://github.com/RedSalamanders/RedSalamander/pull/359) [SMB namespace evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#live-smb-namespace-witness)
- [x] Verify live SFTP Batch Rename and Change Case refusals and the working F2 route. Exact x64 Release passed both warnings, file, case-only and folder F2 Rename, occupied-name refusal and cleanup. [Merged PR #364](https://github.com/RedSalamanders/RedSalamander/pull/364) [SFTP evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_122818_sftp_scheduled_rename_refusal/README.md)
- [x] Verify mixed Local Native and SMB Managed Move on the two approved shares. Exact x64 Release passed 3/0/0 on Samba using corroborated source facts and on Windows SMB using exact identity. Both runs verified destination bytes, absence of both originals and fixture cleanup. [Merged PR #365](https://github.com/RedSalamanders/RedSalamander/pull/365) [Live mixed Move evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_123257_samba_mixed_move/README.md)
- [ ] Finish the remaining mixed Move endpoint scopes and shared remote Rename routes. The controlled Native plus Copy-only Move is recorded separately from the live Native plus Managed SMB cases. Positive SFTP Batch Rename and Change Case require additional provider authority. The other remote Rename and sign-in gates remain open.
- [ ] Finish x64 Full desktop and hosted native qualification. The local Full run recorded standalone S3 multipart and non-interactive ViewerSpace blocked-provider close failures. Both hosted regression jobs reached native test execution on `94516001b`. [Hosted qualification](https://github.com/RedSalamanders/RedSalamander/actions/runs/45890003061)

**Release checkpoint (2026-11-26).** Live single-item SFTP Rename, OneDrive Personal transfer and guarded SMB Recycle passed in x64 Release and merged. The refreshed non-interactive PrePush on `3eed44771` found three outdated tooling checks. The corrected focused checks passed. The frozen PrePush rerun passed all 14 non-interactive entries. Desktop, live provider and hosted native platform gates remain open.

- [x] Run the broad non-interactive x64 Release PrePush after the R0e repair. The frozen `2d41ebc55` run reported 1,801 passed, zero failed and 29 environment skips. The runner marked the change `PASSED` and the repository `NOT_EVALUATED`. [Release validation record](../../TestRuns/4cb089111a23/FileOps/2026-11-26_163210_release_prepush/README.md)
- [x] Capture the refreshed non-interactive x64 Release PrePush verdict on `3eed44771`. It reported 1,800 passes, three Tooling Pester failures, and 29 environment skips. Both recorded verdicts were `FAILED`.
- [x] Repair the gallery registration, mode-aware Graph upload assertion and wrapped Preferences focus assertion. Classify the exact hidden remote Rename fixture. Corrected `4aa1629a5` passed all 275 checks in the two affected Pester suites. [Source check evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-26_192821_release_source_checks/README.md)
- [x] Recheck the source fixes after integrating the new build-tooling baseline. The retired workflow query also needed a source-test and shared-helper update. Corrected `4386c570c` passed both suites with 275 passes and no failures or skips. [Integrated evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-26_192821_release_source_checks/README.md#build-tooling-baseline-integration)
- [x] Rebuild the new tooling baseline and verify native inventory and live SFTP on `1611a9086`. x64 Release had zero warnings and errors. All 152 runner checks and the authorized SFTP rerun passed. The isolated Hello-off profile retired after verifying the installed settings digest. [Integrated native evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-26_192821_release_source_checks/README.md#integrated-x64-release-checks)
- [x] Finish the non-interactive x64 Release PrePush `20261126T183940Z-47540-d7fc757cdfcb4c6db9285c919d3e9c23` on frozen `4aa1629a5`. All 14 scheduled entries passed with 1,803 passes, zero failures and 29 environment skips. The runner marked the change `PASSED` and the repository `NOT_EVALUATED`. The four desktop entries remain pending. [Release evidence in PR #341](https://github.com/RedSalamanders/RedSalamander/pull/341)
- [x] Verify the repaired OneDrive Personal upload route through live Copy, same-drive Move and Recycle. The scoped x64 Release case passed all stages and cleaned its fresh fixture. [Live evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-26_172940_onedrive_personal_upload/README.md)
- [x] Verify guarded Recycle on the approved Samba and Windows SMB fixtures. Both exact x64 Release runs passed 3/0/0 without activation. Samba offered by-name consent and Windows SMB offered exact-identity consent. Skip kept the original bytes and each run removed its fresh fixture. [Live SMB Recycle evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-26_182743_smb_recycle_refusal/README.md)
- [x] Align the User Guide and Main Window shortcut summary with F6 Move, F2 Rename, background F7 creation and the guarded Recycle confirmation.
- [x] Qualify live single-item SFTP Rename and fixture cleanup. Exact x64 Release passed scheduled file, case-only and folder Rename, occupied-name refusal and payload readback without activation. The isolated Hello-off profile retired after verifying the installed settings digest. [Live SFTP evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-26_190427_sftp_inline_rename/README.md)
- [x] Verify the pending-decision pointer route and Start now in the approved x64 Release desktop window. Both exact cases passed on `cf21187a8`, and the owned processes retired in 85 seconds. [PR #328](https://github.com/RedSalamanders/RedSalamander/pull/328) [Desktop evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-26_203625_disclosure_start_now/README.md)
- [ ] Complete the broader desktop walkthrough. The two focused cases passed locally and in the hosted x64 diagnostic. Local non-interactive PrePush left Commands, FileOperations, ViewerPETests.Interactive and ViewerSqliteTests.Interactive outside its scope.
- [x] Repair the Local selection counting case that returned partial Copy in the hosted diagnostic. The new witness cancels before payload admission. Exact x64 Release `59c3d9851` passed the six concurrency scenarios and the 4,097-file counting, cancellation and empty-destination assertions without activation. [Counting evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-26_210415_selection_cancel/README.md)
- [ ] Obtain hosted native x64 and ARM64 Release qualification on the integrated main revision. [Qualification on `09c663878`](https://github.com/RedSalamanders/RedSalamander/actions/runs/45907308619) recorded tooling failures on both architectures and stopped ARM64 during desktop retirement. The earlier ARM64 D10 child receipts prove natural shutdown and unchanged fixture bytes. The parent foreground check failed and recovery refused restoration over a foreign console. In the subsequent [native run](https://github.com/RedSalamanders/RedSalamander/actions/runs/45811712240), x64 finished with one Change Case gallery failure on source that predates the merged repair in PR #325. ARM64 failed the gallery and a separate UIA-route foreground assertion during the original Commands run. The isolated pointer-case retry passed its assertions and its outer desktop restoration refused a foreign console.

- [x] Reconcile the user and developer guides with the scheduled Batch Rename and Change Case execution and task cards.

- [x] Compare steady Stopping allocations on matched x64 Release builds. Both exact no-activation cadence runs passed 3/0/0. The held Local Copy kept its paths, completed bytes and Stopping status. The same rate-update shape allocated or reallocated 25 times per sample on both builds. [Stopping evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-26_154204_popup_stopping_allocations/README.md)

- [x] Recheck the worker archive round-trip after the gallery repair. Exact x64 Release run `20261126T143319Z-57112-b89f3e8ff3b84af485470db353a1309c` passed 1/0/0 without activation. The User Guide and format contract merged in [PR #327](https://github.com/RedSalamanders/RedSalamander/pull/327) at `e2ae2f190`.

- [x] Align the User Guide and archive contract with the two verified Pack formats, ZIP-only Unpack and background scheduling.

- [x] Restore the Change Case documentation gallery and recheck Files-menu admission on x64 Release. Both exact cases passed without activation. The gallery recorded all 45 scenarios and captured the production Change Case prompt. The old failure and remaining native gates are recorded alongside the passes. [Gallery and admission evidence](../../TestRuns/4cb089111a23/Commands/2026-11-26_150252_change_case_gallery/README.md)

- [x] Repair R0e validation for the exact Batch Rename witness. Focused Pester run `20261126T102400Z-47644-5cf79fb848d64d1daba52f58efefb7f9` passed all 20 checks, including the source-mutation controls and the injected archive failure. The revised witness runs before the timer. Paired diagnostics still require matching reviewed fixture projections. [R0e evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-26_112400_r0e_rename_fixture/README.md)
- [x] Qualify the exact worker-admission and archive interlock witnesses on x64 Release. Batch Rename rejected duplicate physical bindings and canceled its modeless worker. The queued identity-less Delete refused a changed pinned tree after parallel Pack publication. The provider recorded `ERROR_REVISION_MISMATCH`, and the archive survived. The combined Files-menu and provider matrix checks passed on the same receipt. [PR #312](https://github.com/RedSalamanders/RedSalamander/pull/312) [Worker witness evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-26_095011_worker_interlock_witnesses/README.md)
- [x] Verify the exact-only Pack directory retention contract in both ZIP and 7z. Both writers published all 16 planned entries and retained the originals without a cleanup prompt or Delete. A directory cleanup route still requires an exact descendant receipt. [PR #317](https://github.com/RedSalamanders/RedSalamander/pull/317) [Directory retention evidence](../../TestRuns/4cb089111a23/Commands/2026-11-26_095442_pack_directory_retention/README.md)
- [x] Verify active progress with a credential-waiting peer using two real Local Copy workers and a controlled credential observer. The popup returned to 100 ms after bounded card motion while the workers remained held. The running history sample recorded zero presentation rebuilds and zero issue scans. Three matching archived `UpdateRates` workload shapes had upper medians of 44, 33 and 33 µs earlier and 29, 28 and 28 µs on the new run. Live authentication remains open. The paired allocation windows are recorded below. [PR #318](https://github.com/RedSalamanders/RedSalamander/pull/318) [Credential peer evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-26_095648_credential_active_peer/README.md)
- [x] Compare popup heap allocations on matched x64 Release Prompting, Running and Paused windows. The pre-credential baseline and integrated runtime both recorded 21, 24 and 23 allocate or reallocate events per sample respectively. Both native runs passed without activation and the raw traces reported zero lost events or buffers. [Paired allocation record](../../TestRuns/4cb089111a23/FileOps/2026-11-26_122011_popup_rate_allocations/README.md)
- [x] Verify the Files-menu guard and the scheduled Rename routes. Exact Windows SMB committed Batch Rename and Change Case. Live Samba refused both before mutation when exact identity was unavailable and passed corroborated F2 Rename. The exact x64 Release Files-menu check observed the intended Dummy and Local command states without activation. [PR #311](https://github.com/RedSalamanders/RedSalamander/pull/311) [PR #315](https://github.com/RedSalamanders/RedSalamander/pull/315) [PR #313](https://github.com/RedSalamanders/RedSalamander/pull/313)

- [x] Verify folder and individual-file selection parity on live Windows SMB, Samba and Local. The exact x64 Release cases passed without activation on the same verified build. Both SMB shares reached four active files for native Copy, host bridge Copy and managed Move. Local Copy reached four under Manual four and controlled Auto, and one under Manual one. All 18 scenarios verified content and the required source outcome. The Windows fixture is empty and the original resource manifest is restored. [Windows SMB, Samba and Local selection evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-26_113221_smb_local_selection_shapes/README.md)
- [x] Recheck Copy-only Move source retention, remediation and repeat decisions on the integrated x64 Release runtime. The three guarded cases each passed setup, validation and cleanup without activation. [Integrated workflow checks](../../TestRuns/4cb089111a23/FileOps/2026-11-26_113221_smb_local_selection_shapes/README.md#integrated-workflow-checks)
- [x] Merge typed Recycle routing in [PR #249](https://github.com/RedSalamanders/RedSalamander/pull/249) at `e103ebcfe`. The exact x64 Debug mixed-selection and Move cases passed. Live remote Recycle coverage remains open.
- [x] Merge Create Folder admission in [PR #252](https://github.com/RedSalamanders/RedSalamander/pull/252) at `ccdab5fe5`. The exact no-activation x64 Debug case passed, including submission after exit consent while a name prompt was open.
- [x] Merge Options → Issues in [PR #255](https://github.com/RedSalamanders/RedSalamander/pull/255) at `95b295b37`. The approved x64 Debug case passed in `20261125T070538Z-18412-72ac67ec16b248b3a2ac40421f7084ff`. It verified opening and hiding Issues and exercising the existing footer choices.
- [x] Merge repeat choices in [PR #251](https://github.com/RedSalamanders/RedSalamander/pull/251) at `4bfdf12d`. The x64 Debug repeat and alert accessibility cases passed without activation. The broad run reported 1,802 passes, one tooling failure and 26 skips. All 18 streaming checks passed on the same head in isolation. [PR #263](https://github.com/RedSalamanders/RedSalamander/pull/263) repaired the timeout fixture and merged at `2e0e2fac`. The corrected streaming checks passed 18 of 18.
- [x] Merge clean Completed grouping in [PR #256](https://github.com/RedSalamanders/RedSalamander/pull/256) at `554039eae`. Integrated head `48478da4b` built x64 Debug with zero warnings and errors, passed all 66 runner-plan checks, and passed all three stages of `FileOps_ClearCompletedScope` without activation in `20261125T092444Z-81072-629893611b2041ec8e910515517df07d`. The preceding broad campaign had 1,802 passes, one tooling timeout and 26 skips. The isolated tooling suite passed all 27 checks.
- [x] Merge explicit-start Custom bandwidth in [PR #259](https://github.com/RedSalamanders/RedSalamander/pull/259) at `6a1fbb5a1`. The approved x64 Debug desktop check `20261125T080539Z-74184-e3d06a910a664a7e9c71c17736f4f791` verified Cancel in Copy and Move and accepted 768 KiB/s in Copy. The broad run recorded one unrelated ViewerSpace failure. Its exact no-activation rerun passed in `20261125T093500Z-82956-eebed5693b494cf3b5d2cc2c454613ca`.
- [x] Merge credential-wait progress in [PR #258](https://github.com/RedSalamanders/RedSalamander/pull/258) at `647c3a20d`. Integrated x64 Debug head `31e78b63e` passed exact cadence in `20261125T100435Z-88876-0992a2b497c24264a69ecfc97d9398af`. Later Release runs verified the actual host prompt and the paired allocations. Live connection authentication remains open.
- [x] Merge the Copy links default in [PR #264](https://github.com/RedSalamanders/RedSalamander/pull/264) at `30e38873f`. The integrated x64 Debug SettingsSchemaTests passed in `20261125T095345Z-72328-559fe48f31b946f082251cb84bd90080`, and all seven persistence inventory checks passed. The later Apply, Cancel and keyboard walkthroughs passed and have their own checklist entries.
- [x] Merge queued Create Folder in [PR #261](https://github.com/RedSalamanders/RedSalamander/pull/261) at `705a978bc`. The exact x64 Debug cancellation, numbered-sibling, and exit checks passed at `dd5626127` in `20261125T101258Z-72684-c10e117bc74b4eac8a21b448265da604`. All 12 focused localization checks passed. The slow Dummy provider and pane-switch controls passed in x64 Release run `20261125T211258Z-89608-9f21f98540ec4a3fb5da3f26b9a5c4fc`. The worker created the second folder under its original parent and kept the pane in the unrelated folder. [Create Folder navigation evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-25_221405_create_folder_navigation/README.md) [Worker timing record](../../TestRuns/4cb089111a23/FileOps/2026-11-25_094731/README.md)
- [x] Merge the exact permanent Leftover Cleanup card in [PR #266](https://github.com/RedSalamanders/RedSalamander/pull/266) at `8c7775db9`. The integrated x64 Debug case passed in `20261125T095557Z-51024-da4bd6fd609a42bd9eb817b90610b81d`. The trace records Cancel, acceptance, and changed-selection refusal without a question. [Exact selection record](../../TestRuns/4cb089111a23/FileOps/2026-11-25_093636/README.md)
- [x] Merge late Move metadata retention in [PR #268](https://github.com/RedSalamanders/RedSalamander/pull/268) at `3da376a2b`. The integrated x64 Debug run `20261125T102020Z-55628-7cfcfa3ddf1b4cb995f8c1242351723c` verified the published content, preserved the source and its Mark of the Web, and recorded no late question. [Paired record](../../TestRuns/4cb089111a23/FileOps/2026-11-25_100231/README.md)
- [x] Merge pre-publication Skip for sensitive metadata loss in [PR #272](https://github.com/RedSalamanders/RedSalamander/pull/272) at `f20f8a949`. Integrated x64 Debug run `20261125T113013Z-81044-5542320ea9104f68889e3984ad2dcc81` passed without activation. Skip preserved both file contents and the source's Mark of the Web, and removed the unpublished stage. [Staged Skip record](../../TestRuns/4cb089111a23/FileOps/2026-11-25_123209/README.md)
- [x] Merge the Copy links Apply fix in [PR #271](https://github.com/RedSalamanders/RedSalamander/pull/271) at `e8f2ef9a3`. Approved x64 Debug run `20261125T111700Z-29824-2b36177d77c441658c4ac645f0e99623` committed Skip on Apply, discarded a Preserve draft on Cancel, and reopened with Skip. The first failed run `20261125T104631Z-64028-0dabdbea2785432fb7fa3c4aed69adeb` remains preserved.
- [x] Merge the Release witness correction in [PR #260](https://github.com/RedSalamanders/RedSalamander/pull/260) at `4fe871577`. The focused x64 Debug cadence passed. The broad run recorded one failed ViewerPE entry. Its ViewerSpace and two ViewerVLC cases all passed on the same receipt in `20261125T093633Z-51556-fcc917bd41194076a2d5ef0825c52548`. Both hosted Release profiles still require qualification.
- [x] Merge exact Curl Rename receipts in [PR #267](https://github.com/RedSalamanders/RedSalamander/pull/267) at `13bbcf027`. The x64 Debug loopback prove-out passed all 49 assertions, the standalone plugin contract suite, and all 66 focused runner-plan tests. [Loopback record](../../TestRuns/4cb089111a23/FileOps/2026-11-25_100737/README.md)
- [x] Merge automatic overlap waits in [PR #276](https://github.com/RedSalamanders/RedSalamander/pull/276) at `5c7235e8c`. Exact x64 Debug Copy, cross-volume Move and review-remediation passed without activation. The repeated tasks kept separate IDs and waited behind the exact predecessor without a question. [Overlap evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-25_131602_overlap_waits/README.md)
- [x] Merge Microsoft Drive Rename provider receipts in [PR #273](https://github.com/RedSalamanders/RedSalamander/pull/273) at `bf9396fe8`. The x64 Debug provider and standalone cases passed, including 259 Microsoft Drive checks. [Microsoft Drive provider record](../../TestRuns/4cb089111a23/FileOps/2026-11-25_123539_graph_rename/README.md)
- [x] Merge Google Drive Rename collision refusal in [PR #274](https://github.com/RedSalamanders/RedSalamander/pull/274) at `e0108c683`. The paired x64 Debug loopback recorded four PATCH and two DELETE requests on collisions before the fix and zero after it. All 46 Google Drive checks and the standalone suite passed. [Google Drive provider record](../../TestRuns/4cb089111a23/FileOps/2026-11-25_124412_drive_rename/README.md)
- [x] Merge the optional same-parent inline Rename route in [PR #279](https://github.com/RedSalamanders/RedSalamander/pull/279) at `513d485ca`. On x64 Debug `7cb3e0b06` the owned Dummy route passed occupied-name Keep both, exact case, cancellation and a stale source with terminal receipts. The provider matrix preserved the default Dummy refusal, and the standalone contracts passed. Batch Rename and live provider admission remain open. [Host Rename evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-25_133545_probed_inline_rename/README.md)
- [x] Verify the OneDrive Personal loopback and provider contracts. Candidate `532215186` passed all 266 Microsoft Drive checks. Its first standalone run failed an independent Curl reader check and the same receipt passed on rerun. Integrated `a6ecbe9a2` built x64 Debug with zero warnings and errors. The owned host Rename passed all three stages in `20261125T133618Z-70668-71612e80240a47bb947774782083ba73` without activation. The integrated standalone plugin suite and 66 runner-plan checks also passed.
- [x] Verify saved sign-in and the approved OneDrive fixture. The user authorized a Windows Hello exception for the isolated copy. The first configured read-only run reached authentication and returned `0x8007052E`. The exact rerun `20261125T134004Z-64848-aeabc524b7aa42d8885d6e9ffb23c5fe` passed without activation after sign-in was renewed. Both runs verified the installed settings digest and retired the isolated settings file. [OneDrive preflight evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-25_144010_graph_live_preflight/README.md)
- [x] Qualify live OneDrive Personal F2 on x64 Debug `fa14fd835`. Run `20261125T141900Z-84592-d0fb0bbd1fd048a28fd376a3f18f03f7` passed setup, Rename, and cleanup without activation. It verified the occupied-name refusal, free file and folder Renames with preserved bytes, case-only spelling, and Recycle of its fresh fixture. The installed settings digest matched. [Live OneDrive Rename evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-25_151900_onedrive_live_rename/README.md)
- [x] Merge OneDrive Personal F2 in [PR #285](https://github.com/RedSalamanders/RedSalamander/pull/285) at `588a1814a`. Integrated `d6946334f` built x64 Debug with zero warnings and errors. Exact Graph loopback `20261125T143542Z-63612-649c44913fa14f6595600a04c72e3b5e` and F7/deferred-exit `20261125T143653Z-37252-6e6bf70668964893b79a8b02ee73e7a3` each passed all three stages without activation. Required PR checks passed.
- [x] Recover the earlier OneDrive Rename fixture. The second bounded recovery run verified Recycle and the exact folder's absence. An independent sync check confirmed it was gone. [Recovery record](../../TestRuns/4cb089111a23/FileOps/2026-11-25_160545_onedrive_recovery/README.md)
- [x] Fix the live OneDrive Personal Copy failure and verify Copy, same-drive Move and Recycle on the corrected route. Exact x64 Release run `20261126T162543Z-72592-86e623ea7ba0477f99504c3487776173` passed 3/0/0 without activation. The earlier native timeout and the measured origin rejection are retained. [OneDrive upload and round-trip evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-26_172940_onedrive_personal_upload/README.md)
- [x] Recover the earlier OneDrive Upload directory. The approved sync folder showed the Upload directory empty. The original exact recovery build then verified provider absence after nonrecursive Recycle. A second sync observation confirmed both historical fixtures were absent. [Recovery evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-25_160545_onedrive_recovery/README.md#2026-11-27-exact-recovery-recheck)
- [x] Verify a Copy-only Move retains the source when its destination gains content proof after admission. The x64 Debug run `20261125T132735Z-27364-4e82094f79da4bb49d7043ffce5bb3fb` passed without activation. One Copy-only question authorized the task. The source and published destination had equal bytes. [PR #282](https://github.com/RedSalamanders/RedSalamander/pull/282) merged at `f33851f7c` after the required checks.
- [x] Verify the registered SFTP fixture and saved sign-in. The authorized isolated x64 Debug read-only run `20261125T143148Z-25884-078b682010554393a6607da084b0432f` passed without activation on `af8eda155`. The installed settings digest matched and the isolated settings file was removed before archival. [SFTP fixture access evidence](../../TestRuns/4cb089111a23/CompareDirectories/2026-11-25_153148_sftp_preflight/README.md)
- [ ] Finish live Batch Rename and Change Case qualification for SFTP, and live Rename qualification for FTP, SCP, Google Drive, OneDrive Business and SharePoint. SFTP single-item F2 passed its own writable fixture and cleanup checks in [PR #337](https://github.com/RedSalamanders/RedSalamander/pull/337).
- [x] Merge the Completed-group tooltip witness in [PR #277](https://github.com/RedSalamanders/RedSalamander/pull/277) at `2612d32b7`. The approved x64 Release case found one tooltip for its one painted badge on the first collapsed frame. The focused Release Clear Completed and provider matrix passed without activation. [Release tooltip evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-25_132817_completed_tooltips/README.md)
- [x] Merge the responsive F7 prompt in [PR #278](https://github.com/RedSalamanders/RedSalamander/pull/278) at `784954b59`. Approved x64 Debug name selection, creation, refresh and Cancel passed. The held-worker and deferred-exit case passed without activation. The controlled x64 Debug Dummy fixture passed on `c3c6f62ac` in `20261125T141259Z-52476-4b56ae7c376e4b0183789dcafa1ec6b8`. The prompt offered the numbered name in 32.660 ms while a direct listing of the same parent took 2,001.464 ms. The worker preserved the occupied folder, created the sibling, and refreshed the pane. [Delayed-provider evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-25_151352_create_folder_slow_provider/README.md). [PR #284](https://github.com/RedSalamanders/RedSalamander/pull/284) merged at `2bc6df430`. [Create Folder evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-25_133107_create_folder/README.md)
- [x] Merge the Pack and Unpack close refusal in [PR #280](https://github.com/RedSalamanders/RedSalamander/pull/280) at `a0c4823a9`. The x64 Debug no-activation run `20261125T125848Z-73424-b648007869814755a389e75125acd4fd` passed both pane refusals in the deferred-exit case. Archive worker migration, cancellation during I/O, and cleanup qualification remain open.
- [x] Merge the Release provider matrix correction in [PR #270](https://github.com/RedSalamanders/RedSalamander/pull/270) at `a0251953b`. The focused x64 Release run `20261125T111418Z-17464-fdd7961fb8f24abc857b1fe6089173ee` passed all seven cleanup decisions and provider conformance. All 66 runner-plan checks passed. [Release matrix record](../../TestRuns/4cb089111a23/FileOps/2026-11-25_121453/README.md)
- [x] Verify the integrated x64 Release Graph and deferred-exit cases on `d6946334f`. The full build had zero warnings and errors. Both focused selections passed all three stages without activation on receipt `cd0a3c103bc7d37df9d4d87dd5df5ec215eff3b34a5a5c8180b8f8af718c355d`. [Integrated Release evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-25_160030_graph_f7_release/README.md)
- [ ] Qualify hosted native x64 and ARM64 Release. [Run 45588264577](https://github.com/RedSalamanders/RedSalamander/actions/runs/45588264577) failed both profiles before the focused corrections. [Run 45610829022](https://github.com/RedSalamanders/RedSalamander/actions/runs/45610829022) also failed both. x64 reported 3,023 passed, five failed and 68 skipped, with the old Completed badge witness, the Copy links Tab sequence and three Tooling Pester failures. ARM64 failed to retire the desktop boundary after a foreign console took foreground. The [first run with both focused corrections](https://github.com/RedSalamanders/RedSalamander/actions/runs/45616935328) completed admission and skipped native qualification. The [Graph merge's coordinator](https://github.com/RedSalamanders/RedSalamander/actions/runs/45636245534) stopped on HTTP 500 before dispatching native tests. On `a2fd204c6`, x64 reported 3,017 passes, three failures and 77 skips. The failures were the Win32 UI dependency audit, the F7 provider-name source guard, and the R4A02 Delete-after-Copy wait witness. ARM64 refused to restore over a foreign console during desktop retirement. Both real native regression jobs are now running for `5d0306834`. [Failed hosted run 45646466945](https://github.com/RedSalamanders/RedSalamander/actions/runs/45646466945) [Current hosted run 45675702457](https://github.com/RedSalamanders/RedSalamander/actions/runs/45675702457)
- [x] Merge the three hosted x64 witness corrections. Candidate `fb35e7126` passed the clean x64 Release build, both approved test selections, the native audit and all 278 focused tooling checks. The F7 source guard passed 1 of 1 and R4A02 passed 3 of 3 with 902,098 µs of explicit Queue wait. Both runs retired their owned processes and released the lease. [PR #300](https://github.com/RedSalamanders/RedSalamander/pull/300) merged at `2b9ac7520`. [Focused Release evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-25_203252_hosted_regression_witnesses/README.md)
- [x] Account for the earlier broader x64 Release PrePush failures from `20261125T182436Z-67896-1c7845a008bc48b8a248fca00d1526f4`. The run finished with 1,799 passed, 3 failed and 29 skipped, with four desktop entries still pending. PR #300 corrected the old Win32 audit finding. The ProcessStreaming assertion looked up a reused PID after the child exited. [PR #301](https://github.com/RedSalamanders/RedSalamander/pull/301) uses the launcher-held handle and passed all 18 focused tests plus 276 source, inventory and governance tests. The exact non-activating Compare cancellation rerun passed in `20261125T195359Z-26512-8ebe7395bd4741568e8a14f688bfdbc6` on the original Release receipt. The broader failed verdict remains in the record. The later non-interactive Release sweep passed the affected entries.
- [x] Verify the Copy links tab order in File Operations Preferences. Approved x64 Release run `20261125T152557Z-64608-a90d2d53758647bb82e632ae00aa8fe6` verified Tab and Shift+Tab through Copy links between the decision-mirror toggle and Parallel transfers. It passed with zero failures and skips and released the desktop lease. [PR #288](https://github.com/RedSalamanders/RedSalamander/pull/288) merged at `a2fd204c6`. [Preferences keyboard evidence](../../TestRuns/4cb089111a23/Commands/2026-11-25_162557_copy_links_tab_order/README.md)
- [x] Merge the Release tooling correction in [PR #287](https://github.com/RedSalamanders/RedSalamander/pull/287) at `673d75c47`. The 296 focused Pester checks passed on `346da7b22` in `20261125T151521Z-34292-37ecdd7ce4834af1a284833cf40db983`. The reviewed provider-matrix projections, exact loopback refusal rule and required PR checks passed.
- [ ] Deliver Batch Rename and Change Case through the qualified shared Rename route. Exact Windows SMB x64 Release passed both host task origins and cleanup with zero failures or skips. The controlled Windows SMB and live RubyGloom checks both verified the pre-mutation refusals. Corroborated F2 passed on live Samba and Windows SMB with controlled missing IDs. Further provider, menu and walkthrough gates remain open. [Scheduled Rename refusal evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-26_083853_smb_scheduled_refusal/README.md) [Windows SMB scheduled Rename evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-26_072635_windows_smb_scheduled_rename/README.md)
- [x] Verify and merge consistent real Cut/Paste generation handling. Isolated candidate `c5ec8fa2f` captures the file list, drop effect and Cut sequence in one clipboard open. It extends the real Cut/Paste journey to check that a new clipboard generation vetoes an earlier Move. The first x64 Debug build at `d43676638` caught two signedness warnings. Corrected candidate `35cf70eef` built with zero warnings and errors. The approved real desktop case `20261125T184421Z-86040-8c95817a3fe74e9ca49a67973b5674a6` passed with zero failures and skips. It verified stale Cut refusal, preserved both files, and restored every saved clipboard format. The runner released the desktop lease. [PR #297](https://github.com/RedSalamanders/RedSalamander/pull/297) merged at `ff6ea9d3c`. Its exact test-enabled x64 Release build at `cbdfdea83` was clean and non-activating Copy-only Move run `20261125T190537Z-8164-ec801774eee240dcbcaa561458f8dc38` passed all three stages.
- [x] Land mixed-endpoint Move consent and verify changed-root refusal in the controlled fixture. Draft `8d3638e65` adds Cancel and Continue coverage for one mixed native and Copy-only Move on two owned Dummy roots. Candidate `ed2d4889d` built x64 Debug on the merged tree with zero warnings and errors. Non-activating run `20261125T192030Z-59472-2aea1d26039d4393ae20de491c131cf7` verified the two planned strategies but failed the outcome check with `ERROR_IO_INCOMPLETE`. Dummy had moved its native item without reporting a typed mutation receipt. Correction `54d9fc69a` reports it only after witnessing the same file at the destination and no item at the old name under the fixture's directory lock. It built x64 Debug with zero warnings and errors, and run `20261125T193637Z-86976-585b1df3b0ef422698098a3c4c4fca98` passed all three stages without activation. Candidate `436b9bf25` adds a worker-side source and destination root check before using the accepted plans. It built x64 Release with zero warnings and errors. Run `20261125T194820Z-82772-b7a37c58b4814e59954306338983d2f5` confirmed that a changed root stopped preparation before mutation. The task incorrectly summarized multiple refused items as a partial Copy. Correction `387e477db` preserves the specific preparation failure. It built x64 Release cleanly. Non-activating run `20261125T200336Z-69136-a1e66005bf7d4aa996fb7e9476b50ea4` passed all three stages and reported `0x800703EE` after one question. Both originals remained and the destination stayed empty. The same Release receipt passed the deferred-exit case in `20261125T200449Z-69136-9c8a2703b4134c7b98346386c89f59b9`. [PR #303](https://github.com/RedSalamanders/RedSalamander/pull/303) merged at `f4de7bfad`. [Mixed Move evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-25_210430_mixed_move_endpoint_consent/README.md)
- [ ] Verify mixed Move across the outstanding provider and endpoint scopes.
- [x] Verify Unpack source retention after a collision skip. Exact x64 Debug run `20261125T154439Z-41676-269af778083b49cf8d6377eaee7534b5` passed without activation. It approved Delete after, kept the selected ZIP and the existing destination when extraction skipped entries, and queued no Delete. [PR #290](https://github.com/RedSalamanders/RedSalamander/pull/290) [Archive retention evidence](../../TestRuns/4cb089111a23/Commands/2026-11-25_164439_unpack_archive_retention/README.md)
- [x] Verify Pack source retention when an item is omitted. Exact x64 Debug run `20261125T161424Z-8584-f111c4d7cf174542b2d63fa66cce8229` passed without activation on `63504c595`. It confirmed that the ZIP held the included files and Pack kept the originals and omitted link without queuing a Delete. The same run passed the Unpack retention checks and archived after fixture teardown. Exact-head x64 Release run `20261125T163138Z-85480-2dddd82cea854964a77d03ceb8eb31a2` passed the same case. [PR #291](https://github.com/RedSalamanders/RedSalamander/pull/291) merged both guards at `64b7c09d2` after the required checks. [Pack and Unpack evidence](../../TestRuns/4cb089111a23/Commands/2026-11-25_171424_pack_unpack_retention/README.md)
- [x] Verify Unpack Delete after against the original archive identity and revision. The x64 Debug build at `5ac478f9b` had zero warnings and errors. Exact non-activating run `20261125T170848Z-64868-04786ec6dd31403b93f90f28f266b6a6` removed the unchanged compressed archive, kept a replacement and the moved original, and kept a same-size edited archive whose last-write time had been restored. [Source identity evidence](../../TestRuns/4cb089111a23/Commands/2026-11-25_180848_unpack_source_proof/README.md)
- [x] Verify Pack Delete after for regular-file selections on x64 Debug. Exact non-activating run `20261125T174113Z-85544-d590850a625548299c60c58c56680533` removed the unchanged 7z source and kept the replaced ZIP source and its moved original. Exact x64 Release head `1c73a2304` built with zero warnings and errors. Non-activating run `20261125T175549Z-76840-aaaac915ddb04500a972c8950a88851f` passed in 1,354 ms with zero failures and skips. [PR #293](https://github.com/RedSalamanders/RedSalamander/pull/293) merged at `527040bf5` after all required checks. [Pack source identity evidence](../../TestRuns/4cb089111a23/Commands/2026-11-25_184113_pack_file_source_proof/README.md)
- [x] Verify one combined Delete after confirmation for an exact regular-file Pack or Unpack selection containing a possible leftover on x64 Debug. Candidate `42beb96fb` built with zero warnings and errors. Non-activating run `20261125T180804Z-31028-e6de09a1101d4a4f83241b40ae1410bd` verified the default Cancel, exact cleanup, and protected replacements with zero failures and skips. The exact test-enabled x64 Release head `e32b09ae4` also passed in non-activating run `20261125T182138Z-91068-2c036bcc81db4e3aab5f3c7e08af799b`. [PR #296](https://github.com/RedSalamanders/RedSalamander/pull/296) merged at `5d0306834` after all required checks. [Combined confirmation evidence](../../TestRuns/4cb089111a23/Commands/2026-11-25_190804_archive_combined_consent/README.md)
- [x] Keep Pack directory selections when their exact cleanup set is unavailable. Isolated candidate `78d2a5c72` removes the unproved recursive Delete route. The x64 Debug retry verified the cached Ghostty and Zig inputs against their exact lock hashes and built with zero warnings and errors. The focused non-activating ZIP case `20261125T185016Z-37604-8000aef41f024e65870c2d6df59b7de3` passed with zero failures and skips. It wrote the complete archive, preserved the originals, and queued no Delete. [PR #298](https://github.com/RedSalamanders/RedSalamander/pull/298) merged at `e6389dc0d` after the required checks and clean x64 Release build at `098e18ef6`. Non-activating Release ZIP run `20261125T190537Z-86108-b236a5e86c724587a306e68f629fc26c` passed in 1,312 ms with zero failures and skips. All 67 runner-plan and 276 source and tooling checks passed in `20261125T184206Z-91560-9b6ffca17a80419a95cc7907efb2f4bf`.
- [x] Merge transactional Pack publication. Candidate `d54ad80eb` uses the shared Local transaction for ZIP and 7z output. Both x64 builds had zero warnings and errors. The exact non-activating Debug and Release cases preserved existing archive bytes on failed publication, retired only the stage and verified the overwrite retry for both formats. [PR #305](https://github.com/RedSalamanders/RedSalamander/pull/305) merged at `af0895dd7`. [Pack publication evidence](../../TestRuns/4cb089111a23/Commands/2026-11-25_211805_pack_publication/README.md)
- [x] Qualify Pack and Unpack worker scheduling and cancellation in the focused archive cases. [PR #308](https://github.com/RedSalamanders/RedSalamander/pull/308) merged the Pack worker at `43568ebb3`. Its stored ZIP and 7z checks preserved the prior output on cancel and protected the published file while exact regular-file cleanup waited. The combined Unpack candidate passed the stored and deflated ZIP cases in x64 Debug and Release without activation. Both approved x64 Release archive dialogs passed and released their desktop leases within the single five-minute window. The Release deferred-exit guard also passed. [Pack worker evidence](../../TestRuns/4cb089111a23/Commands/2026-11-25_220215_pack_worker/README.md) [Unpack worker and dialog evidence](../../TestRuns/4cb089111a23/Commands/2026-11-25_222853_unpack_worker/README.md)
- [x] Obtain exact Pack directory cleanup evidence for ZIP and 7z. The tested worker and durable specifications merged in [PR #384](https://github.com/RedSalamanders/RedSalamander/pull/384).
- [ ] Finish wider archive worker qualification in the integrated native and workflow passes.
- [x] Verify the actual host password prompt with a running peer. Exact x64 Release run `20261126T150958Z-49120-34b2dcfb1ab340b6a6c603e7e14e5901` passed 3/0/0 without activation. The owned Cancel returned no secret. The held Local Copy entered and left credential wait, and the popup recorded seven ticks while the peer stayed Running. [Host prompt evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-26_161031_host_credential_prompt/README.md)

- [ ] Verify live authentication and progress for the remaining approved providers, including configured Windows Hello. The host prompt test exercises the actual dialog with an invented connection profile. Live provider sign-in still needs its endpoint evidence.
- [ ] Complete the workflow walkthrough and closeout inventory. Move the plan to Done once every required gate has evidence or an accepting owner.

**Delivery checkpoint (2026-11-24).**

- [x] Merged ingress, Local Recycle batching and Preparing progress through [PR #246](https://github.com/RedSalamanders/RedSalamander/pull/246).
- [x] Merged exact Rename, Queue/Parallel persistence and pause ownership through [PR #247](https://github.com/RedSalamanders/RedSalamander/pull/247). The x64 Debug Windows SMB and x64 Release popup-cadence cases passed. The [live Windows SMB x64 Release check](../../TestRuns/4cb089111a23/FileOps/2026-11-25_100205/README.md) passed nested and case-only Rename, occupied-name and watcher refusal, the 100-item batch and fixture cleanup.
- [x] Merged corroborated F2 Rename and the Preparing assertion repair through [PR #248](https://github.com/RedSalamanders/RedSalamander/pull/248). Live Samba x64 Debug and Release passed the guarded Rename and fixture cleanup.
- [x] Merged the bounded pinned Zig download fix through [PR #250](https://github.com/RedSalamanders/RedSalamander/pull/250). A fresh x64 Debug solution build verified the archive and executable hashes and emitted zero warnings and errors.
- [x] Passed the authorized x64 Debug decision-mirror focus, queue order and Clear completed checks in `20261124T194522Z-93024-b369633f9e2a41aab5bf8a539d4d9d13`. This pass exercised the repeat-choice build and ended in under three minutes.
- [x] Pass the initial exact x64 Debug typed Recycle and repeat-choice cases. Integration status is above.
- [x] Merge shared bandwidth presets and IEC examples through [PR #253](https://github.com/RedSalamanders/RedSalamander/pull/253). The x64 Debug preset and parser checks and all 12 resource checks passed. Explicit-start Custom editing remains open.
- [x] Pass the initial focused x64 Debug Create Directory admission check.
- [x] Pass 66 focused runner-plan checks for clean Completed grouping. The native result is recorded above.
- [ ] Finish the stage gates below and qualify native x64 and ARM64. [The first main Release run with the repaired download window](https://github.com/RedSalamanders/RedSalamander/actions/runs/45474460298) ended after admission.

**Samba F2 Rename (2026-11-24).** No-activation run
`20261124T162457Z-88408-705225e504c24214a1191b57d8f59ce3` reached an active-callback assertion while
Preparing captured the selected source. The owned process was stopped and the failure artifacts were retained. The task now declares a balanced callback scope around
that synchronous provider call. The assertion remains in place. The test-enabled harness routes CRT
assertions to standard error and its failure report. The 20 focused failure-report checks, including
the native assertion controls, passed in `20261124T164133Z-97372-233c6dc7388c4dcf912e8f73efb7ba85`.
The next run exposed stale write times on metadata-only SMB opens. The provider now captures comparable
timestamps with read-data access, and the host reports a source change before dispatch as a refused
operation. The final x64 Debug run `20261124T173007Z-100696-ce72a4511ad04f9e9076f0aa92525ad2` passed
setup, the live case and cleanup. It verified folder and case-only Rename, Keep both, cancellation,
refusal after a source change and 100 provider renames. The zero-warning, zero-error build and metrics
are recorded in the [Samba archive](../../TestRuns/4cb089111a23/FileOps/2026-11-24_173737/README.md).
The exact [PR #248](https://github.com/RedSalamanders/RedSalamander/pull/248) head also passed live
x64 Release in `20261124T184411Z-73888-00b41fc751a8491c9177872df1459797`. Its 100 Rename operations took
23.912 seconds. The related non-activating Debug PrePush recorded 1,801 passes, one unchanged
ProcessStreaming timeout and 26 environment skips. The original failure is retained. All 18
ProcessStreaming checks passed in isolation on the same head. Native ARM64 and the remaining Rename
entry points are open.

**Queue and pause qualification (2026-11-24).** The x64 Release `FileOps_PopupIdleCadence` run
`20261124T162210Z-62860-6cd00aef04f241ceacedc1d1d503eb17` passed setup, the case and cleanup.
The previous Debug PrePush was stopped to correct the scheduler source test's typed pause boundary.
That source test group passed all 192 checks in `20261124T163442Z-63668-ee7d2800de8c432db58fb14adf325401`.
[PR #247](https://github.com/RedSalamanders/RedSalamander/pull/247) merged at `6c6717856`. Its
non-activating x64 Debug PrePush reported 1,799 passes, one S3 timing failure and 26 environment skips in
`20261124T164838Z-89552-dffa46df698d44afb067f341a3560aba`. PluginContractTests passed on the same
verified binaries in isolated run `20261124T180310Z-86824-456dbfff52324fbea1ca3dccb4c3a99d`. The
final head added a test alignment change and passed the affected settings and hosted checks.
Four broad desktop entries remain open. Both [main native Release jobs for this head](https://github.com/RedSalamanders/RedSalamander/actions/runs/45456742533) failed while downloading the pinned Zig archive before native test execution.

**Follow-up qualification (2026-11-24).** Scoped PrePush on `e3e5ac280` built cleanly. Its standalone settings entry
failed because the Queue default regression fixture attempted to write over seeded legacy settings without the loaded
revision. The correction in `5d35248ee` keeps that revision. Its targeted x64 Debug run
`20261124T155409Z-53376-27107cd039834ad6bf9ddce9a4b2cca7` passed SettingsSchemaTests with zero build warnings or
errors. The source-inventory check also required the Commands one-field persistence case for `queueNewTasks`. The
candidate adds that case using the existing fixture table. Four foreground entries and the full repository verdict
remain open.

**Current work (2026-11-24).** Local Recycle batching reached context-menu, Find and Viewer admission. Find and Viewer
each used one Shell batch for twelve sibling files, and Viewer Cancel kept the files. The repaired same-folder Ctrl+drag
passed through its exact no-activation selection. The branch also contains the first exact Batch Rename handle-lifetime
and throughput changes. Implementation started at `bbb57a572`. Drift from the study baseline affected the ancestry-cleanup
test and its evidence contracts.

[PR #246](https://github.com/RedSalamanders/RedSalamander/pull/246) merged into `main` at `bcaa59d10`.
Its non-activating x64 Debug PrePush executed 14 entries with 1,799 passes, zero failures and 27 environment skips.
[The first main Release attempt](https://github.com/RedSalamanders/RedSalamander/actions/runs/45405656554/attempts/1)
timed out downloading a Ghostty build dependency on both architectures before native regression tests.
[Attempt 2](https://github.com/RedSalamanders/RedSalamander/actions/runs/45405656554/attempts/2) hit the same timeout.
Four desktop-lease entries and the qualification below remain open.

The follow-up x64 Debug PR selection on `842ac811a` executed 12 entries with 1,507 passes and zero failures.
Repository coverage remains `NOT_EVALUATED`. The current candidate also includes Queue/Parallel persistence,
bulk-pause ownership and a dedicated live SMB Rename case. Correcting the SMB child-family registration cleared the
first build failure. The selected x64 Debug build on `6df121e23` completed with zero warnings and zero errors, and
the non-activating popup and repeat-decision selections passed. A guarded Win32 check on both SMB endpoints
showed that a read-data watcher is required to reproduce the sharing violation. The earlier PrePush stopped after
its build to correct the witness. The corrected live Windows SMB case passed on `92c2b2762`.

- [x] **Intake.** Preserve the source study, add the WIP index entry, and check source drift.
- [ ] **S0.** Finish ingress parity and qualification.
  - [x] Verify folder-picker Move, Viewer Recycle wording and safe cancel, and Cancel all in the focused x64 tests.
  - [x] Rerun the same-folder Ctrl+drag guard through its exact no-activation selection after the Windows provider instance fix.
  - [ ] Finish clipboard and context-menu coverage in the workflow walkthrough. The selected known-noncommit Retry and read-only replacement cases passed in the approved x64 Release desktop window.
  - [ ] Verify the implemented short Copy-only Move question across mixed plans, changed endpoints, and Cut/queue
    preservation. The focused source-retention case passed.
  - [x] Route Local Recycle through the existing bulk executor. The focused Find and Viewer runs verified prompt counts,
    preserved files on Cancel, and compared Shell calls and settled resources against the per-item route.
  - [x] Repeat the Local Recycle measurements on x64 Release. Integrated run `20261126T103711Z-74284-0866d2274311474d880dd92b847a9f54` passed 3/0/0. On the twelve-file fixture, per-item settled in 246.715 ms, Find in 67.288 ms, Viewer Cancel in 2.350 ms and Viewer in 96.560 ms. Find and Viewer each used one Shell batch. Cancel used zero batches and preserved the selection.
  - [ ] Exercise the context-menu surface in the walkthrough.
  - [x] Implement typed per-item Recycle availability and direct permanent consent for known-unavailable routes.
  - [x] Qualify the guarded unknown Recycle result on the registered Samba and Windows SMB routes in x64 Release.
    The independent identity probe selected the by-name and exact-identity prompts respectively. Both live runs
    kept the original bytes through Skip and independently confirmed removal of their own fixtures after cleanup.
  - [ ] Complete the mixed-selection Recycle walkthrough on available live endpoints. The focused Local mixed-group check passed exact permanent and Recycle routing after approval and preserved both items after Cancel.
- [ ] **S1.** Qualify Rename by route.
  - [x] Release each exact Batch Rename handle after preflight and validate captured identity on every mutation
    rebind. The local exact-result test passed with a no-DELETE-sharing witness at the preflight boundary.
  - [ ] Qualify exact Local and SMB for nested rows, case-only hops, legacy API fallbacks and blocking watchers. Compare
    single- and 100-item Rename costs, including handle peaks and cancellation.
  - [x] Run the Local exact Rename refusal, legacy fallback, collision and lost-response controls.
    [The x64 Debug record](../../TestRuns/Local-x64/FileOps/2026-11-24_140134/README.md) retains their timings and receipts.
  - [x] Register the approved SMB test directory and confirm a read-only listing and mapping through `remote_file_smb`.
    The case passed in run `20261124T143343Z-95336-b6a780d1e7104d7092573afd510ba4d0`.
  - [x] Exercise the exact Rename eligibility check on the registered SMB endpoint. The Local provider opened the
    owned file and corroborated that exact file identity was unavailable. The Rename case skipped in run
    `20261124T145318Z-86896-cf65eb1fac8447518ff7509d6d3dfe51`.
  - [x] Prepare the separately approved Windows SMB test directory.
  - [x] Compare the namespace guard on both endpoints in diagnostic run `20261124T150306Z-85228-8ce4ce9864d743b0bd18f932d4bb5879`.
    Read-data handles denied competing DELETE opens on the owned parent and ancestor. Absolute no-replace Rename
    preserved the occupied target and moved the file to a free sibling. The first endpoint returned `ERROR_INVALID_PARAMETER`
    for `FileIdInfo`. Windows SMB returned an exact ID. These observations leave the host route and full namespace proof open.
  - [x] Run live exact Rename on Windows SMB with the corrected watcher. The nested, case-only, collision,
    watcher and 100-item checks passed. The [focused record](../../TestRuns/4cb089111a23/FileOps/2026-11-24_161520/README.md)
    retains the run identity and timing. Release, legacy fallback hops and host cancellation remain open.
  - [x] Implement and exercise corroborated F2 Rename on the configured Samba endpoint without exact file identity.
    The non-activating x64 Debug run passed action, refusal, cancel, throughput and fixture cleanup controls.
  - [ ] Verify FTP and SCP with isolated test credentials. The dedicated profile selectors are registered and the
    scoped authentication setup was blocked before a live run.
  - [x] Implement and qualify corroborated same-parent Local/SMB Rename for natural Samba and controlled missing IDs on Windows SMB.
    The retained-parent publication, ancestor decoys and replacement-parent refusal passed on both live shares in [PR #359](https://github.com/RedSalamanders/RedSalamander/pull/359).
  - [ ] Qualify new-name Rename on Curl, Drive, and Graph and its result receipts before enabling each protocol or
    profile. Verify F2, Batch Rename, Change Case and unsupported-route cases.
- [x] **S2.** Complete progress projection and paired measurements.
  - [x] Preserve sampling through decision, live-output, pause and queue transitions. Freeze history while stopping
    and keep a running peer's ETA. The model and paced idle-cadence selections passed.
  - [x] Verify the Preparing elapsed clock at task creation and after directory Delete reservation. The focused x64
    Debug popup and admission cases passed without desktop activation.
  - [x] Cover credential waits and align the popup and taskbar snapshots. The Release idle-cadence selection requires one credential attention item, a running peer with unchanged ETA, and matching taskbar status. The actual host prompt also passed.
  - [x] Collect paired allocation, timing, history and active-peer evidence. The matched Release samples cover Prompting, Running, Paused and Stopping. The controlled worker and real host prompt runs cover active-peer cadence.
- [ ] **S3.** One repeat control, coalesced cleanup and read-only decisions, overlap and sensitive-loss handling.
  Match card and mirror scopes and authority, including stage cleanup and post-publication results.
  - [x] Rerun the existing repeat-decision contract. The x64 Debug selection passed in
    `20261124T145409Z-86896-a97bc8a5503d4a98b153e3aba31f7647`.
  - [x] Land one permanent Leftover Cleanup card for the exact selection in PR #266. The x64 Debug case verified accept, decline, and a changed selection.
  - [x] Land automatic overlap waits in PR #276. The focused cases checked separate task results and preserved storage.
  - [x] Verify metadata preservation, one combined EFS and sparse allocation decision, source retention, and Cancel before publication in the selected x64 Release metadata matrix.
  - [x] Qualify one exact Pack directory cleanup decision, Cancel, and changed-tree retention for ZIP and 7z.
  - [ ] Qualify the remaining predictable sensitive-loss routes.
- [ ] **S4–S5.** Queue controls and persistence, shared cap, completed grouping, mirror focus and Issues entry.
  Recheck I60/I64 ownership and verify interlocks, pause provenance, retention, UIA and focus.
  - [x] Verify that the popup stores Queue/Parallel and a new owner reads the saved default.
  - [x] Verify bulk pause ownership in the engine and popup model. Preparing and queued tasks are included.
    Resume all preserves the earlier individual pause and queue gate. The paired cadence measurements and
    save timings are in the [x64 Debug record](../../TestRuns/4cb089111a23/FileOps/2026-11-24_155152/README.md)
    and [x64 Release record](../../TestRuns/4cb089111a23/FileOps/2026-11-24_163108/README.md).
  - [ ] Verify the Start option's one-submission scope and the new pause behavior through the full interactive
    walkthrough, including focus, UIA, and a pending mutation. Start now kept the neighbor queued and Pause all paused
    and resumed the active transfer in the [exact x64 Release replay](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#completed-card-paint-and-queue-desktop-checks). The task options override still needs its own submission check.
- [ ] **S6–S7.** Options parity, asynchronous Create Directory and archives, and useful results. Preserve saved
  settings. Qualify worker cancellation, exit, refresh and cleanup.
  - [x] Land async Create Folder and the responsive name prompt in PRs #261 and #278. The focused worker and desktop checks passed.
  - [x] Land the explicit-start Custom editor and persisted Copy links Apply and Cancel behavior in PRs #259 and #271.
  - [x] Measure Create Folder with a controlled slow provider. Verify it refreshes the original pane when it stays in place and preserves the new location after navigation.
  - [x] Verify local Pack and Unpack worker and cleanup in x64 Release. The exact ZIP and 7z directory cleanup and refusal cases passed, and the Pack and Unpack round trip passed all three repetitions on the corrected tree. Integrated hosted qualification remains open above. [Exact archive evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-27_080923_restart_qualification/README.md#exact-pack-directory-cleanup)
- [x] **S8.** Synchronize the integrated workflow with the owning specifications, consistency ledger and user guide. The audit covered fixed destinations, Cut/Paste, repeat decisions, options and queue controls, Rename capability and archive cleanup. The Stopping action list now matches the retained read-only path inspection.
- [ ] **Closeout.** Complete the walkthrough, native and live-provider gates, document checks and reviewed history
  admission. Move to Done after all required rows have evidence or an accepting owner for the exact transferred gate.

### Validation checkpoints

- The first exact Rename candidate `ab51279e6` stopped at compilation in
  `20261124T134444Z-100080-5b6b4d2e1e1046718998105aec11af0c`. WIL's final-path adapter requires its owned string type.
  The corrected candidate uses `wil::unique_cotaskmem_string` for the spelling proof.

- `dc30fb46b` built with zero warnings and zero errors using `build.ps1 -Configuration Debug -Platform x64
  -ProjectNames RedSalamander` with tests enabled. Build receipt
  `04985ac959f56c20a9f2a40d8f0613490866797615e6e138c911876fba3b7bd0`.
- Focused x64 Debug runs used the marked local sandbox. The runner reported `NOT_EVALUATED` at repository scope.
  The passed selections and the initial failed witness remain in its completed-run archive.

| Selection | Result | Run ID |
|---|---|---|
| `Commands_FileOpsFolderPickerMoveUsesHostQueue` | Passed | `20261124T080347Z-60876-a29df09ceaeb4316b8894cfd894b2c9e` |
| `folderView_drop_integrity_guards` | Failed on same-folder Ctrl+drag in the first candidate | `20261124T080442Z-60876-402869228c9c4feeb44ff233aad4da72` |
| `FileOpsFamily_BatchRenameExactResult` | Passed, including the post-preflight sharing witness | `20261124T080545Z-60876-bfc378cf17ad4c1ab786cf4527749906` |
| `FileOps_CopyOnlyMoveRetainsSourceWithoutDestinationProof` | Passed | `20261124T080633Z-60876-ff9be60a22e34d6eb6880fdfdf6e8987` |
| `Phase10_PermanentDelete` | Passed, including the Recycle confirmation snapshot | `20261124T080714Z-60876-d01ca01cc0aa43218ba40875855a2e8e` |
| `FileOps_RateWindowDisplay` | Passed | `20261124T080756Z-60876-09d1cdd6f4c049b79b05e52c6d9e390a` |
| `FileOps_CancelAllNeverStarted` | Passed | `20261124T080835Z-60876-ea661fa63ee5419c97544e2d3586433a` |

- The same-folder repair `56a5d9d08` also built cleanly in test-enabled x64 Debug. Its receipt
  is `ec35a64c2ee51216f7ffc47c35477e25558c67d6eb85c3b849788fd5a9544d56`.
  The approved first focus window ended before this build finished.

| Selection on `56a5d9d08` | Result | Run ID |
|---|---|---|
| `FileOps_RepeatDecisionContracts` | Passed without activation | `20261124T081614Z-64872-f5cb2a98b9c542b181b3b8a97ec8192d` |
| `FileOps_IdentityInterlockWaiting` | Passed without activation | `20261124T081715Z-64872-a7b1ace66f8f4fb58da7eb79d89dabd8` |
| `FileOps_PopupIdleCadence` | Failed because the tiny fixture reached publication before a new rate could be measured. The paced rerun passed on the candidate below. | `20261124T081817Z-64872-e3f005a81f5c48649ca3a0dd4aaff921` |

- The full solution built test-enabled x64 Debug on `0211122e8` with zero warnings and zero errors. Receipt
  `782884d4bec483414e035036c1b48f80cd71bd8b1ed4ab10e98f21625ff2a9c4`. The paced
  `FileOps_PopupIdleCadence` selection passed without activation in run
  `20261124T082752Z-59536-98f34ae1390f465989d44c3027b3b8b1`.
- Non-activating `Test-Changes.ps1 -Mode PrePush -Configuration Debug -Platform x64 -SkipBuild` on `0211122e8`
  executed 14 entries. Thirteen passed their entry gates. The run recorded 1,797 passed, two failed, and 27
  skipped cases. The Pester failures matched an obsolete transition tick and a fixed character distance
  in the live-output wait guard. Run `20261124T082910Z-59536-199bace9c93c4fffba708267a19c2e54`,
  final receipt `30e70fc85e83d409cf324d7d60cb2f17aaf93b1d65b04b0fa9ecda6bd23a09ad`.
  Commands, File Operations and both interactive viewer selections remain pending.
- After the guard repair, `TestHarnessSourceContracts.Tests.ps1` passed 243 cases with zero failures or skips
  in Pester 3.4.0. Run `20261124T092118Z-19224-3e7f22bb4d984e5fba1a4bdae44c663e` binds native sources at
  `0211122e8` and the updated test file SHA-256
  `aab6ae96a3e1b4179b64755fd50590d18604b37b13970ef4eb07f271f526b23b`.
  A current-head PrePush rerun remains open.
- Same-machine x64 Debug cadence comparison used baseline `bbb57a572`, verified receipt
  `aad8a6aefb2102d3f24a592ccb4ceb9c620f78f149bed0f39bb773a66376956a`,
  run `20261124T091748Z-75364-035371b12b9344f0a32412ced821a0ac`,
  and candidate `0211122e8`, final PrePush build receipt above,
  run `20261124T091911Z-69464-1f7f62f41fb24db186fcc04be6f8fa9d`.
  Both focused selections passed. In each unchanged synthetic interval of about 1.2 seconds, running
  recorded 11 ticks and 11 renders, paused, prompting and queued each recorded one tick and one render,
  and completed idle recorded zero. The changed live-copy workload requires a matched workload for a
  cost comparison. Allocation, retained history and concurrent active-peer measurements remain open.

- The first Recycle fixture candidate `fcbcea4ce` failed to compile because it passed an integer to the steady-clock
  timer. The corrected candidate `f5aa393f0` rebuilt with zero warnings and zero errors in x64 Debug. Receipt
  `9fb00c721b37714ca6f6f3dff919e0ddb520d01a822043a34ecd0954a3ba6804`.
  Run `20261124T102043Z-60344-b9aaba78a56e498da31ba7e65186af3b` passed the Local Recycle routes, one-prompt checks
  and exact purging. The [paired ingress evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-24_112639/README.md)
  retains the call counts, time, handles, private bytes and focused scope.
- The no-activation policy and all 243 source contracts passed in
  `20261124T102042Z-59168-e61c6ec3728c48cfae1399bc22ef9065`. The native Commands
  `folderView_drop_integrity_guards` selection used the same x64 Debug receipt, reported zero pending interactive
  entries, and passed in `20261124T102827Z-74852-3e73f64a46ac464cad850a29feed1a0f`. Both focused native runs
  reported `NOT_EVALUATED` at repository scope.

- Non-activating x64 Debug PrePush on `ee6b50125` built the full solution with zero warnings and zero errors.
  Run `20261124T103514Z-56072-2c2e7c56ea8d414db4ac0a1efb117d4c`, final receipt
  `0ebf0408ac73876d33fdccf4295efff3dd8917ee68a5b5415c53564d82d8fb6b`. The run recorded 1,798 passed, one failed, and 27 skipped.
  Thirteen of fourteen entries passed their gates. The unchanged Git metadata fixture found a `sleep` process
  when it looked up an exited PowerShell probe by numeric PID. Its focused group passed all three cases on the
  same head in run `20261124T111258Z-31232-3274593ad723444a99aecc4248262b88`. The broad verdict remains `FAILED`.
  Commands, File Operations, and both interactive viewer entries remain pending. The sandbox audit reported
  the same 13 existing findings as the earlier broad run.

- The test-enabled x64 Debug build at `9d4b49d71` used receipt
  `43a396a5e2b5167cea0ff4a6b0182faf99f8f0ab718727869e833ec1e1e24d20`. `FileOps_PopupIdleCadence` and
  `FileOps_ReviewRemediationContracts` passed without activation in runs
  `20261124T115346Z-19596-42d9378fb9594838b3962c7997e4d76a` and `20261124T120125Z-12688-ad787f3e6bed48438c508063879e91db`.
  The [paired Preparing evidence](../../TestRuns/4cb089111a23/FileOps/2026-11-24_130014/README.md) records matching synthetic cadence and a 16-byte increase
  in the measured task object. Repository coverage remained `NOT_EVALUATED` for these selections.

- Non-activating x64 Debug PrePush on `b4bdf0086`, run `20261124T120545Z-65304-4582af4137ca4af0ae165551a14d61f7`,
  used final full-solution receipt `2b5b5c2dbda57d43a2e8e7f927e99782e335afbb224d75eee2899c208dbf8c92`.
  All 14 executed entries passed their gates. The run recorded 1,799 passed, zero failed, and 27 environment skips.
  Tools Pester passed all 1,496 cases, including the earlier Git metadata fixture. Change coverage was `PASSED`.
  Repository coverage remained `NOT_EVALUATED` with Commands, File Operations, and both interactive viewer entries
  pending. Release and live-provider evidence remain open.

Each delivered change updates the implementation, tests, and owning specification together. The current domain contracts
govern shipped behavior until those changes land.

## Target workflow

Submitted actions use the same File Operations task, with the existing delayed card reveal for quick success. Ask only when
the answer changes what happens. For task decisions, the card and Decision Mirror share one question and accept one answer.
Use ordinary progress for waiting and one result for a failure, partial operation, or uncertain outcome. Keep technical
reasons and optional diagnostic facts in the existing Issues details. Provider capabilities never create a new preference,
success notice, or confirmation.

| Action | What the user sees |
|---|---|
| F5, F6, Paste, drag/drop, or a destination picker with a supported route | Start with the captured defaults. A successful native Move on the same FTP/SSH endpoint remains a Move. `Ctrl`+drag within a folder copies under a different name. |
| Routine Move that can only copy | Before queueing, ask “Copy the items and keep the originals?” with **Copy** and **Cancel**. For a mixed selection, show the move count and the copy count with originals kept, with **Continue** and **Cancel**. Include the destination. |
| Shift+F5 / Shift+F6 or an ingress that explicitly requires options or approval | Keep one start surface. Show only relevant options and include any planned Copy-only items in that surface. Name Shift+F6 **Move with Options**. |
| F2, Batch Rename, or Change Case | F2 edits the name and commits on Enter. Batch Rename keeps its preview and Change Case its options. All three execute through the same task and use the same policy for a collision. A folder keeps its contents and is renamed as one item. |
| Delete at a location with Recycle support | Move to the Recycle Bin. When the ingress requires confirmation, as in ViewerSpace, say **Move to Recycle Bin?**. |
| Shift+Delete, or Delete where Recycle is known unavailable | Ask **Permanently delete?** once on the task after preparing the admitted items. Skip the Shell Recycle attempt when the provider has proved it cannot work. |
| A task cannot run yet | Show the reason on its waiting card and allow Cancel. Ask about overlap only in Parallel when the person can actually authorize a different safe outcome. |
| An issue needs a decision | Offer the actions the provider and task can fulfill. Cancel is the safe default where an affirmative action changes data. A decision that would be alone in `More…` is a direct action. |

### Start and delete behavior

- Use one lightweight host confirmation on the existing routine transfer start path for a known Copy-only Move. Ask
  before creating, publishing or enqueueing a task, and before consuming Cut. Bind acceptance to the captured selection,
  plans and endpoints, then revalidate. Accepted Copy-only items stay source-kept even if a later probe enables Move. A
  changed endpoint, selection or plan that would widen the accepted operation needs fresh authority. Cancel leaves the
  cut sequence, task queue and filesystem untouched.
  Explicit options and any already-known risk that can share the same receipt appear in the same confirmation. If the
  preparation or execution proof later downgrades a Move, keep the source and follow the per-item decision gate before
  affected publication where possible. Report the actual source and destination outcome if the downgrade is discovered later.
- Cut, Paste, pane menus, Find, Viewer and Compare keep the source and destination roles they captured. A local folder picker
  returns a local destination explicitly, even from a remote pane. Clipboard Copy/Cut offers Windows file-drop data only
  for representable paths. An unavailable command leaves clipboard contents intact and uses the ordinary pane feedback if
  invoked by shortcut. Same-folder drag Copy validates provider case rules, rejects source aliases and uses the existing
  sibling naming policy.
- Use typed per-item Recycle availability (available, unavailable, or unknown). Keep the qualified bulk Local Recycle
  route for pane, context menu, Find and Viewer when applicable. A mixed selection separates Recycle and permanent groups
  within one request and identifies the items covered by permanent consent before their removal. If Recycle is unknown, the ordinary guarded
  attempt may run. Escalate only on a proved non-commit with the original still present. A known removal or a disputed
  result never grants another delete. Apply this rule to ordinary Delete and the existing Move “Delete source” choice.
- Keep exact or corroborated Delete authority, including its existing by-name consent when required. Explain permanent
  removal and any material by-name limitation in the question itself. Keep method details in Issues. A read-only override
  is an explicit decision only after a provider proves the blocker. `Apply to all` remembers it only for matching typed
  scopes in the task, with a fresh per-item proof. Retire the obsolete blanket `noncached` wording in the File Operations
  spec. A generic access denial and an uncertain mutation never authorize the override.
- Fold Leftover Cleanup and archive delete-after disclosure into their existing confirmation surfaces where the entire
  set can be prepared. Preserve the exact artifact `TouchGuardReceipt`, capture the approved items, and revalidate on the
  worker after waiting and before mutation. Keep the existing by-name handoff rules for Leftover Cleanup and the exact-only
  rule for archive cleanup. Pack/Unpack deletes eligible sources only after archive success and recheck. A changed set needs
  its own current authority or a failed/kept result. UI coalescing grants no additional deletion rights.

### Collision and risk decisions

- Use one repeat checkbox, **Apply to all**, on both card and mirror. Checking it and choosing an eligible action applies
  that action to remaining matching conflicts in this task. When **Skip item** is the only repeatable action, place the
  checkbox with Skip and label it **Apply to all when skipping**. **Retry** attempts only the current item and creates no
  grant. A new question starts unchecked. Remove the redundant generic **Skip all** and omit the checkbox when no later
  eligible item or repeatable action exists. Reuse the existing typed task, class, item and endpoint scopes, including the
  non-mutating failure families, and check that every later action is still allowed.
- Keep genuinely different actions. **Skip items with the same size and date** and **Replace all older** apply their own
  comparison rule. **Delete source** is the existing qualified Move choice with its own scope and consent. A destination
  that is out of space can offer Retry and a scoped skip. Where the provider can only stop work for that destination,
  offer **Skip remaining at this destination** as a direct action. Existing explicit bulk permanent-delete escalation and
  by-name grants remain separately labeled and never default. They are separate operations or scopes with a different result.
- Offer **Keep both** on a name collision when a qualified unique sibling can be made using the provider's collision key.
  An observed file collision offers **Replace** only where the exact approved occupant can be rechecked and consumed,
  with read-only protection and artifact receipts as applicable. Rename never merges folders or offers folder replacement.
  Replacing a link requires the existing exact-bound **Replace link** action and its explicit authority.
  Remove Retry for a path-length failure with no actionable retry. Explain a read-only occupied destination as one fact.
  For F2 with no later item, a single Cancel already leaves the item's name unchanged.
- Combine known risk decisions for the same item before content is published. Keep explicit consent for known sensitive
  loss, including Mark-of-the-Web (`Zone.Identifier`) on Move, EFS, other named streams and extended attributes. I18 owns
  any change to which losses require consent. Use a qualified owned stage when metadata can only be assessed during
  transfer. A direct-final route needs proof that publication satisfies the captured policy. Apply the same consent
  rules to sparse inflation and hydration before either action starts.
  A real pre-publication **Skip item** leaves the source and existing destination intact and accounts for the stage. If sensitive loss
  surfaces after publication without a matching prior grant, keep the Move source, show **Copied; source kept** with the
  reason in details and continue only where safe. Preserve the actual committed destination and any uncertainty.
- Plain-language copy describes the choice or outcome, including cancellation of running and waiting tasks. Use existing
  contextual failure text and put HRESULTs and provider details in Technical details. Normal capability limits, a routine
  wait, and a corroborated successful rename add no separate alert.

### Queue, progress, and results

- In Queue mode or when concurrent overlap cannot be approved, wait behind the captured predecessors using the existing
  overlap edge and `_queueOrderKey`.
  In Parallel, ask only when `CanRunConcurrently()` exposes a choice that changes execution. Revalidate all interlocks and
  capacity independently of the answer. `Start now` can release queue ordering for that task, never a safety or capacity
  wait. If the state changes as it is pressed, the task stays in a recoverable queue position with a precise wait reason.
- Persist **Queue / Parallel** as the default for newly submitted work, with Queue on a fresh profile. Task start options
  override that choice for one submission. Preserve one process-wide transfer budget. Per-job Auto probes storage and the
  provider, and admission caps combined work at the process budget, default four, including internal file workers and
  opaque reservations. Explain this in the existing setting. A separate change to the numeric limit needs paired
  evidence and preserves the shared bound. Use the existing I60 pause, fairness and admission ownership.
- Show Pause while a transfer runs, including during discovery-ahead. Pause all applies to current live tasks, including
  those still preparing or waiting. Preparing may finish nonmutating checks while its mutation remains paused.
  Resume all releases the pauses set by that action and leaves independent prior pauses, decisions, and overlap edges
  in place. Per-task Pause/Resume remains available. Release transfer capacity only after explicit pause reaches its
  quiet point, and reacquire before I/O. A question retains its existing reservation under I60.
- Exclude decision, credential, queue/overlap and live-output wait time from throughput and stall sampling. An actual
  I/O stall during active work still hatches. Publish entry and exit of a blocking interval through a monotonic generation
  in the existing task snapshot, including when the popup is hidden or another task is running. Use one blocked-state
  projection for sampling and the taskbar. A collected collision that lets work continue does not block telemetry. Stopping retains the recorded graph
  while genuine counters can finish updating. Preparing reports real elapsed time and a paused task does not suppress or
  distort the running-only ETA.
- Keep failures, partial/uncertain outcomes, and completed tasks with warnings visible outside the clean Completed group.
  A clean task with only `Information` notes follows the ordinary grouping, auto-dismiss and clear rules. Show any new
  routine notes only in Issues when informational notes are enabled. `Clear completed` acts on only the clean
  completed/canceled IDs covered by its painted header. Failure details and interrupted-Move recovery notices keep their
  existing explicit acknowledgement. Use one Issues entry in the popup's existing Options menu and the existing per-task
  path to details.
- The Decision Mirror stays passive while an edit field, terminal, or another window has focus. Make its existing status
  entry and keyboard route usable for explicit focus. It shares controls, target inspection and scope help with the card.
  Preserve safe focus restore, no accidental acceptance by a held key, and one answer per question.

### Options and other entry points

- Use one bandwidth preset list and Custom parser for Preferences, explicit start options, and the running card. Display
  IEC units and preserve existing saved rates and accepted input equivalents. A File Operations Copy link-policy default
  can override its existing provider default at admission. Preserve inherited behavior on upgrade until a user sets it.
  An explicit per-task Copy choice wins. Native Move retains its provider semantics.
- Capture verification availability in the plan. On routine defaults, indicate unavailable verification in outcome details
  if requested by the default. Explicit Verify stays disabled on an unsupported route, and a newly unprovable required
  verification gets a truthful result, including source retention when the Move requires that proof.
- Create Directory uses one name prompt, then validates and executes asynchronously with the existing task, interlock,
  exit, cancellation and result rules. Refresh only the admitted location, and use the same duplicate-name policy and
  plain errors. Label Create Directory as that operation on the task and in Issues. Use the same worker ownership and
  completion rules when moving Pack/Unpack off the UI thread. Preserve
  the archive format/overwrite options and cleanup authority. `Copy to folder…` and `Move to folder…` can share the existing
  destination picker and typed request when exposed. Routine informational tasks use the existing reveal policy.

## Shared file-system plugin contract

The host owns intent, admission, queueing, consent, progress and presentation. Each provider owns its namespace rules,
metadata, mutation and result proof. Use the typed `IFileSystemRouteCapabilities` query for the exact provider instance,
root, profile, path and operation. Capability JSON and connection names remain diagnostic. Missing or malformed executable
facts disable that route. Use the existing task, plan and conflict broker for every admitted operation.

- For Copy/Move, qualify Copy, native Move, managed publication and source cleanup independently. A Copy-only request
  cannot reach source deletion. For Delete, Recycle, permanent removal, and any read-only or by-name decision are distinct.
  For Create Directory, use the same child-name and namespace contract as Rename. Provider verdicts describe the item and
  phase that occurred, regardless of the ingress or whether the provider implements object binding.
- Make executable Rename capabilities explicit per route. The base Rename claim requires a qualified new-name operation.
  The typed route distinguishes mutation-time refusal of an occupied name, a destination probe, and coexistence by
  item ID. A separate capability qualifies replacement of approved regular-file occupants on a new unbound route.
  Advertise supported scope, path and collision rules, case-only behavior, link kind and cancellation. A provider
  can support new names while declining Replace or recursive Change Case. Admit only a proved bounded or watchdog route.
- Keep the existing `RenamePlan`, worker and scheduler. Exact-object execution uses
  `IFileSystemBoundObject::RenameIfUnchanged`. Give qualified unbound providers one shared typed route contract and one
  versioned provider extension where the existing calls cannot carry the witnesses and result. A new interface uses a new
  IID. Keep existing vtables stable. Use size-versioned records and validate values, lifetimes and callback ownership.
  Rebuild and test the host, every shipped file-system provider and shared ABI consumer in source-tree lockstep when the
  contract changes. Each provider supplies a tested implementation or an explicit unsupported result. Align singular and
  bulk Rename entry points with the same provider rules. A legacy flag-only call never bypasses the checks.
- Reuse the canonical provider name validator, join, case comparison and collision key. Keep the typed source/destination
  and interlock scope through queueing. Revalidate the source, parent and root using the proof required by the admitted
  route. Keep exact source/destination proofs wherever available, and use conservative overlap scopes with unbound
  ancestors. Diagnose observed collisions before touching the occupant. Refuse unsafe aliases and linked parents.
- Require an authoritative no-follow kind for each selected root and discovered child. Recurse through proven directories
  only. Symlinks, name-surrogate reparse objects and Drive shortcuts are leaves. A provider can rename the leaf itself if
  qualified. Preserve the existing classification for proven non-name-surrogate reparses. When kind, parent or containment
  cannot be proved, refuse the affected recursive request before mutation.

### Rename authority and qualification

| Route | New-name Rename | Replacement |
|---|---|---|
| Local and identity-capable SMB | Retain exact no-follow authority and perform the conditional rename. | Preserve the existing exact-bound grants for regular files, read-only files and links when eligible, with the corresponding explicit collision decision. |
| Identity-less Local/SMB | After a successful no-follow open proves `FileIdInfo` unavailable, reuse the by-name Delete corroboration and bounded interlock. Check the captured facts again on the retained source handle. Use an exclusive rename within the admitted parent under the handle rules below. | Deliver new-name Rename first. A later Replace path uses the shared expected-occupant contract and requires its own qualification. |
| FTP, SFTP, and SCP (Curl) | Use the existing same-endpoint R0f-Curl destination probe followed by the native rename command. Qualify observed-collision protection on each shipped protocol route, including SCP's SFTP control path. The server can replace an occupant that appears after the probe. | Keep Replace unavailable while the provider cannot consume an approved occupant and finish cleanly. |
| Google Drive | Rename the resolved item ID without moving children, trashing an occupant, or deleting the source folder. Check actual sibling names as well as exposed names. An ambiguous raw target gets a collision. Preserve exact exposed-name resolution, literal suffix-like names and case. If an external creator adds the same name after the check, preserve both IDs and use Drive's existing disambiguation. | Withhold until its expected-occupant and cleanup path passes independently. |
| OneDrive / SharePoint (Graph) | Use the resolved source ID and prove the provider rejects a conflicting name without changing the occupant. Support exact final case where advertised. Keep `allowDirectoryMerge = false` for Rename. | Withhold until its expected-occupant and rollback/cleanup path passes independently. |
| Dummy, S3, S3 Tables, MTP, 7-Zip, IMAP, and any other loaded profile | Dummy uses the contract for deterministic tests. Other affected profiles advertise unsupported wherever the operation is absent or unqualified. Test those negative routes. Native Move, Copy+Delete or writer staging alone cannot advertise Rename. | Use the same eligibility rule if a future route is implemented. |

#### Curl destination check

Reuse the existing direct Rename path from [R0f-Curl](../../FileSystem/FileSystem_FtpSftpScp.md#conditional-publication-and-identity-safe-recovery)
for F2, Batch Rename and Change Case. Revalidate the admitted source and namespace using the facts this provider exposes.
On the worker, probe the destination fresh immediately before the server rename. An observed occupant returns a known
non-commit collision with no rename or backup command. A failed or ambiguous lookup stops execution. Keep both uses a
different validated sibling and repeats the check for that name. Preserve the existing connection, host-key, deadline,
cancel and unload rules. Standard FTP `RNTO` has no portable no-overwrite selector, and curl's libssh2 quote path supplies
overwrite flags. [FTP command specification](https://datatracker.ietf.org/doc/html/rfc959#section-4.1.3) ·
[curl SFTP quote implementation](https://github.com/curl/curl/blob/curl-8_23_0/lib/vssh/libssh2.c#L1883-L1893)

The Curl host exception covers ordinary same-parent Rename to a probed free name on the same endpoint. If an external
client creates the destination after the probe, the server can overwrite it. Host interlocks only coordinate
RedSalamander work. A probe grants no atomic no-replace, replacement, exclusive temporary-hop or rollback authority.
Qualify the outcome that
each protocol actually produces, including a lost server response. Document the limitation in the provider contract and
user guide. Routine success stays quiet. The same delivery reconciles the general no-replace wording and the R0f-Curl
provider exception while preserving the existing Copy/Move contract. Google Drive has a separate outcome: late duplicates
can coexist by item ID. [Google Drive Files resource](https://developers.google.com/workspace/drive/api/reference/rest/v3/files)

#### Local and SMB corroboration

Capture kind, relevant attributes, size where meaningful, timestamps and available nonzero file-index/volume facts.
The fallback starts only after a qualified successful no-follow open with unsupported `FileIdInfo`. Interface
unavailability, missing objects, access/sharing failures, unknown tags and malformed results
retain their own outcomes. A replacement with identical facts cannot be distinguished by this method.

Qualify the same-directory leaf form of `FileRenameInfoEx` or `FileRenameInfo` on the retained source handle, with
`RootDirectory = nullptr` and replacement disabled. A qualified parent-handle-relative form may be used where supported.
SMB2 requires a zero root on the wire, so test how Windows redirects the leaf on Windows SMB and Samba, including the
working directory, share root and ancestor changes. The existing absolute-path buffer is eligible for this route only
with a proved namespace guard through the call. Refuse when the admitted parent cannot be protected. Record the method
as diagnostic detail where requested. Successful Rename has no by-name alert, and permanent-delete consent stays
explicit. [Windows rename forms](https://learn.microsoft.com/windows-hardware/drivers/ddi/ntifs/ns-ntifs-_file_rename_information) ·
[SMB2 rename structure](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-fscc/52aa0b70-8094-4971-862d-79793f41e6a8)

### Replace and mutation results

- Deliver restoring Rename to a free name independently of adding any new Replace capability. New remote Replace is
  disabled by default. Enable it per provider only after the mutation can consume the approved regular-file occupant's
  provider identity/revision or qualified corroboration facts, immediately before replacement in the provider. Require
  explicit scoped consent, approved kind and read-only checks, and normal clean finalization. `ALLOW_OVERWRITE` carries
  intent only. A changed or missing expected occupant returns a known non-commit, so the same bounded collision question
  can refresh. Unknown results terminate recovery. If unavailable, omit Replace and keep valid alternatives.
- Validate the receipt even when the provider returns `S_OK`. Record the primary mutation outcome, known source
  disposition, and any independently observed rollback or cleanup artifact. Reuse `FileSystemItemMutationResult` where it
  conveys those facts and version it if a required fact has no representation. Provider dispatch publishes a single
  item result with the request identity and attempt it belongs to. A proven pre-mutation refusal can return a known
  non-commit. Missing or contradictory evidence after a mutation-capable call preserves uncertainty and never permits
  replay. Cancellation after dispatch follows the same rule. When a primary commit is proved, keep that commit even
  if cleanup later fails and show the actual cleanup problem once. Preserve unknown or retained artifacts and do not
  infer success, source presence, or clean rollback from HRESULT alone.
- Before enabling provider-native dispatch, qualify the receipt paths reached by new-name Rename. Drive must avoid the
  `DriveTransferTree` merge path in provider Rename and keep a verified primary commit if `TrashReplacedItem` later fails.
  Graph `BuildMoveMutationResult` must distinguish proven pre-mutation refusal, uncertain primary response, failed rollback
  and committed-with-cleanup-debt. Curl must carry `CurlPublicationResult` facts to `ReportItemCompleted`; its existing
  overwrite path always keeps the backup once one is created. Preserve that artifact fact even if the primary result is
  unknown, and cover current Copy/Move callers of the shared helper. Ordinary new-name Rename stays out of those
  overwrite paths. Test replacement-only cleanup with the existing affected callers or in the later replacement slice.
  It does not gate new-name Rename unless they share a changed producer.
- Batch Rename and Change Case validate the full requested mapping, provider collision keys and unrelated destinations
  before first mutation. Preview/preflight reports a known unrelated occupied target with the ordinary row/task result
  and lets the user correct the selection, without partial work. A proved later collision uses the existing task decision
  where the dependency schedule permits that action. Keep mappings immutable once accepted and revalidate each hop.
  Dependency cycles and case-only fallback use the existing exclusive temporary-sibling, rollback and truthful result
  rules where that transport safely supports them. Refuse a mapping that needs an unqualified temporary hop before the
  first mutation. If a later leg is unconfirmed, avoid replay or speculative cleanup.
  A provider may use one atomic case-only operation. Require the final spelling, preserve unchanged contents, and account
  for each skipped, failed, pending or committed row.
- For exact-bound folder work, avoid retaining parent and child mutation handles across incompatible dependency layers.
  Compare ingress snapshots when rebinding and keep required ancestry protection. Restrict legacy rename API fallback to
  qualified known non-commits. A `FileBasicInfo` fallback classifies only ordinary objects unless the same object has a
  trustworthy reparse tag. A watcher-caused failure is a hypothesis until verified per server. Where proved, suspend only
  owned watches that block the mutation, off the UI thread, release their handles before the call, and restore or refresh
  on every outcome against the panes' current navigation generation. Preserve pending I/O cancellation and external changes.

## Acceptance scenarios

Use this table for the product walkthrough and its deterministic self-test counterpart. Inspect the actual source and
destination, task result and prompt count. Simulated provider tests exercise the host path. Live qualification records
the exact server/profile tested.

| Scenario | Expected result |
|---|---|
| F2 on a folder on a qualified NAS, with and without file IDs | The folder has the new spelling and all children remain. Focus and pane contents stay coherent. Success produces no question or method note. Test a watched descendant, a changed parent, and decoy target names in the working directory and share root. |
| F2, Batch Rename and Change Case on qualified FTP, SFTP, SCP, Drive and Graph | A free name works and leaves no temporary artifacts. An observed occupied name preserves both objects. Use Keep both, Skip item where useful, or Cancel. Test provider case rules, folder collisions, and a Change Case link with a target outside the selected tree. |
| A name appears or changes during Rename | A true no-overwrite route preserves the late creator and reports the conflict. Drive preserves both item IDs. Inject a Curl destination just before the fresh probe and just after it. The first case preserves both objects without a rename command. In the second, verify the actual protocol outcome, including replacement on a fixture that permits it. A changed approved occupant invalidates a Replace decision. An unconfirmed mutation yields one uncertain result and no automatic retry. |
| A rename commits, then cleanup fails | Show the committed name and one cleanup issue, including any retained artifact. A lost primary response without proof stays uncertain. Test Drive cleanup, Graph refusal/rollback and cleanup, and the existing Curl retained-backup path in the providers. |
| F6 where only Copy is possible | The short Copy/Cancel question opens before a task exists. Copy submits once and keeps the originals. Cancel leaves the filesystem, queue and cut list untouched. Test mixed selections and endpoint changes while the question is open. Native Move starts directly. An explicit options invocation includes the disclosure in its existing dialog. |
| Context-menu Move from a remote pane to a local folder, or Ctrl+drag within a folder | The named source and destination are the ones captured by the command. Same-folder Copy makes a sibling with no self-overwrite. An unsupported Windows clipboard route writes no file-drop data. |
| Copy 1,000 files with three matching name clashes | The first shows one applicable set of actions. **Apply to all** plus **Keep both** or **Skip item** handles the other matching clashes with no questions. An unrelated read-only/link or different destination scope can still need its own choice. |
| Several files hit the same retryable error | Card and mirror label the checkbox **Apply to all when skipping**. With it checked, **Retry** retries only the current item and creates no grant. Choosing **Skip item** with it checked skips later matching errors within the task. |
| Delete on a NAS with Recycle known unavailable, and Viewer Recycle on Local | NAS shows one **Permanently delete?** decision and makes zero Shell Recycle calls. Viewer says **Move to Recycle Bin?**. Cancel leaves the items unchanged. Read-only override and unknown-Recycle outcomes keep their separate safety tests. |
| Copy or Move would lose known security metadata | A pre-publication Skip leaves the source and final destination unchanged and accounts for any stage. If unexpected loss is discovered after a copy commits, the Move source stays and the result names the actual outcome. |
| Queue overlap, then Start now and Pause all | An ineligible overlap waits with a reason and a cancel action. Start now respects the interlock and capacity. Pause all holds current preparation/queued mutation and active I/O, and Resume all restores only its own pauses. |
| Answer a question after a minute with the popup hidden or while typing in the terminal | One answer is accepted after explicit focus. The editor/terminal keeps its focus until then. Resume preserves the graph history without a false stall; a real provider stall still appears. |
| Clear completed with clean, warning and failed tasks present | Clean covered cards disappear, including one with an `Information` note. That note is available in Issues under the opt-in while the task is retained. Warning, failed, partial, uncertain and recovery records remain available through the existing card/Issues path. |
| F7 against a slow provider | The name is accepted once. The pane stays responsive while the operation runs. Confirm completion/refresh, cancel or exit while blocked, and work in another pane during the wait. |

The provider/FileOps tests must also cover changed source/parents, precise case-only spelling, mixed identity, alias
collisions, lost server responses, failed opens, unknown link kinds, temporary-hop rollback collisions, malformed and
duplicate receipts, failed cleanups with a successful HRESULT, and cancellation before and after dispatch. For provider
refusals assert the observable evidence appropriate to that route. Keep every fixture, including link targets, inside its
governed sandbox. Exercise published hosts and provider calls separately where the behavior changed in both.

## Delivery order

Each row can land with its own tests and owning spec updates. The first Rename delivery is complete for a route once its
free-name path, collision behavior and results are qualified. Replacement capability on another route does not gate
that delivery. The small ingress fixes and the telemetry correction can land independently while Rename work proceeds.

| Order | Work and exit |
|---|---|
| 1 | **S1, exact Rename.** Qualify handle retention, protocol fallbacks, case-only hops, and the watcher hypothesis with exact Local/SMB positives and negatives. Fix the watcher only on the routes where it blocks, with navigation/teardown tests. |
| 2 | **S1, new-name Rename.** Land corroborated SMB first. Then qualify Curl's existing destination-probe path, Drive, and Graph independently under the [route contract](#rename-authority-and-qualification). Wire all three Rename entry points to the same task and hold unqualified profiles disabled. Cover the existing callers of any changed receipt helper. |
| Independently, early | **S0 and the short Copy-only Move question.** Correct Viewer Recycle wording, remote context Move direction, clipboard and drag/drop guards, Cancel all, blocked Retry/read-only copy, and bulk Delete ingress parity. Add known-unavailable Recycle routing and the short Copy-only Move flow. Exercise affected ingresses and prove no extra questions. |
| 3 | **S3 and S2.** Use the single repeat control, coalesce artifact decisions and reconcile read-only consent, simplify overlap and sensitive-loss handling, and publish accurate progress on wait/resume. S2 can land as soon as its integration and rate evidence are ready. |
| 4 | **S4 and S5.** Close the queue controls and persistence, explain the shared cap, preserve problem results, and make the mirror and Issues entry consistent. Coordinate queue work with I60 and collected prompts with I64. |
| 5 | **S6 and S7.** Finish option parity, async Create Directory and archive work, and useful plain-language results. Keep async Create Directory and archive worker changes separately reviewable. |
| Throughout | **S8.** Update end-user docs and owning specs with each slice. Remove the obsolete queued destination selector, cover cut/paste Move, and synchronize options, conflicts, qualified provider behavior and outcome explanations. |

**Optional follow-up scope.** Expand Replace or strengthen Curl's rename primitive separately where a provider qualifies
it. Numerical tuning of the transfer budget needs its own paired evidence. Search, filtering,
export, bulk clear and reveal within Issues, a broader destination editor, and other unsupported-provider features
belong to separate proposals. The existing `Clear completed` action and Issues entry are part of I66.

### Specification and work ownership

| Owner | Required updates and coordination |
|---|---|
| Engine and UI | [FileSystem_FileOperations.md](../../FileSystem/FileSystem_FileOperations.md), [UI_FileOperationsPopup.md](../../UI/UI_FileOperationsPopup.md), [UI_FolderView.md](../../UI/UI_FolderView.md), [Core_FileSystemBridge.md](../../Core/Core_FileSystemBridge.md), and [docs/FileOperations.md](../../../docs/FileOperations.md). Record choices, result rules and user-visible workflows with the behavior change. |
| ABI and providers | [Plugins_PluginAPI.md](../../Plugins/Plugins_PluginAPI.md), [Plugins_VirtualFileSystem.md](../../Plugins/Plugins_VirtualFileSystem.md), [FileSystem.h](../../../Common/PlugInterfaces/FileSystem.h), [FTP/SFTP/SCP](../../FileSystem/FileSystem_FtpSftpScp.md), [Google Drive](../../FileSystem/FileSystem_GoogleDrive.md), [Microsoft Drive](../../FileSystem/FileSystem_MicrosoftDrive.md), and other provider specs whenever their capability truth changes. Reconcile stale provider declarations and contradictory general/provider-specific no-replace rules. |
| Settings and tests | [Core_SettingsStore.md](../../Core/Core_SettingsStore.md), [SettingsStore.schema.json](../../SettingsStore.schema.json), [UI_PreferencesDialog.md](../../UI/UI_PreferencesDialog.md), [Testing_SelfTests.md](../../Testing/Testing_SelfTests.md), and [Testing_PerformanceValidation.md](../../Testing/Testing_PerformanceValidation.md). |
| Existing plans | [I18](UI_ConfirmationPreferencesAndOperationOutcomes_2026-09-08.md) owns confirmation preferences and metadata policy. I66 keeps its sensitive-loss protections and advances the independent pre-publication work. [I64](FileOperations_CollectedCollisionsAndResiduals_2026-11-11.md) owns collected collisions and its popup budget. [I60](PreProduction_ReadinessAndSimplification_2026-11-07.md) owns admission fairness and pause-slot yielding. [I25](FileOperations_S3PathStyleCopyAndHarnessResiduals_2026-09-15.md) and [H3](Operation_FileOperations_LaterFeatures_2026-09-05.md) keep their provider and later-feature scope. [I29](FileOperations_PostMergeQualification_2026-10-05.md) and [I3](FileOperations_RemainingStressValidation_2026-08-04.md) keep their qualification scope. |

### Evidence and closeout

- Run `Test-Changes.ps1 -Explain` and the affected checks during implementation. Use focused FileOps, Commands,
  provider contract and native UI cases where required, in the approved marked sandbox. Use `Test-Changes.ps1 -Mode PrePush`
  before pushing and `Tools/Run-AllTests.ps1 -Suite Full -ValidationMode Fresh` when Full is required. Use zero-warning
  native profiles and the owning build/platform policy. Native ARM64 execution and live-server evidence have separate gates from cross-builds
  and owned fixtures. Keep unavailable gates open or transfer them to an accepting owner with the exact missing evidence.
- Check localized labels, focus, action equality, and UIA at the owning surface's minimum size, including the popup's
  480 DIP at 96/144 DPI and the resource satellites. Use bounded non-activating harness scenarios and invented fixture
  data for UI captures. Ask separately for the time of any test that requires real desktop focus.
- Pair performance runs when implementation affects a measured path, under the existing performance policy. Reuse a
  matching source/profile fixture and instrumentation. A wording or action-list correction uses behavior and layout
  assertions. A new route without a working baseline gets bounded-cost evidence and a positive outcome. Keep paired
  artifacts and source identities in the governed evidence archive.

| Path that can regress | Scenario and measurements |
|---|---|
| Bulk Recycle and Rename | Compare bulk and current Recycle ingress on the same batch, including Shell calls, elapsed time and retained resources. For the existing Rename path and changed batch/discovery work, compare repeatable single and 100-item runs, bind/I/O counts, handle peaks, and cancellation latency. Record the watcher/remote cost only when its path is exercised. |
| S2 telemetry | Compare fixed active, waiting, paused and stopping mixes. Measure timer cadence, `UpdateRates` cost and allocations per sample, and retained history. Require zero new allocations on the same steady-state scenario and zero false stall samples. |
| Queue and popup paths | Where a change affects hot admission predicates, redraws, layout or retention, reuse existing queue/popup metrics and I60/I64 fixtures. Measure wait/admission latency, fairness and permit counts, paint/sync and idle work where affected. Collect fresh paired performance evidence only for the paths changed. |
| Concurrency tuning, F7 and archive workers | For a numeric cap change, compare Release NVMe/SMB workloads with multiple jobs, latency, throughput, handles and memory. For async commands, measure UI dispatch responsiveness, completion and cancellation/exit latency with a stalled controlled provider or worker. |

Each slice includes its dated `Changes/` fragment and passes `Tools/Get-SpecInventory.ps1 -FailOnFindings -NoReports`
and `Tools/Test-Changelog.ps1`. Close I66 after the required delivery rows, workflow scenarios, provider gates and
authoritative updates are accounted for. Move it to `Specs/Plans/Done/`, admit the exact immutable path with its
positive and negative history test, and remove its WIP row. Record the receiving owner for any transferred evidence gate.

### STOP conditions

Stop the affected slice and record the failing case if any of these conditions holds.

- Exact `FileIdInfo` authority would be weakened, a retained-handle rename can escape the admitted parent, or recursive
  Change Case can follow a link.
- A route can modify an observed occupant without qualified Replace authority, `ALLOW_OVERWRITE` alone would grant
  replacement, or a new route relies on the Curl exception outside the documented R0f-Curl scope.
- A missing receipt or unknown mutation could trigger retry or cleanup, or a committed destination could be reported
  as `Skipped`.
- Cleanup would bypass the artifact `TouchGuardReceipt`, provider proof or the required worker-side revalidation.
- The queue change needs a second ordering key or permits `Start now` to bypass safety or capacity.
- A fix needs a DxUi dependency pin change. Route the producer fix and consumer qualification through their owning
  work. Keep the current pin in I66.

---

## Source study at ac619f860

The following sections describe the checkout before implementation. They retain the source anchors and original observations.
The target workflow and plugin contract above govern implementation choices. Server-specific behavior and UI timing that were
inferred by source inspection need their own reproduction and qualification.

### Drift check

Run from the repository root before starting any slice:

```powershell
git diff --stat ac619f860..HEAD -- Common RedSalamander Plugins Specs docs Tools Tests
git status --short
```

All line numbers and symbol anchors in this document were verified by read-only inspection of `ac619f860` (plus
RedSalamander's pinned DxUi revision `7fb2ab71852087d93b3a3e5b3ba504c37f56267c` in `Dependencies/DxUi.lock.json`) on
2026-11-23. Re-anchor by symbol name before editing.

---

### 1. Source findings

#### Q1. Is the File Operations workflow always good and coherent?

**Honest answer: No — not yet.**

RedSalamander's **safety and truth core is genuinely strong** and worth preserving verbatim:

- Immutable typed plans (`TransferPlan`, `RenamePlan`, `DeletePlan`), exact-object no-follow binding, atomic `.rs-tmp` /
  `.rs-bak` staging and conditional publication, and structured per-item outcome axes (`PublicationState`,
  `VerificationState`, `SourceDisposition`, `OwnedStageDisposition`, `ItemCompletion`).
- Durable interrupted-Move breadcrumbs that stay standalone across restarts and are never wiped by `Clear completed`.
- Non-activating File Operations popup, off-UI Preparing/discovery/shutdown drain, storage-medium-aware auto-concurrency
  (`HDD = 1`, `USB = 1..2`, `SSD = 4`, `NVMe / Network = 8`), and event-driven keep-awake with a 2-minute quiet release.
- Generation-tokened `FileOperationsConflictBroker`, FAT 2-second timestamp tolerance, DST 1-hour detection, in-pane
  conflict inspection with rollback, and the `I63` D-6 "never show a menu with only 1 item" polish rule.

However, the **end-user workflow built on top of that engine is NOT coherent yet**. Because the safety core, popup UI,
start confirmations, Decision Mirror, Issues Pane, Queue scheduler, Preferences page, and non-Local providers evolved in
separate waves (`I35`, `I63`, `I18`, `I64`, and `I60`), end users hit sharp seams:

1. **Three incompatible confirmation paradigms** coexist (blocking sync modals for `Shift+F5`/`Shift+F6`/Copy-only Move/Compare Move, async card prompts for `Shift+Delete`, and zero options/confirmation for routine `F5`/`F6`/Paste/Drag&Drop).
2. **Broken Rename on identity-less SMB/NAS AND across all non-binding providers (FTP, SSH [SFTP/SCP], Google Drive, OneDrive)**, plus folder-specific handle-retention defects and SMB protocol fallback gaps.
3. **Redundant and contradictory conflict controls** (`Skip item` + `Apply to all` checkbox + `Skip all`, dead checkboxes on 4(+2) failure buckets, `Apply to all` and `* all` offered on 1-file tasks and F2 inline rename, `Delete source` buried inside `More...`).
4. **Useless, double/triple, or technically opaque prompts** (Viewer Recycle lying *"Permanently delete"*, `MetadataLoss` firing *after* committing copies of browser-downloaded files with MOTW, `SameHostOverlap` interrupting even in Queue mode, `PathTooLong` offering `Retry`, double/triple prompts on Leftover Cleanup, Pack/Unpack Delete-After, and Shift+Delete of read-only files).
5. **Throughput graph grey-zone corruption** when waiting on or resuming from user questions (and during `Stopping` / `WaitForLiveOutputPublisher` / `PromptForConnectionSecret`).
6. **Hidden cliffs in Parallel mode, `Start now`, `Clear completed`, Decision Mirror focus-stealing, Issues Pane discoverability, Bandwidth/Preferences inconsistencies, sync UI-thread F7/Pack/Unpack, and cross-provider Move asymmetry.**

---

#### Q2. Why can't I rename a folder on SMB, FTP or SSH?

The current host requires object binding. Its coverage misses host Rename on SMB and non-binding remote providers.
Source inspection also identifies handle-lifetime and protocol paths that can fail. Server-specific behavior needs
reproduction on the affected profile before it establishes the cause of a reported failure.

##### Cause 1. The host requires bound authority
- **What `BindObjectAuthority` actually returns (`FolderWindow.FileOperations.cpp:1172-1195`)**:
  `ObjectBindingState::Unsupported` is returned in **two completely different situations** that current code conflates:
  1. **Local/SMB binding provider (`IFileSystemObjectBinding` IS implemented)**: `FileSystem::BindObject` (`Plugins/FileSystem/FileSystem.cpp:7272-7404`) opens the item no-follow, then calls `GetFileInformationByHandleEx(file, FileIdInfo, ...)` (`:2997, 7353`). On NAS/Samba shares that do not support `FileIdInfo` (or lack `FileIdInfo` on directories/share roots even when files return it — documented at `FileSystem.cpp:173-174` and `FolderWindow.FileOperations.Tests.IdentityLess.cpp:1493-1504`), `BindObject` returns `ERROR_NOT_SUPPORTED` (`0x80070032`). *(Note: a failed open MUST stay `Indeterminate` / `ERROR_OPEN_FAILED`, never `Unsupported`.)*
  2. **Non-binding providers (`QueryInterface(IFileSystemObjectBinding)` returns `E_NOINTERFACE` at `FolderWindow.FileOperations.cpp:1173-1178`)**: FTP, SFTP, SCP, Google Drive, and Microsoft Drive advertise Rename on supported profiles and implement `IFileSystem::RenameItem` (`FileSystemCurl.CopyMove.cpp:2988`, `FileSystemGoogleDrive.cpp:2684`, `FileSystemMicrosoftDrive.cpp:6214`). These providers lack object binding, so the host's binding request returns `ObjectBindingState::Unsupported` (`0x80070032`).
- **Why these production Rename requests fail closed (`0x80070032` `ERROR_NOT_SUPPORTED`)**:
  Unlike Copy (`bindingRequired = false` + retained-handle stage nonce + `I63` `IFileSystemExpectedOccupantPublication`),
  Delete (`I35` `IFileSystemIdentityDelete` for identity-less Local/SMB + `DeletePlan::nativeAuthority` for FTP/SSH/cloud),
  and Native Move (`MoveItem` with `bindingRequired = false`), **Rename was excluded from both identity-less Local/SMB
  fallbacks (`Common/PlugInterfaces/FileSystem.h:1129`, `FileSystem_FileOperations.md:2067`) AND `nativeAuthority`
  provider-native dispatch (`IFileSystem::RenameItem` has ZERO production callers; `FileSystemRenameBatch.cpp` is
  `#ifdef ENABLE_TESTS`-only via `ChangeCase.cpp:619`)**. Every rename hits up to five hard `Bound`-only gates:
  1. `Task::PrepareMutationInterlockScopes` (`FolderWindow.FileOperations.State.cpp:12171`): calls `builder.Add(..., parent, WriteSource, /*bindingRequired=*/true, false)`. `MutationInterlockScopeBuilder::Add` (`State.cpp:11220-11514`) binds **the parent AND every ancestor up to `\\server\share` (or remote root)**. If the provider lacks `IFileSystemObjectBinding` (FTP/SSH/Drive/OneDrive) OR ANY SMB directory/share root lacks `FileIdInfo`, Preparing aborts with `0x80070032` (*"This operation is not supported for this item or location."*).
  2. `Task::ExecuteInlineRename` (`State.cpp:15261-15266, 15331-15343, 15382-15399, 15434-15450`): requires `Bound` source, parent revalidation, `Bound`/`Missing` destination, `RenameIfUnchanged`, and post-rename `RevalidateObjectAuthority`.
  3. `Task::PrepareBatchRenameAdmission` (`FolderWindow.FileOperations.cpp:3123-3134`): requires `Bound` ingress snapshot for every row.
  4. `ChangeCase::BuildOrApplyToPaths` (`ChangeCase.cpp:277-284`): requires `Bound` roots when `includeSubdirs` is enabled.
  5. `Task::ExecuteBatchRename` / `ExecuteBatchRenameMutation` (`State.cpp:15995-16002, 15666-15775`): requires `Bound` for every source and destination.

##### Cause 2. Folder-specific handle retention
- **2A. `ExecuteBatchRename` retains folder and child handles** (`FolderWindow.FileOperations.State.cpp:15983-16053, 15841`):
  Pre-binds EVERY item in the batch simultaneously with `FILESYSTEM_BIND_RENAME` (`FILE_READ_ATTRIBUTES | DELETE`) into
  `authorities[]`, and after renaming each child executes `sourceAuthority = std::move(returned.authority)` (line `15841`),
  keeping that child's `DELETE` handle open while moving to the next dependency layer to rename its parent folder:
  - A live child handle can block renaming its parent. Measure which combinations and servers exhibit `ERROR_ACCESS_DENIED` or `ERROR_SHARING_VIOLATION` in this path.
  - *Hypothesis on Samba/strict SMB*: holding a `DELETE` handle on the parent folder while Layer 0 renames children can also cause `ERROR_SHARING_VIOLATION` (`0x80070020`), matching the Samba behavior previously recorded in `FileSystem_FileOperations.md:1186`.
- **2B. Opposite-pane `DirectoryWatch` holds an open `FILE_LIST_DIRECTORY` handle on the folder** (`Plugins/FileSystem/FileSystem.Watch.cpp:123-129`, `RedSalamander/DirectoryInfoCache.cpp`):
  If the opposite pane is displaying the target folder or any subfolder inside it, `DirectoryWatch` holds an open
  `FILE_LIST_DIRECTORY` handle with pending `ReadDirectoryChangesW`. The handle grants `FILE_SHARE_DELETE` and Rename does
  not suspend the watch before mutating. Whether this handle blocks the actual rename is an endpoint-specific hypothesis.
- **2C. Ancestor `FILE_READ_ATTRIBUTES` handles retained throughout rename**:
  `PrepareMutationInterlockScopes` retains open `FILE_READ_ATTRIBUTES` handles on the parent and every ancestor up to the
  share root across `ExecuteInlineRename` / `ExecuteBatchRename` (*hypothesis*: strict NAS lease setups may reject folder
  renames while ancestor handles remain open).
- *(Ruled out by static audit: `FileSystem.DirectoryOps.cpp` closes enumeration handles immediately; `DirectoryInfoCache` sets
  `allowOffscreenWatch = false`; `IconCache` and `NavigationView` retain no filesystem handles.)*

##### Cause 3. SMB protocol paths in bound Rename
- **3A. Incomplete `FileRenameInfoEx` -> `FileRenameInfo` fallback** (`Plugins/FileSystem/FileSystem.cpp:4642-4658`):
  Only falls back on `ERROR_INVALID_PARAMETER`, `ERROR_INVALID_FUNCTION`, or `ERROR_NOT_SUPPORTED` — missing
  `ERROR_CALL_NOT_IMPLEMENTED` and `ERROR_ACCESS_DENIED` that `Common/LocalFileTransaction.h:328-329` explicitly handles for
  filesystems/redirectors rejecting `FileRenameInfoEx`.
- **3B. Missing two-hop case-only rename fallback in `RenameIfUnchanged`** (`Plugins/FileSystem/FileSystem.cpp:4603-4700, 4948-5200`):
  Unlike `MovePathInternal` (`FileSystem.FileOps.cpp:8968-8976`, which calls `RenameCaseOnlyWithTemp`), `RenameExactTo`
  attempts only a single `SetFileInformationByHandle(..., replaceIfExists=FALSE)`. On SMB/Samba servers where direct
  case-only `FileRenameInfo` returns `ERROR_ALREADY_EXISTS` or `ERROR_ACCESS_DENIED`, Inline Rename enters a dead-end
  conflict prompt (`destinationIsSource == true`, Replace disabled) and Batch Rename / Change Case fail.
- **3C. Unconditional `FileAttributeTagInfo` requirement in `BindObject`** (`Plugins/FileSystem/FileSystem.cpp:7360-7379`):
  No fallback exists if an SMB server does not implement `FileAttributeTagInfo`. **Safety constraint**: `BindObject` at
  `:7367-7371` requires `ReparseTag` for `IsReparseTagNameSurrogate(...)`. `FILE_BASIC_INFO` / `GetFileInformationByHandle`
  supply attributes **without** `ReparseTag`, so any fallback is legal **only for non-reparse items
  (`(FileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) == 0`)**; reparse items MUST either resolve a trustworthy reparse tag
  separately or fail closed.
- **3D. Reconciliation failure escalates straight to `ERROR_IO_INCOMPLETE` (`0x800703E4`)** (`FileSystem.cpp:4670-4694`,
  `State.cpp:15453-15458, 15845-15849`):
  Reports `"Outcome unknown"` and aborts without offering Retry.

##### Test coverage
- The only SMB rename self-test (`R0fSmb_LoopbackReadWriteCreateDelete` in
  `FolderWindow.FileOperations.Tests.Phases05_06.cpp:7592-7604`) calls the **unused legacy `IFileSystem::RenameItem`** on a
  single file (`rename-me.bin`). Zero self-tests run `AdmitInlineRename`, `AdmitBatchRename`, or `AdmitScheduledRename`
  over SMB (file or folder), and zero host-level tests run Inline F2 / Batch Rename / Change Case over FTP or SSH (SFTP/SCP).

##### Baseline Rename and Move by provider

| Entry Point | Identity-less SMB/NAS (Local, no `FileIdInfo`) — File / Folder | Identity-capable SMB — File | Identity-capable SMB — Folder | FTP / SSH (SFTP, SCP) (`FileSystemCurl`) — File / Folder | Google Drive / OneDrive — File / Folder |
|---|---|---|---|---|---|
| **Inline F2** | Refused without required binding | Exact route, risks 3A–3D | Exact route, risks 2B and 3A–3D | Refused without binding | Refused without binding |
| **Batch Rename** | Refused without required binding | Exact route, risks 3A–3D | Exact route, risks 2A/2B and 3A–3D | Refused without binding | Refused without binding |
| **Change Case** | Refused without required binding | Exact route, risk 3B | Exact route, risks 2A/2B and 3B | Refused without binding | Refused without binding |
| **F6 Move to sibling folder on the same endpoint** | Native route can use `MoveFileWithProgressW` | Native route | Native route, in-place rename blocked by `SanityCheckBothPanes` | Native when admitted by the connection (`RNFR`/`RNTO` or SFTP `rename`) | Native when admitted by the provider (`files.update` / Graph `PATCH`) |

---

#### Q3. Why do some questions repeat or misstate the action?

Complete audit of every ingress dialog and all 26 `ConflictClass` buckets identified **10 distinct classes of defective or
low-value user questions**:

1. **BUG — ViewerSpace Recycle Bin delete lies "Permanently delete ... This cannot be undone."**
   - `FolderWindow::ExecuteViewerAction` (`FolderWindow.FileOperations.cpp:3883-3891`) passes `FILESYSTEM_FLAG_USE_RECYCLE_BIN` + `requireConfirmation = true`.
   - `StartOperation` (`State.Runtime.cpp:1317, 1417-1434`) takes the `!deleteBypassesRecycleBin` `else` branch and formats **`IDS_FMT_FILEOPS_CONFIRM_PERMANENT_DELETE`** (*"Permanently delete {0}: from: {1}\r\n\r\nThis cannot be undone."*) for a Recycle Bin operation.
2. **Single-file tasks and Inline F2 Rename offer `Apply to all`, `Skip all`, `Replace all older`, and `Skip all with the same size and date`.**
   - `BuildConflictDecisionScope` (`State.cpp:5289-5375`) and `RequestBridgeConsent` (`State.cpp:16976`) never check whether the task can possibly hit a second item (`_sourcePaths.size() > 1u` or directory root). Only `RecycleFailed`/`RecycleFailedByName` (`State.cpp:24453, 24885`) check `_sourcePaths.size() > 1u`.
   - Renaming **one file** with F2 onto an existing name shows the `Apply to all` checkbox AND `Skip all` (or `More...` containing `Skip all` + `Replace all older` / `Skip all with the same size and date`), plus both `Skip item` and `Cancel` (which do the exact same thing on a 1-file rename).
3. **Double and triple prompts where the user already explicitly stated intent (plus a normative spec contradiction on read-only delete):**
   - **Leftover Cleanup** (`FolderWindow.LeftoverCleanup.cpp:1232-1328`): User selects leftovers -> clicks `Delete Selected...` -> Prompt 2: `ConfirmArtifactTouch` modal (*"This request touches N item(s) that may be left over..."*, `State.Runtime.cpp:1121`) -> Prompt 3: `PermanentDeleteConfirmation` card (`State.cpp:12670`) if permanent. `ConfirmArtifactTouch` also mints the `TouchGuardReceipt` checked on the worker at `State.Runtime.cpp:2329`. See [Start and delete behavior](#start-and-delete-behavior) for the combined confirmation.
   - **Pack / Unpack "Delete after"** (`FolderWindow.FileSystem.Commands.cpp:12548-12566, 12748-12767`): User checks *"Delete files from disk after packing"* / *"Delete archive after unpacking"* -> clicks OK -> Prompt 2: `ConfirmPermanentDeletePaths` modal *before* packing/unpacking -> Prompt 3: `ConfirmArtifactTouch` *after* packing/unpacking if any path matches an artifact pattern (`DeleteHandoffForPlan` returns `None` at `State.Runtime.cpp:211`).
   - **Shift+Delete on Read-Only items (`State.cpp:12670` -> `State.cpp:27544, 27571, 5592`) & Spec Self-Contradiction**: After confirming `PermanentDeleteConfirmation`, hitting any read-only file (`.git/objects/`, etc.) pops a second `ReadOnlyRegularFile` confirmation (`Permanently delete read-only`). Furthermore, **`FileSystem_FileOperations.md` contradicts itself**: line `1537` states read-only permanent-delete consent is *"noncached"*, whereas line `3922` explicitly permits task-scoped `Apply to all` caching for provider-proven `ReadOnlyDeleteBlocked` (matching `IsCacheableConflictDecision(DeleteReadOnly) == true` at `State.cpp:5783`). [Start and delete behavior](#start-and-delete-behavior) spells out the two grants and the repeat scope.
4. **`MetadataLoss` on Move interrupts everyday downloaded files AFTER already writing and committing the copy.**
   - `State.cpp:21046, 21088-21112`: On any bridge Move that cannot preserve `MOTW` (`Zone.Identifier`), ADS, or EA (e.g., moving almost any browser-downloaded file from NTFS to FAT32, exFAT, SMB/NAS, SFTP, or S3), RedSalamander writes and **commits** the destination file first (`PublicationTransaction::State::Committed`), then stops to ask *"Moving this item would lose security-significant metadata such as Mark-of-the-Web, alternate streams, or extended attributes"* (`[Move anyway] [Copy and keep source] [Cancel]`).
   - Publication may already have replaced an existing file. Cancel leaves the published destination and stops the rest of the Move. A real Skip requires the decision before publication. See [Collision and risk decisions](#collision-and-risk-decisions).
5. **`SameHostOverlap` asks `[Queue after these tasks] [Don't start]` even in Queue mode and when concurrent execution is forbidden.**
   - `State.cpp:12854-12933`, `State.Queue.cpp:1722-1823`, `FileOperationsInternal.h:3044`: `CanRunConcurrently()` is `false` for all transfer-vs-transfer overlaps (`serializedTransfers`), leaving only `[Queue after these tasks]` and `[Don't start]` — and fires even when `New tasks: Queue` is already active!
6. **`PathTooLong` offers a useless `Retry` button** (`State.cpp:4987-4993` — `IsRetryableConflictBucket` fails to exclude `PathTooLong`).
7. **`ReadOnlyRegularFileExists` hides the fact that the destination file already exists** (`State.Diagnostics.cpp:475-476` renders only `IDS_FILEOPS_CONFLICT_READONLY`: *"Item is read-only."*).
8. **Pre-publication deferred consents (`EfsPlaintext`, `SparseInflation`, `PlaceholderHydration`) offer NO `Skip item` button** (`State.cpp:5755`), forcing all-or-nothing Cancel of a multi-file folder transfer (and leaving dead `RetainSource` branches at `State.cpp:20171, 20531`).
9. **Full options modal for a routine Copy-only Move** (`State.Runtime.cpp:1174, 1275-1289`). Any planned Copy-only item forces the transfer options surface even when the caller requested the defaults. This follows the admitted strategy. A supported native Move within the same endpoint does not take this branch just because it uses FTP or a cloud provider.
10. **Destructive default on Popup `Cancel all`, wall-of-text jargon, and raw hex `0x...` errors:**
    - Popup `Cancel all` (`Popup.cpp:13511`) sets **`OK` as the Enter default** and says *"pending"* (`IDS_MSG_FILEOPS_CANCEL_ALL_POPUP`) while canceling running tasks too — contradicting App/Compare Close (`Cancel` default at `FolderWindow.FileOperations.cpp:4436`).
    - `PermanentDeleteByNameConfirmation` / `RecycleFailedByName` append a **92-word internal-implementation essay** (`IDS_FILEOPS_CONSENT_BY_NAME_CHECKED`, `RedSalamander.rc:2079`) and label the button *"Delete by name"* instead of *"Permanently delete"*.
    - Pane/command error alerts (`RedSalamander.rc:1406-1407, 1469, 1473, 1647-1648, 1660, 1669, 1928`) print raw hex `Error: 0x{0:08X}`.

---

#### Q4. Why do Skip and Skip all overlap?

##### Equivalent decisions in the broker
- `ConflictAction::SkipAll` is host-only (`FolderWindow.FileOperationsInternal.h:661`); the plugin ABI `FileSystemIssueAction` has only `Skip` (`Common/PlugInterfaces/FileSystem.h:217-228`).
- In `BuildConflictActionPolicy` (`FolderWindow.FileOperations.State.cpp:5755-5762`), for all 10 non-deferred buckets (`RegularFileExists`, `ReadOnlyRegularFileExists`, `TypeMismatch`, `DestinationLink`, `ReadOnlyRegularFile` [non-delete], `NameNotRepresentable`, `SharingViolation`, `AccessDenied`, `PathTooLong`, `NetworkOffline`), whenever `applyToAllEligible` is `true`, the engine adds **both** `ConflictAction::Skip` AND `ConflictAction::SkipAll` while **also** keeping `applyToAllEligible = true` (which renders the checkbox `IDS_FILEOPS_CONFLICT_APPLY_TO_ALL`: *"Apply this choice to remaining similar conflicts in this task"*).
- In `FileOperations::ConflictBroker::TryAcceptLocked` (`RedSalamander/FileOperationsConflictBroker.cpp:53-70`):
  - Clicking **`Skip all`** produces `Decision{ConflictAction::Skip, /*applyToAll=*/true, generation}`.
  - Checking **`Apply to all`** and clicking **`Skip item`** produces the **exact same** `Decision{ConflictAction::Skip, /*applyToAll=*/true, generation}` — stored in the exact same `decisionCache` slot, logged identically, returning the exact same `FileSystemIssueAction::Skip`.

##### History of the duplicate action
1. **Original (`00b353e6e`)**: Only `Skip` + `Cancel` + `Apply to all` checkbox existed.
2. **Post-v7.0.48 (`d9536a480`)**: Added `SkipAll` as an overflow (`More...`) shortcut. For `DiskFull` (`State.cpp:5603-5613`), the author cleanly removed both the checkbox and one-item `Skip` (*"Skip All is its own explicit task grant; no checkbox or one-item Skip is needed."*), but kept BOTH `Skip` + checkbox AND `SkipAll` on all 10 other non-deferred buckets.
3. **`I63` Slice C (`ded63e43b`, `State.cpp:5530-5539`)**: Added single-overflow promotion (*"a menu needs two commands"*). Whenever `SkipAll` was the sole overflow action, it popped directly into the primary button row right beside `Skip item`: `[Skip item] [Skip all] [Cancel]` directly underneath `[ ] Apply this choice to remaining similar conflicts in this task`!
4. **`I63` B4 & B6 (`ded63e43b`, `State.cpp:5656-5675, 5767-5777`)**: Added `DiscardSource` (*"Delete source"*), `SkipSameSizeAndDate` (*"Skip all with the same size and date"*), and `ReplaceOlder` (*"Replace all older"*). Because redundant `SkipAll` already occupied an overflow slot, adding `DiscardSource` bumped `overflowActionCount` from 1 to 2 — **hiding the genuinely distinct `Delete source` action inside `More...` solely because redundant `Skip all` sat beside it** (`State.cpp:9265-9294`)!

##### Failure prompts with a duplicate repeat choice
On **`AccessDenied`**, **`SharingViolation`**, **`PathTooLong`**, and **`NetworkOffline`** (`State.cpp:5338-5342, 5750-5779`), the prompt displays:
- Checkbox: `[ ] Apply this choice to remaining similar conflicts in this task`
- Buttons: `[Retry] [Skip item] [Skip all] [Cancel]`

Checking the box does nothing for `Retry` (`ConflictBroker.cpp:68`), nothing for `Cancel`, nothing for `Skip all`, and makes `Skip item` 100% identical to `Skip all`. The exact same degeneracy happens on **`TypeMismatch`** and **`NameNotRepresentable`** whenever `Keep both` is not offered.

##### Bulk action inventory

| Action | Buckets | Coexists with Checkbox? | Distinct From `Apply to all`? | Verdict |
|---|---|---|---|---|
| `SkipAll` (*"Skip all"*) | 10 non-deferred buckets | **YES** | **NO** (identical to Checkbox + `Skip item`) | Use one repeat control |
| `SkipAll` (*"Skip all"*) | `DiskFull` (`State.cpp:5603`) | NO (no checkbox, no `Skip item`) | Sole skip/repeat action | **Keep** (clean pattern!) |
| `PermanentDeleteAll` (*"Delete all remaining permanently / by name"*) | `RecycleFailed`, `RecycleFailedByName` when `>1` item | NO (no checkbox) | Explicit task-scoped escalation grant | **Keep** (coordinate button adjacency with `I64` RU8) |
| `SkipSameSizeAndDate` (*"Skip all with the same size and date"*) | `RegularFileExists` when Same size & date | YES (checkbox silently ignored) | **YES** — conditional policy (skips only matching pairs) | **Keep**, fix placement/1-file gating & clarify vs checkbox |
| `ReplaceOlder` (*"Replace all older"*) | `RegularFileExists` when Newer + Replace | YES (checkbox silently ignored) | **YES** — conditional policy (replaces only older pairs) | **Keep**, fix placement/1-file gating & clarify vs checkbox |

*(Also note: `More...` is currently rendered **leftmost**, before all primary buttons, on both Card (`Popup.cpp:12616-12633`) and Decision Mirror (`DecisionMirror.cpp:668-679`), and `Delete source` silently ignores `Apply to all` when dates/sizes differ (`ConflictBroker.cpp:62-69`) with its caveat hidden only in the Card checkbox tooltip (`Popup.cpp:9237-9247`) and completely absent on the Decision Mirror.)*

---

#### Q5. Why can the throughput graph advance during a question?

##### Rate sampling and graph ownership
- `DxUi::ThroughputGraph` (`DxUi/src/Controls/DxUi.Controls.cpp:5696-6360`) holds a 180-sample (`18.0 s` at `100 ms/sample`) ring buffer with **no** internal ticker and **no** grey-zone renderer.
- The grey hatched zone is drawn by RedSalamander's `FileOperationsGraphStallBand` (`RedSalamander/FolderWindow.FileOperations.Popup.cpp:265-333`), which hatches `palette.disabledText` at `0.45` alpha over every contiguous run where `history.stalledSamples[i] != 0u`. It was designed solely for real I/O stalls (`I35 R10.2`), and `Specs/UI/UI_FileOperationsPopup.md:2034-2036, 2054-2056` explicitly forbids stall/rate progression during a user question.
- Samples are produced by `FileOperationsPopupState::OnTimer` (`Popup.cpp:14976-15026`) -> `UpdateRates()` (`:6350`) -> `BuildRateSnapshot()` (`:6229`) -> `AdvanceRateHistory()` (`:4340-4525`) -> `AppendResampledRateSamples()` (`:4186-4222`).

##### Sample and wait transitions to verify

1. **Root Cause A1 (CRITICAL — Missing `MarkRateSamplingStateChanged()` on ALL conflict/decision transitions + `OnTimer` skipping `UpdateRates()` while hidden/minimized):**
   - `Task::MarkRateSamplingStateChanged()` (`State.cpp:14495-14498`) is called **only** in `SetPaused` (`:14315`), `SetWaitingInQueue` (`:14420`), and `SetQueuePaused` (`:14483`).
   - **Zero** conflict/decision transitions call it: `SetConflictPromptMetadataLoadingLocked` (`State.cpp:6213`), `SetConflictPromptLocked` (`:6142`), `WaitForConflictDecision` (`:6689-6889`, including `ConflictBroker::ResetLocked` at `:6855`), `ClearConflictPrompt` (`:6107`), or `SetAwaitingUserDecision` (`:14793`).
   - `OnTimer` skips its regular `UpdateRates` call for a hidden or minimized popup (`Popup.cpp:15001-15005`). `RefreshTelemetryTimer` can still synchronize rates once on an aggregate waiting transition (`:7147-7163`), so the exact gap depends on which tasks are active and which transitions the sampler observes.
   - If an enter/exit pair is missed and the state-change tick is unchanged, the next `AdvanceRateHistory` can skip the reset (`Popup.cpp:4364-4390`). The maximum resample window is 18 seconds (`:4189-4190`), enough to displace the prior history. Test hidden and visible paths with both a single waiting task and a simultaneously active peer.

2. **Root Cause A2 (CRITICAL — Visible-popup 1 Hz waiting-timer gap on answer):**
   - Even with the popup visible the entire time, `RefreshTelemetryTimer` slows to `1000 ms` (1 Hz) during a loaded conflict (`Popup.cpp:7117-7213`).
   - When the user clicks an answer `600..999 ms` after the last 1 Hz tick, `prompt.active` clears without calling `MarkRateSamplingStateChanged()`.
   - At the next `100 ms` tick, `nowTick - stallReference` is already `700..1099 ms > 600 ms` (`RateStallThresholdMs`) -> immediately flashes false `"Speed: — · Waiting for the destination…"` and injects ~10 grey-hatched 0-rate buckets (~1.0 s of grey zone) for the user's decision time.

3. **Root Cause B (CRITICAL — Prompts & post-question waits that NEVER set `_conflictArbiter.prompt.active` keep the graph VISIBLE and LIVE-ADVANCING grey at 10 Hz!):**
   - **`WaitForLiveOutputPublisher`** (`State.cpp:12996-13004`, `State.Queue.cpp:2163-2191`): After the user answers `SameHostLiveOutput` with *"Queue until other task finishes"*, `WaitForConflictDecision` clears `prompt.active = false` and blocks on `_queueCv.wait(lock)` without setting `_waitingInQueue`, `_queuePaused`, `_waitForOthers`, or `_paused`. The throughput graph immediately reappears as `Running` and scrolls a grey hatched zone across the entire plot at 10 Hz, wiping 100% of real history in 18 seconds.
   - **`PromptForConnectionSecret`** (`Plugins/FileSystemCurl/FileSystemCurl.Shared.cpp:1181, 1292`): Modal FTP/SFTP/SCP password or SSH key passphrase prompt blocks the worker while `prompt.active == false` and `_awaitingUserDecision == false` -> same live 10 Hz grey-zone wipe of all 18 s of real history.
   - **`BatchRenameArtifactPrompt`** (`State.Runtime.cpp:2218`): Sets `_awaitingUserDecision`, which `BuildRateSnapshot` (`Popup.cpp:6318`) ignores.
   - **Pre-prompt slow `BindObject` I/O in `Task::FileSystemIssue`** (`State.cpp:10423-10448`): Runs before `prompt.active` becomes `true`, painting a grey stall band whenever remote metadata binding takes `> 600 ms`.
   - **`waitingForOthers` omitted** from `IsRateSamplingBlocked` (`Popup.cpp:3712`), `GetPopupTelemetryWork` (`State.Runtime.cpp:2781`), and `MarkRateSamplingStateChanged` (`State.cpp:14380`).

4. **Related Telemetry & Progress Bugs Discovered in the Same Audit:**
   - **`Stopping` erodes real throughput history to 0 over 18 s**: `task.stopping` is absent from `IsRateSamplingBlocked` (`Popup.cpp:3712`), so `AppendResampledRateSamples` keeps appending decaying-to-0 samples at 10 Hz — violating `UI_FileOperationsPopup.md:488` (*"Stopping retains actual counters, paths, result facts and throughput history"*). Also `taskMayAnimate` (`Popup.cpp:9304`) still allows marquee animation because `StatusCountsAsRunning(Stopping)` is true.
   - **Manually `Paused` Task B kills the Footer Aggregate ETA for concurrently running Task A**: `Paused` keeps `includeInAggregateCohort = true` (`Popup.cpp:1462`) with `transferBytesPerSec == 0.0`, which trips `etaInputsComplete = false` at `Popup.cpp:2330`.
   - **`Preparing... ({0} elapsed)` is stuck at `0s`**: `StatusTextForTask` (`Popup.cpp:1834-1836`) reads `task.operationStartTick`, which stays `0` until `ExecuteOperation` (`State.cpp:23856`).
   - **Why existing tests missed Root Cause A**: `CheckRateBlockedStates()` (`FolderWindow.FileOperations.Tests.Throughput.cpp:1016-1059`) tests `waitingForDecision = true` for 10 s and returns without ever transitioning `waitingForDecision` back to `false` or simulating a 1 Hz / hidden-popup gap.
   - **Forward coordination with `I64` Slice G**: When collected collisions (`prompt.active && prompt.collected`) land, the task keeps transferring while the collected prompt is shown — so `IsRateSamplingBlocked` will need `prompt.active && !prompt.collected` (and `ResolveWholeTaskProgressPresentation` will need to account for collected prompts).

---

#### Q6. What else is inconsistent?

Additional issues beyond Q2–Q5 are listed below. See [Specification and work ownership](#specification-and-work-ownership)
for related plans.

1. **[P0 BUG] Context-menu `Move...` (`FolderView::MoveSelectedItems`) inverts source/destination on non-local panes & lacks a Copy counterpart** (`FolderView.FileOps.cpp:1295-1379` omits `sourceContextSpecified = true` and uses local-only `IFileOpenDialog`; `FolderWindow.FileOperations.cpp:3541-3573` interprets `!sourceContextSpecified` on a non-local pane as an inbound local drop/paste).
2. **[P1 BUG] `FolderView::CopySelectionToClipboard()` writes raw non-local virtual paths into Win32 `CF_HDROP`** (`FolderView.FileOps.cpp:763-783` lacks the `SupportsFileDropCut()` guard present right below it in `CutSelectionToClipboard()` at `:789`).
3. **[P1 BUG] Ctrl+Drag Copy into the same folder is silently cancelled**, unlike `Ctrl+C` + `Ctrl+V` (`FolderView.DragDrop.cpp:701-711`), and uses `AreSameLocalVolume`/`EqualsNoCasePath` even for case-sensitive remote internal drops (`:696, 701-706`).
4. **[P1] Delete `ExecutionMode` mismatch degrades Shell Recycle Bin outside `Del`**: `CommandDelete` (`FolderWindow.FileOperations.cpp:4656`) uses `BulkItems` (and misses `EnsureFileOperations()` at `:4610`), while context-menu Delete (`:3713`), Find Results Delete, and Viewer Recycle (`:3963`) pass `PerItem` (1 STA thread + 1 shell call per item).
5. **[P1 UX Proposal] `Clear completed` permanently removes Failed, Partial, Warnings, and Indeterminate cards + their Issues diagnostics in one unconfirmed click** (`Popup.cpp:2394-2397, 2442-2445, 10365-10369, 15383-15386`, `State.Runtime.cpp:3303-3331`):
   - *Verified baseline*: `_completedGroupExpanded` starts **`true`** (`Popup.h:2115`), automatic collapse happens ONLY after an accepted `Cancel all` when every member ended clean (`UI_FileOperationsPopup.md:234-240`), and automatic dismissal (`IsAutoDismissableFileOperationCompletion`) vs explicit user `Clear completed` are separate contracts.
   - *UX friction*: Failed/Partial/Indeterminate cards are grouped inside `Completed (N)` alongside clean Done/Canceled tasks (unlike standalone interrupted-Move notices), so explicit `Clear completed` erases problem cards and their Issues rows without confirmation (and manual collapse folds problem cards behind counts).
6. **[P1] Parallel mode global serialization & broken `Start now` dead-end trap**: Any folder transfer, overwrite, or non-Local/identity-less transfer sets `activeExclusive = true` (`State.cpp:11563-11633`, `State.Queue.cpp:2415-2448`), serializing completely unrelated drives/providers. Worse, clicking `Start now` on a queued task (`State.Queue.cpp:2558`, `State.Runtime.cpp:2844-2864`) clears `_waitForOthers`, immediately slams into `activeExclusive`, **makes `Start now` disappear**, leaves the task stuck waiting, and removes it from the reorderable queue with no way back!
7. **[P1] No `[Pause]` button on cards during concurrent discovery-ahead** (`Popup.cpp:12641-12828`) even while files are actively copying.
8. **[P1 Contract Evolution — Owner Approved] `maxParallelTransfers` scope clarity vs storage-adaptive `Auto` budget**:
   - *Verified baseline*: `FileSystem_FileOperations.md:2295-2299` **intentionally** defines `fileOperations.maxParallelTransfers` (default `4`, range `1..16`) as one process-wide file budget covering both jobs AND their per-item workers (`State.cpp:14376, 25954`). It is a documented resource guarantee, not an accidental bug.
   - *UX/policy gap*: `IDS_PREFS_FILEOPS_PARALLEL_TRANSFERS_DESC` (*"Share this limit across file-operation jobs"*) obscures that it also caps a single job's storage-adaptive `Auto` budget (clamping NVMe `8` and SMB `8` to `4` in `FileSystem.cpp:1689, 1714`), and interacts confusingly with session-only `Queue` mode and per-plugin concurrency settings. See [Queue, progress, and results](#queue-progress-and-results) for persistence and the shared budget.
9. **[P1] Queue vs Parallel mode (`_queueNewTasks`) is NOT persisted** (`FolderWindow.FileOperationsInternal.h:3333` vs `Common/SettingsStore.h:349-380`), silently reverting to Queue on restart, and has inconsistent scope between Start Modal (per-task) and Popup Footer (session-global).
10. **[P1] Decision Mirror steals keyboard focus mid-keystroke** (`DecisionMirror.cpp:272-283, 727` checks only top-level window foreground, not active text inputs/filter/address bar/inline rename/embedded terminal) and lacks Popup Card parity (Off/On selector vs Checkbox, no `Show source`/`Show destination`/path copy, missing `Delete source` scope help).
11. **[P1] Issues Pane discoverability & missing basic tools**: No `Issues` button in Popup Footer or Options menu (`Popup.cpp:10220-10290, 13582-13642`); pane lacks search/filter, inline `Show informational notes` toggle, `Export`, `Clear`, and `Reveal in pane` (`IssuesPane.cpp:1509-1572`). Also `OperationStringId` maps `FILESYSTEM_CREATE_DIRECTORY` to `"Unknown"` (`IssuesPane.cpp:86`) vs `"Copy"` (`PromptText.cpp:294`, `Popup.cpp:11177`), and `FormatDiscoveryProgressText` (`Popup.cpp:529-532`) has a duplicate unreachable `FILESYSTEM_DELETE` branch.
12. **[P1] Bandwidth speed-limit presets disagree across all 3 surfaces + mixed IEC/SI copy**:
    - Preferences (`Preferences.FileOperations.cpp:44-52`): `1..500 MiB/s` + Custom (**missing `1 GiB/s`**).
    - Start Modal (`HostServices.cpp:2141-2150`): `1..500 MiB/s, 1 GiB/s` (**no Custom input**).
    - Running Card (`Popup.cpp:13686-13693`): `1..100 MiB/s, 1 GiB/s` + Custom (**missing `500 MiB/s`**).
    - Custom hints advertise `KB`/`MB`/`GB` while parsing binary IEC units (`TryParseBinaryThroughputText`).
13. **[P1] Sync UI-thread bypasses**: F7 Create Directory (`FolderWindow.FileSystem.Commands.cpp:10596-10760` blocks the UI thread on slow SMB/cloud, skips `RefuseFileOperationForPendingExit`, and reports invalid names as *"does not support creating folders"*) and Pack/Unpack (`:12548-12850` run synchronously on the UI thread, Local-only, ad-hoc Skip/Replace modal).
14. **[P2] Hidden/silent options & command-label lies**:
    - `IDS_CMD_MOVE_WITH_OPTIONS` (`RedSalamander.rc:2190`) says **"Move/Rename with Options"**, yet Shift+F6 has no editable destination/name field.
    - `Links` policy (`Preserve`/`Skip`) is prominent in Shift+F5 yet absent from `Preferences > File Operations` (stored only in hidden per-plugin JSON `reparsePointPolicy`, `State.cpp:1437-1475`).
    - Enabled `verifyAfterCopy` is **silently turned off** on routine F5 to unsupported destinations (`State.Runtime.cpp:1237-1244`) with zero note.
    - `InformationalTask` bypasses the 500 ms reveal delay (`State.Runtime.cpp:3630`); `Pause all` skips Preparing/Waiting tasks (`:3001-3015`); `IDS_FILEOPS_CHECK_AUTODISMISS_ON`/`_OFF` duplicate identical text.
15. **[P2] Cross-provider & SMB cliffs**:
    - The Local route advertises Recycle for SMB/UNC (`FileSystem.cpp:7102`). Where the location cannot recycle, that sends plain Delete through a Shell refusal before the escalation question.
    - Move FROM any non-Local provider (S3, Google Drive, OneDrive, FTP/SFTP/SCP, MTP) to Local always downgrades to Copy-only (while Local -> S3 is Managed Move!), S3 advertises `rename: false`, OneDrive advertises `delete: false`, MTP advertises `delete: false, rename: false`.
16. **[P2] End-user documentation drift (`docs/FileOperations.md`)**:
    - Lines `48-50` document a removed queued-card Destination selector (`UI_FileOperationsPopup.md:947` removed it).
    - Lines `53` vs `241` expose the mismatched bandwidth presets.
    - Lines `322, 328` describe `Ctrl+V` as copy-only, omitting `Ctrl+X` Move.
    - Lines `141-144` document the redundant `Skip all` + `Apply to all` overlap.

---

### 2. Baseline reference inventories

#### 2.1 Mutation ingress

| Ingress | Code Entry Point | Confirmation Surface | Options Exposed | Async / Interlock | Delete/Transfer Mode |
|---|---|---|---|---|---|
| `F5` Copy | `FolderWindow.FileOperations.cpp:4775-4918` | None | None | Yes | `PerItem` |
| `Shift+F5` Copy | `FolderWindow.FileOperations.cpp:4775-4918` | Sync modal (`HostServices.cpp:2070+`) | Links, Verify, Start, Bandwidth | Yes | `PerItem` |
| `F6` Move | `FolderWindow.FileOperations.cpp:4919-5065` | None (except Copy-only -> sync modal) | Only if Copy-only | Yes | `PerItem` |
| `Shift+F6` Move | `FolderWindow.FileOperations.cpp:4919-5065` | Sync modal (no rename field!) | Verify, Start, Bandwidth | Yes | `PerItem` |
| `Del` (pane shortcut) | `FolderWindow.FileOperations.cpp:4590-4670` | None (escalates on card if needed) | None | Yes (missing `EnsureFileOperations()`) | `BulkItems` |
| `Del` (context menu) | `FolderWindow.FileOperations.cpp:3713` | None | None | Yes | **`PerItem`** (regression) |
| `Del` (Find results) | `FolderWindow.FileOperations.cpp:3963` | None | None | Yes | **`PerItem`** (regression) |
| `Del` (ViewerSpace Recycle) | `FolderWindow.FileOperations.cpp:3883-3891` | **BUG: Sync modal saying "Permanently delete"** | None | Yes | **`PerItem`** |
| `Shift+Del` | `FolderWindow.FileOperations.cpp:4671-4770` | Async Task Card (after Preparing) | None | Yes | `PerItem` |
| `Ctrl+C` / `Ctrl+V` | `FolderView.FileOps.cpp:763` / Paste | None | None | Yes (`Copy` misses local guard!) | `PerItem` |
| `Ctrl+X` / `Ctrl+V` | `FolderView.FileOps.cpp:785` / Paste | None (except Copy-only) | None | Yes | `PerItem` |
| Drag & Drop | `FolderView.DragDrop.cpp:455-765` | None (except Copy-only) | None | Yes (same-folder Copy bug) | `PerItem` |
| Context `"Move..."` | `FolderView.FileOps.cpp:1295-1379` | Win32 Folder Picker | None | Yes (**inverts non-local!**) | `PerItem` |
| `F2` Inline Rename | `FolderWindow.FileOperations.cpp:2849` | Inline edit | None | Yes | Rename (fails on identity-less SMB & all non-binding providers!) |
| Batch Rename / Change Case | `FolderWindow.FileOperations.cpp:3350`, `FileSystem.cpp:4016` | Preview window / options dialog | Command-specific | Yes | Rename (same) |
| `F7` Create Directory | `FolderWindow.FileSystem.Commands.cpp:10596-10760` | Modal name prompt | None | **NO (Sync UI thread!)** | Sync |
| Find Copy / Move | `FolderWindow.FileOperations.cpp:3960+` | None | None | Yes | `PerItem` |
| Compare Sync Copy / Move | `FolderWindow.FileOperations.cpp:4070` | Copy = None; Move = Sync modal | Move only | Yes | `PerItem` |
| Pack / Unpack | `FolderWindow.FileSystem.Commands.cpp:12548-12850` | Own modal + pre-op Delete modal | Format, mask, Skip/Replace, Delete-after | **NO (Sync UI thread!)** | Sync + async Delete |

#### 2.2 Conflicts and deferred consent

| `ConflictClass` | Primary Buttons | `More...` Overflow (if `>=2`) | Default / Escape | Checkbox? | Key Defect(s) |
|---|---|---|---|---|---|
| `RegularFileExists` | `[Keep both] [Replace] [Skip item] ([Skip all] if 1 overflow) [Cancel]` | `>=2` of `Delete source`, `Skip all`, `Skip all with same size & date`, `Replace all older` | Cancel / Cancel | YES | Redundant `Skip all`; hides `Delete source`; shown on 1-file & F2 Rename |
| `ReadOnlyRegularFileExists` | `[Keep both] [Replace read-only] [Skip item] ([Skip all]) [Cancel]` | Managed Move: `[Delete source, Skip all]` | Cancel / Cancel | YES | Text says only *"Item is read-only."* (omits exists); redundant `Skip all`; 1-file leak |
| `ReadOnlyRegularFile` (Delete) | `[Permanently delete read-only] [Skip item] [Cancel]` | None | Cancel / Cancel | YES | Separate permanent-delete and read-only grants. See [Start and delete behavior](#start-and-delete-behavior) for scoped caching. |
| `ReadOnlyRegularFile` (Transfer/Rename) | `[Replace read-only] [Skip item] [Skip all] [Cancel]` | None | Cancel / Cancel | YES | Redundant `Skip all` + checkbox; 1-file leak |
| `TypeMismatch` | `[Keep both (if eligible)] [Skip item] [Skip all] [Cancel]` | None | Cancel / Cancel | YES | Redundant `Skip all`; dead checkbox when `Keep both` unavailable |
| `DestinationLink` | `[Keep both] [Replace link (if bound)] [Skip item] [Skip all] [Cancel]` | None | Cancel / Cancel | YES | Redundant `Skip all` + checkbox |
| `NameNotRepresentable` | `[Keep both (if eligible)] [Skip item] [Skip all] [Cancel]` | None | Cancel / Cancel | YES | Redundant `Skip all`; dead checkbox without `Keep both` |
| `TargetConflict` | `[Skip item] [Cancel]` | None | Cancel / Cancel | NO | No bulk skip for broken symlinks in a tree |
| `SharingViolation` | `[Retry] [Skip item] [Skip all] [Cancel]` | None | Cancel / Cancel | YES | **Zero-effect checkbox** (`Skip item`+box === `Skip all`) |
| `AccessDenied` | `[Retry] [Skip item] [Skip all] [Cancel]` | None | Cancel / Cancel | YES | **Zero-effect checkbox** |
| `DiskFull` | `[Retry] [Skip all (if qualified)] [Cancel]` | None | Cancel / Cancel | NO | Clean reference pattern (no checkbox, no redundant `Skip item`) |
| `PathTooLong` | `[Retry] [Skip item] [Skip all] [Cancel]` | None | Cancel / Cancel | YES | **Useless `Retry`** + **zero-effect checkbox** |
| `NetworkOffline` | `[Retry] [Skip item] [Skip all] [Cancel]` | None | Cancel / Cancel | YES | **Zero-effect checkbox** |
| `UnsupportedReparse` | `[Skip item] [Cancel]` | None | Cancel / Cancel | NO | Jargon (*"reparse point"*); no bulk skip |
| `Unknown` | `[Retry] [Skip item] [Cancel]` | None | Cancel / Cancel | NO | No bulk skip on recurring errors |
| `RecycleFailed` (Delete) | `[Permanently delete] [Skip item] ([Delete all remaining permanently]) [Cancel]` | None | Cancel / Cancel | NO | Clean (properly checks `_sourcePaths.size() > 1`) |
| `RecycleFailed` (Move `Delete source`) | `[Permanently delete] [Keep source] [Cancel]` | None | Cancel / Cancel | NO | Clean task-scoped escalation (second step after `Delete source`) |
| `RecycleFailedByName` | `[Delete by name] [Skip item] ([Delete all remaining by name]) [Cancel]` | None | Cancel / Cancel | NO | 92-word internal essay + *"Delete by name"* jargon |
| `PermanentDeleteConfirmation` | `[Permanently delete] [Cancel]` | None | Cancel / Cancel | NO | Followed by ReadOnly prompt on cases such as `.git` objects. See [Start and delete behavior](#start-and-delete-behavior). |
| `PermanentDeleteByNameConfirmation` | `[Delete by name] [Cancel]` | None | Cancel / Cancel | NO | 92-word internal essay + *"Delete by name"* jargon |
| `SameHostOverlap` | `[Queue after these tasks] ([Run at the same time]) [Don't start]` | None | Queue after / Don't start | NO | Interrupts even in Queue mode & when `CanRunConcurrently() == false` |
| `SameHostLiveOutput` | `[Queue until other task finishes] [Delete/Rename/Replace anyway] [Skip item] [Cancel]` | None | Queue until / Cancel | NO | Choosing Queue triggers live 10 Hz grey stall bug (`WaitForLiveOutputPublisher`) |
| `EfsPlaintext` | Copy: `[Copy without encryption] [Cancel]`; Move adds `[Copy and keep source]` | None | Cancel / Cancel | YES | Pre-publication: missing `Skip item`; checkbox shown on 1-file tasks |
| `SparseInflation` | `[Continue] [Cancel]` | None | Cancel / Cancel | YES | Pre-publication: jargon; missing `Skip item`; dead `RetainSource` branch |
| `PlaceholderHydration` | `[Download and continue] [Cancel]` | None | Cancel / Cancel | YES | Pre-publication: jargon (*"placeholder"*); missing `Skip item`; dead `RetainSource` branch |
| `MetadataLoss` | `[Move anyway] [Copy and keep source] [Cancel]` | None | Cancel / Cancel | YES | The question can occur after the destination commits (`State.cpp:21088`). See the target publication rules above. |
| `InsufficientSpace` / `SpaceUnknown` | `[Continue] [Cancel]` | None | Cancel / Cancel | YES | **Dead / unwired in production** (self-test only) |

#### 2.3 Throughput and progress

| State | Timer | `IsRateSamplingBlocked` | `MarkRateSamplingStateChanged` | Graph & Samples | Verified Bug(s) |
|---|---|---|---|---|---|
| **Paused** | 1000 ms | Yes | Yes (`SetPaused`) | Visible (`Paused` overlay); history preserved | Stays in footer cohort with rate `0` -> **kills aggregate ETA for running peers** (`Popup.cpp:2330`) |
| **Conflict (arbiter)** | 100 ms (loading) / 1000 ms (loaded) | Yes (`prompt.active`) | **NO** | Hidden while active | **Bug A1**: hidden/DecisionMirror wait backfills up to 180 grey samples on restore; **Bug A2**: 1 Hz gap injects ~1 s grey hatch on resume |
| **Non-arbiter Prompt / `WaitForLiveOutputPublisher`** | 100 ms | **NO** | **NO** | **Visible & live-advancing grey at 10 Hz** | **Bug B**: wipes 100% of 18 s real history into solid grey stall band |
| **Waiting** (`queuePaused`, `waitingInQueue`, `waitingForOthers`) | 1000 ms (except `waitingForOthers`: 100 ms) | Partial (`waitingForOthers` missing) | Partial (`SetWaitForOthers` missing) | Hidden | `waitingForOthers` omitted from `IsRateSamplingBlocked` & `GetPopupTelemetryWork` |
| **Preparing** | 100 ms | No (`started == false`) | No | Hidden; 0 samples | `Preparing... ({0} elapsed)` stuck at `0s` (`Popup.cpp:1834-1836`) |
| **Discovering** | 100 ms | No | No | Hidden pre-transfer; visible concurrent with marquee + provisional ETA | Missing `[Pause]` button on card while copying during discovery-ahead (`Popup.cpp:12641-12828`) |
| **Verifying** | 100 ms | No | No | Visible (`Verifying` overlay) + secondary series + 50/50 bar | Clean |
| **Stopping** | 100 ms | **NO** | No | Visible, unhatched | **Appends decaying-to-0 samples at 10 Hz**, wiping throughput history over 18 s (`UI_FileOperationsPopup.md:488`); marquee stays animated |
| **Stalled** (real I/O silence) | 100 ms | No | No | Visible + grey hatched `FileOperationsGraphStallBand` | Intended design (`I35 R10.2`) |

#### 2.4 Provider capabilities and workflow limits

| Provider | Object Binding (`IFileSystemObjectBinding`) | Copy / Move / Rename / Delete / Recycle | Cross-FS Move (`export/import`) | Verification | Notable Workflow Cliff |
|---|---|---|---|---|---|
| **Local NTFS/ReFS** | Yes (`FileSystem.h:120`) | All `true` | `["*"]` / `["*"]` | `hostReadback` + `blake3` | The default process budget caps single-job NVMe auto-concurrency from eight to four. See [Queue, progress, and results](#queue-progress-and-results). |
| **SMB / UNC** | Yes, but server may return `ERROR_NOT_SUPPORTED` for `FileIdInfo` | Advertises all `true` including Recycle | `["*"]` / `["*"]` | Bound-only | Rename fails closed on identity-less NAS. The Recycle route may fall through to escalation for locations without a bin. The global budget caps the effective Auto budget. |
| **FTP / SSH (SFTP, SCP)** (`FileSystemCurl`) | **No (`E_NOINTERFACE`)** | `copy/move/rename/delete: true`, `recycle: false` | `[]` / `[]` | `none` | **Inline F2, Batch Rename, and Change Case all fail (`0x80070032`)** even though `FileSystemCurl::RenameItem` is implemented; Cross-FS Move Copy-only; Verify silently dropped on F5; always serialized |
| **Google Drive** | **No (`E_NOINTERFACE`)** | All `true` | `[]` / `[]` | `writer-digest` | **Inline F2, Batch Rename, and Change Case all fail (`0x80070032`)** despite `RenameItem`; Cross-FS Move Copy-only; cross-root Copy fails (`I35` XP-09) |
| **Microsoft Drive (OneDrive)** | **No (`E_NOINTERFACE`)** | `copy: false`, **`delete: false`**, `rename/move/recycle: true` | `[]` / `[]` | `writer-digest` | **Inline F2, Batch Rename, and Change Case all fail (`0x80070032`)** despite `RenameItem`; same-endpoint Copy (`H3`) and Permanent Delete unsupported; Cross-FS Move Copy-only |
| **S3** | **No (`E_NOINTERFACE`)** | `rename: false`, `recycle: false` | `["*"]` / `["*"]` | `writer-digest` | F2 Rename unsupported; **asymmetric Move** (Local->S3 deletes source; S3->Local Copy-only!) |
| **MTP** | **No (`E_NOINTERFACE`)** | `nativeMove/delete/rename/recycle: false` | `[]` / `[]` | `none` | Cannot delete or rename; Move always Copy-only (`H3`); concurrency 1 |
| **7z / IMAP** | **No (`E_NOINTERFACE`)** | Read-only mount | Export copy only | `none` | Mounted 7z is read-only; Pack/Unpack are separate sync local-only commands |
