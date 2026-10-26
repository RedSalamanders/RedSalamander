# I48 MSVC 14.52 and Release optimization

Status: ACTIVE implementation and qualification
Date: 2026-10-23
Repository: RedSalamander
Base: aad3cff1211f55d01f7cec4efb47fa1f7ccb19c0

## Checklist

- [x] Isolate this work on ej/msvc1452-release-20261023; preserve the original checkout.
- [x] Retain the original x64 Release build, image sizes/hashes and frozen runtime closure under .build/msvc1452.
- [x] Audit all six profiles before changing policy. Both sides select installed/default MSVC 14.52.37025; there is no retained 14.51 installation, so no compiler-version speedup is claimed.
- [x] Require MSVC 14.52.x and enable Release /Gy /GL /OPT:REF and full /LTCG with non-incremental linking.
- [x] Forward the exact compiler and WPO policy to the unchanged adopted DxUi API pin; give the changed archive policy a separate fingerprint/output root.
- [x] Verify evaluated native profiles, transitive archive settings and broken-policy controls; update build/consumer specs and documentation.
- [ ] Run affected tooling, all six builds and x64 native tests in Debug/Release/ASan Debug. Native ARM64 and interactive acceptance remain distinct.
- [ ] Compare final image/archive sizes, build duration and unchanged-fixture runtime/resource samples. Retain original failures, interrupted/noisy attempts and activity observations.
- [ ] Publish local reviewed evidence and settings differences, fold the changelog and validate. Move to Done only after required local qualification is complete.

## Measurement and scope

The DxUi plan of the same name coordinates the five repositories. This repository owns its application measurements and its existing dependency API adoption. The original and candidate binaries must use identical fixtures and the same exact toolset; the Release compiler policy is the intended experimental difference. An initial sample collected alongside other builds is diagnostic, not a quiet-host/non-regression pass. Retain frozen baseline images and every attempt; do not silently rebaseline.

Resource-only projects have no C++ bodies. Independently owned pinned repositories and DLL boundaries remain separate build/optimization units. Existing feature/performance holds are not closed by this compiler-policy change.

Generated raw evidence: .build/msvc1452/audit-before, baseline-build.json, baseline-bin and initial native benchmark reports. Review and retain selected local evidence before closeout.
The initial focused baseline passed both large-folder icon enumeration tests through VSTest. Its native sandbox/results use the already initialized E:\RedSalamander.Perf root; the original TRX and stdout remain bound by .build/msvc1452/baseline-performance.json. RedMonitor is an independently pinned owner behind compatibility projects; this plan does not silently change that external repository.

## Qualification workspace and current gates

The user authorized E:\RedSalamander.Perf on 2026-10-23 for worktrees/builds/evidence. Final qualification lives in E:\RedSalamander.Perf\msvc1452 on ej/msvc1452-qualification-20261023; original Z: baselines remain retained and hash-verified frozen copies exist under E:\RedSalamander.Perf\msvc1452-frozen. The original optimizer-only phase retained all four consumer API pins. This repair phase explicitly adopts the local API-4 backport recorded below. Source edits during initial builds were correctly rejected by attestation; retain those attempts and rerun against frozen inputs.

The user approved interactive tests when ready. The user also approved DxUi's visible 30-second automatic-start countdown; desktop restoration checks stay enabled. Native ARM64 execution is unavailable on this x64 host. Hosted CI provisions and verifies the required 14.52 family through .github/actions/setup-msvc; the download/install branch remains pending a real GitHub-hosted execution. Local -VerifyOnly and refusal controls must pass without installing software. All required gates stay open until their receipts are reviewed.

- [ ] Qualify the narrow MSVC 19.52 ARM64 AWS hash /Ob1 overlay through canonical dependency/product builds. Preserve all C1001 attempts and do not infer native ARM64 performance from cross-compilation.

## Reviewed local qualification checkpoint

The first Debug Fresh Full run is failed: Compare 253 passed / 1 failed / 21 skipped; Commands 939 / 14 / 6; FileOps 5 / 2 / 315 unreached after admission/setup failure. Ordinary tooling reports 1004 passed / 1 failed / 1 skipped, including a stale local RedMonitor source-path check. ViewerSpace export boundary assertions assume the executable is outside the marked sandbox, which is false in this approved E: worktree. An owned junction fixture also prevents final archival; the failed run is retained in place. These results do not close qualification or establish a Release-LTO regression.

The original image-size pair is test-enabled. The additional production-only pair is 8,099,840 to 8,401,920 bytes (+3.73%), with RSBuildEnableTests=false and identical source path/compiler/version on both sides. Final runtime/resource results and remaining profiles are kept in the external report linked by Measurements/Toolchain/2026-10-23-msvc1452/README.md so receipt-bound source remains frozen during testing. Native ARM64 runtime and actual hosted CI provisioning remain open.

## Requested failure remediation

The user requested all five repositories become green after the initial toolchain report. Original failed receipts and frozen compiler-policy comparisons remain preserved. The current phase repairs confirmed product, fixture and source-contract defects; it must not label old native-source-identical measurements as measurements of later native repairs.

- [ ] Reproduce and classify every retained RedSalamander failure, including broad-order contamination.
- [ ] Repair causes and add behavioral/negative controls with current domain contracts and documentation.
- [ ] Pass focused and predecessor-cluster checks before canonical Fresh Full in Debug, Release and ASan Debug.
- [ ] Refresh ARM64 cross-builds after native repairs and report native-runtime/hosted-runner availability separately.
- [ ] Reconcile the five-repository report and current source/artifact evidence; preserve every original failure.


Remediation iteration 1: full-solution Release rebuild passed; focused reruns confirmed stale Copy discovery expectations, wrapped Markdown source-guard matching, a successful UIA selection mutation returning UIA_E_NOTSUPPORTED, and inherited FileOps state. The next candidate restores live default shortcuts after Preferences Apply and strengthens fixture isolation and diagnostics. Local x64 tests plus ARM64 cross-builds are the agreed completion scope; native ARM64 execution and hosted CI remain separate.

Iteration 2: the source/spec guard passed. The Issues row retained model generation 3 and presentation generation 1 across its successful selection mutation; a standalone DxUi regression reproduced UIA_E_NOTSUPPORTED when its delegate relayouts a sibling. FileOps broad-order failure reported exactly five completed entries with no active work or popup; completion-recovery's five scenarios omitted dismissal. Iteration 3 adds exact owned-result cleanup and correct sandbox-boundary fixture construction. The earlier speculative Issues-producer popup cleanup was removed because the trace disproved that cause.

Iteration 3: all 11 popup cases and the live Mouse/shortcut Apply case passed. Move breadcrumb's native assertions passed; canonical archival rejected its leftover junction, so that run remains failed and retained. Completion-recovery cleanup removed the five inherited results: Review reached its next predicate, where pending clipboard admission failed; an HRESULT diagnostic is now added. Original Release F7 trace proves foreign PID 73160 replaced foreign PID 84624 while the test PID was 48124; F7 now reuses the existing process-owned observer while retaining all storage and timing limits.

Local API-4 DxUi backport candidate: 06205d6242be408bbb7bb624fb61174a89a97854, parent f65e4a4b9311370b3e3070088e2a8abe29adcb27. The canonical standalone Debug regression passes both sibling relayout and same-pointer rebinding controls. The backport folds its base's historical fragments and changes only Grid/provider native behavior plus the regression and docs. It is unpublished; local qualification uses a process-local Git URL rewrite under Build_Toolchain.md. The normal main/CI update command and publication requirement are unchanged. Fresh exact-pin consumer validation is pending.

Iteration 4: Release file-operation Commands passed 65/66; Issues UIA removal now passes on the exact local DxUi backport. Batch Rename and F7 admission passed. Preferences passed 175/176; its footer focus predicate still fails, while the backdrop case passed this repeat without establishing the original cause. Broad FileOps admission fails with ERROR_ALREADY_EXISTS because two phases reused one synthetic clipboard sequence; iteration 5 allocates fresh per-invocation identities and adds a same-sequence replay refusal control.

Standalone native PR entries and all 1005 Pester assertions passed, but durable qualification rejected one skipped vcpkg database check. The exact pinned checkout was missing from this isolated worktree; it has been restored without editing the source checkout, and the real floor/override check passes directly. Do not count the earlier PR receipt as green. Compare native assertions passed, but archival found five owned links left behind; all five now have no-follow RAII cleanup and exact target-byte checks. The failed archive remains retained.

Iteration 5 also corrects 30 FAILED(bool try_query_to(...)) guards, with five independent classifier controls and zero remaining invalid direct guards; the rebuild must confirm C4878 is gone. Find's selection fixture now establishes a selected current row through Grid requests. Pending disclosure and footer failures now report their actual geometry, animation and focus predicates. A deterministic late-theme notification regression is added before its product repair; its expected failing receipt will establish the preview race before correction. No timeout, original failure, archive refusal, or strict idle predicate is relaxed.

Closed Compare perf traces are 27–35 GB each. Storage-only NTFS compression is being verified against exact hashes/timestamps, without deleting raw evidence or creating reparse points. Compression does not convert an old failed receipt into a pass. Original compiler-policy measurements remain distinct from this native repair phase.

## Candidate 5b (2026-10-24)

Candidate5a's queued-admission repeat and predecessor cluster passed. All five Compare link fixtures passed and canonical archival succeeded. The eight-case Commands batch passed documentation-gallery ownership and ViewerSpace held-disposal close; it reproduced VLC timing, first Compare paint, single Find current-row setup, posted thumbnail command, initial scroll quiescence and the intentionally injected late-theme preview failure. Passing repeats do not establish the original close/gallery failure causes.

The next candidate preserves valid retained Preferences focus on a same-page refresh and overlays edited UI preview fields after an incoming theme notification. The single-result Find helper reuses the multi-result Grid request path. Disclosure routes wait for painted height settlement; Compare waits for its first frame; the slider waits for its posted command; scroll measurement reuses the existing warm-perf quiescence helper. Assertions and original performance limits remain intact.

VLC timing now starts before the focus command and ends at the matching publication snapshot, before trace I/O and focus-diagnostic work; it records that administrative time separately and keeps the 900 ms limit and held-retirement gate. This instrumentation must establish whether any remaining delay occurs inside the product transition. It is not a claim that the VLC defect is already fixed.

Cadence's declared known-byte fixture now publishes the required exact-byte knowledge. A separate unknown-byte paused fixture retains the observed repaint regression: indeterminate progress still animates while paused. Its resource ceiling is unchanged. Generic animation suspension in canonical DxUi and explicit local consumer adoption are pending, with the indeterminate accessibility state preserved. Final Fresh Full x64 and ARM64 cross-build qualification remain open.

The agreed completion scope is local x64 tests and ARM64 cross-builds. Native ARM64 execution and actual hosted provisioning remain separate handoff gates and do not block this task's local closeout. Historical failed receipts and raw traces remain preserved; regular NTFS compression and byte-verified hardlink sharing change storage allocation only.
## Candidate 6: progress motion and VLC stage attribution

Candidate5b passed eight selected Commands cases and the reparse Compare cluster; VLC same-plugin publication remained 1,241,781 us against the unchanged 900 ms limit. The mixed comma-separated cross-suite list was not recognized by the native FileOps selector, which takes one phase, family or prefix; that failed receipt stays retained. Candidate6 uses the separate FileOps_PopupIdleCadence selector, explicitly adopts unpublished API-4 pin 0def3d03edae58b099b7d1918d6de114ae8a2609, and pauses hosted motion without changing enabled/indeterminate state. It adds a direct semantics assertion and native VLC UI-stage metrics; it does not yet claim a causal VLC repair. Final Full validation remains open.

## Final local qualification candidate

Candidate6 passed all nine selected Commands cases and the reparse Compare cluster. The separately selected FileOps cadence family passed three phases, including actual enabled/indeterminate paused semantics: unknown-paused renders fell from the retained 48 in 1.2 seconds to one with one timer tick. Foreground movement by another authorized test was reported as environmental, not attributed to the non-activating fixture. VLC repeated at 176,855 us with set_media at 123,285 us; this diagnostic-only repeat does not establish the earlier 1,241,781 us failure's cause. The original 900 ms limit stays unchanged. One remaining D4 fixture converted Boolean try_query_to to HRESULT; it now handles Boolean availability and reports E_NOINTERFACE on failure. A full source search found exactly this remaining assignment. Final six builds, repeatability coverage and Fresh Full x64 qualification follow with frozen inputs.

## Confirmed VLC media-switch repair

The five-repeat final-source probe reproduced one failure: publication 1,157,277 us, with 1,104,064 us specifically inside libvlc_media_player_set_media on the owner UI thread. Four passing repetitions do not erase that result. Full qualification is held; all six pre-repair builds and the shipping size remain historical candidate receipts.

Transfer sole VlcState ownership to the existing module-pinned load worker for file/config validation, media construction and player media replacement. The UI keeps its HWND/drawable alive through the pending-work count and never queries a worker-owned player. Only the current request/window result may attach or play. Compatible navigation preserves the instance/player. While that state is in flight, admit one latest pending path/configuration and cancel obsolete preparation; never enqueue an additional worker per selection. Stop, close, forced HWND destruction and post/submission failures must cancel/drain ownership without synchronous stop/release or stale playback. Add a held-worker gate, player identities and queue counters to prove responsiveness, coalescing and reuse independently of libVLC timing. Keep the 900 ms publication bound; decoder completion is a separate predicate.

The first Release held-worker pilot passed all five repetitions with the original 900 ms publication deadline and actual instance/player reuse. Real-player negative controls now cover media submit/post failure, failed-post latest-request recovery, keyboard Stop and retained-HWND Close. All final builds/tests remain pending for this expanded candidate; earlier receipts stay historical.

Source review found that the delayed loading overlay still hid the real Stop HUD, despite its logical availability flag. Preserve the completed pre-visibility matrix, then keep the HUD shown and above that panel during worker-owned replacement. The real-player controls now force the delayed overlay and assert actual visibility, height and sibling order. New final builds and Fresh Full remain required.

Final-source probes passed ViewerSpace Close at its original 100 ms limit and two earlier search-service cases. The two-root startup lifecycle case still took 16,018 ms against 15 seconds: its tiny roots selected an entire-volume MFT seed. A same-binary existing traversal-control diagnostic passed in 913 ms. The fixture now explicitly isolates synthetic-root lifecycle via that case-scoped seed control, restores the prior environment and adds a persisted currentness-unproven safety assertion. Production seeding and dedicated native MFT/journal tests remain unchanged; no timeout or skip is added. Ordinary focused and Fresh Full qualification follow.

## P3 broad outcome and P4 focused candidate

The P3 six-profile build matrix passed. Debug Fresh Full finished failed on 2026-10-24 at 07:22:33 UTC: tooling 1006/0/0, Compare 254/0/21, Commands 951/2/6, FileOps 278/4/40. FileOps includes 12 unreached cases and 28 actual capability skips. ViewerPE interactive also failed its real Explorer inspector splitter checks. Three Commands junctions prevented final archival; the failed raw run remains retained. Release and ASan Full have not run for this candidate.

P4 corrects the VLC worker source audit, adds three Commands junction RAII guards with exact target checks, and aligns four FileOps fixtures with existing typed policy and accounting. Gallery foreground and ViewerSpace splitter diagnostics preserve the original predicates. These changes are not qualified by the earlier builds. Next: actual Debug focused checks, cause repairs where diagnostics still fail, all six final builds, Fresh Full in all three local x64 configurations, then repair-impact measurements and report closeout. No compiler performance claim follows from fixture repairs or a passing retry.

P4 Debug built successfully. All five VLC source-audit repetitions, six junction-case observations and their archive, the Copy label case, and exact deletion-accounting case passed. The provider matrix passed all eight cleanup decisions, then reached an outdated 7z separator expectation. P5 aligns that one expectation with the existing dual-separator contract; it adds detailed gallery observation, splitter pointer/coordinate and deferred-consent field diagnostics. Gallery still failed three of five broad-route repetitions; Explorer inspector capture failed in all six churn cycles. The consent matrix remains failed. Diagnostics are not behavior fixes or qualification.

The reached R4-A19 independent-volume control was blocked because alternate-volume execution was not enabled. The existing initialized, marked C:\RedSalamander.Perf sandbox is available; the host resource manifest now declares it explicitly. Subsequent focused and Full runs must opt in with -AllowAlternateVolumeTestRoot, retain marker/volume checks and prove the real case. No new drive-root directory was initialized. Builds and main evidence remain on E:. Ten Phase10 cases remain unreached behind the consent failure. Final six builds, three Fresh Full profiles, measurements and closeout remain pending.

P5 Debug built and native diagnostics passed the provider matrix, all five gallery repetitions, and the exact real independent-volume case with the declared C: root. Gallery's earlier cause remains unexplained. Consent fails specifically at EFS source retention because the synthetic task defaults to Copy; P6 makes its intended Move context explicit and adds an actual Copy refusal/recovery control. A same-binary fake-MTP control fails with the caller profile and passes with a private profile; its fixture now owns and restores that lease across asynchronous phases.

Inspector churn still fails in all six cycles: logical/physical pointer coordinates agree, but a Windows Ghost surface intercepts the click. The parent blocks in WaitForSingleObject while owning a warning window. P6 first adds a hidden parent-timer regression without changing that wait, and an artifact-rotation regression against the unchanged cleanup algorithm. Their expected red receipts must precede repairs; final cross-builds and Fresh Full remain pending. Historical failed runs and their raw traces are untouched.

P6 Debug native receipts reproduce the hidden-timer failure and rotation accepting locked-file cleanup failure against the old algorithms. The read-only/junction positive control already passed; it does not prove the cause of earlier retained output. Consent and the full R4-A19 baseline, including private-profile MTP and the actual alternate-volume case, pass. Phase10 now reaches a stale verification presentation fixture: the synthetic unfinished Copy omits its known-byte flag. The candidate sets that precondition and adds an unknown-byte negative control, pumps the bounded parent queue during the original child timeout, and makes rotation fail closed through the existing no-follow cleanup helper. Green probes, all six final builds and three Fresh Full runs remain required.
