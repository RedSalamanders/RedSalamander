# I56 Hosted acceptance of unified CI

Status: ACTIVE - implementation is delivered; hosted qualification, account capabilities and publication readiness remain explicit.
Priority: P1.
Owner: organization maintainer for account/visibility/secrets; project owners for native failures; CI maintainer for receipts, workflow integration and this checklist.
Prepared: 2026-10-31 from main 738745a8b986d6fdb021f20dd553f45064b71932 in ej/ci-artifact-closeout-20261031.
Scope: RedSalamander, DxUi, RedXe, RedMonitor, RedPrism and RedTilde.

## Authority and separation

The owner explicitly selected "Close implementation and track hosted acceptance separately" on 2026-10-31.
[I55](../Done/CI_OrganizationStandardization_2026-10-29.md) therefore records implementation closeout, not all-six qualification.
This is the single successor for the remaining hosted CI acceptance. Product/runtime findings stay with I50/I51
and their existing owners; no duplicate native remediation queue or consumer-pin adoption is created here.

The same owner removed all hosted publication workflows from DxUi. Both SDK ZIPs remain internal qualification
artifacts, including the native extracted-consumer checks. DxUi has no Prepare, public-release, winget or gallery
publication action. The five applications retain explicit artifact-only publication.

Engine deployment is phased under the owner's public-project continuation: the four public projects target v1.0.7
source e32cf9f9b29f4e817decb09e1be44be7fc1a94de; the two private projects retain v1.0.6 source
19710a8aeac979a82e83bec82e8b866cee3d3deb unchanged. Each project's manifest and reusable callers must agree.
A receipt from an older engine identity never qualifies its changed deployment. Actual rollout and evidence are recorded below.

Normative owners: [hosted CI](../../Build/Build_UnifiedCI.md), [build identity](../../Build/Build_Toolchain.md),
[test evidence](../../Testing/Testing_ValidationEvidence.md), [self-tests](../../Testing/Testing_SelfTests.md),
[portable ZIPs](../../Installer/Installer_PortableZip.md), [winget](../../Installer/WingetIntegration.md),
and the [shared engine contract](https://github.com/RedSalamanders/.github/blob/main/Specs/CI.md).

## Checklist

- [ ] Resolve RedSalamander's actual regression and evidence-upload failures with the existing native/test owners; retain original failing cases and receipts.
- [x] Resolve RedXe's executable/catalogue source-role mismatch without changing pins or lowering dependency/native acceptance requirements. Both hosted Release test-profile restore/build steps passed in 41207884174; complete native regression/production acceptance remains a separate item.
- [ ] Owner resolves the RedPrism/RedTilde billing/visibility prerequisites; confirm real compiler-free CI, CodeQL and dependency review execute.
- [ ] Apply and verify the common required checks, organization-owner PR-only emergency route, supported dependency/security settings and absence of duplicate managed native CodeQL in those two repositories.
- [ ] Complete original successful main qualification against each approved deployed engine on native x64 and ARM64 for all six; verify actual regression counts, active failure probes, global barrier, production roles and ZIP/SDK smoke.
- [ ] Independently verify matching retained descriptors, receipts, SHA-256 ZIP bytes and actual job ordering. Admission-only green runs never satisfy this item.
- [ ] Execute artifact-only Prepare against each of the five applications' current successful source/engine qualifications; preserve refusal tests and keep all public writes skipped.
- [x] Verify repository credential metadata and execute the authorized promotion preflight: existing token authentication passed; organization-secret administration was refused before any secret changed.
- [x] Owner creates organization WINGET_TOKEN; verify the updated public-repositories-except-DxUi access policy and remove the two repository overrides after explicit authorization. Private application access remains part of their pending public transition.
- [x] Execute the trusted-main read-only credential check against the effective organization token; retain authentication/scope evidence separately from actual publication/native-install readiness. Run 41200876308 passed.
- [ ] Re-run the central compiler-free dependency observer to readable six-of-six coverage and current immutable engine pins.
- [ ] Reconcile this checklist and move this acceptance record to Done only after its actual applicable evidence exists.

## Last observed hosted state before v1.0.6 integration

| Project | Source/run | Observed evidence and remaining work |
|---|---|---|
| RedSalamander | 738745a8b986d6fdb021f20dd553f45064b71932; [41147854938](https://github.com/RedSalamanders/RedSalamander/actions/runs/41147854938) | Both native builds reached regression. Regression steps failed; Git fixture long-path errors and an upload-artifact root-directory mismatch were observed. The global barrier and production were skipped. Diagnose complete failed-case evidence; these observations are not a claim that all failures share one cause. |
| RedXe | 98e6f8c94c1957f7822b28065ca27ec6d1b2d6ca; [41147447636](https://github.com/RedSalamanders/RedXe/actions/runs/41147447636) | Both test-role builds failed restoration: the pinned database has no curl 8.23.0 or yyjson 0.13.0 entry. Earlier successful 41136579577 does not qualify these later inputs. |
| DxUi | 39f3400893dd1fb71278fb075d402f2e3fb95627; [41147350129](https://github.com/RedSalamanders/DxUi/actions/runs/41147350129) | Both native regression/probe, global barrier, production SDK and extracted-consumer stages passed with v1.0.5. Repeat for the deployed v1.0.6 identity; no publication check applies. |
| RedMonitor | c1a46d1fab31e98ceb8386781f45e87ee1d5bfaf; [41147644322](https://github.com/RedSalamanders/RedMonitor/actions/runs/41147644322) | Both native regression/probe, barrier, production and ZIP smoke passed with v1.0.5. Current artifact-only Prepare 41196277999 passed and its downloaded bytes were independently verified. Repeat for v1.0.6. |
| RedPrism | main 045581b4ca636c7d1f2515c834eab214a38c2eea; PR 23 merged | Private. Main CI 41147305566 failed before any job step because of billing/spending limits. No hosted native qualification exists; the retained local x64 proof is narrower. |
| RedTilde | main ac87b7ae55f349febb83a7c28dc11004f6ed87a7; PRs 33 and 35 merged | Private. Main CI 41195289886 failed before any job step because of billing/spending limits. No hosted native qualification exists; the retained local x64 proof is narrower. PR 33 used the owner's recorded merge-check exception; this does not waive native/publication gates. |

## Artifact-only preparation and credentials

The 2026-10-31 exercises are complete as positive and negative implementation controls:

- [RedMonitor Prepare 41196277999](https://github.com/RedSalamanders/RedMonitor/actions/runs/41196277999) succeeded against original qualification 41147644322, version 0.1.60. Both ZIPs and complete receipts were downloaded and revalidated; release, installation and submission were skipped.
- [RedXe Prepare 41196280434](https://github.com/RedSalamanders/RedXe/actions/runs/41196280434) correctly refused failed qualification 41147447636 before any public write.
- [RedSalamander Prepare 41196282912](https://github.com/RedSalamanders/RedSalamander/actions/runs/41196282912) correctly refused failed qualification 41147854938 before any public write.
- RedPrism/RedTilde have no successful hosted qualification to select. Billing-blocked or invented run IDs cannot establish Prepare acceptance.
- DxUi is deliberately excluded from publication and Prepare.

Evidence is retained under Z:/RedSalamander.Perf/evidence/unified-ci/closeout-20261031. The verified RedMonitor
x64/ARM64 ZIP SHA-256 values are 7c72c3c5894972c21f20f88af5c5cceda6f46509820a7674bc87d096dc201350 and
1ce2cdb33c79e25000eb051b8b4d534f4dc65c732c09adafe78905f0e7d0ac98 respectively.

Repository secret metadata initially showed WINGET_TOKEN in RedSalamander and RedXe only. The current local gh
credential cannot inspect organization secrets (HTTP 403, missing admin:org scope). The owner explicitly requested
promoting the existing RedSalamander token to an organization secret. The trusted-main
promote-publication-credential.yml action checks that token's organization-secret permission, encrypts it directly
through gh stdin, and verifies selected access for RedSalamander, RedXe, RedMonitor, RedPrism and RedTilde only.
No plaintext secret value is exported or logged; existing repository overrides remain intact. Record its actual
execution result below. A missing permission stays an operator prerequisite; never grant broader scopes automatically.

## Execution, verification and Done criteria

Refresh each remote main, inspect ownership/dirty worktrees, and compare its full Git input identity with retained
ci-state before execution. Automatic unchanged failed inputs require the explicit retry action; never start a second
concurrent native batch or cancel active heavy work. Keep five quiet minutes, compiler-free PR/timer routes and
exact full-SHA engine pins. Use the real canonical project runners and active failure probes; a cross-build is not
native execution. Preserve test/production separation and the global AND barrier.

Debug/ASan, paired performance, interactive desktop, provider/device and product-specific runtime holds remain with
their existing owners. They are not closed by a Release batch or by this implementation split. No new sandbox,
foreground takeover, product visibility, paid-account change, dependency adoption or public release is authorized here.

Done requires complete applicable source/engine-bound hosted evidence across all six, current five-application
Prepare/credential readiness, six-of-six metadata coverage, real branch/security capability verification and updated
records. Actual public release and upstream winget submission remain separately selected manual actions; do not
publish merely to close this plan. Stop the affected qualification on invalid identity, failed/missing evidence,
unsupported hosting, ownership drift or an unreviewed dependency change; continue independent work.

## Implementation integration and credential execution - 2026-10-31

I55 is now immutable implementation history on main through [PR 180](https://github.com/RedSalamanders/RedSalamander/pull/180), merge a9c02b78847131e3dfdd1a2b08b0479d6b5ff526. Its required compiler-free CI, CodeQL and dependency-review checks passed. All six project adapters now use reviewed engine v1.0.6 source 19710a8aeac979a82e83bec82e8b866cee3d3deb. The native v1.0.6 acceptance items above remain open; no source- or engine-mismatched old receipt was promoted.

DxUi [PR 152](https://github.com/RedSalamanders/DxUi/pull/152), merge f2940cc265db79458b6615492779b839c5c1b94a, removed both SDK and gallery publication workflows; canonical fold PR 153 merged as be99b682ff67b847417b3dd27b56294a72e20d4e. GitHub's active workflow inventory confirms only the six project workflows (CI, formatting, main qualification, extended qualification, manual benchmark and compiler-free security), plus system-managed entries. No publisher remains. Local metadata validation now checks the existing independent executable and ports revisions without changing either pin; its 32 behavior controls and 26 documentation controls passed.

The other v1.0.6 implementation PRs and their canonical folds are merged: RedXe 69/70, RedMonitor 23/24, RedPrism 29/30 and RedTilde 36/37. Private implementation/fold PRs retain exact-head merge-check exceptions under the owner's closeout split; account-blocked checks remain failed and native/publication gates remain enforced.

[Credential promotion 41198028463](https://github.com/RedSalamanders/RedSalamander/actions/runs/41198028463) executed from trusted main a9c02b78847131e3dfdd1a2b08b0479d6b5ff526. The existing publication token authenticated successfully. GitHub then refused the read-only organization-secret permission preflight with HTTP 403: admin:org or the corresponding fine-grained Actions-secret permission is missing. Execution stopped before gh secret set, so no organization or repository secret was changed. No secret value was exported or logged. The owner explicitly chose to create the organization secret personally after this refusal. Organization-secret creation is therefore assigned to the owner, not completed or automatically retried; no broader permission is requested for the publication token. Existing repository overrides remain available as before.

The retained log is Z:/RedSalamander.Perf/evidence/unified-ci/closeout-20261031/organization-token-promotion-41198028463.log. Requested organization access remains exactly RedSalamander, RedXe, RedMonitor, RedPrism and RedTilde; DxUi is excluded. Credential authentication does not prove a public release, native winget installation, cleanup or upstream submission.

The owner-authenticated, read-only dependency inventory observed all six immutable default-branch trees with engine 19710a8aeac979a82e83bec82e8b866cee3d3deb. Its report is Z:/RedSalamander.Perf/evidence/unified-ci/closeout-20261031/owner-six-project-dependency-status.json. This confirms coordinated deployment using owner read access; it does not claim that the central workflow token can read the two still-private repositories, so the hosted six-of-six observer item remains open.

## Organization credential follow-through - 2026-10-31

The owner created the organization WINGET_TOKEN and then changed its access policy to all current public organization repositories except DxUi. The two private applications need separate work before they become public; their visibility and account capabilities remain unchanged. At 12:14 UTC, inherited-secret metadata (updated 12:13:38 UTC) confirmed access for RedSalamander, RedXe, RedMonitor and the public shared .github repository, with no access for DxUi, RedPrism or RedTilde. This matches the revised owner policy. Selection must be reconciled when a private project becomes public; it is not automatic publication readiness.

After rechecking that RedSalamander and RedXe were public and inherited the organization secret, the explicitly authorized repository WINGET_TOKEN overrides were deleted. Subsequent repository-secret listings confirmed both overrides absent. The retained metadata-only reports are organization-token-access-current.json and organization-token-overrides-removed.json under the existing closeout evidence directory. No token value was read, exported or logged. The remaining public application, RedMonitor, had no repository override.

The existing promote-publication-credential.yml path now performs a trusted-main, manual read-only authentication/scope check. The obsolete secret-writing promotion operation is removed because the owner has completed organization setup. Its actual PowerShell run block passed seven isolated controls: public_repo, repo, missing scope, missing scope header, rejected authentication, missing credential and multiline credential. The controls allowed only GET /user and checked that neither token nor account data appeared in diagnostics. actionlint accepted the workflow. These local controls are distinct from the live execution below.

The v1.0.6 RedXe qualification [41198036552](https://github.com/RedSalamanders/RedXe/actions/runs/41198036552), source 229689e298cfabf5efd5eb285e6b4a5589a8060a, reproduced the missing curl 8.23.0 and yyjson 0.13.0 version-database entries on both native platforms. Production stayed skipped. Its original failed log is RedXe-v106-41198036552-failed.log. RedSalamander 41198411091, DxUi 41198401034 and RedMonitor 41198030903 were still executing native test-profile builds when observed; no active heavy job was cancelled and none is recorded as a pass.

## Live credential verification - 2026-10-31

[RedSalamander PR 183](https://github.com/RedSalamanders/RedSalamander/pull/183), reviewed source f9a71a4939debcb5a799305ff3e3ebb42b538718, merged as e361efec1e27ce0a980334309c900c6f285161f5 after all three required compiler-free checks passed. The live manual [credential verification 41200876308](https://github.com/RedSalamanders/RedSalamander/actions/runs/41200876308) executed that exact trusted-main source. Its actual authentication/scope step passed. The job had no checkout, compiler or write operation, and no token/account data was printed or retained. This establishes authentication and declared public repository scope for the effective organization token; no fork creation, exercised write permission, native installation or upstream submission is claimed.

Post-execution metadata reconfirmed the organization secret's unchanged 2026-10-31T12:13:38Z update identity, access for the four current public repositories except DxUi, no access for the two private repositories, and no repository WINGET_TOKEN overrides in any of the seven repositories. Credential result, execution log and metadata are retained as credential-verifier-result.json, credential-verifier-execution.log and credential-post-execution-access.json under the existing closeout evidence directory. The shared contract update [engine PR 11](https://github.com/RedSalamanders/.github/pull/11) merged as ee96f1377304a343f91fb8520f758ca86ecce791; its Windows/Linux behavior and dependency checks passed. This documentation change leaves every deployed v1.0.6 pin unchanged.

RedSalamander's compiler-free static admission and all 46 local engine controls passed. Test-Changes.ps1 -Explain selected full native work conservatively for the workflow edit and retained four interactive entries as pending; no native or focus-taking tests were launched by that command. The explanation is retained as credential-affected-plan.json and credential-affected-plan.log. This does not waive the open native/runtime acceptance above. This proof-only follow-up changes prose/fragments, so it must preserve the merged verifier's native input digest and does not retry an existing admitted native identity.

## Public-project continuation - 2026-10-31

The owner explicitly requested completing I56 for the public projects only: RedSalamander, RedXe, RedMonitor and
DxUi. RedPrism and RedTilde remain private and unchanged, including their deployed v1.0.6 pins. Their account,
security, native and publication items above remain pending; this continuation does not close the whole six-project plan.

The public projects are moving together to reviewed engine v1.0.7 source e32cf9f9b29f4e817decb09e1be44be7fc1a94de,
[engine PR 12](https://github.com/RedSalamanders/.github/pull/12), merge 904e5daf2dc29b9bc824a9cd9b623d71656c7a38.
It separates sandbox evidence from workspace diagnostic uploads in automatic and manual qualification. The original
failure combined both roots, calculated a Windows drive-relative common ancestor, and rejected external sandbox files.
The repair preserves the sandbox and every native/publication gate. All 47 local engine controls, Windows/Linux hosted
behavior controls, dependency review and eight workflow syntax checks passed. Native evidence must match each newly
deployed public pin; earlier v1.0.6 passes are retained history.

- [x] Deploy the exact v1.0.7 manifest/caller pins to all four public projects; retain both private main commits and v1.0.6 pins unchanged.
- [x] Reconcile and execute the public repair's local tooling/format/static controls without treating them as native acceptance.
- [x] Verify the four public rulesets require the same three compiler-free checks and retain the owner-only PR-merge emergency route.
- [x] Re-run the central compiler-free observer to four-of-four public coverage; retain the two private access refusals as incomplete all-six coverage.
- [ ] RedSalamander: fix current contract failures and complete both Release regression/probe branches, barrier, production and ZIP smoke.
- [ ] RedXe: restore the independent pinned ports catalogue without changing either revision or lowering floors; complete the same native pipeline.
- [x] RedMonitor: qualify the v1.0.7 deployment on both architectures, retaining its v1.0.6 success separately. Original run 41206180882 passed; downloaded canonical receipts, ordering and both ZIPs were independently verified.
- [ ] DxUi: qualify both native regression/probe and SDK consumer branches on v1.0.7; no publication or Prepare applies.
- [ ] Independently validate source/engine-bound descriptors, counts, hashes and job ordering for these four public projects.
- [ ] Complete artifact-only Prepare for RedSalamander from its current successful qualification; retain no public writes.
- [ ] Complete artifact-only Prepare for RedXe from its current successful qualification; retain no public writes.
- [x] Complete artifact-only Prepare for RedMonitor from its current successful qualification; retain no public writes. Run 41208578017 passed and its downloaded plan/ZIPs were independently revalidated; release, install and submit were skipped.

The fresh RedMonitor v1.0.6 qualification 41198030903 passed both native roles and all six scenarios per architecture.
[Prepare 41203473510](https://github.com/RedSalamanders/RedMonitor/actions/runs/41203473510) passed for source
40db036f625a9746f50c6dfe508d6bde1c20ad51, version 0.1.68. Downloaded complete receipts and ZIP bytes were revalidated
with Read-CiPublicationPlan. x64 SHA-256 is de7316cb2cdf160ad9abb70b517c9fd7b3abad41600909117129b30582f6e0ce;
ARM64 is bb44fa7fe52d8124145ea7a70f7d2bb26ad8aebf614226a3d7743f701d2a94ad. Release, installation and submission
were skipped. This does not qualify the upcoming v1.0.7 identity.

The RedXe failure is a source-role mismatch, not a requirement to lower floors or make pins equal: the executable
checkout's older version database was being used for the newer manifest baseline. The repair restores independent
clean catalogue sources, selects them with --vcpkg-root and retains the existing tool, baseline and dependency floors.
Actual Git controls cover distinct revisions, warm restore, port metadata, dirty/foreign/non-clone refusal, invalid
identities, caller environment and deep port paths. Native dependency and runtime execution remain required.

RedSalamander's failed tooling checks included obsolete pre-unification matrix/schedule assertions, a retired plan
route, the former Dependabot layout, inline bootstrap details and winget logic now moved into tested modules.
The guards are reconciled to current authority while preserving required suites, manual-only diagnostics, immutable
pins, real native install receipts and owned cleanup. The existing I50 H6 foreground failure remains a real native gate;
its admission diagnostics now enter the case failure summary instead of being available only in a separate trace.
No case is removed or counted as executed after failed foreground admission.

Evidence for this continuation is retained beneath Z:/RedSalamander.Perf/evidence/unified-ci/public-acceptance-20261031.
The unrelated dirty review-remediation checkout is preserved. No active heavy workflow is cancelled and no private
repository, product visibility, account setting, credential or public release is changed.

### Public repair integration and local verification

RedMonitor [PR 25](https://github.com/RedSalamanders/RedMonitor/pull/25), merge
df1b3aacf37498e077dfa850e66f88edf8c7c72c, and its separate canonical fold PR 26, merge
7e3e8bdf991b174359537b3d9de9cae9096800b1, passed the normal three required checks. Its original v1.0.7
qualification 41206180882 is executing; x64 has actual six-of-six scenario execution and an active probe pass.
ARM64, the global barrier and production remain pending at this observation.

DxUi [PR 154](https://github.com/RedSalamanders/DxUi/pull/154), merge
07ff5741cc0cbcd10f7aa5e9e120d0c4b0c16d92, and canonical fold PR 155, merge
125dd9de766a861ab95ba6ca05863b763650b5ae, also passed normal checks. Native v1.0.7 run 41206683035 is
executing. Its v1.0.6 run 41198401034 completed both regressions, barrier and SDK/extracted-consumer roles
successfully; this historical pass does not replace v1.0.7 acceptance. No DxUi publisher was added.

RedXe [PR 71](https://github.com/RedSalamanders/RedXe/pull/71), source
81101ac3209e1af5256b9920d95768c3bec6cfd1, contains the independent catalogue repair and matching v1.0.7
callers. The complete independent tooling invocation and compiler-free Static adapter passed. Actual pinned
vcpkg dry runs resolved both x64 and ARM64 at curl 8.23.0, yyjson 0.13.0 and WIL 1.0.260126.7; executable,
baseline, consumer lock and dependency floors remain unchanged. Compiler detection in that dry run is not a
product build or native qualification. Original failed/incomplete-clone evidence is retained beside the final
resolution log; no incomplete clone was silently reused or removed.

RedSalamander's current Pester profile, excluding the RequiresBuildToolchain tag, finished with 1112 passed,
zero failed and three conditional skips: two require the product's pinned vcpkg checkout, and one requires a
built ProductUiTests.exe. Those prerequisites remain native obligations. TEMP and TestDrive were allocated inside the existing marked
sandbox. An earlier invocation used ordinary temporary paths and seven containment guards correctly refused
them; its failed receipt is preserved. The successful rerun changed no sandbox guard. Static admission passed;
affected planning conservatively selected native work and retained interactive obligations as pending.
These local checks do not establish native runtime acceptance or waive I50 H6.

### Public deployment and metadata checkpoint - 14:04 UTC

RedSalamander PR 186's reviewed source 4719744ed84ffd1b67b26ef2b94141318a55ca61 merged as
1f434930303bc878888df2787014d166d972b134; canonical fold PR 187 merged as
c0d5a4f6391991c37a17c7576b5e68fda07ad5aa. RedXe PR 71 merged as
59bdd293af016fb9440bb4a4c0435057e225be23 and fold PR 72 as
964cc69143f57b07a509b604ab253c806b1da693. Every merge used the exact reviewed head and all three normal
required checks, with no administrator override. RedXe native run 41207884174 is executing from its folded main.
RedSalamander's earlier v1.0.6 run 41198411091 remains active and was not cancelled; new candidates serialize behind it.

The public ruleset snapshot public-required-checks-current.json confirms active enforcement of checks / ci-gate,
security / codeql and security / dependency-review, with OrganizationAdmin bypass limited to pull_request.
No automation or native/publication bypass was introduced.

The original [central observer 41208002879](https://github.com/RedSalamanders/.github/actions/runs/41208002879)
read all four public immutable trees at v1.0.7. Its actual overall result is failure because the automatic central
token receives HTTP 404 for the two private siblings. Their unavailable rows remain explicit; the all-six observer
criterion is still open. The owner-authenticated, read-only GitHub CLI transport separately returned six-of-six
metadata coverage. It confirms unchanged private source d53625460ea7a177f02d4e666acac84e91505469 for RedPrism,
978c41fefddec1a1c64c2e09ac36dde91d02b70c for RedTilde, and unchanged engine
19710a8aeac979a82e83bec82e8b866cee3d3deb in both. No secret value was read and no private repository was changed.

RedMonitor's current run 41206180882 passed both six-scenario Release branches and active probes. Downloaded
regression receipts passed Test-Qualification -Phase Regression. Actual job timestamps prove both production jobs
started after the successful global barrier. Its x64 production ZIP independently passed receipt, size and SHA-256
verification: 82f28858e9fa7011c80599a2a46f8fc94457aaf2ea6b6bb640ac25bc249e5c5e, 2725178 bytes, version 0.1.70.
ARM64 production and current Prepare remain pending at this checkpoint. RedMonitor PR 27, merge
9aac0e5df4e9782b3390fbb50d5b12af374a568d, reconciles the credential specification without changing its native digest.

### RedMonitor v1.0.7 native acceptance

Original [qualification 41206180882](https://github.com/RedSalamanders/RedMonitor/actions/runs/41206180882),
source df1b3aacf37498e077dfa850e66f88edf8c7c72c, completed successfully at version 0.1.70 with the exact v1.0.7
engine. Both native Release test roles discovered and executed six scenarios, with zero failures/skips and active
failure-probe passes. Downloaded canonical receipt SHA-256 values match their regression references. The actual
global barrier completed after both regressions; both production jobs started later and passed ZIP smoke with
test mode disabled. Test-Qualification -Phase Complete independently accepted the downloaded descriptor,
regression/production receipts and ZIP bytes.

The production ZIP SHA-256 values are x64 82f28858e9fa7011c80599a2a46f8fc94457aaf2ea6b6bb640ac25bc249e5c5e
(2725178 bytes) and ARM64 50caf0a11e63cdff05770526ec5dc0692e1754771465877c4fac65b838c5f107 (3127817 bytes).
Proof files are RedMonitor-v107-independent-artifact-proof.json and RedMonitor-v107-ordering-proof.json under
the public-acceptance evidence directory. Native evidence and ZIPs are retained there separately from the
historical v1.0.6 receipts.

RedMonitor's credential documentation PR 27 and separate canonical fold PR 28 are merged; current main is
7823b333908dd120382bf86ce474411d3ff3596e. Their prose/fragment changes preserve native digest
5eb7ec9cce22ba51faf11fbf646987661efa4127a9533748f54b25af8ce40bad.

Manual [Prepare 41208578017](https://github.com/RedSalamanders/RedMonitor/actions/runs/41208578017) passed against
qualification 41206180882 from trusted main 7823b333908dd120382bf86ce474411d3ff3596e. The downloaded prepared
plan passed Read-CiPublicationPlan, including both original production ZIP hashes/sizes above and exact source,
run, engine and version binding. Actual release, install and submit jobs were all skipped. The separate local
proof is RedMonitor-v107-prepare-proof.json; prepared-publication bytes, job metadata and execution log are retained.
This completes RedMonitor's public I56 Release/Prepare subset, without claiming Debug/ASan, live ETW, native
winget installation, exercised write permission or upstream publication.

The subsequent prose/fold [CI run 41208545297](https://github.com/RedSalamanders/RedMonitor/actions/runs/41208545297)
succeeded and explicitly retained qualification 41206180882 with no native work dispatched. Its original log is
RedMonitor-docs-no-native-coordinator.log. The two credential-documentation merges therefore did not compile the
same native inputs again; this is observed exclusion/reuse behavior, separate from the intentional test/production roles.

### RedXe restore/build confirmation

Original v1.0.7 [run 41207884174](https://github.com/RedSalamanders/RedXe/actions/runs/41207884174), source
964cc69143f57b07a509b604ab253c806b1da693, passed the actual test-enabled Release restore/build step on x64
at 14:13:49 UTC and ARM64 at 14:15:41 UTC. This confirms the original catalogue blocker is repaired on both
native hosts without changing the independent tool/baseline revisions or manifest floors. Both regression/probe
steps are still executing at this checkpoint; no production, complete qualification or current Prepare pass is claimed.
