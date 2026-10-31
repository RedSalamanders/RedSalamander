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

The current engine is v1.0.6, reviewed source 19710a8aeac979a82e83bec82e8b866cee3d3deb, delivered by
[engine PR 10](https://github.com/RedSalamanders/.github/pull/10). Every project's manifest and reusable callers must
agree. Prior v1.0.5 native receipts remain historical evidence and do not qualify the new engine identity.

Normative owners: [hosted CI](../../Build/Build_UnifiedCI.md), [build identity](../../Build/Build_Toolchain.md),
[test evidence](../../Testing/Testing_ValidationEvidence.md), [self-tests](../../Testing/Testing_SelfTests.md),
[portable ZIPs](../../Installer/Installer_PortableZip.md), [winget](../../Installer/WingetIntegration.md),
and the [shared engine contract](https://github.com/RedSalamanders/.github/blob/main/Specs/CI.md).

## Checklist

- [ ] Resolve RedSalamander's actual regression and evidence-upload failures with the existing native/test owners; retain original failing cases and receipts.
- [ ] Resolve RedXe's coherent vcpkg baseline/floor mismatch without silently lowering dependency or native acceptance requirements.
- [ ] Owner resolves the RedPrism/RedTilde billing/visibility prerequisites; confirm real compiler-free CI, CodeQL and dependency review execute.
- [ ] Apply and verify the common required checks, organization-owner PR-only emergency route, supported dependency/security settings and absence of duplicate managed native CodeQL in those two repositories.
- [ ] Complete original successful v1.0.6 main qualification on native x64 and ARM64 for all six; verify actual regression counts, active failure probes, global barrier, production roles and ZIP/SDK smoke.
- [ ] Independently verify matching retained descriptors, receipts, SHA-256 ZIP bytes and actual job ordering. Admission-only green runs never satisfy this item.
- [ ] Execute artifact-only Prepare against each of the five applications' current successful source/engine qualifications; preserve refusal tests and keep all public writes skipped.
- [ ] Execute and verify the requested organization WINGET_TOKEN promotion for exactly the five applications; record credential availability/authentication separately from actual publication permission and native installation proof.
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
