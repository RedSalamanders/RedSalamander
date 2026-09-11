# DxUi shared-library adoption and release plan

## Progress checklist

Scope: **all I19 implementation, tests and specs authorized on 2026-09-09**. Current slices: library qualification, RedXe native CI and the full RedSalamander solution migration. The full checklist below remains open until each gate is actually satisfied.

- [ ] Current Fresh Full runs: Debug in `Z:\RsI19Ci`, ASan Debug in `Z:\RsI19`. Debug tooling reports
  767 pass/1 stale formatter-write assertion. Compare reports 217 pass/25 fail/30 skip: the synthetic Fake Phone
  journal has exhausted all 64 quarantine names under the real user profile. Preserve that evidence and fix fixture
  isolation after the receipt-bound checkouts are idle. Commands and the remaining suites are still running.
- [x] RedXe a812cec at DxUi ecad707 passes all six native CI profiles (34374884161), including the repeated
  Process Viewer lifecycle correction and native ARM64 sanitizer detection.
- [x] Isolated repair checkout `Z:\RsI19Fix`: formatter policy and MTP case-isolation source guards pass
  222 focused checks. All sixty MTP registrations now enter a fresh journal sandbox per attempt and restore
  the prior environment afterward. C++ compilation and poisoned-journal runtime proof are still pending.
- [x] The unchanged original Debug executable fails `mtp_property_fetch_is_batched` with the copied 64-slot
  saturated journal and passes with fresh scratch state. Both receipts are retained as `mtp-baseline-poisoned`
  and `mtp-baseline-clean` under the I19 evidence folder. This establishes a pre-existing isolation failure.
- [ ] Correct MonitorTest's progress display to use seconds for its seconds/rate fields; the existing diverse
  message path supplied milliseconds, while its final summary already used seconds. Native validation is pending.
- [ ] [blocked] New RedXe resource-fixture CI at 80b02a3 cannot start: GitHub reports an account payment or
  spending-limit restriction (34377362060 / 34377358597). Organization-owner resolution is required; this is
  unrelated to public DxUi access. The preceding six-profile product qualification remains valid for a812cec.
- [ ] Quiet resource acceptance is paused while the independently running original Release Monitor consumes
  about 20.7 GiB and Windows reports less than 1 GiB free. Authorization to close that instance is pending.
- [ ] Paired RedXe AV resource fixture: identical two-view harness added to original and candidate checkouts;
  Debug smoke passes, paired quiet-machine measurement and acceptance remain open.

- [x] Source gap analysis: origin records, live direct-header/project consumers, post-extraction changes, named test dispositions and baseline PNG comparison documented.
- [ ] C0 runtime/module closure, targeted reproductions, test replacement mapping, and retained baselines complete.
- [x] User decisions recorded: use `DxUi.lib`; no G5 legacy compatibility work; standalone Slider behavior is authoritative; six build configurations are mandatory; G4 separates native clipboard capacity while retaining validation and responsiveness.
- [ ] One canonical library implementation and public consumption path; all product modules/adapters accounted for.
- [x] G4 native large-selection copy/paste and failure cases pass while the embedded text ceiling remains enforced; no UI-thread retry sleeps.
- [ ] Every project supports Debug/Release/ASAN Debug on x64/ARM64, with matching native instrumentation, working sanitizer detection and native ARM64 qualification.
- [ ] DxUi standalone matrix and both product candidate gates provide current, traceable evidence.
- [ ] External dependency changes invalidate build/test receipt reuse and are visible in shipped provenance.
- [ ] RedSalamander migration preserves required product interaction, accessibility, localization, settings and resource behavior while adopting the accepted standalone Slider and diagnostics changes.
- [ ] Advisory notices and the manual upgrade/fix/retest loop work; consumer CI and rollback are rehearsed.
- [ ] Required hardware/manual/native-platform gates pass; any out-of-scope capabilities remain explicitly unqualified under their existing owner.
- [ ] Legacy implementation removed only after complete migration; docs/specs/tooling/index closeout complete.

### Active library work

- [x] G4 behavior approved: native large selections, validated memory/Unicode, no retry sleeps, separate embedded ceiling.
- [x] Existing implementation, tests and owning contracts inspected.
- [x] Retain matching Debug/Release baselines before library edits (including retained noisy and repeated Release runs).
- [x] Add a regression that fails on the current native clipboard ceiling.
- [x] Implement native capacity separation and failure-safe handling.
- [x] Verify native boundary/large selections and unchanged embedded rejection in x64 Debug NativeTextInput; remaining configurations stay open.
- [x] Update library specs, usage docs and focused regression evidence; final matrix and paired acceptance remain open.
- [x] G4 additionally passes x64 ASan Debug NativeTextInput with a working sanitizer detection probe.
- [x] G1 tooltip fix and G7 offscreen Grid selection fix pass focused x64 Debug suites; Embedded remains green.
- [x] G2 neutral public helper headers compile/link in the standalone consumer test; exact-pin relocation and product adapters remain open.
- [x] Native module collision reproduced: DLL copies lost animation and used the executable's menu class. Module-owned registration at DxUi ecad707 passes all 12 external EXE/two-DLL probes, rendering and ten negative pin/build checks in all six native CI profiles (34371134464), including both ASAN annotation policies. Local x64 Debug/Release/ASan Debug suites pass. Consumer candidates now pin this correction; product qualification is in progress.
- [x] G3 x64 Debug localized empty/no-match rendering and geometry regressions pass.
- [x] G8 standalone six-configuration mapping and native sanitizer detection qualify at 3dae073 (CI 34355763655); consumer matrices remain open.
- [ ] Complete required matched performance acceptance. All fourteen quiet repeats and 32 alternating CPU-subset repeats are retained. Repeated-baseline comparisons still cross thresholds; the prior Debug composition signal drops to +0.61% across the eight paired medians. Results remain inconclusive; no threshold or baseline was relaxed.

### Consumer implementation progress

- [ ] Qualify both consumer candidates at DxUi ecad707, which corrects native menu/animation ownership in statically linked plugin modules. Previous consumer pins and receipts remain retained.
- [x] RedXe ecad707 passes complete local x64 Debug/Release/ASan Debug suites. A cumulative Process Viewer counter assertion exposed in duplicate CI is reproduced by repeating the lifecycle in-process; corrected delivery-delta checks pass all three local suites at a812cec. Fresh native CI remains pending.
- [ ] RedSalamander Fresh Full Debug is running on the corrected Preferences candidate; its full build has zero warnings/errors. Fresh Full ASAN is building at 2f964a9a, including the corrected ViewerSqlite helpers and CI exit propagation. Other product gates remain open.
- [x] RedXe unchanged x64 Debug and Release suites pass; previous Release package retained before adoption.
- [x] RedXe candidate 8c548fe passes full x64 Debug, Release and ASan Debug suites, including sanitizer detection and module provenance.
- [x] RedXe candidate upgraded to DxUi 52da33d passes full local x64 Debug, Release and ASan Debug suites with zero build warnings/errors and all six native CI profiles at RedXe 1bf6662 (34368430957), including ARM64 sanitizer detection.
- [x] RedXe native x64/ARM64 Debug, Release and ASan Debug CI passes at 2ef0bcb (run 34359693390), including product tests, real sanitizer detection and provenance. Manual/hardware and paired resource acceptance remain open.
- [x] DxUi native x64/ARM64 suites, external fixtures, both ASAN annotation policies and gallery pass all six configurations at 3dae073 (CI 34355763655).
- [x] Relocated ASAN consumer with RedSalamander's STL compatibility policy passes rendering and ten negative pin/build checks.
- [x] RedSalamander original Debug and Release Full/Fresh baselines captured. Neither is green: Debug has one tooling failure, three Commands failures and one FileOperations failure; Release has one tooling failure, ten Commands failures and one FileOperations failure. Full reasons are retained in `Z:\RedSalamander.Perf\evidence\I19-20260909\baseline-results.json`; diagnosis and clean comparison remain open.
- [x] RedConfigureTests Debug links the exact external archive, with zero warnings/errors and an attested module sidecar.
- [x] Focused dependency tests reject malformed pins, missing/dirty source, wrong archive inputs and incompatible identities.
- [x] RedConfigureTests x64 Debug passes all 42 cases, including the actual Japanese satellite and all twelve combo-box/TagPicker inputs. App and remaining matrix/resource qualification stay open.
- [x] Monitor x64 Debug builds with the external archive and zero warnings/errors; runtime/resources remain open.
- [x] Full RedSalamander x64 Debug solution builds against 3dae073 with zero warnings/errors, including main, every viewer, Terminal, Monitor, RedConfigure and PerformanceTests2.
- [x] Remove obsolete ARM64 ASAN rejection from runner/deployment preflights and make native Commands inventory use its selected profile. All six accepted profiles and invalid-name rejection pass in 54 focused runner-plan tests. Native ASAN execution remains a separate CI gate.
- [x] CI failure diagnosis: e8e5e606 ASAN compilation failed in ViewerSqliteTests because shared theme-test UIA helpers were hidden from the sanitizer build. A subsequent matrix command incorrectly replaced the build exit code. Corrected helpers compile in x64 ASAN with zero warnings/errors; 19 build-evidence/workflow checks pass, including a real failing-script exit propagation test. Fresh native Full qualification remains open.
- [ ] Full product regressions and the remaining five native profiles pass. Candidate Debug exposed 53 Commands failures before a modal-test stall. A focused Preferences run reported 137 pass/39 fail; debugger snapshots prove retained dialog state and destroyed HWND registrations. The Preferences wheel subclass removes its saved procedure before forwarding WM_NCDESTROY, bypassing owner cleanup. Fix and regression qualification are in progress; the 128-window library bound is unchanged.
- [x] Reproduce the Preferences leak with a 48-cycle regression: the unfixed build retains 10 hosts/attachments after its first close versus a baseline of 7. Preserve the failing trace in `Z:\RedSalamander.Perf\evidence\I19-20260909\preferences-lifetime-before-context.json`.
- [x] Corrected Preferences subclass forwarding and page-host teardown build in the full Debug solution with zero warnings/errors. The 48-cycle regression passes: both counters return to 7 after every close, versus 10 after the first unfixed close. Paired retention evidence: `Z:\RedSalamander.Perf\evidence\I19-20260909\preferences-lifetime-retention.json`.
- [x] Correct the second teardown defect: pane models/delegates now disconnect before page-host Detach destroys their controls. The expanded 48-cycle regression rotates through all fourteen categories and returns hosts/attachments to 7/7 after every close. Preferences family: 176 passed, zero failed, one foreground-dependent skip; focused PowerShell 7 harness checks: 202 passed. Retained evidence: `Specs/TestRuns/Local-x64/Commands/2026-09-09_154457-DxUiPreferences/`.
- [ ] Qualify the teardown fixes in Fresh Full and the other five native profiles; family diagnostics do not replace those gates.
- [x] Initial candidate Full run passes FileOperations (180/0/20 skipped), CompareDirectories (242/0/30 skipped), ProductUi, plugin contracts, both noninteractive viewers and other standalone suites. Two long-path dependency fixtures and the interactive ViewerPE class-name expectation require harness corrections; Commands and final Full acceptance remain open.
- [x] ProductUiTests x64 Debug builds and passes all 23 retained product cases plus two viewer adapter regressions. CI/Full runner mapping and 48 runner-plan, seven inventory, and 40 build-evidence/workflow tests pass.
- [x] Sandbox cleanup is scoped to its fixture IDs; 48 plan-helper and four dependency/profile tests pass before the product runner migration.
- [x] Product source-policy guards rehomed: NavigationView input and credential modal-quit, child-pointer teardown and joined UIA workers; 193 source-policy checks and 19 tooling-governance/impact checks pass. The remaining 61 library dispositions are recorded in DxUi Specs/Testing/SourcePolicyDispositions.json at 76e3d48: runtime replacements, retired helper spelling/G5 diagnostics and explicit ownership reviews. Four mixed runtime cases and the corrected Menu interactive fixtures pass all six native configurations at DxUi 52da33d (CI 34363702073).
- [x] Relocated native sandbox regression passes Debug/Release/ASan Debug on x64 and reproduces the original package-helper failure with the old header. Complete package smoke remains open.
- [x] CI managed-checkout root cause identified: d0b3067 native jobs all stop before compilation because root `vcpkg/` violates source snapshots. Stage the tool/cache in `.build/vcpkg-tool`; 57 fingerprint/build/dependency tests pass, including a real nested-Git fixture. New native CI remains pending.
- [ ] Existing receipt reuse checks reject dirty/missing dependency source; linked-module sidecars are attested and packaged. Full Debug build emits all selected module sidecars; package and full-suite qualification remain open.

Work proceeds in `Z:\RsI19` on `codex/dxui-i19` while the original checkout supplied the completed baseline runs. This checklist is synchronized to the
requested original WIP folder; bring the implementation back after candidate validation.
DxUi anonymous Git and Actions API access is verified: neither consumer needs `DXUI_READ_TOKEN`.
The six-job RedXe workflow uses the automatic job token for advisory API requests. Native CI at RedXe 2ef0bcb passes all six configurations (34359693390) after correcting a Windows code-integrity test assumption and replacing a personal Codex-dependent validator with repository-owned validation. No native or hardware qualification is inferred from cross-compilation.

Status: **ACTIVE, I19 — standalone native suites pass all six configurations; RedXe local matrix and package rollback pass; RedConfigure pilot qualification and broader adoption remain open**.

Priority: **P1**. Effort: **L**, delivered in bounded changes. Risk: **High** for RedSalamander migration; **Low–Medium** for the advisory version check and consumer upgrade tooling.

Owner: this plan owns the cross-repository adoption sequence, dependency qualification, and RedSalamander migration. DxUi owns shared controls and library contracts; each product owns its adapters, behavior, packaging, and release decision. The user authorized implementation, tests and spec updates on 2026-09-09. G4 is the first active slice; publication and consumer pin changes are separate lifecycle steps.

Planned on 2026-09-09 against these clean local checkouts. These are audit identities, not new dependency pins:

| Repository | Local checkout | Audited commit |
|---|---|---|
| [DxUi](https://github.com/RedSalamanders/DxUi) | `Z:\src\DxUi` | `d192e474e540adc2656e69b8e8150b0f7a05d63f` |
| RedXe | `Z:\src\RedXe` | `75545887d002b5075d170d98b7bfd3a77d8d7f61` |
| RedSalamander | `Z:\src\RedSalamander` | `3f0df2f2c28239eb658a68f85b9c7f69c2356af4` |

This is a **non-normative proposal**. Existing contracts remain in force until implementation, tests, and authoritative updates land. Local paths are discovery context; build automation must support relocated checkouts.

## 1. Accepted direction

1. Keep **one canonical DxUi implementation**, built from its repository. Extend RedXe's existing exact-pin consumption; migrate RedSalamander from its in-tree implementation.
2. Use **`DxUi.lib`**. The user selected static linking on 2026-09-09 after reviewing the sizing estimates. DLL implementation and experimentation are outside this plan.
3. Use the simple consumer-driven update loop: merge DxUi with green library tests, show an advisory new-version notice in consumer builds/PR checks, then deliberately update that consumer's pin, build and run regressions locally, and push/merge its PR after consumer CI passes. If DxUi causes a failure, fix and test it upstream, merge it, then update the consumer pin and repeat.
4. Use **two independent compatibility gates**: library behavior tests in DxUi, and candidate-versus-current product tests in RedXe and RedSalamander. Neither replaces the other.
5. Preserve a tested previous pin and product package for rollback. Track which DxUi revision each built and shipped module actually contains; an updated source lock alone is insufficient evidence of adoption.
6. **Accepted after the gap audit:** adopt current DxUi diagnostics/scheduling without legacy compatibility work (G5). The standalone DxUi Slider is the correct behavior (G6); adapt consumers and test expectations to it.
7. **Accepted G4 clipboard policy:** native copy/paste supports large selections independently of the embedded editor's 65,536-unit ceiling. Keep allocation-bounded reads, Unicode/size validation, explicit allocation/clipboard failure and the single open attempt; no UI-thread retry sleeps. Preserve the separate embedded edit/snapshot ceiling.
8. **Mandatory for every project:** Debug, Release and ASAN Debug builds on both x64 and ARM64 across DxUi, RedXe and RedSalamander. Deliver missing support; ASAN is not optional and must not silently map to unsanitized Debug.

The first useful milestone is one RedXe upgrade through this loop plus a RedSalamander consumer/drift inventory and retained baseline. Each product adopts on its own schedule; normal builds continue using its existing exact pin.

### 1.1 Required build matrix

| Configuration | x64 | ARM64 |
|---|---|---|
| Debug | Required | Required |
| Release | Required | Required |
| ASAN Debug | Required | Required |

Apply this matrix to every library, application, plugin, sample and test project in the three repositories, with consistent mappings for resource/package projects. Preserve the established MSBuild name `ASan Debug` where used. ASAN builds must instrument native project code and link the matching instrumented DxUi archive; a configuration label or ordinary Debug fallback does not satisfy the requirement.

Implement and verify project/solution mappings, compiler/CRT/sanitizer runtime, build/test CLI accepted values, dependency restore/imports, isolated outputs/fingerprints and CI/test coverage for all six combinations. Native ARM64 execution qualifies runtime behavior. Missing toolchain/runtime/runner support is a delivery dependency to resolve, not an optional skip or a waived platform. This is the accepted target; the audited current code does not yet establish it.

## 2. Current state and gaps

The completed [codebase gap analysis](DxUi_SharedLibrary/GapAnalysis_2026-09-09.md) accounts for all 58 origin records, 79 live consumer source files with 96 direct DxUi includes, and the inherited test cases. It confirms a missing tooltip fix, private-header consumption gaps (including header-only ViewerVLC), untranslated built-in empty/no-match strings, and changed native clipboard behavior. G4's native capacity separation is implemented and passes all six native configurations; G5 is accepted without legacy preservation; G6's standalone Slider behavior is authoritative. G8 owns delivery of the mandatory six-configuration build matrix. C0 remains open for retained baselines, targeted reproductions and runtime qualification.

| Evidence inspected | Current fact | Required follow-up |
|---|---|---|
| DxUi [consumption contract][dx-build], [project][dx-project], and [capabilities][dx-capabilities] | One static target, public headers, exact API/target validation, consumer MSBuild imports, isolated outputs. Native and embedded mechanisms already exist. | Reuse these entrypoints; do not repeat extraction or introduce consumer-maintained source lists. |
| RedXe [lock][rx-lock], [integration contract][rx-integration], [restore][rx-restore], and `Build/RedXe.DxUi.props` / `.targets` | RedXe already pins the audited DxUi HEAD. `RedXe`, `AVControl`, and `AVControlTests` reference the archive. Host and plugin keep DxUi C++ ownership inside their respective modules. | Add an advisory version check and rehearse a deliberate pin-update PR using existing tests. Preserve host/plugin ownership; no initial integration rewrite is needed. |
| RedXe [AV plan][rx-av-plan] | Synthetic text/UIA and WARP tests exist. Real IME, screen-reader/touch, matched text/UIA resource acceptance, and AV release gates remain open. No checked-in `.github/workflows` YAML was found in this checkout. | Establish product CI around existing tests; confirm actual repository check configuration. Keep AV release work with its existing owner. |
| RedSalamander [legacy project](../../../Common/DxUi/DxUi.vcxproj) and its project references | Production references found in RedSalamander, RedSalamanderMonitor, RedConfigure, ViewerText, ViewerSqlite, ViewerSpace, ViewerImgRaw, ViewerPE, ViewerWeb, and Terminal; test references in DxUiTests and RedConfigureTests. | The source audit adds ViewerVLC as a header-only typography consumer and maps all direct DxUi includes. Qualify the transitive build/runtime closure and actual loaded modules during C0/C4/C5. |
| RedSalamander [legacy public header](../../../Common/DxUi/DxUi.h) | Includes the product's `PlugInterfaces/Viewer.h`; implementation also uses consumer `WindowMessages.h`. | Replace consumer coupling with adapters; never move viewer/plugin contracts into DxUi. |
| DxUi [origin record][dx-origin] and RedSalamander Git history | Extraction records a September 5 source baseline. Subsequent RedSalamander changes include tooltip clock handling and test updates in `817b7e2d`, `f0710f9b`, and `3f0df2f2`. | Audit disposition: 817b7e2d tooltip host/dispatcher fix and regression are missing upstream; f0710f9b adds product UTF utility coverage; 3f0df2f2 adds the product runner input warning. Resolve the shared fix in DxUi and retain the product cases/runner policy. |
| DxUi [validation contract][dx-testing] and [CI][dx-ci] | Foundation, inherited controls, embedded WARP, public consumer fixtures, galleries, and native x64/ARM64 Debug/Release jobs are defined. | Inspect actual receipts when qualifying a candidate. A workflow definition or historical pass is not a fresh execution result. |
| RedSalamander [test coverage](../../Testing/Testing_TestCoverage.md), [suite plan](../../../Tools/Modules/Testing/TestSuitePlan.psm1), and [validation evidence](../../Testing/Testing_ValidationEvidence.md) | Full already includes DxUi families and product integration tests; build/source/runtime receipts govern reuse. | Extend this machinery for external DxUi inputs; do not create a competing test runner or count old-library tests as new-library proof. |
| DxUi [migration HOLD][dx-rs-hold] and consumption prose | RedSalamander migration remains on HOLD. A trailing paragraph still says RedXe needs its first pin/bridges, contrary to the inspected consumer code and newer integration contract. | At implementation activation, reconcile stale prose and make the HOLD record route to this consumer owner. Do not claim real IME/AT completion while correcting the pin status. |

Before each implementation slice, repeat `git status --short --branch` and compare its owning paths against the audited identities:

```powershell
git -C Z:\src\DxUi diff d192e474e540adc2656e69b8e8150b0f7a05d63f..HEAD -- include src Build Tools Specs
git -C Z:\src\RedXe diff 75545887d002b5075d170d98b7bfd3a77d8d7f61..HEAD -- Dependencies Build RedXe Plugins Tests Specs
git -C Z:\src\RedSalamander diff 3f0df2f2c28239eb658a68f85b9c7f69c2356af4..HEAD -- Common RedSalamander RedConfigure RedSalamanderMonitor Plugins Tests Tools Specs
```

Inspect worktree changes separately. Use focused branches; never reset another checkout. Keep the human-managed scratch note outside this work.

## 3. Packaging decision: DxUi.lib

### 3.1 Decision and scope

**Accepted by the user on 2026-09-09: use the static library.** Build `DxUi.lib` from an exact source pin with matching public headers and compiler/CRT settings, then link the dependent EXEs/plugins. This matches the existing supported consumption path and avoids adding an exported binary interface and runtime dependency.

The sizing inspection below informs the choice: Release static duplication is a few MiB across the product, while most control/model/graphics allocations would remain with either packaging. Retain those estimates as decision context, not a future DLL work queue. A later packaging change would require a separate decision.

### 3.2 Validation still required for static adoption

Retain a baseline of the current in-tree/static implementation and compare the shared-library candidate using the same product scenarios and fixtures. Preserve existing performance/resource gates for startup, input/frame latency, preparation/composition, allocations, cache/surface retention and idle work. Check cold/warm build behavior and unintended rebuilds during integration; add optimization work only for an observed problem. This is migration validation, not a prerequisite to reconsider the accepted packaging choice.

### 3.3 Initial estimate from current RedSalamander binaries (2026-09-09)

This is a sizing estimate for the current **in-tree RedSalamander DxUi**, using x64 Release binaries from the September 9 build at the audited RedSalamander commit. It is not a measured static/DLL comparison, a measurement of the standalone DxUi candidate, or completion of C0/C1. No product was launched and no DLL was built for this inspection.

The existing build receipt records `tests_enabled=false`, compiler `19.51.36256.0`, and receipt ID `3a4434b69e1af54e524fdddbddf549e6ad4acbbaecaafc973861886c4674cf80`. All ten inspected PE file hashes match that receipt, and each PE CodeView GUID matches its PDB. The library compiler log contains `/O2 /GL /Gy /MD`; the inspected ViewerText linker log contains `/OPT:REF /OPT:ICF /LTCG:incremental`. Dead-code removal is already active in that consumer.

Read-only method: use `llvm-pdbutil dump --modules --section-contribs` on each matching PDB, select object modules under `Common/DxUi`, and sum the union of their contribution ranges in each section. Parse only the final section-contribution table, excluding the module-summary contributions to avoid double-counting. Read PE file sizes separately. Attribution covers `.text`, `.rdata`, and `.data`; linker-created unwind/relocation records, alignment, cross-module inlining, and shared template attribution prevent treating these sums as exact removable bytes. See [LLVM's PDB inspection tool](https://llvm.org/docs/CommandGuide/llvm-pdbutil.html).

| Current Release module | Whole file, MiB | PDB-attributed DxUi code/data, MiB |
|---|---:|---:|
| RedSalamander.exe | 6.561 | 0.644 |
| RedSalamanderMonitor.exe | 1.267 | 0.353 |
| RedConfigure.exe | 1.381 | 0.537 |
| ViewerText.dll | 1.270 | 0.523 |
| ViewerSqlite.dll | 0.945 | 0.506 |
| ViewerSpace.dll | 1.045 | 0.418 |
| ViewerImgRaw.dll | 1.046 | 0.463 |
| ViewerPE.dll | 0.924 | 0.463 |
| ViewerWeb.dll | 1.635 | 0.462 |
| Terminal.dll | 0.606 | 0.009 |
| **Total, these ten files only** | **16.680** | **4.380** |

Exact totals are 17,489,920 file bytes and 4,592,672 attributed bytes. The Release archive itself is **121.461 MiB**, a build artifact containing compiler/link inputs; it is not shipped as 121 MiB in every consumer and does not represent RAM use. Debug is a separate, non-shipping comparison: its archive is 95.046 MiB and the main executable is 77.724 MiB. Debug sizes should not drive the Release packaging decision.

For a first DLL estimate, taking the largest observed contribution of each DxUi object/section across consumers gives a **0.685 MiB proxy for one copy**. This is neither a true union of required functions nor a DLL size prediction: consumers retain different subsets, and exporting the public API changes optimization. Allow roughly **0.8–1.5 MiB for a Release DLL**, then remaining inline/template code and adapters in consumers. The resulting decision ranges are:

| Cost | Current static delivery | DLL scenario estimate | Expected difference |
|---|---|---|---|
| Installed DxUi-related code/data across these modules | Approximately 4.4–5 MiB, including a rough allowance beyond direct attribution | Approximately 1–2 MiB across the DLL and retained consumer code | **About 2.5–4 MiB less disk space**; medium confidence. This is not the total installed application or compressed package size. |
| Resident code/read-only memory in ordinary use | Depends on pages touched, not archive size or all mapped image bytes | One copy can serve modules in the process | **Budget around 0–1 MiB saved**, low confidence; main-only use may be neutral or slightly worse. |
| Resident code/read-only memory after heavily exercising several viewers, optionally Monitor/RedConfigure | More duplicated implementation pages can become resident | Same-version DLL image pages can be shared where eligible | **Perhaps 1–3 MiB saved**, low confidence; not an expected reduction in private heap or GPU memory. |
| Direct writable global data | 23,817 attributed bytes across these ten modules, excluding heap allocations | Fewer copies within one process; separate data per process | Only tens of KiB of direct globals; heap/resource retention needs separate measurement. |
| Control/model heap, per-window textures and swap chains | Determined by active controls/windows | Substantially the same for unchanged ownership | **Assume zero saving** from packaging alone. |
| CPU/FPS, threads, timers | Current rendering/input work | Same work; DLL requires no new worker/timer | Expect broadly neutral steady-state cost; do not promise an FPS gain. |
| Implementation-only development edit | Can relink ten dependent production modules | Potentially relink only the DLL if the interface/import library stays stable | Potential developer-time benefit; no timing delta measured. Header/API changes still rebuild consumers. |

The memory ranges assume that only part of each roughly 0.4–0.65 MiB implementation copy is resident in ordinary use, and more overlapping code is exercised in a heavy session. They are illustrative scenario budgets, not confidence intervals from live sampling. DLL image sharing does not share ordinary heap objects or globals across processes. [Microsoft DLL data rules](https://learn.microsoft.com/en-us/windows/win32/dlls/dynamic-link-library-data).

One resource observation retained from the comparison: [WindowHost graphics sharing](../../../Common/DxUi/DxUi.WindowHost.cpp) stores D3D11/D2D/DWrite resources in a static map keyed by thread ID. Static linkage gives each module its own map. A common DLL could consolidate compatible same-thread buckets across host/plugins, avoiding some duplicated device/context/factory resources. It would not share buckets across different threads/processes or remove per-window back buffers. No driver-memory saving is priced into the estimates above. As scale context only, two 1920×1080 BGRA8 buffers contain **15.82 MiB** of pixels before driver overhead, regardless of library linkage; this is not a claim that every current control allocates full-screen buffers.

**Decision following this estimate:** use `DxUi.lib`. The estimated DLL savings do not justify additional packaging and ABI work for this adoption. The resource estimates remain unpaired and do not waive migration performance validation.

## 4. Dependency lifecycle and release workflow

**Accepted approach, 2026-09-09:** library PR, advisory availability notice, deliberate consumer upgrade, local build/regressions, consumer PR/CI, merge. Product maintainers drive adoption. The tooling described here is planned, not implemented by editing this document.

### 4.1 The update loop

1. **Merge DxUi:** implement the change and its library regression tests in a DxUi PR. Run all required DxUi checks, including applicable performance/docs/gallery checks, and merge. An available upgrade means an exact merged `main` commit with successful required CI for that revision; a pending or failed merge build is not advertised as ready. A commit SHA is enough to identify the version; release tags are optional.
2. **Notify in the consumer:** RedSalamander/RedXe builds and existing PR checks compare their pinned SHA with the available DxUi revision and show an advisory notice. They still compile the pinned revision. No dependency file changes automatically.
3. **Upgrade locally:** create or use a focused consumer branch, change `Dependencies/DxUi.lock.json` to the selected exact SHA, then use the normal restore/build/test commands. Restore into that consumer's isolated dependency output; do not reset the developer's sibling DxUi checkout. Include any necessary adapter changes in the same consumer branch.
4. **Run regressions:** build and run the consumer's required regression and performance gates. Use the current pin as the retained baseline when comparison is required. Library-green means the library tests passed; the product tests establish whether this product can adopt it.
5. **If green:** push the consumer change and open/update its normal PR. Record old/new DxUi SHA and local validation results in the PR. Consumer CI validates the actual PR change before merge; if the tested inputs change, rerun the affected required checks. Merge and ship through the normal product process.
6. **If red:** identify whether the defect belongs to the consumer adapter or DxUi. Fix an adapter defect in the consumer branch. For a DxUi defect, add a reproducing regression test and fix in a DxUi PR; after green tests and merge, update the consumer branch to the new merged SHA and repeat steps 3–5. Keep the product regression too. The consumer default branch remains on its previous working pin throughout.
7. **If a regression appears after adoption:** revert the consumer upgrade, including paired adapter edits if necessary, and rebuild/test the previous pin. Repair the issue through the same loop. Keep the previous shipped package available for normal product rollback.

Once an upgrade starts, test a fixed candidate SHA. A newer DxUi merge may produce another notice but does not invalidate a passing upgrade or require chasing a moving target. RedXe and RedSalamander can adopt independently; one product's failed upgrade does not change the other product's pin or block unrelated consumer work.

### 4.2 Where the notice belongs

| Place | Behavior |
|---|---|
| Root consumer build/restore entrypoint | Check once per invocation, not once per project/compiler step. Show current and available short SHAs, a compare link, and the lock path to update. Use a short bounded lookup and cached result where useful; show when the result was last checked. |
| Existing PR validation | Run when a PR opens or is updated, so the maintainer sees the notice before deciding to merge. Put it in the build log/check summary; no automatic PR comments are needed. |
| Existing default-branch build after merge | Reuse the same check if that build already runs. It is a useful reminder but should not be the first place an upgrade is noticed. |

Example: `DxUi update available: pinned <old-sha>, available <new-sha>. Update Dependencies/DxUi.lock.json on a branch and run the product regression suite.`

This is an **advisory build notice**, not a compiler warning or failing MSBuild diagnostic; `/WX` must not turn it into a failure. An unavailable network/API, missing read access, or a rate limit reports **update check unavailable**, not “up to date”, and does not fail a build whose pinned dependencies are already available. A real failure to restore/validate the pinned dependency still fails normally. The check uses existing read access and needs no cross-repository write permissions.

Verify that the candidate is newer along the normal `main` history before calling it an upgrade. A same pin produces no notice; a pending/failed candidate is not recommended; a divergent pin is reported for review rather than silently replaced. Test these cases plus offline behavior, fixed-pin builds, and one notice per invocation. Keep lookup behavior in one reusable helper rather than duplicating it in build scripts and CI.

The deliberate tradeoff is that an idle product may not notice a new DxUi version until its next build or PR. That is acceptable for this workflow. A maintainer can still initiate an upgrade immediately for an urgent fix. There is no periodic monitoring service or adoption deadline in this plan.

### 4.3 Small set of rules to keep

- **Exact pin and normal source build:** keep the existing lock shape and library MSBuild imports. Build matching headers and `DxUi.lib` from that pin. No source enumeration or floating branch dependency.
- **Isolated, compatible outputs:** preserve per-consumer/configuration/architecture outputs and existing source/toolchain/CRT/build-flag identity checks. Changing the pin must invalidate stale artifacts and test receipts. All three repositories require the six configurations in section 1.1, including matching instrumented ASAN Debug libraries on x64 and ARM64; never silently substitute unsanitized Debug.
- **Existing tests and evidence:** keep the library and product suites, required supported-platform coverage, and current result/receipt formats. Add the external DxUi identity to missing build/test inputs rather than introducing a second qualification system.
- **Compatibility notes and rollback:** describe observable/API changes and any required adapter edits in the DxUi PR. Preserve minimum supported Windows behavior and normal notices/symbols/packaging requirements. Each product can retain its previous tested pin while an upgrade is repaired.

Automatic lock-update PRs, cross-repository pre-merge testing/dispatch, a per-product qualification catalogue, new receipt schemas, scheduled reconciliation, a dedicated binary-package cache, and a new version/support-line program are not part of this implementation. Revisit only a demonstrated maintenance problem; ordinary source restore, build, test and PR review are sufficient to start.

## 5. Tests owned by DxUi

Extend the existing suites and fixtures; do not replace them with a new framework. Map each supported capability and every inherited test disposition to an owner. Shared behavior discovered through a consumer defect should gain a minimal synthetic reproduction in DxUi while the product retains its end-to-end regression.

| Layer | Required evidence |
|---|---|
| Public API/build boundary | Compile/link/run consumers using public headers only; exact-pin negative cases; no application includes, consumer source lists, or incompatible flags. Preserve the standalone public samples and relocated fixture. |
| Control semantics | State/layout, event order and counts, preview/commit/cancel, validation/read-only/disabled behavior, keyboard traversal, pointer capture/cancel, nested/reentrant callbacks, deletion during callbacks, Unicode editing, grids/trees and virtualization. |
| Native hosting/input/UIA | HWND lifetime, nested modal/popups, focus restoration, DPI, IME/TSF text-store revisions, clipboard abstraction and G4 large native selections (including 100,000 UTF-16 units), independent embedded overflow rejection, failure-without-text-loss and no retry sleeps, stale/disconnected UIA providers, selected offscreen items, destruction and queued callbacks. |
| Embedded rendering | Supplied device/context, hostile pipeline state, clean/dirty/hidden/zero-size states, clipping/alpha, multiple views, failed preparation, device loss/recovery, independent pools and generation teardown. |
| Visual contract | Existing test baselines plus gallery coverage of relevant themes, high contrast, reduced motion, DPI and states. Review intended pixel changes; semantic assertions remain mandatory. |
| Resource contract | Paired complex-UI baseline/candidate, preparation versus composition, zero clean-composition allocations, cache/surface/queue bounds, hidden idle work, failure/recovery peaks and retention. Hardware presentation remains a separate gate. |
| Test/tool integrity | Case registration/disposition accounting, result/skip parsing, negative fixtures, test-owned data paths, and detection of missing suites or forged/stale receipts. Run required ASAN Debug coverage on x64 and ARM64, prove sanitizer detection with a controlled fixture, and retain applicable fault/stress coverage. |

Library CI must build and test Debug, Release and ASAN Debug on x64 and ARM64. The existing external-consumer fixture currently selects x64; extend it to the required six-configuration matrix, with native ARM64 execution instead of reporting an unexecuted fixture as green. Required unavailable capabilities produce a blocked gate with a reason. Optional skips remain visible and cannot erase a formerly failing test.

For native tests that need an interactive desktop, reserve that resource and keep deterministic fixtures separate from human acceptance. Cover real IME, screen readers, touch, system clipboard, and hardware/device behavior in a named release checklist when affected. Synthetic messages and WARP cannot establish these results. Do not alter normal user settings, clipboard, audio defaults, camera state, or application data through unattended fixtures.

When a test harness changes, rerun the previous implementation with that same harness before comparing. Use the existing independent [performance contract][dx-perf]; standalone reviewed evidence belongs in DxUi `Measurements`, while product measurements stay in the respective product repository.

## 6. Tests owned by each product

### 6.1 RedXe

Use [test.ps1][rx-test], `HostPluginTests`, `AVControlTests`, and `PluginContractTests` as the integration gate. Every RedXe project must support the six configurations in section 1.1. Build the host and every affected plugin/test against the same candidate pin and matching configuration, and verify that the launched binaries came from those outputs. Adopt the standalone Slider contract directly; product tests verify its wiring rather than restoring legacy semantics.

Protect tile and raised-overlay lifecycle; preview/commit/cancel; mouse/keyboard/touch routing; focus transfer; host TSF/clipboard/UIA bridges; clipping and DPI; page changes; resize and device loss; hidden/zero-size surface release; multiple widget instances; shutdown with outstanding callbacks/providers; and unchanged non-DxUi widgets. Keep host-owned scheduling and the COM/POD ABI intact. Test settings read/write and packaging using synthetic AV services and isolated data.

Compare the current pin and candidate with the same product and fixture inputs. Record real presented latency, total preparation plus composition, allocation/resource peaks, idle wake-ups, and module retention after UIA publication. Candidate checks must neither silently waive existing AV HOLD gates nor turn new unrelated AV features into an adoption prerequisite. Current unqualified capabilities stay explicitly unqualified.

### 6.2 RedSalamander

Reuse [Run-AllTests.ps1](../../../Tools/Run-AllTests.ps1), its stable suite/case identities, [test-enabled build evidence](../../Build/Build_Toolchain.md), and the current [validation evidence contract](../../Testing/Testing_ValidationEvidence.md). The external lock, imported build tools, source/content and artifact hashes must enter build attestation, `source.dxui`, runtime closure, affected-set selection, and receipt invalidation. An external source change must not reuse receipts for the old in-tree library.

| Product surface | Protected behavior |
|---|---|
| RedConfigure and Preferences | Draft/apply/cancel/persistence, validation, localized strings, theme preview, control events, keyboard traversal, real focus after modal close, UIA close/reopen. Adapt the alpha slider to the accepted standalone Slider contract; intended Slider visual/event differences are not migration regressions. |
| Main window / FolderView | Selection, navigation, rename/edit, menus/popups, drag/drop and cancellation, Find/Compare dialogs, offscreen selected-grid UIA semantics, large/virtualized models and existing performance budgets. |
| File Operations UI | Existing confirmation, pause/cancel and prompt focus, progress updates, popup lifetime and teardown. Keep current file/provider outcomes and identity rules unchanged. |
| ViewerText, Sqlite, Space, ImgRaw, PE, Web | Open/close and repeated reopen, shared controls/chrome, search/navigation, focus, menus, theme/DPI, malformed/large input and cancellation using existing isolated fixtures. |
| Terminal | Prompt/tool-window focus, launch/close, plugin lifetime and existing input behavior. Preserve the admitted engine pin and current Terminal ownership. |
| Monitor | Control/host lifecycle, theme/DPI, text/menu input, resource retention, and declared selftest availability in test-enabled builds. |
| Packages | Fresh portable extraction plus MSI/MSIX install/update checks as applicable, all dependent EXEs/plugins present, correct architecture and dependency identity, startup on minimum supported OS. |

During migration, keep each inherited case accounted for: shared runtime cases in DxUi, product/adapter cases in RedSalamander, or a reviewed exclusion with reason. Preserve case identity/provenance when moving tests. Convert old private-header/test-macro assumptions into library-owned tests or legitimate public diagnostics; do not enable consumer-only library flags to manufacture compatibility.

Use focused cases while developing, then a **Fresh Full** for admission/closeout, plus paired test-enabled Release performance with enforced budgets. Existing Debug Full is not Release performance proof. Require Debug, Release and ASAN Debug builds for every project on x64 and ARM64, with matching instrumented dependencies for ASAN and native ARM64 execution before claiming runtime qualification. Full with documented environment skips remains limited to those tested capabilities.

All product test data and runtime evidence must use the governed, already initialized, marker-owned `X:\RedSalamander.Perf` sandbox; `.build` holds build outputs and tool-generated inventories, not runtime test data. Archive reviewed receipts through the existing `Specs/TestRuns` workflow. Preserve the runner's profile leases, no-activation behavior, directed-input warning, and prohibition on killing independently launched applications.

### 6.3 Evidence in the existing consumer workflow

Use each repository's existing build/test receipts and the consumer PR description. Record the old/new DxUi SHA, tested consumer change, selected configurations, regression/performance results, and any required manual or environment gates. Add DxUi identity to existing build/test inputs where it is missing; no shared cross-repository receipt schema is needed.

Verify that a pin change rebuilds the library and affected consumer modules, invalidates incompatible prior results, and appears in existing build/package provenance. A green test of the old archive or a generic sample cannot qualify the updated product. Keep this check in the current runner/build machinery.

## 7. Implementation sequence and deliverables

The packaging decision is complete; implementation checkboxes remain open. Planning observations and sizing estimates do not complete migration or test gates.

| Slice / owner | Work and concrete deliverables | Exit gate / dependency |
|---|---|---|
| **C0 — consumer and library maintainers** | Use the completed source gap audit; retain its 58-file/test dispositions and extend them to runtime/module evidence. Route G1–G3/G7/G8 gaps and the accepted G4 implementation; G5 requires no legacy preservation and G6's standalone Slider is accepted. Characterize tooltip/localization/clipboard and existing Grid UIA gaps, name current consumer owners, and retain in-tree product/library baselines and first cost receipt. | No unmapped consumer or unexplained behavior difference. Baselines reproducible. Required before changing binaries. |
| **C1 — build owners across the three repositories** | Deliver Debug/Release/ASAN Debug × x64/ARM64 for every project, with matching instrumented DxUi dependencies, solution/CLI/CI support and isolated configuration identity. Validate exact pins, public imports, toolchain/CRT and rebuild behavior through existing entrypoints. | All six build configurations and sanitizer detection qualify; public consumer fixtures pass and reject mismatches/stale outputs. No DLL experiment or new binary package system. |
| **C2 — DxUi test owner** | Fill the tooltip-clock, public-helper and localized-text gaps; implement accepted G4 native capacity separation with bounded validation, explicit failure, one open attempt and unchanged embedded text limits; coordinate the shared Grid UIA fix and public consumer fixtures. Retain native/embedded resource gates, existing receipts and the library PR record. | Required standalone matrix and paired evidence pass; no consumer checkout dependency. Depends on C0 gap mapping. |
| **C3 — RedXe owner** | Establish/reuse product CI and current-versus-candidate tests. Change the lock on a consumer branch and rehearse an upgrade/rollback with existing artifact provenance, preserving AV release claims. | RedXe gate and rollback receipt; actual pin matches linked and packaged modules. Can proceed with C2 after C0. |
| **C4 — RedSalamander build/adapter owner** | Add product lock, isolated restore/imports, attested dependency closure, mandatory sanitizer matrix, theme/message/viewer adapters, and direct adoption of current DxUi diagnostics without legacy compatibility layers. Pilot the separately launched RedConfigure process plus RedConfigureTests; keep a retained legacy baseline. | Pilot behavior and performance pass, clean/relocated build works, legacy and external definitions cannot enter the same PE. Depends on C0/C2. |
| **C5 — RedSalamander UI/plugin owners** | Migrate remaining process families in bounded changes: Monitor, then the main host with its viewer/Terminal plugins, header-only ViewerVLC typography, and integration tests. Update imports/includes/test dispositions and verify complete runtime module identity. | Fresh product validation, paired Release/hardware evidence, packaging and applicable manual acceptance. Depends on C4; coordinate active UI owners before touching their surfaces. |
| **C6 — consumer build owners** | Add the single advisory availability check to normal builds/PR validation and document the manual lock-update/fix/retest loop. Rehearse same-pin, newer-green, pending/failed, divergent and offline notices. | Each integrated product can notice an update, adopt a fixed SHA after local tests/CI, reject a regression, and revert. Can begin with RedXe during C3; RedSalamander follows C5. |
| **C7 — RedSalamander and DxUi owners** | Remove the obsolete in-tree implementation/project only after every consumer has moved; reconcile durable specs/docs, case ledger and HOLD routing, archive this plan through repository closeout rules. | No duplicate shared implementation, all supported consumers accounted for, required evidence current, deferred items assigned explicitly. |

Migration granularity is **process/ownership compatible**, not an arbitrary file count. Do not mix old/new DxUi definitions in one PE. Before allowing old/new modules in one process, prove window-class registration, message/token registries, COM/provider lifetime and callback ownership cannot cross incompatible implementations. If that cannot be proved, switch the host and all its DxUi-using plugins atomically as one process family. A temporary compatibility header may forward to public APIs, but must have an owner/removal gate and contain no second control implementation.

Until removal, a necessary fix to the legacy tree must have an upstream disposition in the ledger. Shared code is maintained in DxUi; application adapters stay in the consumer. No ongoing bidirectional source-copy workflow is introduced.

Coordinate with RedSalamander I4 for test ownership, I5 for performance infrastructure, I9/I10 for CI/tooling, I15 for Terminal, I18 for Preferences/File Operations, and H0 for confirmed shared accessibility residuals. This plan does not reopen retired plans or take over their unrelated work. At activation, the DxUi migration HOLD record should become a routing record to this owner rather than a second execution queue.

## 8. Verification and closeout

### 8.1 Existing entrypoints to reuse

These are commands for future implementation qualification, **not tests executed while authoring this plan**. Run from the respective repository root with governed fixture roots and retained baseline paths. Inspect current help and drift before execution; add missing workflow capabilities through normal tooling governance.

```powershell
# DxUi: run matched baseline/candidate measurements as described in its performance contract.
.\validate-skills.ps1
.\validate-specs.ps1
.\validate-dependencies.ps1
.\format.ps1 -Check
.\test.ps1 -Configuration Debug -Platform x64 -PerformanceBaseline <retained-debug-receipt>
.\test.ps1 -Configuration Release -Platform x64 -PerformanceBaseline <retained-release-receipt>
.\build.ps1 -Configuration Debug -Platform ARM64
.\build.ps1 -Configuration Release -Platform ARM64
.\test-consumer.ps1
# Also deliver and run ASan Debug on x64/ARM64 through these same entrypoints.
# Run all three ARM64 test configurations on a native ARM64 runner; see section 1.1.
# For changed visuals: .\gallery.ps1 -PublishDocs, with visual review.

# RedXe: existing restore/build/test entrypoints, per selected candidate checkout.
.\restore-dxui.ps1 -Platform x64
.\test.ps1 -Configuration Debug -Platform x64
.\test.ps1 -Configuration Release -Platform x64
# Deliver and run all six configurations, including ASan Debug x64/ARM64.
# Execute all three ARM64 product-test configurations on a native ARM64 runner.

# RedSalamander: full product admission, including test-enabled build evidence.
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
.\Tools\Get-ToolInventory.ps1 -Validate
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -Configuration Debug -Platform x64
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -Configuration Release -Platform x64
# Deliver and run all six configurations, including ASan Debug x64/ARM64.
# Add enforced Release performance, native ARM64, package and manual gates.
```

The command examples above show existing entrypoints; they do not constitute the full matrix or claim missing ASAN CLI support is already implemented. Extend those entrypoints and run every section 1.1 configuration during implementation.

No Full run, runtime benchmark, or dependency upgrade is required to validate this documentation-only delivery. Validate its links and exact WIP index coverage with the spec inventory and inspect the diff. No compiled inputs or gallery visuals change. Native and paired resource gates apply when implementation begins.

### 8.2 Durable authorities to update during implementation

- **DxUi:** [build/consumption][dx-build], [testing][dx-testing], [performance][dx-perf], input/accessibility, controls/layout and hosting contracts where behavior changes; `capabilities.json`, public docs and gallery as required. Update supported/pending status only with evidence.
- **RedXe:** [DxUi integration][rx-integration], performance/resources, plugin ABI and AV contracts, build/test guidance, and end-user docs if behavior changes. Retain the existing AV owner for unmet hardware/human gates.
- **RedSalamander:** [toolchain](../../Build/Build_Toolchain.md), [shared helpers](../../Core/Core_SharedHelpers.md), [DxUi design](../../UI/UI_DxUiWinUIDesign.md), [shared grid](../../UI/UI_DxUiSharedGrid.md), [test coverage](../../Testing/Testing_TestCoverage.md), [selftests](../../Testing/Testing_SelfTests.md), [performance](../../Testing/Testing_PerformanceValidation.md), [validation evidence](../../Testing/Testing_ValidationEvidence.md), [tooling governance](../../Testing/Testing_ToolingGovernance.md), affected UI/plugin and installer contracts, `AGENTS.md`, and developer documentation. Add a dedicated consumption contract if needed, with consistency-ledger anchors.

Any new `Tools/` command/module requires inventory, help, caller, focused-test, cache-closure and normative updates. On final closeout, update `Specs/NormativeConsistency.json`, move this plan to Done, remove its active row, and add the exact new Done path and positive/negative admission test required by [plan lifecycle policy](../../README.md). Do not modify frozen Done history.

### 8.3 Completion checklist and stop conditions

The authoritative progress checklist is at the beginning of this document.

Stop promotion of the affected slice for an unexplained behavior difference, missing/changed baseline, confirmed resource regression, untested required capability, stale/mismatched evidence, incompatible module ownership/ABI, or an occupied output owned by an independently launched application. Report the exact missing evidence and preserve the previous pin. Optimize, narrow the slice, repair the regression, or obtain advice for a measured tradeoff; never silently rebaseline or classify a failing required test as optional. Independent planning and unaffected work can continue.

[dx-build]: https://github.com/RedSalamanders/DxUi/blob/d192e474e540adc2656e69b8e8150b0f7a05d63f/Specs/Build/Build_ToolchainAndConsumption.md
[dx-project]: https://github.com/RedSalamanders/DxUi/blob/d192e474e540adc2656e69b8e8150b0f7a05d63f/src/DxUi.vcxproj
[dx-capabilities]: https://github.com/RedSalamanders/DxUi/blob/d192e474e540adc2656e69b8e8150b0f7a05d63f/capabilities.json
[dx-origin]: https://github.com/RedSalamanders/DxUi/blob/d192e474e540adc2656e69b8e8150b0f7a05d63f/Specs/Done/SourceImport/source-origin.json
[dx-testing]: https://github.com/RedSalamanders/DxUi/blob/d192e474e540adc2656e69b8e8150b0f7a05d63f/Specs/Testing/Testing_Validation.md
[dx-ci]: https://github.com/RedSalamanders/DxUi/blob/d192e474e540adc2656e69b8e8150b0f7a05d63f/.github/workflows/ci.yml
[dx-perf]: https://github.com/RedSalamanders/DxUi/blob/d192e474e540adc2656e69b8e8150b0f7a05d63f/Specs/Core/Core_PerformanceAndResources.md
[dx-rs-hold]: https://github.com/RedSalamanders/DxUi/blob/d192e474e540adc2656e69b8e8150b0f7a05d63f/Specs/Plans/WIP/RedSalamanderMigration.md
[rx-lock]: https://github.com/RedSalamanders/RedXe/blob/75545887d002b5075d170d98b7bfd3a77d8d7f61/Dependencies/DxUi.lock.json
[rx-integration]: https://github.com/RedSalamanders/RedXe/blob/75545887d002b5075d170d98b7bfd3a77d8d7f61/Specs/Core/Core_DxUiIntegration.md
[rx-restore]: https://github.com/RedSalamanders/RedXe/blob/75545887d002b5075d170d98b7bfd3a77d8d7f61/restore-dxui.ps1
[rx-av-plan]: https://github.com/RedSalamanders/RedXe/blob/75545887d002b5075d170d98b7bfd3a77d8d7f61/Specs/Plans/WIP/RFC_Plugins_AVControl.md
[rx-test]: https://github.com/RedSalamanders/RedXe/blob/75545887d002b5075d170d98b7bfd3a77d8d7f61/test.ps1

Implementation checkpoint: the pilot's root build and module attestation pass, and its full focused
runtime suite passes. The first Japanese test correctly exposed two nested TagPicker inputs; the
product now supplies localization through the public logical control tree. All native project
configuration declarations are present; instrumentation and native execution are qualified separately.
The original Debug baseline recorded three Preferences page-settling failures and one Pester failure;
those precede the dependency switch and are retained for candidate comparison.

Baseline package check: clean-extraction PluginContractTests failed for the attempted original x64 Release
package. Publication correctly discarded the failed staging archive and preserved the existing July package.
That existing archive is not the current baseline; do not use its size or contents as adoption evidence.
The failure log is retained; diagnose the original assertion before claiming rollback qualification.

First RedSalamander native CI failed before compilation: Git rejected long paths in historical test evidence.
Enable long paths before checkout in all CI jobs and the reusable build. Native qualification remains open.

The existing CI auto-format job changed 341 files, including archived evidence. Its exact automated commit is
reverted in the follow-up; CI now checks only changed owned C++ files with a pinned formatter and read-only access.
The local Full/Fresh run remains isolated at 83b703b7 while CI infrastructure changes are validated separately.

Package blocker diagnosed: TestSupport tried to find a checkout before reading the explicitly marked root.
A clean extraction has no checkout, so the existing local-writer package probe failed before testing the provider.
The helper now validates the explicit root first and discovers a checkout only for its default; package runtime
verification remains open. DxUi 3dae073 passes all 18 local suites in Debug, Release and ASan Debug.

The relocated native sandbox regression passes x64 Debug, Release and ASan Debug; the same ASAN fixture exits 12 with the original header, proving the checkout-discovery defect. Full package verification remains open.

RedSalamander CI reaches dependency restore after the long-path/formatter fixes. Its old vcpkg registry cannot resolve the already-requested BLAKE3 1.8.7. Both registry/tool pins advance to 11159b9 (the first BLAKE3 1.8.7 update); all twelve requested versions and the exact AWS override exist there. Dependency floors and the AWS override are unchanged. Native rebuilds qualify the resulting dependency graph.
