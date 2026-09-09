# DxUi shared-library adoption and release plan

Status: **ACTIVE, I19 — planning complete; implementation not started**.

Priority: **P1**. Effort: **L**, delivered in bounded changes. Risk: **High** for RedSalamander migration and any DLL ABI change; **Medium** for dependency automation.

Owner: this plan owns the cross-repository adoption sequence, dependency qualification, and RedSalamander migration. DxUi owns shared controls and library contracts; each product owns its adapters, behavior, packaging, and release decision. The request authorizes this WIP plan, not implementation or publication.

Planned on 2026-09-09 against these clean local checkouts. These are audit identities, not new dependency pins:

| Repository | Local checkout | Audited commit |
|---|---|---|
| [DxUi](https://github.com/RedSalamanders/DxUi) | `Z:\src\DxUi` | `d192e474e540adc2656e69b8e8150b0f7a05d63f` |
| RedXe | `Z:\src\RedXe` | `75545887d002b5075d170d98b7bfd3a77d8d7f61` |
| RedSalamander | `Z:\src\RedSalamander` | `3f0df2f2c28239eb658a68f85b9c7f69c2356af4` |

This is a **non-normative proposal**. Existing contracts remain in force until implementation, tests, and authoritative updates land. Local paths are discovery context; build automation must support relocated checkouts.

## 1. Recommended direction

1. Keep **one canonical DxUi implementation**, built from its repository. Extend RedXe's existing exact-pin consumption; migrate RedSalamander from its in-tree implementation.
2. Retain **`DxUi.lib` as the first delivery format**. Static linking is not inherently too expensive. Measure build time, shipped code duplication, startup, and memory before deciding whether a DLL earns its additional ABI and deployment work.
3. Make updates routine: qualify an immutable library revision, automatically propose a lock update in each consumer, run that product's tests, then merge and release independently. A library release must never silently change an installed product or a developer's sibling checkout.
4. Use **two independent compatibility gates**: library behavior tests in DxUi, and candidate-versus-current product tests in RedXe and RedSalamander. Neither replaces the other.
5. Preserve a tested previous pin and product package for rollback. Track which DxUi revision each built and shipped module actually contains; an updated source lock alone is insufficient evidence of adoption.

The first useful milestone is a repeatable RedXe candidate-upgrade rehearsal plus a RedSalamander consumer/drift inventory and retained baseline. RedSalamander adoption must not wait for a speculative DLL redesign.

## 2. Current state and gaps

| Evidence inspected | Current fact | Required follow-up |
|---|---|---|
| DxUi [consumption contract][dx-build], [project][dx-project], and [capabilities][dx-capabilities] | One static target, public headers, exact API/target validation, consumer MSBuild imports, isolated outputs. Native and embedded mechanisms already exist. | Reuse these entrypoints; do not repeat extraction or introduce consumer-maintained source lists. |
| RedXe [lock][rx-lock], [integration contract][rx-integration], [restore][rx-restore], and `Build/RedXe.DxUi.props` / `.targets` | RedXe already pins the audited DxUi HEAD. `RedXe`, `AVControl`, and `AVControlTests` reference the archive. Host and plugin keep DxUi C++ ownership inside their respective modules. | Add automated candidate qualification and upgrade proposals; preserve existing host/plugin ownership. No initial integration rewrite is needed. |
| RedXe [AV plan][rx-av-plan] | Synthetic text/UIA and WARP tests exist. Real IME, screen-reader/touch, matched text/UIA resource acceptance, and AV release gates remain open. No checked-in `.github/workflows` YAML was found in this checkout. | Establish product CI around existing tests; confirm actual repository check configuration. Keep AV release work with its existing owner. |
| RedSalamander [legacy project](../../../Common/DxUi/DxUi.vcxproj) and its project references | Production references found in RedSalamander, RedSalamanderMonitor, RedConfigure, ViewerText, ViewerSqlite, ViewerSpace, ViewerImgRaw, ViewerPE, ViewerWeb, and Terminal; test references in DxUiTests and RedConfigureTests. | Inventory direct references, transitive links, header-only use, app-specific helpers, and actual loaded modules. This initial list is not a completeness proof. |
| RedSalamander [legacy public header](../../../Common/DxUi/DxUi.h) | Includes the product's `PlugInterfaces/Viewer.h`; implementation also uses consumer `WindowMessages.h`. | Replace consumer coupling with adapters; never move viewer/plugin contracts into DxUi. |
| DxUi [origin record][dx-origin] and RedSalamander Git history | Extraction records a September 5 source baseline. Subsequent RedSalamander changes include tooltip clock handling and test updates in `817b7e2d`, `f0710f9b`, and `3f0df2f2`. | Compare both histories and behavior. Determine whether each change is already present, needs an upstream regression/fix, or is product-only. Do not overwrite either tree or blindly cherry-pick. |
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

## 3. Static library versus DLL decision

### 3.1 Costs to distinguish

The `.lib` file size is not the amount added to a product or loaded into RAM. Linker reachability and optimization determine retained code. Inspect effective Release flags and linked images: `/OPT:REF` can discard unused packaged functions/data and `/OPT:ICF` can fold identical COMDATs. Do not assume these options are active solely because the configuration is named Release, or enable address-folding without checking code that relies on function identity. [Microsoft linker documentation](https://learn.microsoft.com/en-us/cpp/build/reference/opt-optimizations?view=msvc-170).

Static DxUi also does **not** require a static CRT: the current library uses the DLL CRT (`/MDd` or `/MD`). A DLL can reduce repeated code across modules and eligible shared image pages across processes, but introduces a separate runtime dependency. Product-local DLL copies and different versions are not an automatic system-wide memory saving. A compatible DLL can be replaced without relinking; that does not establish behavioral compatibility or make untested replacement safe. [Microsoft dynamic-linking documentation](https://learn.microsoft.com/en-us/windows/win32/dlls/advantages-of-dynamic-linking).

| Option | Benefits | Costs / limitations | Proposed disposition |
|---|---|---|---|
| Source-pinned `DxUi.lib` | Existing supported path; exact product build identity; simple packaging; ordinary C++ API within each module. | Each linked EXE/DLL can contain its own used DxUi code and module state. Updates need relinking and shipping the product. | Default for adoption. |
| Cached or prebuilt `DxUi.lib` plus matching headers | Reduces repeated compilation while preserving static deployment. | Must match source, toolchain, CRT, flags, and dependencies; cache integrity and symbols need governance. | First optimization if build time is the problem; source rebuild remains available. |
| Product-local `DxUi.dll` with a matched C++ interface | Potentially removes duplicated code inside a product; implementation-only edits may reduce relinking. | Export discipline, class layout, inline/template code, CRT/STL agreement, loader/package/unload tests; still a coordinated binary set. | Valid measured experiment, not supported today and not a stable independent ABI. |
| `DxUi.dll` behind versioned C handles or COM interfaces | Can support an explicit long-lived binary contract and factory-owned lifetime. | Substantial facade, capability/version negotiation, error/ownership rules, compatibility fixtures, maintenance cost. | Only if independent binary servicing or measured duplication justifies it. |

Matching CRTs alone does not solve C++ object layout, allocator ownership, exception, or callback-lifetime compatibility. A DLL design should keep allocation and destruction with the owner and avoid STL or DxUi control ownership crossing a stable ABI. The existing RedXe COM/POD plugin boundary remains intact. [Microsoft CRT boundary guidance](https://learn.microsoft.com/en-us/cpp/c-runtime-library/potential-errors-passing-crt-objects-across-dll-boundaries?view=msvc-170).

For binary reuse, record actual compiler/linker versions, not merely `v145`. MSVC has documented binary compatibility restrictions; `/GL` and `/LTCG` inputs require exactly matching build tools, including minor versions. `stdcpplatest`, public layouts, build macros, and iterator-debug settings make conservative exact matching appropriate here. [Microsoft binary compatibility guidance](https://learn.microsoft.com/en-us/cpp/porting/binary-compat-2015-2017?view=msvc-170).

### 3.2 Required experiment and decision receipt

Before implementation, retain the existing static baseline. First audit Release dead-code removal, unnecessary references, translation-unit granularity, and incremental-build invalidation. Do not introduce `/WHOLEARCHIVE` or a new binary package system to solve an unmeasured problem.

Measure the same scenario, fixtures, source behavior, architecture, configuration, toolchain, hardware, and power policy serially. If a DLL prototype is needed, keep its API changes isolated from functional improvements and measure the complete EXE/plugin/DLL set.

| Cost | Measurement |
|---|---|
| Developer/CI time | Cold restore and full build; warm no-op build; product-only edit; private DxUi `.cpp` edit; public-header edit. Separate restore, compile, link, and test wall time. |
| Disk/distribution | Per-module PE sections and linker map attribution; total installed bytes and compressed package bytes; debug symbols reported separately. Compare actual loaded subsets as well as all shipped plugins. |
| Memory | Private bytes, private/shareable working set, committed image pages where measurable, GPU/cache/surface resources, steady/peak/end values. Do not sum shared working sets as if each were private. |
| Responsiveness | Cold/warm startup and first control, first plugin load, input-to-visible latency, preparation/composition and presented frame percentiles. |
| Lifecycle | Multiple views/plugins, hidden/minimized idle CPU and wake-ups, theme/DPI changes, repeated open/close, device loss, UIA references surviving window close, module shutdown. |

Use RedXe's host plus AVControl and a representative RedSalamander session with viewer plugins, Monitor, and RedConfigure. Include single-product and simultaneous-product runs when claiming cross-process savings. Count per-module registries, caches, and callback/image lifetime: neither linkage choice automatically shares device resources safely.

**Decision rule:** keep static unless repeatable measurements show a meaningful benefit against budgets agreed before the experiment, with no unresolved behavior/performance regression and acceptable implementation/support cost. Missing paired evidence means **undecided**, not DLL approval. Record deltas, uncertainty, fixture identities, and rationale. Existing DxUi investigation bands are noise checks, not allowed degradation; confirmed regressions require developer advice under its performance contract.

If DLL is selected, first revise the DxUi consumption contract and validators that currently require the single static target. Specify exports/version negotiation, thread affinity, reentrancy, factory destruction, module pinning for outstanding UIA/COM references, and safe initialization outside loader-lock work. Use product-local staging and atomic package upgrades; no global shared install or runtime self-updater. Test missing/wrong-architecture/wrong-version DLLs, install/update/rollback, and old-client compatibility only to the extent explicitly promised. RedSalamander staging must use [RuntimeDependencies.props](../../../RuntimeDependencies.props) and its existing portable/MSI/MSIX paths.

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

One resource opportunity deserves a separate experiment: [WindowHost graphics sharing](../../../Common/DxUi/DxUi.WindowHost.cpp) stores D3D11/D2D/DWrite resources in a static map keyed by thread ID. Static linkage gives each module its own map. A common DLL could consolidate compatible same-thread buckets across host/plugins, avoiding some duplicated device/context/factory resources. It would not share buckets across different threads/processes or remove per-window back buffers. No driver-memory saving is priced into the estimates above. As scale context only, two 1920×1080 BGRA8 buffers contain **15.82 MiB** of pixels before driver overhead, regardless of library linkage; this is not a claim that every current control allocates full-screen buffers.

**Initial opinion:** Release static linking is affordable at the observed scale. A DLL can plausibly save a few MiB on disk and some incremental link work, but the 121 MiB archive is not a reason to redesign the ABI. Keep the static adoption path; prioritize a DLL prototype only if installation size, iteration time, or measured duplicate graphics pools is an actual problem. These estimates do not waive the paired measurement gate.

## 4. Dependency lifecycle and release workflow

### 4.1 Dependency identity and developer experience

- Keep one reviewed `Dependencies/DxUi.lock.json` per product. The existing lock shape is authoritative for current fields; proposed metadata additions require schema/validator/tool updates together. Exact source SHA remains the restore authority; human-readable release tags and release notes are additional labels.
- Resolve the canonical repository to an isolated checkout and isolated build/vcpkg outputs. The current RedXe restore can seed from a sibling revision without modifying it. Reuse this behavior and the library's `Build/DxUi.Consumer.props` / `.targets`; never enumerate DxUi `.cpp` files in consumers.
- Extend the evaluated fingerprint to bind full source/content identity, public headers, API revision, compiler/linker, SDK, architecture, configuration, CRT, iterator-debug level, sanitizer/LTO flags, WIL/vcpkg locks, and imported build tooling. A string containing only a major toolset is not sufficient binary-cache identity.
- Provide an explicit development-candidate mode that selects a local DxUi tree into isolated outputs and records dirty/untracked content. It must be visibly non-release and must not modify the release lock. Clean exact-pin candidate validation remains the release path.
- Keep Debug/Release, x64/ARM64, test-enabled/product, and sanitizer identities distinct. RedSalamander has ASan configurations absent from the current standalone project: qualify a real supported mapping/build mode before migration, never silently map ASan to unsanitized Debug.
- Start with source restore and incremental builds. Add an immutable content-addressed binary cache only after timings justify it; validate matching headers, archive, symbols, licenses, manifest and hashes before use. Cache misses rebuild from the pin; mismatches fail rather than falling back to an unverified binary.
- Test missing/mismatched/dirty pins, wrong repository/API/architecture/CRT, moved roots and spaces, short fixture paths, parallel product builds, interrupted restore, warm offline restore, corrupt cache, and stale outputs. Warm offline success must require a complete verified dependency set.

### 4.2 Proposed update sequence

The sequence below describes future CI and bot behavior; no workflow, scheduled task, repository write, or release is created by this plan.

1. **Library change:** a DxUi PR declares API/behavior/visual changes and affected capabilities. It carries a regression test, release-note/migration entry, docs/gallery review, and retained performance baseline where applicable. Ordinary behavioral fixes should not silently change a product default.
2. **Independent qualification:** DxUi runs its required native/test/consumer-fixture matrix without either application checkout. Publish immutable receipts tied to the exact candidate SHA and fixture/build identities.
3. **Downstream preview:** dedicated product jobs test that candidate against the current integration branches and the maintained release baselines in a small reviewed consumer matrix. Each job checks out an exact product SHA and substitutes the candidate lock only in its disposable checkout. Product tests remain owned and executed in that product's workflow.
4. **Promotion:** classify the candidate as library-qualified, RedXe-qualified, RedSalamander-qualified, or blocked. Before RedSalamander migration its consumer result is explicitly **not integrated**, never a compatibility pass. Supported consumers must pass before marking a revision generally qualified; an intentional breaking change requires a coordinated migration entry and branch policy. A library-only result can still be retained for development.
5. **Product PRs:** after a qualified immutable revision is available, an updater opens or refreshes one lock-update PR in each integrated consumer. Include old/new SHA, API/behavior changes, tested product SHA, evidence links, skips/manual gates, and rollback pin. No floating `main` in product builds and no unattended merging initially.
6. **Consumer admission:** run the product gate on the actual lock-update change. Re-run or prove receipt identity after any source/workflow/base change. Evidence for a PR head does not automatically qualify a different squash/merge SHA. Serialize updates per consumer branch and supersede stale proposals explicitly.
7. **Release verification:** stamp the actual DxUi source/build identity into module diagnostics and package provenance; inspect the linked artifacts and package, then ship using the product's existing release process. Both products may ship different qualified pins.
8. **Recovery:** revert the lock and associated adapters if necessary, rebuild/requalify, and redeploy the retained product package. Keep old source/dependency artifacts and symbols obtainable. Do not overwrite a running DLL or assume settings changes are reversible; pair any schema change with its own downgrade policy.

Use explicit opt-in PR dispatch for early downstream preview, automatic proposals when qualification completes, and a scheduled reconciliation check to catch missed events. The check records current pins, latest qualified revision, open update/failure, and time since qualification. Proposed operating target: open a proposal within one working day; triage failed or stale proposals within two; critical fixes get immediate owner review. These are proposed team targets, not an automation already installed.

Prefer a narrowly scoped GitHub App or equivalent repository-scoped credentials for cross-repository dispatch and PR creation; default repository tokens must not be assumed to grant cross-repository access. Keep PR-writing credentials out of candidate build/test jobs, restrict trusted workflow execution, and bind results to the requested commits. Cache keys and receipts include the complete build/workflow input closure. Reuse current CI runners and test entrypoints, with a distinct native ARM64 lane; do not substitute cross-compilation.

### 4.3 Version and support policy

- Define source/API compatibility and observable-behavior compatibility separately. Static consumers require recompilation; matching API revision does not prove unchanged layout, focus, events, accessibility, or performance.
- Propose release labels with patch/minor/major intent: corrective behavior within contract, compatible additions, and breaking changes. The lock still pins a commit; a label never bypasses tests.
- Retain the current qualified line and the previous supported line during a breaking migration, with an explicit sunset criterion. A deprecation needs a replacement, consumer inventory, migration note, and removal gate after supported consumers have moved.
- Before promising a release, define the supported Windows, compiler, architecture, and feature matrix. Preserve RedXe's Windows 10 baseline and RedSalamander's Windows 11 policy; gate optional OS features. A build with newer Windows imports is not proof it runs on the minimum OS.
- Ship dependency notices, provenance, symbols, and a source-to-package map. Track failed-upgrade rate, time to adoption, CI time, and rollback time; use them to simplify the workflow after the first upgrades.

## 5. Tests owned by DxUi

Extend the existing suites and fixtures; do not replace them with a new framework. Map each supported capability and every inherited test disposition to an owner. Shared behavior discovered through a consumer defect should gain a minimal synthetic reproduction in DxUi while the product retains its end-to-end regression.

| Layer | Required evidence |
|---|---|
| Public API/build boundary | Compile/link/run consumers using public headers only; exact-pin negative cases; no application includes, consumer source lists, or incompatible flags. Preserve the standalone public samples and relocated fixture. |
| Control semantics | State/layout, event order and counts, preview/commit/cancel, validation/read-only/disabled behavior, keyboard traversal, pointer capture/cancel, nested/reentrant callbacks, deletion during callbacks, Unicode editing, grids/trees and virtualization. |
| Native hosting/input/UIA | HWND lifetime, nested modal/popups, focus restoration, DPI, IME/TSF text-store revisions, clipboard abstraction, stale/disconnected UIA providers, selected offscreen items, destruction and queued callbacks. |
| Embedded rendering | Supplied device/context, hostile pipeline state, clean/dirty/hidden/zero-size states, clipping/alpha, multiple views, failed preparation, device loss/recovery, independent pools and generation teardown. |
| Visual contract | Existing test baselines plus gallery coverage of relevant themes, high contrast, reduced motion, DPI and states. Review intended pixel changes; semantic assertions remain mandatory. |
| Resource contract | Paired complex-UI baseline/candidate, preparation versus composition, zero clean-composition allocations, cache/surface/queue bounds, hidden idle work, failure/recovery peaks and retention. Hardware presentation remains a separate gate. |
| Test/tool integrity | Case registration/disposition accounting, result/skip parsing, negative fixtures, test-owned data paths, and detection of missing suites or forged/stale receipts. Run ASan/fault/stress coverage where supported. |

Library CI must retain x64 Debug/Release execution and ARM64 Debug/Release builds and native execution. The existing external-consumer fixture currently selects x64; extend architecture coverage where promised instead of reporting an unexecuted ARM64 fixture as green. Required unavailable capabilities produce a blocked gate with a reason. Optional skips remain visible and cannot erase a formerly failing test.

For native tests that need an interactive desktop, reserve that resource and keep deterministic fixtures separate from human acceptance. Cover real IME, screen readers, touch, system clipboard, and hardware/device behavior in a named release checklist when affected. Synthetic messages and WARP cannot establish these results. Do not alter normal user settings, clipboard, audio defaults, camera state, or application data through unattended fixtures.

When a test harness changes, rerun the previous implementation with that same harness before comparing. Use the existing independent [performance contract][dx-perf]; standalone reviewed evidence belongs in DxUi `Measurements`, while product measurements stay in the respective product repository.

## 6. Tests owned by each product

### 6.1 RedXe

Use [test.ps1][rx-test], `HostPluginTests`, `AVControlTests`, and `PluginContractTests` as the integration gate. Build the host and every affected plugin/test against the same candidate pin and verify that the launched binaries came from those outputs.

Protect tile and raised-overlay lifecycle; preview/commit/cancel; mouse/keyboard/touch routing; focus transfer; host TSF/clipboard/UIA bridges; clipping and DPI; page changes; resize and device loss; hidden/zero-size surface release; multiple widget instances; shutdown with outstanding callbacks/providers; and unchanged non-DxUi widgets. Keep host-owned scheduling and the COM/POD ABI intact. Test settings read/write and packaging using synthetic AV services and isolated data.

Compare the current pin and candidate with the same product and fixture inputs. Record real presented latency, total preparation plus composition, allocation/resource peaks, idle wake-ups, and module retention after UIA publication. Candidate checks must neither silently waive existing AV HOLD gates nor turn new unrelated AV features into an adoption prerequisite. Current unqualified capabilities stay explicitly unqualified.

### 6.2 RedSalamander

Reuse [Run-AllTests.ps1](../../../Tools/Run-AllTests.ps1), its stable suite/case identities, [test-enabled build evidence](../../Build/Build_Toolchain.md), and the current [validation evidence contract](../../Testing/Testing_ValidationEvidence.md). The external lock, imported build tools, source/content and artifact hashes must enter build attestation, `source.dxui`, runtime closure, affected-set selection, and receipt invalidation. An external source change must not reuse receipts for the old in-tree library.

| Product surface | Protected behavior |
|---|---|
| RedConfigure and Preferences | Draft/apply/cancel/persistence, validation, localized strings, theme preview, control events, keyboard traversal, real focus after modal close, UIA close/reopen. |
| Main window / FolderView | Selection, navigation, rename/edit, menus/popups, drag/drop and cancellation, Find/Compare dialogs, offscreen selected-grid UIA semantics, large/virtualized models and existing performance budgets. |
| File Operations UI | Existing confirmation, pause/cancel and prompt focus, progress updates, popup lifetime and teardown. Keep current file/provider outcomes and identity rules unchanged. |
| ViewerText, Sqlite, Space, ImgRaw, PE, Web | Open/close and repeated reopen, shared controls/chrome, search/navigation, focus, menus, theme/DPI, malformed/large input and cancellation using existing isolated fixtures. |
| Terminal | Prompt/tool-window focus, launch/close, plugin lifetime and existing input behavior. Preserve the admitted engine pin and current Terminal ownership. |
| Monitor | Control/host lifecycle, theme/DPI, text/menu input, resource retention, and declared selftest availability in test-enabled builds. |
| Packages | Fresh portable extraction plus MSI/MSIX install/update checks as applicable, all dependent EXEs/plugins present, correct architecture and dependency identity, startup on minimum supported OS. |

During migration, keep each inherited case accounted for: shared runtime cases in DxUi, product/adapter cases in RedSalamander, or a reviewed exclusion with reason. Preserve case identity/provenance when moving tests. Convert old private-header/test-macro assumptions into library-owned tests or legitimate public diagnostics; do not enable consumer-only library flags to manufacture compatibility.

Use focused cases while developing, then a **Fresh Full** for admission/closeout, plus paired test-enabled Release performance with enforced budgets. Existing Debug Full is not Release performance proof. Preserve x64 Debug/Release, ARM64 build coverage and native execution before claiming runtime support, and qualified ASan coverage. Full with documented environment skips remains limited to those tested capabilities.

All product test data and runtime evidence must use the governed, already initialized, marker-owned `X:\RedSalamander.Perf` sandbox; `.build` holds build outputs and tool-generated inventories, not runtime test data. Archive reviewed receipts through the existing `Specs/TestRuns` workflow. Preserve the runner's profile leases, no-activation behavior, directed-input warning, and prohibition on killing independently launched applications.

### 6.3 Shared qualification receipt

Define a small machine-readable receipt that references existing raw results rather than duplicating them. It should bind library/product source identities, dirty state, lock digest, build/configuration/toolchain/dependency identities, tested executable/module/package hashes, exact test plan and fixture/baseline identities, OS/hardware, pass/fail/skipped/blocked counts and reasons, performance comparison, manual acceptance, and artifact locations.

Reject missing/stale/incompatible evidence. Tests should prove that changing the lock, adapter, library binary, build import, fixture, or workflow invalidates reuse. The package manifest must show that every module intended to consume DxUi uses the admitted revision; also detect accidental inclusion of the legacy implementation. A green generic sample cannot qualify the product package.

## 7. Implementation sequence and deliverables

All implementation checkboxes start open. Planning observations in section 2 do not complete these gates.

| Slice / owner | Work and concrete deliverables | Exit gate / dependency |
|---|---|---|
| **C0 — consumer and library maintainers** | Complete consumer/module/header/service/test inventory; reconcile post-extraction changes and known accessibility/tooltip issues; name supported consumer branches and owners; retain old-pin/in-tree product and library baselines. Produce a migration/test-disposition ledger and first cost receipt. | No unmapped consumer or unexplained behavior difference. Baselines reproducible. Required before changing binaries. |
| **C1 — DxUi build owner** | Audit static build/link cost; implement only justified cache/incremental fixes; perform bounded DLL experiment only if C0 shows a problem. Publish static/DLL decision and compatibility policy. | Measured decision and authoritative contract updates for any chosen new packaging. Static migration may proceed while an optional DLL investigation remains deferred. |
| **C2 — DxUi test owner** | Fill identified public API, native/embedded, resource, drift-regression and consumer-fixture gaps; define qualification receipt and release notes. | Required standalone matrix and paired evidence pass; no consumer checkout dependency. Depends on C0 gap mapping. |
| **C3 — RedXe owner** | Establish/reuse product CI, candidate override, exact artifact provenance, and current-versus-candidate tests. Rehearse an upgrade and rollback without changing existing AV release claims. | RedXe gate and rollback receipt; actual pin matches linked and packaged modules. Can proceed with C2 after C0. |
| **C4 — RedSalamander build/adapter owner** | Add product lock, isolated restore/imports, attested dependency closure, sanitizer strategy, theme/logging/message/viewer adapters and candidate mode. Pilot the separately launched RedConfigure process plus RedConfigureTests; keep a retained legacy baseline. | Pilot behavior and performance pass, clean/relocated build works, legacy and external definitions cannot enter the same PE. Depends on C0/C2. |
| **C5 — RedSalamander UI/plugin owners** | Migrate remaining process families in bounded changes: Monitor, then the main host with its viewer/Terminal plugins and integration tests. Update imports/includes/test dispositions and verify complete runtime module identity. | Fresh product validation, paired Release/hardware evidence, packaging and applicable manual acceptance. Depends on C4; coordinate active UI owners before touching their surfaces. |
| **C6 — repository release owners** | Add cross-repository candidate jobs and qualified-revision lock PRs, missed-event reconciliation, support/deprecation rules and rollback runbook. Rehearse both success and rejected-candidate paths. | One admitted update per integrated product, one deliberately rejected regression/stale-evidence case, and successful rollback rehearsal. Initial automation can begin after C3; full acceptance depends on C5. |
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
# Run the matching ARM64 test.ps1 configurations on a native ARM64 runner.
# For changed visuals: .\gallery.ps1 -PublishDocs, with visual review.

# RedXe: existing restore/build/test entrypoints, per selected candidate checkout.
.\restore-dxui.ps1 -Platform x64
.\test.ps1 -Configuration Debug -Platform x64
.\test.ps1 -Configuration Release -Platform x64
# Build ARM64 Debug/Release; execute product tests on a native ARM64 runner.

# RedSalamander: full product admission, including test-enabled build evidence.
.\Tools\Get-SpecInventory.ps1 -FailOnFindings
.\Tools\Get-ToolInventory.ps1 -Validate
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -Configuration Debug -Platform x64
.\Tools\Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -Configuration Release -Platform x64
# Add the current enforced Release performance, ARM64/native, ASan, package and manual gates.
```

No Full run, runtime benchmark, DLL prototype, or upgrade is required to validate this documentation-only delivery. Validate its links and exact WIP index coverage with the spec inventory and inspect the diff. No compiled inputs or gallery visuals change. Native and paired resource gates apply when implementation begins.

### 8.2 Durable authorities to update during implementation

- **DxUi:** [build/consumption][dx-build], [testing][dx-testing], [performance][dx-perf], input/accessibility, controls/layout and hosting contracts where behavior changes; `capabilities.json`, public docs and gallery as required. Update supported/pending status only with evidence.
- **RedXe:** [DxUi integration][rx-integration], performance/resources, plugin ABI and AV contracts, build/test guidance, and end-user docs if behavior changes. Retain the existing AV owner for unmet hardware/human gates.
- **RedSalamander:** [toolchain](../../Build/Build_Toolchain.md), [shared helpers](../../Core/Core_SharedHelpers.md), [DxUi design](../../UI/UI_DxUiWinUIDesign.md), [shared grid](../../UI/UI_DxUiSharedGrid.md), [test coverage](../../Testing/Testing_TestCoverage.md), [selftests](../../Testing/Testing_SelfTests.md), [performance](../../Testing/Testing_PerformanceValidation.md), [validation evidence](../../Testing/Testing_ValidationEvidence.md), [tooling governance](../../Testing/Testing_ToolingGovernance.md), affected UI/plugin and installer contracts, `AGENTS.md`, and developer documentation. Add a dedicated consumption contract if needed, with consistency-ledger anchors.

Any new `Tools/` command/module requires inventory, help, caller, focused-test, cache-closure and normative updates. On final closeout, update `Specs/NormativeConsistency.json`, move this plan to Done, remove its active row, and add the exact new Done path and positive/negative admission test required by [plan lifecycle policy](../../README.md). Do not modify frozen Done history.

### 8.3 Completion checklist and stop conditions

- [ ] C0 inventory, source drift, test dispositions, and retained baselines complete.
- [ ] Static/DLL decision measured and recorded; any deferred DLL exploration has explicit rationale.
- [ ] One canonical library implementation and public consumption path; all product modules/adapters accounted for.
- [ ] DxUi standalone matrix and both product candidate gates provide current, traceable evidence.
- [ ] External dependency changes invalidate build/test receipt reuse and are visible in shipped provenance.
- [ ] RedSalamander migration preserves interaction, accessibility, localization, settings, rendering and resource behavior.
- [ ] Upgrade proposals, failed-candidate handling, supported-version policy and rollback are rehearsed.
- [ ] Required hardware/manual/native-platform gates pass; any out-of-scope capabilities remain explicitly unqualified under their existing owner.
- [ ] Legacy implementation removed only after complete migration; docs/specs/tooling/index closeout complete.

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
