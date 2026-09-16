# DxUi codebase gap analysis — 2026-09-09

Supporting audit for [I19: DxUi shared-library adoption](../DxUi_SharedLibraryAdoptionAndReleasePlan_2026-09-09.md). **Source analysis complete; migration and local x64 Debug/Release qualification are complete on the feature branch; ARM64/ASan are user-deferred.** This is a non-normative audit, not a second implementation plan.

Current implementation checkpoint (2026-09-13): RedSalamander native candidate `1232d181` consumes canonical
DxUi `13788e95`; the duplicate library and tests have been removed. G1-G4 and G7 have upstream
implementations and regressions, G5/G6 use the accepted standalone behavior, and all six build configurations remain supported; final local qualification covers
x64 Debug/Release. ARM64/ASan follow-up is explicitly user-deferred. Final qualification, accepted resource costs and explicit platform
deferrals belong to the opening I19 checklist. The findings and action wording below describe
the audited September 9 source identities; they are retained to explain the migration, not as
an additional list of unimplemented work.

## 1. Result

Standalone DxUi is the correct canonical implementation for the accepted **DxUi.lib** direction. It contains substantial work beyond the RedSalamander copy, including supplied-device hosting, text services, accessibility integration, additional controls, and resource changes. It is **not yet a drop-in replacement** for RedSalamander.

The first adoption candidate needs four concrete gaps addressed: the missing tooltip clock fix, supported access to helpers currently used through private headers, localization of built-in empty-state strings, and implementation of the accepted native clipboard policy. Build/test integration and product regression evidence then qualify that candidate. None of this requires changing the accepted simple manual dependency-update lifecycle.

**User decisions after this audit, 2026-09-09:** G4 separates native clipboard capacity from the embedded editor ceiling while retaining validation and responsiveness; G5 is an accepted change with no legacy diagnostics/scheduling compatibility to preserve; the standalone DxUi Slider behavior in G6 is authoritative; every project in DxUi, RedXe and RedSalamander must support Debug, Release and ASAN Debug on both x64 and ARM64. G8 tracks delivery of that mandatory six-configuration matrix. These decisions do not claim the builds or migration have already been implemented.

## 2. Scope and evidence

| Role | Repository / source | Audited identity |
|---|---|---|
| Current product | RedSalamander, Common/DxUi and Tests/DxUiTests | 3f0df2f2c28239eb658a68f85b9c7f69c2356af4 |
| Candidate library | DxUi, include/DxUi, src and Tests | d192e474e540adc2656e69b8e8150b0f7a05d63f |
| Extraction ancestor | RedSalamander source recorded by DxUi | 5f83dc4b0b7c5d66de4f96895da43298532dd046, imported 2026-09-05 |

Local roots: Z:/src/RedSalamander and Z:/src/DxUi. Code was unchanged during this audit; the RedSalamander working changes are the I19 planning documents. RedXe is context for an existing consumer, not a third source fork compared here.

Method: reconcile all **58 origin records** against both current trees; examine post-extraction Git changes including imported support files; compare public declarations and implementation differences after separating namespace, host-name, formatting, and diagnostic-guard changes; count named test definitions and reconcile them with the archived case ledger; compare all eight inherited PNG files; inspect project references and source includes across Common, RedConfigure, RedSalamander, RedSalamanderMonitor, Plugins, and Tests. Historical Specs snapshots, PoC files, build outputs, and vendored dependencies are excluded from the live consumer count.

The [origin ledger][origin] contains 54 owned records, two consolidated FrameRuntime records, and two retired project files. All 56 current paths exist. Historical original hashes describe the import checkout, including its line endings; they are not constraints on current source. Appendix A accounts for every record.

Findings below are confirmed source/API differences or explicitly identified qualification gaps. No product build, runtime regression, GUI session, or paired performance benchmark was run for this documentation-only audit. Source/test counts do not establish passing tests or exhaustive behavioral equivalence.

## 3. Findings and dispositions

### G1 — Missing tooltip timing fix — high, DxUi fix before adoption

RedSalamander commit **817b7e2d** changed SetTooltipDelayed, BeginTooltipHideDelay, and DebugAdvanceTooltipDelayForTest to use AnimationDispatcher::GetCurrentTickMs(). Its dispatcher helper was added in the same change. Standalone still uses the host's possibly stale last animation tick, falling back to GetTickCount64(), and has no GetCurrentTickMs helper. A delay scheduled after a host has been idle can therefore be based on an old tick and expire too early when animation resumes.

Evidence: [legacy host](https://github.com/DualTail/RedSalamander/blob/3f0df2f2c28239eb658a68f85b9c7f69c2356af4/Common/DxUi/DxUi.WindowHost.cpp), [legacy dispatcher](../../../../RedSalamander/Ui/AnimationDispatcher.h), [candidate host][host], and the missing regression **TestTooltipDeadlinesUseCurrentDispatcherClockAfterIdleHostTick** in [legacy tooltip tests](https://github.com/DualTail/RedSalamander/blob/3f0df2f2c28239eb658a68f85b9c7f69c2356af4/Tests/DxUiTests/DxUiTests.Tooltip.cpp). It is the only new named case in the current legacy suite relative to the 941-case extraction ledger.

Action: port the regression and fix upstream, preserving the new dispatcher's thread-local lifetime. Resolve native deadlines against the native scheduler's clock; handle embedded timing through its host-owned clock contract instead of introducing a native timer into embedded mode. Test delayed show and hide after an idle/stale host tick, first use, clock restart, and the corresponding embedded path. A blind three-line cherry-pick is insufficient because the helper and host modes differ.

### G2 — Consumer APIs are stranded in private headers — high, compile/integration blocker

The supported consumer import exposes include/, while several live RedSalamander headers now exist only under standalone src/Controls. Merely changing the project reference will not make those includes supported. The production/test source scan found **79 distinct files with direct DxUi includes**, outside the legacy library and its test directory. Their 96 include sites are accounted for here:

| Legacy include, beneath DxUi/ | Direct include sites | Current standalone position and recommended treatment |
|---|---:|---|
| DxUi.h | 58 | Public include/DxUi/DxUi.h exists. Adapt RedSalamander::DxUi references to DxUi; compile all custom controls and delegates. |
| DxUi.FrameRuntime.h | 3 | Public replacement is DxUi/FrameRuntime.h. FolderView, the product animation dispatcher, and Monitor use it. |
| DxUi.Typography.h | 23 | Private. Promote the shared typography surface actually used by consumers to a supported public header, with coherent cache and metrics ownership. |
| DxUiNativeMenuInterop.h | 6 | Private. Main CompareDirectories menu plus ViewerText/Space/ImgRaw/PE/Web use native-menu conversion and hosting helpers. Expose the reusable bridge with its own public dependencies; keep application command policy in adapters. |
| DxUi.FocusRestore.h | 2 | Private. Main application and CompareDirectories use it. Decide a supported shared focus helper or reuse the matching existing product helper. |
| DxUi.PointerInput.h | 1 | Private, with out-of-line functions. NavigationView requires PointerInputEvent and routing helpers. Publish the needed neutral API; an include rename alone is insufficient. |
| DxUi.AccessibilityTextUnits.h | 1 | Private. TerminalAccessibility uses the shared Unicode/text-unit policy. Publish the required helper contract and preserve Terminal regression ownership. |
| DxUi.Win32Hooks.h | 1 | Private. Preferences uses GetStoredWndProc/InstallWndProcHook/CallStoredWndProc. Prefer the existing product Win32CallbackHelpers where semantics match; review restoration/focus behavior. |
| DxUi.Internal.h | 1 | Private. Commands selftests include it. Move implementation-specific assertions to DxUi tests or use a legitimate public diagnostic boundary. |

**ViewerVLC is an additional header-only consumer.** It includes Typography and uses CreateTextFormat/MakeUiTextSpec, despite having no direct DxUi project reference. It was absent from the initial ten-binary link inventory. Include it in migration and regression coverage; do not infer an extra archive-linked binary or revise the earlier disk estimates from this source observation.

The 12 direct project references remain ten production targets—RedSalamander, Monitor, RedConfigure, Terminal, ViewerText, Sqlite, Space, ImgRaw, PE, Web—and DxUiTests/RedConfigureTests. Shared Common/ModalWindowShell and product headers also expand the transitive use inside those projects.

Public-header adaptation also includes **MakeThemePaletteFromViewerTheme(ViewerTheme)** becoming **MakeThemePalette(ThemeColors)**. Six viewer families call the old conversion. Keep ViewerTheme in the plugin ABI and copy its values into the neutral ThemeColors in consumer code, setting sizeBytes correctly; do not reinterpret-cast the structures. Preserve MakeAppThemeDxPalette and MakeFolderContentDxPalette application policies.

WindowHost remains a public alias for ControlHost, so changing every host variable/type spelling is unnecessary. Namespace and changed class layouts still require recompilation of all affected modules. A temporary forwarding header may ease source migration, but must not expose src/Controls, compile shared .cpp files into consumers, or retain a second implementation.

### G3 — Built-in strings lost product localization — high, user-visible adaptation gap

Legacy LoadDxUiString loads product resources 1304/1305 through LoadStringResource; standalone [LoadDxUiString][theme] only caches the English fallback. ComboBox uses it for “No matches”; Tree and Grid use it for “No data.” The French product resources already contain “Aucun résultat” and “Aucune donnée,” for example.

Grid has SetEmptyStateText, which can supply localized text. The inspected public API has no equivalent empty-message setter for Tree or no-match setter for ComboBox, and no public string resolver for these defaults. Thus the adapter boundary is incomplete for those controls. Independently selected labels/placeholders already supplied by the application are unaffected by this particular gap.

Action: add small library-owned text configuration/setter APIs and feed them product-localized strings. Prefer explicit text ownership to importing product numeric resource IDs. Verify filtering with no matches, empty tree/grid, fallback behavior, language initialization/change, multiple module instances, and sizing for longer translated text. Keep language resources in the consumer. Do not restore LoadStringResource as a library dependency.

### G4 — Separate native clipboard capacity — decision accepted, implementation required

Standalone [WindowsTextClipboard][clipboard] is now shared by native ControlHost and application-owned text services. It rejects writes above **65,536 UTF-16 code units**, rejects embedded NUL on write, bounds/validates Unicode reads, and attempts OpenClipboard once. The legacy native implementation had no equivalent 65,536-unit cap, read unbounded NUL-terminated text, and retried clipboard acquisition up to 20 times with 10 ms waits. Its test-only fallback behavior also differed.

This has a concrete native path: Grid::OnCopy calls host.CopyTextToClipboard(BuildSelectionTsv()). A valid selection producing 65,537 code units could previously proceed to clipboard allocation, but the candidate rejects it before opening the clipboard. A transiently busy clipboard can also fail sooner. Unicode bounds/validation and avoiding UI-thread sleeps are useful improvements, but they do not establish product compatibility.

**Why this changed:** commit [2ca8cc0: application-owned text services][clipboard-change] introduced the shared TextClipboard backend and made native ControlHost use it too. The accompanying [input contract][input-contract] explicitly requires one clipboard-open attempt, bounded reads, valid UTF-16 and a 65,536-unit ceiling. The embedded text snapshot/edit contract already uses the same ceiling. This is a text-service design choice, unrelated to static linking or DLL packaging.

There are three separate choices:

| Choice | Reason established by code/contracts | Assessment for adoption |
|---|---|---|
| Honor the clipboard allocation extent, require termination and validate Unicode | Prevent reading past supplied memory or importing malformed text. | Keep these checks; they do not require a 65,536-unit global limit. |
| Cap accepted text at 65,536 UTF-16 units (128 KiB of UTF-16 payload, excluding the terminator) | The ceiling is explicit in the text-service contract and tests. It bounds accepted text/work in that service. | The inspected commit/contracts do not explain why this exact limit should also govern native Grid copy/paste. It rejects a valid 100,000-character selection even when allocation could succeed. |
| Attempt OpenClipboard once | The contract explicitly removes retry sleeps. The old helper could sleep 19 times for 10 ms while called on the UI thread. | Keep the UI responsive. If transient contention needs retries, design deferred/cancellable retries with a clear result rather than restoring sleeps. |

**Accepted by the user, 2026-09-09:** keep memory/Unicode validation and responsiveness, and separate native clipboard capacity from the embedded editor's 65,536-unit text ceiling.

- Native Grid/TextField copy/paste must support valid selections above 65,536 UTF-16 units, including the 100,000-character regression fixture. The embedded text ceiling must not impose that limit on native clipboard transport.
- Bound reads by the actual clipboard allocation, require termination, validate Unicode, check size arithmetic and handle allocation failure explicitly. Any additional native resource budget must be independent of the embedded editor limit and justified by product/resource evidence.
- Keep the current single clipboard-open attempt and explicit failure result. Do not restore UI-thread retry sleeps or add a retry service to this slice; deferred/cancellable retries would need an observed product requirement.
- Preserve the separate embedded snapshot/edit ceiling and its rejection semantics. Increasing clipboard transport capacity must not allow an oversized paste to bypass the embedded editor's validation.

Decision complete; implementation started under I19/C2 on 2026-09-09. The progress checklist at the beginning of the owning I19 plan tracks code, tests, specs and qualification.

Action: implement the accepted policy and add native Grid/TextField cases at 65,535/65,536/65,537 and 100,000 UTF-16 units, plus malformed/unterminated input, checked size/allocation failure and injected clipboard contention. Retain embedded oversized-edit rejection tests. Keep clipboard failure distinguishable from successful copy and preserve selected text when a cut cannot copy. Use injected clipboard fixtures rather than changing the desktop clipboard during automated tests.

### G5 — Diagnostics and scheduling — accepted, no legacy compatibility work

**User decision, 2026-09-09:** adopt the standalone mechanisms. G5 is not an adoption problem and does not require retaining legacy logging behavior, metric spellings, trace-file environment variables, window properties or the old dispatcher implementation.

The factual differences remain useful implementation context: standalone uses optional synchronous, thread-local Diagnostics::sink/performanceSink, the DxUi.Diagnostics property, current metric names and a thread-local native animation dispatcher. Integrate current APIs directly where the product needs diagnostics. Update or remove old diagnostic consumers and assertions instead of adding compatibility aliases or translation layers.

Application-owned scheduling for non-DxUi surfaces still has to work, but its legacy implementation is not a requirement. Use the appropriate current scheduling/host contract and normal shutdown/plugin-unload tests. Shared message numbers do not imply shared token registries; ordinary module ownership and teardown correctness remain engineering requirements, independent of G5 compatibility.

Disposition: no separate preservation work or G5 compatibility gate. Retain the existing product resource/performance checks using the adopted interfaces.

### G6 — Changed controls, rendering and input need product checks — medium/high by surface

| Candidate difference | Product implication and qualification |
|---|---|
| Slider has Preview/Commit/Cancel, restores the initial value on capture cancellation, adds RequestValue and animated visual/value state, and changes touch/paint geometry. SetOnValueChanged remains available. | **Accepted by the user: this standalone behavior is correct.** Adapt RedConfigure's alpha-slider consumer and test expectations to it; verify draft/apply wiring, event counts, cancellation and layout/DPI. Do not reproduce the old Slider behavior or treat its intended visual/event differences as regressions. The main thumbnail menu uses the separate MenuItemKind::Slider implementation. |
| Primary buttons preserve system selection color pairs in high contrast. | Intended accessibility improvement; compare product dark/light/high-contrast themes, disabled state and focus. |
| PageIndicator, FontRole::IconLarge and per-instance ComboBox minimum popup row height are new. | Additive capabilities. They do not require adoption by existing screens. Default and configured popup geometry still need qualification. |
| Grid/Tree animation invalidation and native/embedded host cache trimming changed. | Measure active/settling animation frames, clean/hidden idle behavior, cache retention and many hosted controls. Cache thresholds are trimmed at a subsequent paint/preparation; do not interpret them as an unconditional never-exceeded allocation count. |
| Native TSF focus identity/callback invalidation and shared clipboard/text services were revised. Popup loops prioritize popup input/paint. | Recheck real focus after menus/modal close, root replacement during callbacks, native text editing, IME, accessibility provider shutdown and plugin unloading. Synthetic tests do not establish real IME/screen-reader acceptance. |
| Missing optional D3D SDK debug layers have an explicit fallback. | Beneficial for Debug portability; retain both hardware and WARP checks and do not weaken other device-creation failures. |

Standalone also adds supplied-device GraphicsDevice/EmbeddedHost, preparation/composition, revision-checked text snapshots, application-owned text services, embedded UIA, and hidden/zero-size surface release. These serve RedXe and extend the shared control engine. **RedSalamander's native HWND hosts need not migrate to embedded rendering.** Native product non-regression remains required even when a change originated in embedded work.

### G7 — Shared pre-existing offscreen Grid UIA issue — high, existing defect rather than a delta

Both copies allow SelectionItem pattern acquisition for a grid row that is visible **or selected**, but get_IsSelected and get_SelectionContainer require SnapshotContainsGridRow, which checks the visible-row list. A selected offscreen row can therefore expose the pattern while its getters reject it. Standalone's additional embedded-provider logic does not remove this condition.

Evidence: legacy Accessibility.cpp at SnapshotContainsGridRow/get_IsSelected/get_SelectionContainer and candidate [Accessibility.cpp][accessibility] at the same functions (candidate lines 1550, 6907, 6947; pattern selection around 4966). This matches the residual already routed through RedSalamander H0/TW-9/TW-10.

Action: carry a shared selected-offscreen-row runtime regression and repair through the DxUi owner coordinated by I19/H0. Verify getter results, selection container identity, bounds/offscreen state, removal and stale providers. Keep it explicitly separate from migration regressions; neither a library switch nor a passing retained suite proves it resolved.

### G8 — Build, sanitizer and evidence integration — high for admission

Both implementations are static libraries using the MSVC/WIL/DirectX stack. Standalone consumer imports already enforce an exact lock, API/target checks and isolated output; reuse those mechanisms. The product must stop finding Common/DxUi headers through broad include paths once a module moves.

**Mandatory target, per user decision:** every project across DxUi, RedXe and RedSalamander must provide all six combinations. This includes libraries, applications, plugins, samples and test projects; native source must be instrumented in ASAN Debug. Resource/package projects must map consistently to the selected solution configuration.

| Configuration | x64 | ARM64 |
|---|---|---|
| Debug | Required | Required |
| Release | Required | Required |
| ASAN Debug | Required | Required |

Use the existing MSBuild spelling `ASan Debug` consistently where already established. The requirement is real sanitizer instrumentation, not just a configuration name. Build the matching instrumented DxUi archive and affected native project code; do not silently map ASAN Debug to ordinary Debug.

The audited legacy project declares ASan Debug configurations on both architectures, while standalone currently declares Debug/Release. Deliver the missing project/solution configurations, toolchain/runtime support, CLI validation, isolated output/fingerprint inputs and CI/test lanes in each repository. Verify an intentionally failing sanitizer fixture is detected, then run the relevant regression suites. Native ARM64 execution is required to claim runtime qualification. A missing compiler/runtime/runner is an implementation dependency to resolve, not permission to drop a matrix cell or describe it as optional. Configuration declarations alone are not runtime coverage.

Wire the external pin, headers, implementation/build inputs and archive identity into existing build attestations and source.dxui/runtime dependency closure. Test-enabled Release and normal Release must not mix stale headers or artifacts. Fixed DXUI_ENABLE_DIAGNOSTICS=1 replaces the legacy ENABLE_TESTS layout switches; consumer flags must not change library layouts. Test actual updated binaries and preserve architecture/compiler/CRT consistency.

Reuse the existing Full runner, receipts, sandbox, leases and directed-input warning. A consumer upgrade remains a normal manual pin-update PR. No bot, cross-repository release service, extra receipt framework or DLL is needed.

## 4. Post-extraction changes on the RedSalamander side

| Commit / files | Classification | Disposition |
|---|---|---|
| 817b7e2d: WindowHost, AnimationDispatcher, Tooltip tests | Shared behavior fix missing upstream | G1. Bring the fix and regression into DxUi with appropriate native/embedded clocks. |
| f0710f9b: Theme tests | Product utility coverage; added reusable strict UTF conversion assertions inside an existing test | Retain in RedSalamander. The test is already one of the 20 product-owned exclusions in the import ledger; do not copy FileOps/general string utilities into DxUi. |
| 3f0df2f2: DxUiTests main | Product test-runner integration; directed foreground-input warning and failure handling | Preserve under RedSalamander's runner. Standalone can retain its independent test UX; migrating a runner must not bypass the product warning/desktop lease contract. |

Review of the mapped support paths found no additional post-extraction behavioral change beyond the dispatcher part of 817b7e2d. This audit does not claim that unrelated RedSalamander application changes belong in the shared library.

## 5. Test coverage reconciliation

The archived [test-port ledger][testport] records **941 named legacy cases: 853 ported and 88 deliberately excluded**. Current legacy contains **942** named void Test...() definitions, adding the tooltip regression. Current standalone Controls contains **880** such definitions: 853 retained contracts (including four renamed theme cases) plus 27 additions. The four TestViewerTheme... names became TestThemeColors...; this is not four lost tests.

Of the 27 added definitions, eight are in TextInputServicesTests.h; the remaining 19 are across existing .cpp families. The 880 count excludes gallery entrypoints and the Foundation/Embedded suites, which use a different Check-based structure. It is a source inventory, not a count of fresh passing assertions or coverage percentage.

### 5.1 Inherited case accounting

| Family | Original | Ported | Source-text exclusions | Product exclusions | Current standalone .cpp | Current legacy |
|---|---:|---:|---:|---:|---:|---:|
| Accessibility | 67 | 48 | 19 | 0 | 48 | 67 |
| Animation | 33 | 33 | 0 | 0 | 36 | 33 |
| ComboBox | 27 | 25 | 2 | 0 | 26 | 27 |
| Controls | 40 | 32 | 8 | 0 | 32 | 40 |
| Grid | 62 | 58 | 4 | 0 | 58 | 62 |
| Menu | 66 | 57 | 6 | 3 | 57 | 66 |
| MultilineText | 106 | 106 | 0 | 0 | 106 | 106 |
| NativeTextInput | 119 | 116 | 3 | 0 | 117 | 119 |
| NewControls | 68 | 67 | 1 | 0 | 79 | 68 |
| ReadOnly | 24 | 24 | 0 | 0 | 24 | 24 |
| Rendering | 27 | 27 | 0 | 0 | 27 | 27 |
| TextField | 80 | 72 | 8 | 0 | 72 | 80 |
| Theme | 100 | 83 | 0 | 17 | 84 | 100 |
| Tooltip | 12 | 12 | 0 | 0 | 12 | 13 |
| Tree | 32 | 28 | 4 | 0 | 28 | 32 |
| WindowHost | 78 | 65 | 10 | 3 | 66 | 78 |
| **Total** | 941 | 853 | 65 | 23 | 872 | 942 |

The **65 source-text assertions** are explicitly excluded because they test implementation spelling. Their general “replaced by runtime contracts” reason is not a one-to-one proof of replacement coverage. Before retiring the product DxUiTests project, map any still-required behavior/structural policy to the relevant runtime or source guard under I4, especially accessibility lifetime, input, popup capture and host shutdown. Preserve legitimate architecture checks instead of deleting them merely for reading source.

The **23 product/test-utility exclusions** must remain in a product test target when the legacy library test executable is removed. They cover FolderView behavior, color/theme/window/viewer helpers, strict Unicode conversion, paths/formatting/IO/cloud/JSON policies, payload coalescing, and sandbox/test-support behavior. Exact names are listed in Appendix B. None should silently disappear in a project-reference cleanup.

### 5.2 Visual and newer test evidence

Seven of eight inherited baseline PNGs are byte-identical. **advanced_controls_dark.png differs**, consistent with the standalone slider-geometry/paint commits, most recently 13625b6. The user has accepted the standalone Slider behavior and visuals; update affected product expectations to that contract without another legacy-compatibility decision. Continue checking unrelated visuals and resource regressions; this acceptance does not authorize blanket baseline replacement.

Existing standalone tests add PageIndicator interaction/root replacement, touch-friendly Slider geometry/animation/grab behavior, ComboBox row sizing, high-contrast primary buttons, native text callback invalidation, optional-debug-layer fallback, and application text/clipboard transactions. Embedded tests exercise capture cancellation, invalidation/revision behavior, supplied-device composition and text/UIA. These are useful assets to reuse; this audit did not rerun them.

Minimum missing/required adoption cases: G1 idle tooltip deadlines; G3 translated empty/no-match content; G4 native large-selection clipboard and contention; public-only consumer fixtures for G2 helpers; selected offscreen Grid getters for G7; actual RedConfigure/viewer/Terminal/Monitor interaction and teardown; native/embedded paired resource evidence. Keep real IME/assistive-technology/native ARM64 gates explicit.

## 6. Recommended sequence within I19

1. Complete C0's retained native product baseline and targeted reproductions for G1/G3/G4/G7. This audit supplies source inventory and gap dispositions, but does not complete the baseline or runtime gate.
2. Use focused DxUi PRs for the tooltip regression/fix, the small public helper surface, localized control text, and the accepted G4 clipboard policy. Include the relevant shared tests, docs/gallery updates where required, and normal library CI.
3. Deliver the mandatory Debug/Release/ASAN Debug × x64/ARM64 matrix across all projects in the three repositories. Add RedSalamander's exact lock/imports and adapters, adopt current diagnostics directly, and qualify the evidence path. Pilot the separate RedConfigure process, adapting its alpha slider to the accepted standalone behavior and preserving localized empty/filter states.
4. Migrate Monitor and then the main process/plugins with public-header fixtures and existing product regressions. Include header-only ViewerVLC, Terminal's text units, and the product animation dispatcher's FrameRuntime usage. Do not mix old/new headers within a PE.
5. Run required Fresh Full and matched Release/resource/platform/manual qualification. Only then remove Common/DxUi and retire/rehome its tests. Keep the 23 product cases and reconcile the 65 source-policy exclusions.

For each shared fix: merge green DxUi, update the fixed consumer SHA, run local regressions, push the consumer PR and merge after its CI passes. If a product regression remains, fix its owning repository and repeat. This is the lifecycle already selected by the user.

## Appendix A. Complete imported-file disposition

All paths below are repository-relative. “Different” means different current normalized text/bytes, not necessarily different behavior. Text comparison normalizes CRLF/LF, RedSalamander::DxUi to DxUi, the WindowHost type token to ControlHost, and horizontal whitespace; it does not treat different tests as equivalent. Retired projects are intentionally replaced. Current sources and API/test findings above determine migration work; the historical import record is not an instruction to copy files.

| Legacy source | Current standalone path | Disposition / current comparison |
|---|---|---|
| Common/DxUi/DxUi.Accessibility.cpp | src/Controls/DxUi.Accessibility.cpp | Owned; different |
| Common/DxUi/DxUi.AccessibilityTextUnits.h | src/Controls/DxUi.AccessibilityTextUnits.h | Owned; different |
| Common/DxUi/DxUi.ComboBox.cpp | src/Controls/DxUi.ComboBox.cpp | Owned; different |
| Common/DxUi/DxUi.Controls.cpp | src/Controls/DxUi.Controls.cpp | Owned; different |
| Common/DxUi/DxUi.cpp | src/Controls/DxUi.cpp | Owned; different |
| Common/DxUi/DxUi.FocusRestore.h | src/Controls/DxUi.FocusRestore.h | Owned; same after normalization |
| Common/DxUi/DxUi.FrameRuntime.cpp | src/Foundation/FrameRuntime.cpp | Consolidated; different |
| Common/DxUi/DxUi.FrameRuntime.h | include/DxUi/FrameRuntime.h | Consolidated; different |
| Common/DxUi/DxUi.Grid.cpp | src/Controls/DxUi.Grid.cpp | Owned; different |
| Common/DxUi/DxUi.h | include/DxUi/DxUi.h | Owned; different |
| Common/DxUi/DxUi.Internal.h | src/Controls/DxUi.Internal.h | Owned; different |
| Common/DxUi/DxUi.Menu.cpp | src/Controls/DxUi.Menu.cpp | Owned; different |
| Common/DxUi/DxUi.NativeTextInput.cpp | src/Controls/DxUi.NativeTextInput.cpp | Owned; different |
| Common/DxUi/DxUi.PointerInput.cpp | src/Controls/DxUi.PointerInput.cpp | Owned; same after normalization |
| Common/DxUi/DxUi.PointerInput.h | src/Controls/DxUi.PointerInput.h | Owned; different |
| Common/DxUi/DxUi.Scrollbar.cpp | src/Controls/DxUi.Scrollbar.cpp | Owned; same after normalization |
| Common/DxUi/DxUi.SingleLineTextEditing.cpp | src/Controls/DxUi.SingleLineTextEditing.cpp | Owned; different |
| Common/DxUi/DxUi.TextInput.cpp | src/Controls/DxUi.TextInput.cpp | Owned; different |
| Common/DxUi/DxUi.TextStoreACP.cpp | src/Controls/DxUi.TextStoreACP.cpp | Owned; different |
| Common/DxUi/DxUi.Theme.cpp | src/Controls/DxUi.Theme.cpp | Owned; different |
| Common/DxUi/DxUi.Tree.cpp | src/Controls/DxUi.Tree.cpp | Owned; different |
| Common/DxUi/DxUi.Typeahead.cpp | src/Controls/DxUi.Typeahead.cpp | Owned; same after normalization |
| Common/DxUi/DxUi.Typography.h | src/Controls/DxUi.Typography.h | Owned; different |
| Common/DxUi/DxUi.vcxproj | Replacement DxUi-owned project | Retired |
| Common/DxUi/DxUi.Win32Hooks.h | src/Controls/DxUi.Win32Hooks.h | Owned; different |
| Common/DxUi/DxUi.WindowHost.cpp | src/Controls/DxUi.WindowHost.cpp | Owned; different |
| Common/DxUi/DxUiNativeMenuInterop.h | src/Controls/DxUiNativeMenuInterop.h | Owned; different |
| Tests/DxUiTests/Baselines/advanced_controls_dark.png | Tests/Controls/Baselines/advanced_controls_dark.png | Owned; different |
| Tests/DxUiTests/Baselines/core_controls_dark.png | Tests/Controls/Baselines/core_controls_dark.png | Owned; identical bytes |
| Tests/DxUiTests/Baselines/core_controls_high_contrast.png | Tests/Controls/Baselines/core_controls_high_contrast.png | Owned; identical bytes |
| Tests/DxUiTests/Baselines/core_controls_light.png | Tests/Controls/Baselines/core_controls_light.png | Owned; identical bytes |
| Tests/DxUiTests/Baselines/menu_popup_acrylic_light.png | Tests/Controls/Baselines/menu_popup_acrylic_light.png | Owned; identical bytes |
| Tests/DxUiTests/Baselines/page_transition_dark.png | Tests/Controls/Baselines/page_transition_dark.png | Owned; identical bytes |
| Tests/DxUiTests/Baselines/popup_and_bars_acrylic_light.png | Tests/Controls/Baselines/popup_and_bars_acrylic_light.png | Owned; identical bytes |
| Tests/DxUiTests/Baselines/popup_and_bars_dark.png | Tests/Controls/Baselines/popup_and_bars_dark.png | Owned; identical bytes |
| Tests/DxUiTests/DxUiTestHelpers.h | Tests/Controls/DxUiTestHelpers.h | Owned; different |
| Tests/DxUiTests/DxUiTests.Accessibility.cpp | Tests/Controls/DxUiTests.Accessibility.cpp | Owned; different |
| Tests/DxUiTests/DxUiTests.Animation.cpp | Tests/Controls/DxUiTests.Animation.cpp | Owned; different |
| Tests/DxUiTests/DxUiTests.ComboBox.cpp | Tests/Controls/DxUiTests.ComboBox.cpp | Owned; different |
| Tests/DxUiTests/DxUiTests.Controls.cpp | Tests/Controls/DxUiTests.Controls.cpp | Owned; different |
| Tests/DxUiTests/DxUiTests.cpp | Tests/Controls/DxUiTests.cpp | Owned; different |
| Tests/DxUiTests/DxUiTests.Gallery.cpp | Tests/Controls/DxUiTests.Gallery.cpp | Owned; different |
| Tests/DxUiTests/DxUiTests.Grid.cpp | Tests/Controls/DxUiTests.Grid.cpp | Owned; different |
| Tests/DxUiTests/DxUiTests.Menu.cpp | Tests/Controls/DxUiTests.Menu.cpp | Owned; different |
| Tests/DxUiTests/DxUiTests.MultilineText.cpp | Tests/Controls/DxUiTests.MultilineText.cpp | Owned; same after normalization |
| Tests/DxUiTests/DxUiTests.NativeTextInput.cpp | Tests/Controls/DxUiTests.NativeTextInput.cpp | Owned; different |
| Tests/DxUiTests/DxUiTests.NewControls.cpp | Tests/Controls/DxUiTests.NewControls.cpp | Owned; different |
| Tests/DxUiTests/DxUiTests.ReadOnly.cpp | Tests/Controls/DxUiTests.ReadOnly.cpp | Owned; same after normalization |
| Tests/DxUiTests/DxUiTests.Rendering.cpp | Tests/Controls/DxUiTests.Rendering.cpp | Owned; different |
| Tests/DxUiTests/DxUiTests.TextField.cpp | Tests/Controls/DxUiTests.TextField.cpp | Owned; different |
| Tests/DxUiTests/DxUiTests.Theme.cpp | Tests/Controls/DxUiTests.Theme.cpp | Owned; different |
| Tests/DxUiTests/DxUiTests.Tooltip.cpp | Tests/Controls/DxUiTests.Tooltip.cpp | Owned; different |
| Tests/DxUiTests/DxUiTests.Tree.cpp | Tests/Controls/DxUiTests.Tree.cpp | Owned; different |
| Tests/DxUiTests/DxUiTests.vcxproj | Replacement DxUi-owned project | Retired |
| Tests/DxUiTests/DxUiTests.WindowHost.cpp | Tests/Controls/DxUiTests.WindowHost.cpp | Owned; different |
| .clang-format | .clang-format | Owned; same after normalization |
| RedSalamander/Ui/AnimationDispatcher.h | src/Support/AnimationDispatcher.h | Owned; different |
| Common/TestWindowActivationGuard.h | Tests/Support/TestWindowActivationGuard.h | Owned; different |

Additional current library files outside this origin map are Configuration, ControlCatalog, Diagnostics, Embedded, EmbeddedAccessibility, TextInputServices and ThemeColors public headers; control catalog/text clipboard/text-store-target/text-service implementation; supplied-device renderer and shaders; independent diagnostics/utilities/payload/messages support; and the canonical project/private forwarding header. Foundation and Embedded test projects, TextInputServicesTests and PerformanceCapture are additional test infrastructure. Samples, gallery, benchmarks, public consumer fixtures, build/restore/imports and CI are independent delivery assets, not missing RedSalamander source.

## Appendix B. Product-owned test cases to retain

These 23 names come from the archived exclusion ledger. Their present owner is RedSalamander; choose the surviving product target before deleting DxUiTests.vcxproj.

- TestFolderViewIncrementalSearchKeepsContainsHighlightButUsesPrefixFocus
- TestFolderViewInactiveVisualStateDimsNormalTextAndIcons
- TestFolderViewEmptyPlaceholderMetricsUseCurrentEmptyLayout
- TestSharedColorRefArgbAndTruncatingBlendGoldenValues
- TestHsvColorRefNegativeHuePoliciesRemainExplicitAndDistinct
- TestSharedLuminanceMathSeparatesLinearizedAndEncodedPolicies
- TestWindowSizingDpiRoundingAndPopupGeometryGoldenValues
- TestViewerFileComboPopupInputAndHeightGoldenValues
- TestViewerTitleBarThemeGoldenValues
- TestUtfConversionPoliciesRemainExplicitAndPreserveEmbeddedNulls
- TestOrdinalTrimAndEnvironmentPoliciesRemainExplicit
- TestWindowsPathClassesKeepDriveRelativeAndFullyAbsolutePoliciesDistinct
- TestBinaryThroughputGrammarPreservesBoundaryPoliciesAndRoundTrips
- TestCompactFileMetadataFormattingPreservesTimeAndAttributePolicies
- TestPercentEncodingAndHandleIoPoliciesCoverForwardProgressAndAtomicSiblings
- TestCloudPaginationGuardBoundsProgressAndCancellation
- TestYyjsonMemberPoliciesDistinguishPresenceTypeRangeAndCoercion
- TestAppThemeDxPaletteRefreshesAccentVariantsAfterAssigningAccent
- TestFolderContentDxPalettePreservesItsNamedProfile
- TestAppThemeSelectionResolutionIsSharedAndDeterministic
- TestContiguousPostedPayloadCoalescingPreservesQueueOrderAndOperationKeys
- TestSharedTestSupportPreservesSandboxAndEnvironmentPolicies
- TestSharedTestSupportPumpsMessagesAndBoundsSnapshotPolling

## 7. Verification and remaining limits

Static checks performed: both Git identities and working-tree scope; all origin paths/dispositions; post-extraction changes in mapped source/tests/support; public/private include reachability and direct project references; inherited/new case accounting; all eight baseline byte comparisons; focused source inspection of the findings. WIP links/index and Markdown whitespace are validated with the documentation delivery.

Still required before adoption: compilation of a real candidate against public headers; fresh standalone and product runtime tests; exact test execution/replacement mapping where exclusions remain; paired product resource/performance measurements; the mandatory six-configuration build/sanitizer matrix, packaging, native architecture and manual input/accessibility qualification. These are unperformed qualification steps, not failures observed during this source audit.

[origin]: https://github.com/RedSalamanders/DxUi/blob/d192e474e540adc2656e69b8e8150b0f7a05d63f/Specs/Done/SourceImport/source-origin.json
[testport]: https://github.com/RedSalamanders/DxUi/blob/d192e474e540adc2656e69b8e8150b0f7a05d63f/Specs/Done/SourceImport/test-port.json
[host]: https://github.com/RedSalamanders/DxUi/blob/d192e474e540adc2656e69b8e8150b0f7a05d63f/src/Controls/DxUi.WindowHost.cpp#L1383
[theme]: https://github.com/RedSalamanders/DxUi/blob/d192e474e540adc2656e69b8e8150b0f7a05d63f/src/Controls/DxUi.Theme.cpp#L177
[clipboard]: https://github.com/RedSalamanders/DxUi/blob/d192e474e540adc2656e69b8e8150b0f7a05d63f/src/Controls/TextClipboard.cpp#L10
[accessibility]: https://github.com/RedSalamanders/DxUi/blob/d192e474e540adc2656e69b8e8150b0f7a05d63f/src/Controls/DxUi.Accessibility.cpp#L6907


[clipboard-change]: https://github.com/RedSalamanders/DxUi/commit/2ca8cc03aae08f943d00abdfeddfe6f87f7800cd
[input-contract]: https://github.com/RedSalamanders/DxUi/blob/d192e474e540adc2656e69b8e8150b0f7a05d63f/Specs/UI/UI_InputAndAccessibility.md#L114
