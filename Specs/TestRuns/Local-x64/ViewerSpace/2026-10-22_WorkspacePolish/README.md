# ViewerSpace native workspace qualification

Status: final x64 Debug and isolated Release builds and focused workspace/capture/keyboard checks passed. Repaired resource/documentation contracts passed 68/68. The original broad affected run failed; no Full repository pass is claimed.

## Source and build binding

The workspace is based on `95f7114b892484414dd9daac7f3a89c1d1fd68ae` with the reviewed uncommitted ViewerSpace changes. The exact changed native file hashes are in `logs/isolated-final-native-source-match.json`; the primary and isolated checkout bytes match for these files. Different physical checkouts have different receipts; this is not automatic cross-root Full evidence reuse.

The final primary x64 Debug build reported zero warnings and errors, receipt `f2184c8a376f0719c145714af9a6ed7c23a6aa6bba0ccc20ef15aa7df375e362` (`logs/native-workspace-build9.txt`). The capture and initial regression build is Debug receipt `999bc54c5eccaaee478ebede3fd51b142957f04ffba4fa205a77cc80630690b7`. Its binary hashes remain in each unchanged capture provenance record. Two subsequent changes update an accessible exact-byte tooltip when its rounded label does not change and suppress an unavailable synthetic count share; the final qualification will identify the build separately.

## Implemented scope

Persistent discovered totals; current-folder shown totals; compact inspector; Copy path; session exclusions with restore; one layout menu; manual reviewed cleanup through the normal host operation workflow; current-folder name/type/age filters; optional file-count weighting and color/date modes; Problems; bounded atomic CSV/JSON export. Viewer settings are in global Preferences, Plugins, Space. The host interface includes active-pane focus and other-pane open without assigning a new other-pane shortcut.

Reliable allocation and hard-link accounting, metadata-based scan acceleration, incremental reuse, fair scan scheduling, the broad provider/platform matrix and end-to-end latency/retention gates remain in I33. No allocation values are fabricated. I34 owns unrelated qualification failures. Neither plan is declared complete.

## Focused results before final repetition

Debug8: three fresh 2000-direct-file workspace cases passed; lifecycle captures, desktop observation, ambient desktop changes, listing validation, hostile-provider scanning and blocked-provider/close races passed. The exact approved keyboard/menu/Escape case passed under the canonical desktop lease, with exact saved focus restored. Source contracts: 211 passed, zero failures. Spec inventory: zero blocking findings. The changelog is folded and valid.

The first 200,000-file/20,000-directory summary case missed its deadline during concurrent Release compilation. Its unchanged quiet repetition passed. Both logs are retained; this is not evidence that the loaded failure never happened, nor a demonstrated model-correctness defect. The initial longer isolated checkout failed AWS SDK compilation on paths over the Windows legacy path length. The short isolated checkout then built the same dependency and full solution with zero warnings/errors. Compiler logs are retained; no product source workaround was introduced.

## Performance evidence and limitations

All 96 baseline and 96 Debug8 sibling-query samples are bound in `metrics/sibling-query-comparison.json`. Median query spans were 33,550 ns and 33,700 ns respectively. These were sequential blocks, not counterbalanced runs; no whole-viewer speedup or general latency conclusion follows.

`metrics/selection-bindings.json` binds raw SHA-256, byte/row counts and every original selected line. Selection keeps every row in each named workspace/exclusion/export, analysis/layout, paint, input and drain family, in original order, including overruns. Full raw streams remain under the marked sandbox. `debug8-metric-summary.json` is a reproducible diagnostic aggregation. Exclusion binding reached 4,221 us, fact preparation 3,137 us and export preparation 1,351 us. These rows do not close the 1.2-ms turn targets, wide/skewed corpora, whole-process memory or 150/310-ms end-to-end gates. No failed or slow sample is discarded.

## Application captures

`captures/baseline` is the original synthetic native harness frame. `captures/candidate` contains partial, done, right-panel, excluded and cleanup-review states from Debug8. These are actual owned embedded client frames, with no operating-system chrome. Every image retains its original scenario/build/DPI/fixture provenance. No manual desktop capture, personal path or redaction placeholder is published. The separate browser prototype remains a proposal.

## Final qualification

Final-source x64 Release receipt `087391f6c2bed292086b9edffbbe1beaeefa9007b1343c65bfee603a60a1b496`: full solution build, zero warnings/errors. Three fresh workspace cases and seven additional no-activation cases all passed, including the 200,000-file/20,000-directory summary case. The exact approved interaction case also passed with exact saved foreground and keyboard focus restored. Result/binary bindings are in `logs/final-source-release-results.json` and the full receipt is retained. `captures/candidate-release` retains final-source done/excluded/review pixels and original provenance.

Final Release diagnostic maxima were 2,278 us exclusion binding, 1,952 us derived facts and 1,411 us export preparation. Retained complete metric families include overruns. This is correctness/scenario evidence, not a turn-budget pass or performance improvement claim.

The affected deployment writer failed with MSVC C2471 while updating `FileSystemMtp/vc145.pdb` (102 compiler errors including the stop limit). The preceding full Debug build and final isolated Release build passed. `logs/affected-targeted-deployment-failed.txt` preserves the full failure. The canonical runner restored full x64 Debug receipt `ceb899c971aa0968691dfdcb6999c5d6cbcb9095b779b793c0adb734fdddb9e8` with zero warnings/errors before further readers. The original failed deployment case remains failed. No MTP source fix or relaxed check was introduced.

PrePush Explain could not verify PR equivalence and selected the full local obligations. Five broader interactive families remain pending; the exact approved ViewerSpace case does not qualify their entire groups. The final-source Debug repetition passed the same three fresh workspace runs and seven no-activation cases, followed by the exact approved interaction case with exact saved foreground and keyboard focus restored. Bindings are in `logs/final-source-debug-results.json`. The complete changed native source matches the final Release checkout. Exact affected terminal outcome and archive checks will be added here. Scoped passes never imply repository Full qualification.

The broader plugin contract entry reported FileSystem debug selftests 189 passed/1 failed and S3 221 passed/3 failed. Other native selftest groups have their own outcomes in the canonical run. These failures are assigned to I34 for qualification triage; this packet does not establish their cause or qualify them as ViewerSpace regressions.

## Broad outcome and scoped contract repairs

Canonical affected run `20261022T164835Z-53288-f0a45ee600ba4b96b45c2a404f9a00ed` ended FAILED (process exit 1). All 15 selected entry outcomes are projected with original digest bindings in `logs/affected-run-projection.json`: 12 PASSED and 3 FAILED (deployment tooling, general tooling, plugin contracts). ViewerPE and ViewerSqlite noninteractive, CompareDirectories, performance, product UI and the other selected entries passed their recorded scopes. Five broader interactive families remain pending. Final archive retention also rejected the Compare junction alias; no containment guard was bypassed. No Full repository pass is claimed.

The general tooling batch initially recorded 982 passed, 18 failed, 1 skipped. Four failures were in this change: shared gallery hash/provenance binding, exact Space menu order, missing/duplicate static menu access keys and Escape documentation wording. These were repaired and the five owning focused test files passed 68/68 with zero skips. The menu/caption resource and documentation-only closeout changes require new final receipts; C++ implementation bytes are unchanged.

The other 14 tooling failures remain separately scoped: six symlink-privilege cases, four sealed identity/workflow hash cases, three Ghostty diagnostic-field cases and a first-party export scan entering a generated WIL dependency tree. These observations and the FileSystem/S3 assertions belong to I34; they have not been repaired or causally attributed by this work. Original failed results remain failed.

Published guide pixels now come from the canonical broad ViewerPE capture case, in the separately validated [native capture archive](../../UI/2026-10-22_ViewerSpaceNativeCaptures/README.md). The gallery and all sidecars bind the actual receipt and original native provenance; earlier Debug8 and final-source Release captures in this packet remain historical evidence for their exact builds. All thirteen new publication frames were visually reviewed without editing their pixels.

Final menu-resource Debug build `f03efff201e9db0e6be38444cd2733b5619a213899ce9752a95c2e0766a83b41` passed with zero warnings/errors. Fresh workspace analysis/export, all 13 lifecycle captures and the exact approved keyboard/menu/Escape case passed; exact saved focus was restored. The earlier three-process/large-summary evidence stays bound to its own receipts. `menu-closeout-native-delta.json` proves that only five native RC files changed after that evidence; all C++/header/project inputs remained byte-identical.

Final menu-resource Release build `46dd52ba30046d8fe5435584feab8a9dc2fd8f01e01cadd2627181c6f7c7e466` also passed with zero warnings/errors. Its fresh workspace analysis/export, all 13 lifecycle captures and exact approved keyboard/menu/Escape case passed, with exact saved foreground and keyboard focus restored. Root Debug and isolated Release remain separate physical receipt domains. The isolated Release checkout is retained for review; the earlier long-path failed checkout was archived after preserving its compiler evidence.

Final PrePush accounting was Explain only: PR equivalence could not be verified, full local obligations remain selected, and the five broader foreground families remain pending. The 15-entry affected run and 68-case scoped repair repetition are the actual executions recorded here. No hosted CI, broader interactive, ARM64/live-provider matrix or whole-viewer performance gate was qualified.

## Publication preparation

Before publication, the repository-pinned clang-format 22.1.3 found spacing differences in nine changed native files. Formatting preserved line counts and every non-whitespace character; `logs/merge-format-source-delta.json` binds the exact before/after bytes. All 21 changed native source/header files then passed the pinned formatter. This is a formatting-only delta after the final Debug/Release receipts above; those original runtime results have not been rebound to a new build or claimed as a new execution.

After formatting, the six owning source/resource/documentation/tooling contract files passed 279/279 with zero skips (`logs/merge-format-source-contracts-results.json` and the full log). Specification inventory again reported zero blocking findings. Changelog folding/checking remained valid with no pending fragments. The final publication PrePush invocation remains Explain-only; the original affected failures and broader qualification limits still apply.
