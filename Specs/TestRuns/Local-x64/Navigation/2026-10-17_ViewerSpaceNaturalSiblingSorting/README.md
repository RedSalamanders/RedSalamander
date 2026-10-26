# ViewerSpace breadcrumb natural sorting follow-up

Date: 2026-10-17. Owner: the bounded breadcrumb follow-up in active I33. The completed I39 NavigationView record remains immutable.

A sorting audit found a private ViewerSpace breadcrumb folder menu still using CompareStringOrdinal. NavigationView's regular/full-path sibling menus already use Windows natural sorting. Main panes, Find text columns, address autocomplete, Batch Rename text columns and connection labels retain their separate lexical policies; this follow-up changes only ViewerSpace breadcrumb folder choices.

The plugin now uses StrCmpLogicalW with an ordinal tie-breaker for equivalent names. Literal names, generation-qualified targets, the existing 512-choice cap, file/synthetic-item exclusion and no-provider-I/O boundary remain intact. The existing nonactivating analysis-host fixture adds twelve retained mixed-case/numeric/leading-zero/multiple-digit-run/dot/Unicode directories under a retained child. The root's original 2002-row UIA/lifetime fixture remains unchanged.

| Native diagnostic phase | Outcome |
| --- | --- |
| Debug old comparator | Expected failure: four ordering checks fail; navigation, names/tokens, no-I/O, stale-generation and UIA/lifetime checks pass |
| Debug natural comparator | Ordinary and synthetic ambient-desktop cases both pass |
| Release natural comparator | Ordinary and synthetic ambient-desktop cases both pass |

summary.json records exact per-phase runs, check counts, test names and receipts. All four targeted builds completed with zero warnings/errors. Baseline test build receipt: 0a66197beddf05eade37ca6bf167d48554bbc333c107584e32a1dae48a6e52a3. Debug plugin receipt: ff01abe52754f0db033a336615eac60cda4c4f1ac6214aa62e0062f35024c55c. Release plugin receipt: be0d7bd4149f95d50df1d8d38c027a421e0388c52e13ea42681d5e8d87dfb67c. Release test receipt: b2f01c4e6165fe9f242734bd45bdcbdca28a703c3a4bf58dacd3a42c822a07fc. The baseline DLL hash matches its retained original app-build receipt. The identical Debug harness binary is used on both sides of the comparison. Current C++ hashes are in tested-code.json; phase results record exact test/plugin binary hashes.

The focused native diagnostics invoke the existing harness with --no-activate under the repository's artifact profile lock, contained-process launcher, marked Z:\RedSalamander.Perf root and isolated LOCALAPPDATA/TEMP profile. They are not a canonical Run-AllTests aggregate or Fresh Full. The ambient case perturbs only read-only desktop observations; it generates no physical input. Every candidate trace records zero owned foreground transitions. No desktop screenshot was used.

| Debug mixed-12 query, n=32 each | Old ordinal | Natural candidate |
| --- | --- | --- |
| p50 | 25.1 us | 28.8 us |
| nearest-rank p95 | 60.0 us | 41.2 us |
| maximum | 62.0 us | 63.6 us |

All samples, including the first and maximum, remain in the traces and timing-comparison.json. These are same-machine/same-fixture observations under ordinary host load: baseline coincided with scope explanation and candidate with Release preparation. The small median increase and differing tails do not establish an isolated optimization or regression; no popup-presentation, whole-viewer or paired Release performance claim is made.

There are 228 passing viewer/source/documentation contract checks with zero failures/skips. Changed-range clang-format and git diff --check pass. The first Release attempt was blocked by an independently running application and left it untouched. After the user agreed to close it, the successful plugin/test builds and native runs above completed. release-blocked-observation.txt is an observation of that guard output, not a completed build log.

Original traces, results, build receipts/logs and scope explanation are retained with SHA256 copy bindings in artifact-bindings.json. Build/test sources were not edited during compilation. Subsequent metadata edits only complete the bounded I33 subsection and preserve this evidence. Full, ASan, native ARM64, broad provider/DPI/locale/human qualification and pre-existing I33 gates remain separate; I33 stays ACTIVE.

Final validation: specification inventory has zero blocking findings; explicit pre-stage archive validation passes all 37 packet files, and whole-inventory archive validation passes 13844 checked-in files. Final Release test/plugin binary hashes and both current C++ source hashes match the passing native results. git diff --check passes.
