# Natural breadcrumb sibling sorting: focused Debug evidence

Date: 2026-10-17. Scope: Windows x64 Debug, same host and 12-/1000-folder fixtures.
The mixed fixture includes numeric runs, mixed case, dot names, leading zeros and Unicode.
The native query uses a warmed DirectoryInfoCache; the case proves no query-time enumeration.
Each phase retains all 32 samples per fixture, including the first sample and maximum.
Nearest-rank p50/p95 are recomputed from the original traces; timing-comparison.json retains every sample.
These measurements cover the cached query, not menu painting, startup, Release, or whole-application performance.

| Phase | Case result | Build receipt | Native run |
| --- | --- | --- | --- |
| Old lexical comparator | FAILED as expected: Folder10 before folder2; 0 passed, 1 failed, 0 skipped | a4b1a6740e6dffb90da43da559cafc405a127ef31debf4facad22b6ac95a4543 | 20261017T160547Z-85008-039b803c38b84950ae38676d58c8a639 |
| Windows natural comparator and bounded heap | PASSED: 1 passed, 0 failed, 0 skipped | 45ba644e093f9c38bb8a9fa2d721895e2f9e2b5b63606d2a62b28a9096ab7a10 | 20261017T161239Z-96908-1c8cf00273f545189e4cb8e5623dcdb9 |

| Fixture | Old p50 / p95 / maximum (ms) | Candidate p50 / p95 / maximum (ms) |
| --- | --- | --- |
| Mixed 12 folders | 0.2060 / 0.3983 / 0.4017 | 0.1927 / 0.2777 / 0.3021 |
| Bounded 1000 folders | 9.5746 / 10.3279 / 10.9642 | 7.4102 / 8.1884 / 8.4112 |

The combined comparator/selection change improves this measured query. The data does not isolate the
individual contribution of Windows comparison versus the heap. Retention remains bounded to 99 paths;
current-folder priority affects retention only, and final display order is natural. The regression covers
empty/exact-99/oversized sets, files excluded, multi-part numbers, Unicode, leading zeros, current folder
inside/outside the first 99, an absent current folder, and case-insensitive current-folder matching.
Both sibling popup builders call the same production query; the test-only entry creates no HWND and
never installs provider menu callbacks. No physical input or desktop screenshots are used by this case;
the Commands runner retains its governed warning/desktop lease.

Original native results, aggregate results, traces and source/runtime bindings are retained separately
for both phases. artifact-bindings.json verifies their byte-preserving copies; .gitattributes disables
text normalization for this archive. tested-code.json records final C++ file hashes.
The initial fixture build failed on the Windows small macro before a native run; its diagnostic is
retained as initial-fixture-build-failure.txt. It is not presented as the valid native baseline.

Final source/documentation contracts: 222 passed, 0 failed, 0 skipped. Changed-range clang-format and
git diff --check pass. Release remains deferred because the user chose to keep that instance running;
no Full, ARM64, ASan or broader product qualification is claimed. Both native runs reported 41 disk-audit
issues in the surrounding sandbox; neither scoped case is presented as sandbox-audit closeout.
Final validators: Get-SpecInventory -FailOnFindings reports zero blocking findings. Test-TestRunArchive passes explicit pre-stage validation for all 19 packet files and whole-inventory validation for 13844 checked-in files.
