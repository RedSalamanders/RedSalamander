# V49 root review addendum

V48 was frozen at 82 logical / 81 physical inputs but did not compile. Its original Debug build log remains under `debug-v48-build/build.log`; no native test or acceptance was produced.

V49 changes in the primary checkout:
- Both new path-detail attachment assertions pass `GetCurrentThreadId()` to the pinned DxUi debug API, which requires a `DWORD` argument.
- Three `std::array` initializer lists use the inner aggregate braces required to avoid C5246.
- `ClosePathDetail` explicitly clears the host default/cancel aliases after resetting interaction state and before hiding the detail panel. The pinned host revalidates button tree membership, so the earlier code did not establish a dangling-pointer crash; immediate alias retirement makes the close/destruction boundary explicit.

Root independently compared every V49 frozen input with both the candidate checkout and the frozen source copy. `v49-root-frozen-verification.json` reports no mismatches, 82 logical / 81 physical entries and manifest SHA-256 `830183957ddc633079e9cee5649f4bc373568cfa811c46a0dfdc6462b94e1112`.

Root also reviewed `run-v48-commands-predecessor-clusters.ps1` SHA-256 `959e6928ed1a3a23ea0dfe7018e047be4cab4b45d712d9564e8ca766c496f0bc`. Its 24 exact case identities match the retained V47 six-failure/predecessor map. The Commands runner constructs selected execution from declared canonical registration order, with repetition as the outer loop; no shuffle is requested. The external driver requires all 48 executions to pass, each identity at repeat indexes 1 and 2, with no skip and a clean sandbox. Execution must override the script's historical V48 defaults with V49 roots. The synthetic 10k-item predecessor remains enabled; its optional 50k mode is not part of this diagnostic.

These are source and provenance reviews. V49 build, path-detail native behavior, predecessor clusters, broad suites, current gallery and paired resources are pending; this addendum does not grant acceptance.
