# IMAP structural attributes: Debug candidate

Focused result: **208/208 assertions pass**, three runner cases passed and none
failed/skipped. All 12 existing standalone IMAP helper functions pass after
verification of all 177 receipt artifacts. Source/resource policy: 197/197,
no skips. The [broad C0 run](../2026-09-06_c0_imap_listing_attributes_broad_debug/README.md)
passes 18/0/0 separately.

Base: `f5858fbe0a56c99fd825958ae6434fda0e37e394` plus `tested.patch`
(LF-normalized Git diff) and exact working-byte `tested-source-hashes.json`.
Snapshot: `57bc7d9f6a1ddc2e80539eec07b74433272555dedc3d9cc4d842d78a580c756c`.
Full test-enabled Debug receipt:
`a3b75f8ef4e4590e2f4fb84014dd9839ba368322e37cdd2cb071b35485a2d319`.
Build: 187,108 ms, zero warnings/errors. No source/spec/index changes occurred
between build and runtime qualification; curated evidence is snapshot-excluded.

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild
-CaseFilter C0_CurlImapListingTruth -TestRoot C:/RedSalamander.Perf
-RunId c0-imap-attributes-candidate-debug-20260906`.

The [failing baseline](../2026-09-06_c0_imap_listing_attributes_baseline_debug/README.md)
passes 110/134, with 24 failures. This candidate retains all 134 and adds 74
controls: thirteen further four-check shapes, three four-check partial-summary
repairs and ten numeric-range checks. The native trace reports the exact count.

The parser borrows complete top-level attribute values; nested, quoted, literal
and opaque body data never supply metadata keys. Duplicate metadata and damaged
framing are rejected without resynchronizing into a fake FETCH inside a literal.
UID/sequence and size/literal numbers use their respective protocol ranges.
Missing flags are not unread, partial summaries trigger bounded repair, valid
empty headers are complete, and payload text is removed from anomaly logging.

`baseline-comparison.json` pairs all 38 original scenario measurements; the
candidate has 54. For complete 1/201/4,097-message listings, requests remain
1/2/37 and summary peaks remain 1/200/200. Debug durations are
1,069/5,992/115,815 us before and 1,012/5,418/113,100 us after: single observations,
not a throughput or percentile claim. Four quoted/literal/prefix/nested UID cases
now require one request instead of three. The three original incomplete-summary
cases now perform two requests instead of accepting incomplete metadata after one.
Persistently invalid sizes now use five requests instead of one, within the
existing per-listing repair budget. Rejected duplicate/truncated records use three
requests and retain no summary rather than manufacturing metadata.

Raw results/traces and ordered selected metrics retain provenance. The runner
reports 204 disk-audit findings, unchanged by cleanup here. This is scripted
parser/repair evidence, not live transport/TLS, whole-response memory, provider
mutation or final Full qualification.

Post-test checkpoint review: `checkpoint-verification.json` records 268 passing
checks over all five archives, exact raw-copy hashes, ordered metric projections,
receipt/results/native counts, twelve baseline/candidate source-manifest entries,
unchanged tested code and helper/policy evidence. Final plan and FETCH-range
wording annotations are recorded separately from runtime source identity.
`spec-inventory-summary.json` reports zero blocking findings and 52 informational
historical link findings; its full digest-bound report remains in the governed
test root. This is specification hygiene, not a whole-provider runtime verdict.
