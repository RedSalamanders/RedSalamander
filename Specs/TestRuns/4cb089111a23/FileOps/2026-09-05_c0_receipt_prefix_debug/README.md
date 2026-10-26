# C0 bounded mutation receipt — Debug

Run `c0-receipt-prefix-debug-20260905`, `C0_`: 6/6 passed, no skips
(four behavior cases plus setup/cleanup). Build receipt
`6a2e447ae39f17f3ec187024c842d93e2b9eb25f91bb46ef5324df83f13535a2`;
zero warnings/errors. Base `58f22d0e` plus working changes is runner-bound.

The new actual-host callback witness places an exact V1 prefix immediately before
an inaccessible page. Five unsupported size headers are also placed at the page
edge. Full-stage copying, malformed Boolean/stage rejection, missing-receipt
clearing, future-suffix normalization, and malformed-callback error ignored by the
provider are exercised. The last case remains an immutable uncertain result.

`FileOps.SelfTest.C0MutationReceipt.PrefixBoundary`: 249 us; five rejected header
shapes and five host callbacks. This single sample is bounded correctness evidence,
not a callback percentile or a performance improvement claim.

Existing Copy/Delete committed-response-loss cases retain one request/commit each;
the two Local failure tasks retain exactly one ordinary overwrite prompt each and
preserve source/destination bytes plus independent result axes. The prior
`2026-09-05_c0_native_receipts_verified_debug` archive supplies the same-machine
three-case baseline; the additional prefix case has no runtime RED baseline.
The original unsafe full-struct copy was found by source review.

All four raw result/trace/aggregate/metric files are retained intact. The disk audit
flags 25 retained earlier run directories; they were not deleted, and no clean
test-root audit is claimed. Release, full provider coverage and final qualification
remain required.
