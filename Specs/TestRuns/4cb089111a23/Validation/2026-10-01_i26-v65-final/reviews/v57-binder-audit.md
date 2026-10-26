# V57 acceptance binder audit

Scope: read-only review of `bind-v57-acceptance-results-draft.ps1` against the frozen V55/V56/V57 source reviews and V55 gate-routing evidence. No runs were started. The final V57 Full was still running, so its result was not evaluated.

The binder does several important things correctly: it pins the reviewed 26-case map and V55/V56/V57 source manifests/reviews; verifies the V55→V56→V57 frozen byte comparisons and reviewed delta counts; requires each mapped case entry to be passed; binds Commands entries to the Full run ID; and explicitly marks performance/resource and whole-facet acceptance as unestablished.

## Findings

- **P1 — Full build receipt is not bound to the V57 frozen source snapshot (lines 153–158, 180–189, 318).** The supplied post-Full receipt is accepted based on git head, platform/configuration/target, test enablement, exit code, and matching opaque `build_receipt_id`. Its `source_snapshot_id` is only copied to output; it is never compared with an expected snapshot derived from the V57 frozen inputs or checked against an independently verified build attestation. Thus a receipt for a different dirty workspace at the same head can satisfy this binder while the 26 case results are presented beside the V57 source identity. **Repair:** require the final Full v2 result and post-Full receipt to name the same source snapshot, then verify that snapshot against the V57 frozen manifest/attestation (or require the runner’s verified receipt record that already proves this exact source identity). Keep the existing pin and profile checks.

- **P2 — The carried V55 Debug receipt can pass with an incomplete suite (lines 207–208, 298–303).** It checks only suite name, zero failures, and at least one pass, plus the fixed run ID. It does not require the root-reviewed 197 passed / 0 failed / 20 skipped counts from `v55-final-gates-routing.json`, the complete expected result set, or the clean-sandbox/build receipt provenance recorded there. A truncated or altered `results.json` with one passing case and no failures could therefore still be labeled a carried Debug FileOps receipt. **Repair:** bind the expected counts and clean-sandbox state from the reviewed V55 routing record; require the full case/status/count consistency and the matching V55 build receipt identity. Pin or independently verify the archived result bytes if the artifact store does not guarantee immutability.

## No finding

The V56 fixture/source carry-forward is checked as source equivalence: the V56 and V55 manifests are pinned, copied source bytes are rehashed, changed paths are limited to the six reviewed Commands/docs paths, and the pinned V57 review ties to the V56 review hash. That proves the declared source carry-forward; it does not claim the V56 seven-case execution ran, and this binder’s output should continue to describe it as source equivalence only. The V55 gallery checks also require their exact named case to pass. The external resource-pair and whole-facet acceptance remain explicitly false/unestablished in the output.



## Follow-up review — binder SHA-256 `b5efcb6cd41e32534c2c9464e67b91a0b713275f6a1060408fd36686e5890fa5`

**Both prior findings are repaired in the reviewed checks. No remaining false-pass or provenance bug found in these changes.**

- The post-Full receipt now requires the reviewed build-receipt schema and V57 snapshot `dcd29a8210761877b254e93acd35f9224a7738c04f078067841da7a8daf445b8`; Full v2 `snapshot_id` must equal that receipt’s `source_snapshot_id` (lines 155–160, 182–185). An archived Full v2 result and matching post-Full receipt confirm these field names and equality are the actual schema contract. The final V57 receipt/result were still unavailable, so the fixed V57 snapshot value itself remains to be matched when that run completes.

- V55 Debug is now bound to the exact archived bytes: the native `results.json` hash is `51995765954c6df5cbb1ec850fb691d873e5c658120121554956f23bc230bf51`, and its actual case array is 217 entries (197 passed, 20 skipped, zero other statuses). The runner `invocation.json` hash is `ec356eae1d2e2641042d3c7eda1758dc14f8f3cc6901881ec1455d1083d7c744`; its run ID, clean sandbox, counts, before/after manifest and snapshot, and receipt ID/SHA all match the archived V55 evidence. The pinned receipt SHA `233cbdddec2a05e083e4037e434476cf1a2f03043989b9be3ba3328733f37d51` also matches `debug-v55-build/build-receipt.json`, whose receipt ID and source snapshot match the invocation. The native result has no receipt or sandbox fields, as expected; those claims correctly come from the pinned invocation.

The gate remains appropriately narrow: these checks establish V55 Debug provenance and the listed case bindings, while the binder still does not claim overall facet acceptance or resource acceptance.

