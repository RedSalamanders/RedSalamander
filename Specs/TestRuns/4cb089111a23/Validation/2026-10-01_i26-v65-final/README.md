# I26 V65 final evidence packet

This is the root-reviewed V65 x64 evidence scope. The packet preserves original historical outcomes and does not itself close separately owned gates.

## Current V65

- Release Full/Fresh: PASSED 2211/0/46; run `20261001T115346Z-104056-8e2cfb6a0b8049218c622dd915a2697e`; [terminal result](full/run-all-tests-results.json). Native results are ZIP member `full/native-results.json` in the [lossless member manifest](compressed/run-text-evidence-manifest.json).
- Source snapshot `f2e0a18a04fbc8a95368adc0c946bafb5ebccdf68d81b4a476d0fefbd323593c`; post-Full receipt `7442adedcaf17d954a58b726f49756312beabb2c2426589e4fd71af42584fdac`; [receipt](full/build-receipt-post-full.json).
- Separate tests-enabled x64 Debug and Release build receipts: [Debug](build/debug/build-receipt.json), [Release](build/release/build-receipt.json).
- Debug hosted decision keys: 2/0/0; run `20261001T113214Z-63820-2c8d88dff6c54bda82569d94542cb2ec`; [root review](validation/v65-hosted-keys/debug/root-review.json).
- Release hosted decision keys: 2/0/0; run `20261001T112207Z-80192-1f5f873df736462f91655cb76638d1a9`; [root review](validation/v65-hosted-keys/release/root-review.json).

## Acceptance and resources

- Reviewed 39-ID case map ZIP member `reviews/v65-acceptance-case-map.json` in the [lossless member manifest](compressed/run-text-evidence-manifest.json), [actual case binding](reviews/v65-acceptance-results-root-reviewed.json) and [independent binding root review](reviews/v65-acceptance-binding-root-review.json). The mapping does not establish whole-facet or performance acceptance.

- [Matched Embedded A78/B19 root interpretation](integrated-pair/root-review.json) and [descriptive comparison](integrated-pair/comparison-summary.json) cover the named equal-feature presentation/resource fixture, not every product path. Final raw comparison CSVs are exact ZIP members under `integrated-pair/`. The [original ProgramDatabase pair](historical/programdatabase-pair/root-review.json) retains its initial review as historical evidence; its comparison CSVs, summary and README are exact ZIP members under `historical/programdatabase-pair/` in the [lossless member manifest](compressed/run-text-evidence-manifest.json). Its build profile differs from Full and grants no final resource acceptance.

## Carried and historical evidence

- Original V55 broad Debug FileOps 197/0/20; run `20261001T010532Z-89108-35279440c388486683c7646e58d45f24`; original runner evidence is ZIP member `carried/v55-debug/fileops/runner-results.json` in the [lossless member manifest](compressed/run-text-evidence-manifest.json). [Gallery provenance](gallery/README.md) links the separately published original PNGs; V65 did not recapture them.
- Original V56 seven-case cluster: 21/0/0 across three repetitions; run `20261001T032939Z-69528-0f2e124ea9dd4721a752ccd97d5b9d93`; [root review](carried/v56-cluster/root-review.json).
- Historical V55, V57 and V58 failed Full runs remain failures; the [V64 Full root outcome](historical/v64-full/root-review.json) retains its original verdict and has no V65 acceptance effect.

## Packet boundaries

- The [lossless raw-result ZIP member manifest](compressed/run-text-evidence-manifest.json) records original member paths, byte counts and SHA-256 for [the archive](compressed/run-text-evidence.zip). The original run trees remain outside this compact packet.

- Native 192-DPI and external AT/screen-reader execution remain separately conditional. HWND message tests do not prove OS-injected physical keyboard/pointer input. I18 confirmation authority and H4/consumer ARM64/ASan adoption are separately owned.

- The generated curation-manifest.json records every packaged file and its source identity.
