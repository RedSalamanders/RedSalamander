# C0 listing allocation candidate: Release lookup

Native suite: **3 passed / 0 failed / 0 skipped**, 19.8 seconds; lookup **249/249**.
All 12 fixture serializer checks and the original 237 route/framing/refusal checks
pass. The fixed 64-render witness produces and compares **14,159,232 bytes** in
**23,748 us**. No same-configuration pre-optimization Release generator measurement
was collected; comparing this with Debug is not a Release speedup claim.

Test-enabled full-solution Release build: **247,604 ms**, zero warnings/errors.
Receipt `ad73628fbd5c9796ee1e377429641c52ba0d91c37010b8c097b2dd510500f967`;
snapshot `80931b760b2b6757c44172031d2ed2f143530d37d6497481791b2d0e1b0b576d`.
This is the same source snapshot as the [Debug candidate](../2026-09-05_c0_curl_listing_lookup_debug/README.md),
which retains the tested patch, raw file hashes and passing 197-check policy XML.
All builds complete before Release runtime measurements; no socket sampling runs.

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter C0_CurlEntryLookupTruth -RunId c0-curl-listing-candidate-lookup-release-20260905
```

Result/trace/fingerprint copies are exact and digest-bound in `file-provenance.json`.
All Curl/C0 rows remain in source order with raw bytes/rows/hash, selected line
numbers and selected hash. Build-log normalization and both hashes are explicit.
Raw data remains under that initialized sandbox run; reproduction needs the
matching source/build and a new run ID. Normal settings, TLS/SSH, credentials and
live endpoint mutation authority are unchanged. The broad Release failure remains
in its [separate archive](../2026-09-05_c0_curl_listing_broad_release/README.md).
