# C0 fake LIST byte oracles: unchanged-generator Debug baseline

Based on `d63fbe5338b731beb6544de860e550bceb4729cd`. This run precedes the
parent/leaf/row allocation optimization. Native suite: **3 passed / 0 failed /
0 skipped**; lookup **249/249**. All previous 237 checks remain and the 12 new
serializer checks pass against the original implementation.

Eleven hand-authored byte oracles cover empty/root/nested listings, direct children,
ordering, UTF-8/percent/arrow names, zero/unknown/overridden sizes, newly added files,
and raw malformed LIST/NLST overrides. They call the same translation-unit-local
fixture serializer used by the network handler, not a second encoder. The twelfth
check renders every one of 4,097 files 64 times and compares each complete payload.

`FileOps.Curl.FixtureListing.Wide`: **918,520 us**, **64 renders**, **14,159,232
bytes**, S_OK. Timing includes output construction, full comparison and per-render
destruction; setup and metrics are outside the timed loop. It is a fixture-only
measurement, not production/provider throughput. No build, policy or OS socket
sampling overlaps this runtime run. The earlier broad Debug timeout remains
preserved in the [same-generator broad baseline](../2026-09-05_c0_curl_framing_wide_debug/README.md).

Test-enabled full-solution Debug build: **200,009 ms**, zero warnings/errors.
Receipt `6b89258ede712df430502e0d714202ad4559a17afd05dc7be737c3bae0b28a52`;
snapshot `aecba1336f356f438047d8e5144d701b401f2a6658745e0839cdddcc8c0f6468`.
The original source/localization policy XML records 197/197, zero skips,
16.3540461 seconds. Policy overlaps the build, not the measured runtime test.
`baseline.patch` retains the pre-optimization test/spec delta against the base;
`baseline-source-bytes.json` records the tested file hashes. The baseline patch
must not be applied atop the subsequent implementation without review.

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter C0_CurlEntryLookupTruth -RunId c0-curl-listing-oracles-baseline-debug-20260905
```

Reproduction requires the matching baseline build and a new run ID. Raw evidence
remains beneath the named `C:/RedSalamander.Perf/runs/` directory. Results, traces,
fingerprints and policy XML are copied exactly with hashes in `file-provenance.json`.
The metric projection retains all Curl/C0 rows in source order, binding original
bytes/rows/hash, selected line numbers and selected hash. Build-log normalization
and both hashes are explicit. No settings or credentials are archived.

C0 is not closed: broad Release allocation failures, broad Debug completion and
the remaining provider/host/plan gates need independent proof. Limits, server
topology, mutation policy, production parsing and normal connections are unchanged.
