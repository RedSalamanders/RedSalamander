# C0 fake FTP listing allocation candidate: Debug lookup

Based on `d63fbe5338b731beb6544de860e550bceb4729cd`, with the uncommitted
borrowed-parent/leaf and direct-output formatting change. Native suite: **3 passed,
0 failed, 0 skipped**; lookup **249/249**, including all 12 new serializer checks.
The nine compile-time slice/self-prefix checks also compiled successfully.

Same-machine fixed 64-render / 4,097-file witness: **589,488 us** versus
**918,520 us** in the [original-generator baseline](../2026-09-05_c0_curl_listing_baseline_debug/README.md),
35.8% less elapsed time in this pair. Both render and compare **14,159,232 bytes**.
This is one fixture-only comparison, not production throughput or percentile evidence.
All eleven literal-byte rows pass. Source/resource policy passes **197/197**, zero
skips, in 17.3220974 seconds. Policy overlaps the subsequent Release build, not
either measured Debug runtime. No build or socket sampling overlaps this lookup.

Test-enabled full-solution Debug build: zero warnings/errors, **187,944 ms**.
Receipt `15c07d7e8de2a32e9b3d33023e48c93c858e8685cec91686004e1f9e9b8ef21e`;
snapshot `80931b760b2b6757c44172031d2ed2f143530d37d6497481791b2d0e1b0b576d`.

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter C0_CurlEntryLookupTruth -RunId c0-curl-listing-candidate-lookup-debug-20260905
```

Reproduction needs the matching build and a new run ID. Exact result/trace/policy
copies and hashes are in `file-provenance.json`. The ordered Curl/C0 projection
binds raw bytes, rows, hash, selected line numbers and selected hash. Build-log
normalization is explicit. Raw artifacts remain under the named governed run.
Normal settings, secrets, TLS/SSH policy and live mutation authority are unchanged.
`candidate.patch` and `candidate-source-bytes.json` preserve the exact tested delta
and raw file hashes before the final evidence-only plan/attribute updates. The
same source snapshot was built and tested in Release; see the linked sibling runs.
The [broad Debug failure](../2026-09-05_c0_curl_listing_broad_debug/README.md) remains
red; this focused pass does not close C0.
