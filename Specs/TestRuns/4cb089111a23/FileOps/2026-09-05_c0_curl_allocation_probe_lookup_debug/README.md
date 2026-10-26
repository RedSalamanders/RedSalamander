# C0 allocation probe: Debug control

Native suite: **3 passed / 0 failed / 0 skipped**, 22.0 seconds; exact lookup
**250/250**. The new control obtains port 58653 and preserves the seeded Winsock
error 10050. The prior 249 checks remain unchanged. The test-enabled full-solution
Debug build completes in **186,436 ms**, zero warnings/errors, receipt
`f3a00106a30aa38f9dfa7ddd5aa3c4c91b07020ad469691ea5469df4ef4d8520`.
It shares snapshot `3a6f609f3361388a361e02c57d3f1dea590403b3ddbb414f3aa34908be75a74d`
with the [Release control](../2026-09-05_c0_curl_allocation_probe_lookup_release/README.md),
which preserves the tested patch, source hashes and 197/197 policy result.

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter C0_CurlEntryLookupTruth -RunId c0-curl-allocation-probe-lookup-debug-20260905
```

Results, traces and fingerprints are exact copies; the ordered Curl/C0 metric
projection binds raw bytes/rows/hash and selected line numbers/hash. Build-log
normalization is explicit. No build or other test overlaps runtime. Raw evidence
remains under this governed run; its 132 disk-audit findings are not removed.
Broad Debug was not repeated for this failure-only diagnostic step; the earlier
wide-copy timeout remains unresolved. This focused pass is not C0 or Full closeout.
