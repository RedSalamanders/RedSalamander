# C0 fresh-loopback connection controls: Debug

Native suite **3 passed / 0 failed / 0 skipped**, 21.9 seconds; exact lookup
**252/252**. Implicit and explicit-bound probes both connect successfully and
preserve Winsock error 10050. Observed local ports are 60904/60906 and durations
729/12,808 us. These are isolated controls, not a throughput comparison.

Test-enabled full-solution Debug build: **187,110 ms**, zero warnings/errors;
receipt `5aa5ca13ac691f3ee847dea141ebc8ec914186afe98a1e06484566a014d404df`.
It shares snapshot `71cf0be399fa0a06a45529a4a01eb4ddd39c6d22d825b8925a6a46efb2b7b926`
with the [Release control](../2026-09-06_c0_curl_connect_probe_lookup_release/README.md),
which retains the tested patch/source hashes and 197/197 policy result.

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter C0_CurlEntryLookupTruth -RunId c0-curl-connect-probe-lookup-debug-20260906
```

Exact results/traces/fingerprints have source hashes. The ordered Curl/C0 metric
projection binds raw bytes/rows/hash and selected lines/hash; build-log normalization
is explicit. No build or other test overlaps runtime. Raw data and 138 disk-audit
findings remain in the governed run. Broad Debug was not repeated for this
diagnostic-only step; its earlier wide-copy timeout remains unresolved. No C0 or
Full closeout is claimed.
