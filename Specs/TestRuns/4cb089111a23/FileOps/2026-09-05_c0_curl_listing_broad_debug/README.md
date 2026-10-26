# C0 listing candidate: broad Debug remains red

Native suite: **15 passed / 1 failed / 1 skipped**. Copy has **141/142** assertions
and all **46 scenarios** finish. Entry lookup is explicitly not reached after the
Copy failure; its separate exact run passes 249/249. Delete traversal 127/127,
directory size 85/85 and Move preflight 97/97 pass. These are focused C0 results,
not Full repository qualification.

The source/receipt is the [same Debug candidate](../2026-09-05_c0_curl_listing_lookup_debug/README.md).
Only test-server allocation behavior changed. Corpus, full scans, mutation proof,
topology, 120-second deadline and 64-listener cap are unchanged.

| Wide Copy | Earlier original generator | Candidate |
|---|---|---|
| Parallel, concurrency 4 | 4,097/4,097; 107.378122 s | 4,097/4,097; 71.420825 s |
| Serial, concurrency 1 | 3,675/4,097; timed out at 119.997979 s | 3,828/4,097; timed out at 119.987399 s |
| Parallel destination LIST generation + mutex wait | 99.063986 s; 8,195 LISTs | 57.010681 s; 8,195 LISTs |
| Serial destination LIST generation + mutex wait | 53.059216 s; 7,353 LISTs | 38.623266 s; 7,657 LISTs |

Baseline: [framing broad Debug](../2026-09-05_c0_curl_framing_wide_debug/README.md).
These same-machine runs have no overlapping build, policy check or OS socket
sampling. Parallel output is complete in both; serial corpus completion differs,
so its total cost is not an equal-work comparison. Server timings are nested
inside client phases; never add them to operation time or sum overlapping workers.
The serial staging-name/promotion phases still consume 51.059058/53.418168 s.
This change helps the fixture but does **not** repair the timeout or prove a
production throughput gain. Release connection-allocation diagnosis is separate.

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter C0_ -RunId c0-curl-listing-candidate-broad-debug-20260905
```

Raw evidence stays under that governed run. Result/trace/fingerprint copies are
exact with provenance hashes. All Curl/C0 metrics remain in original order with
raw byte/row/hash and selected-line/hash binding. Failure and skip reasons are
preserved. No settings, credentials or live endpoint mutations are involved.
