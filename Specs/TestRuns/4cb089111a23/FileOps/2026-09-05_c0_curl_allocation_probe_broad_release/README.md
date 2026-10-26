# C0 allocation probe: fresh binds succeed after connection failures

Native suite: **15 passed / 1 failed / 1 skipped**, 81.0 seconds. Copy finishes
all **46 scenarios** at **134/142** checks. Lookup is not reached after failure;
its separate exact control passes 250/250. The [same Release build and source](../2026-09-05_c0_curl_allocation_probe_lookup_release/README.md)
retain the original topology, corpus, order, deadlines and listener cap.

Both wide copies publish and byte-check all 4,097 files: parallel 7.073610 s and
serial 23.703416 s. Four subsequent depth-128 Move scenarios fail with cursor Curl
7 / Windows 10048. All four have unavailable easy-handle local/remote ports (-1),
response code zero, and zero reported publications. Their expected success/content/
receipt checks fail. Depth-129 refusal cases pass. No throughput gain or transport
repair is claimed; the earlier [unprobed failure](../2026-09-05_c0_curl_listing_broad_release/README.md)
remains preserved.

| Failed route | Fresh bound port | Probe duration | Probe result |
|---|---:|---:|---|
| MoveItem, concurrency 1 | 49239 | 56 us | S_OK |
| MoveItem, concurrency 4 | 49176 | 624 us | S_OK |
| MoveItems, concurrency 1 | 49201 | 619 us | S_OK |
| MoveItems, concurrency 4 | 49277 | 575 us | S_OK |

Each probe follows its failure row on the same thread. This contradicts the simple
claim that every fresh local bind was unavailable at those observations. It does
**not** prove which implicit-connect tuple failed, rule out transient pressure,
or establish a leak. Concurrent releases and the probe's small allocator cost
remain caveats. Next distinguish implicit connect from explicit wildcard binding
using only owned loopback diagnostics; do not change machine ranges, enable socket
reuse speculatively, or treat this observation as permission to replay a mutation.
No build, policy check or OS socket-table sampling overlaps this runtime.

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter C0_ -RunId c0-curl-allocation-probe-broad-release-20260905
```

Raw artifacts remain under that governed run. Exact result/trace/fingerprint copies
preserve all failures and skipped reasons. The complete ordered Curl/C0 projection
binds raw bytes/rows/hash, selected line numbers and selected hash. C0 stays open;
normal settings, secrets and real server data are unchanged.
