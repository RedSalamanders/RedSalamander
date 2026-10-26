# C0 rejected loopback candidate: all owned TCP states

Same frozen source and Release receipt as the [main checkpoint](../2026-09-05_c0_curl_loopback_topology_broad_release/README.md).
Broad C0 again fails: **15 passed / 1 failed / 1 not reached**, Copy **134/142**.
All 46 scenarios finish. Both wide copies publish exactly 4,097 files in 9.563265 /
33.961110 seconds. Four depth-128 Move routes fail and emit Windows 10048; depth-129
refusal cases pass. This is retained failure evidence, not a successful repair.

The serial-wide marker triggers one read-only query filtered to the observed
PID, verified against executable path and process start. Unlike the earlier
loopback-peer-only sample, this retains all 1,092 owned rows, including 362 BOUND
rows. The raw sample records trigger and collection interval. It is not a complete
lifecycle trace, does not attribute PID-0 TIME_WAIT sockets and does not prove
a leak or exhaustion. Sampling overlaps the timing; no throughput claim is made.

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter C0_ -RunId c0-curl-loopback-topology-owned-states-release-20260905
```

Use a new run ID for a repeat and the exact archived candidate bytes. That candidate
is now reverted, so rebuild before testing the current source. Raw evidence stays
under this governed run. `file-provenance.json` binds exact copied results, traces,
fingerprints and socket rows; the ordered metric projection binds raw/selected
counts, hashes and original line numbers. See the main checkpoint for remaining
work, authority boundaries and build/policy provenance.
