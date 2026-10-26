# C0 listing candidate: broad Release connection failures persist

Native suite: **15 passed / 1 failed / 1 skipped**, 80.7 seconds. Copy has
**133/142** assertions; all **46 scenarios** finish. Entry lookup is not reached
after failure; the separate exact lookup passes 249/249. Delete traversal 127/127,
directory size 85/85 and Move preflight 97/97 pass.

The [Release candidate receipt](../2026-09-05_c0_curl_listing_lookup_release/README.md)
binds the original `127.0.0.1` endpoint topology and the same source as Debug.
Both wide Copy cases publish and byte-check all **4,097 files**: parallel
**7.255687 s**, serial **22.202512 s**. Each destination handles 8,195 LISTs;
generation/mutex-wait totals are 2.233449/2.498240 seconds. These nested timings
are not additional operation time. No build, policy check or socket sampling
overlaps runtime. No general production throughput improvement is claimed.

Five `FileOps.Curl.ConnectionFailure` rows report cursor `CURLE_COULDNT_CONNECT`
(7), available OS error **10048**, unavailable local/remote ports (-1), and zero
response code. Four depth-128 Move scenarios fail, followed by the parallel bulk
Copy depth-129 refusal returning connection failure instead of its expected limit
error. The first Move has one publication and returns `ERROR_PARTIAL_COPY`; its
scenario content/receipt checks fail. Preserve these independent failures rather
than treating connection diagnosis as successful mutation proof. All endpoint
server statuses in those scenario rows are S_OK.

This reproduces the unresolved allocation failure, not the deterministic refused
endpoint's 10061 control. The available metrics do not identify the failed socket
tuple or prove the cause of allocation pressure. An isolated pass would not close
this broad-order failure. Corpus, full-list validation, limits and mutation policy
are unchanged; only the fake listing allocation optimization is retained.

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter C0_ -RunId c0-curl-listing-candidate-broad-release-20260905
```

Exact result/trace/fingerprint copies preserve failures and skips with source
hashes. The ordered Curl/C0 projection binds raw bytes, rows, hash, selected lines
and selected hash. Raw data remains under that governed run. No normal settings,
credentials or real server data were changed. C0 remains open.
