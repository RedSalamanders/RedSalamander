# Metadata sink repair — wide run exposes loopback connection exhaustion

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_CurlCopyTraversalTruth -FailFast -RunId c0-curl-copy-metadata-sink-debug-20260905`.

Based on `cf5b9b8e`, with the explicit constant-storage size-probe body/header sink,
compile-time count/overflow checks, diagnostic phases/codes and corrected test
aggregate braces. The corpus, timeout and exact-content requirements are unchanged.
Build: **03:01, zero warnings/errors**.
Receipt: `b1bf92ed8839e86a12256d448eab717baf8076bc7f2dafd0dceeb2a97e1b9299`.
Workspace: `d6eed57c28e50cb74f1c1b096ed2edb1cf4569498773e4d5b245477ab3a449be`.

Result: **121/139 assertions; 32/46 scenarios pass**, wrapper 2 passed / 1 failed /
0 skipped, 131.2 s wall time, native case 123.5 s. All 32 shallow/literal/strict
listing/early-late refusal/cancel scenarios pass before the wide corpus.

The immediate raw-code-23 metadata failure no longer occurs. Parallel wide Copy
publishes 3,622/4,097 files in 108,113,260 us before ERROR_CONNECTION_REFUSED;
failure stages are upload, source-download and reserve-staging-name. Its source
endpoint also reports ERROR_SEM_TIMEOUT for an idle control session. Serial wide
Copy then fails to connect, publishing zero files in 1,052,887 us. All 12 following
depth cases encounter connection refusal before their intended boundary checks;
they are failed coverage, not passing depth evidence or an established recursion
regression. Copy completion matches actual HRESULT; Move expected-success receipts
remain failed assertions. Do not reduce or disable any of these checks.

Windows System/Tcpip **event 4227 at 16:20:32.7266208 UTC**, inside the parallel
wide run, reports outgoing local-endpoint reuse/high connection churn. The bounded
event and later aggregate socket snapshot are retained in `tcpip-diagnostic.json`.
The post-run count is not a peak measurement. Unlike the earlier TIME_WAIT-only
observation, this event directly witnesses a local connection-allocation failure.
No network setting was changed, no unrelated process was stopped and no retained
test data was deleted.

Next separate step: stop the fake FTP endpoint from allocating a new passive
listener port for every transfer, retain fresh data sockets and stale-request
rejection, and handle intentionally idle control sessions as protocol closure.
Measure listener allocation/reuse and preserve the identical 4,097-file corpus.
Repeated full parent-directory publication lookups still need bounded traversal;
neither this sink fix nor a fixture repair proves their memory/throughput contract.

Full raw perf: 196,679 bytes / 530 rows; SHA-256
`83f80bb371bf97a89c54dff88d41934a419adc2b419f41ded398c02903f04068`.
Results/reasons and both traces are retained. Source/localization: 197/197,
18.69 s; spec inventory: zero blocking findings. Disk audit: 71 issues, none deleted.
The metadata dependency is repaired; the complete gate and Release remain open.
