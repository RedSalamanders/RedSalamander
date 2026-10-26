# Wide Copy diagnosis — FTP metadata writes to the default stream

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_CurlCopyTraversalTruth -FailFast -RunId c0-curl-copy-wide-diagnostic-debug-20260905`.

Based on `cf5b9b8e`, with content-free transport-code/phase diagnostics, parallel
wide Copy before serial, and completion consistency checked against the actual
HRESULT. The independent expected-success and exact-content assertions are intact.
No transfer behavior was repaired in this run.

Build: 03:08, zero errors, two C5246 aggregate-initializer warnings at the new
test concurrency-order declaration. The complete small build log is retained as
`build-warning.txt`; these warnings must be fixed, not suppressed.
Receipt: `0872afbfb94a2e892aafdf36c0a263c91e62ef456b4393ac9a24b1cd44113a00`.
Workspace: `58b3b12f8934ed6e2a1d20e3ac6ffda3418375ea4f2859bb739861929a2f2821`.

Result: **137/139 assertions; 44/46 scenarios pass**. Wrapper: 2 passed / 1 failed /
0 skipped, 35,795 ms; native case 27,828 ms. Both wide cases fail before completing
their 4,097-file corpus:

| Order / concurrency | Duration (us) | Published files | Failure |
|---|---:|---:|---|
| First / 4 | 105,565 | 34 | CURLcode 23, source-commitment, E_FAIL |
| Second / 1 | 303,806 | 49 | CURLcode 23, source-commitment, E_FAIL |

Both endpoints report S_OK and both completion callbacks match the actual failed
HRESULT. Thus a preceding serial timeout is not required to reproduce the failure.
This run does not establish the cause of earlier 120-second serial timeouts.

The failing shared size probe sets NOBODY but no WRITEFUNCTION. The bundled
[curl 8.21 FTP implementation](https://raw.githubusercontent.com/curl/curl/curl-8_21_0/lib/ftp.c)
explicitly routes synthesized SIZE/REST metadata through the body write callback
(`client_write_header`, including Content-Length and Accept-ranges). The default
[write callback is fwrite](https://curl.se/libcurl/c/CURLOPT_WRITEFUNCTION.html),
whose short return becomes CURLE_WRITE_ERROR (23). This connects the observed
source-commitment failure to the unintended default-output dependency; the exact
CRT stream condition was not separately captured.

Next: give every size-probe attempt an explicit constant-storage metadata sink,
retain error/size proof, repair the aggregate braces, and rerun the unchanged wide
corpus. Investigate any remaining long-run lookup/session limitation separately;
do not attribute this raw-code-23 failure to TIME_WAIT pressure.

Full raw perf: 198,291 bytes / 532 rows; SHA-256
`5bac8c34f61e76ac7383ea00e94b2ed646303ff8666ee0084da9566bae345ea1`.
All results/reasons and traces are retained. Disk audit: 70 issues, none deleted.
This diagnostic is failed evidence, not C0 or plan closeout.
