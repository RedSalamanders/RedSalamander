# C0 fresh-loopback connection controls: Release

Based on `70d41c2335b8db18ea119fc608e73f57c7f8d53d`. Native suite **3 passed,
0 failed, 0 skipped**, 20.0 seconds; exact lookup **252/252**. Implicit and
explicit-bound connects both reach a fresh owned listener, return S_OK/nonzero
local port, and preserve the seeded Winsock error 10050. Durations are 10,707 and
14,906 us respectively. These two observations are controls, not throughput claims.

Each probe creates and closes its own listener/client, sends no protocol bytes,
and attempts one nonblocking connect. A pending connect uses one select wait of
at most 100 ms, checks SO_ERROR and distinguishes setup/bind/connect/wait/completion
stages. Both modes run after the existing bind-only observation when a test-enabled
Curl attempt reports 10048. Neither reconnects the original peer or changes retry,
mutation, socket-reuse or machine policy. Caller Winsock error is restored after
all probe owners close. Normal non-test builds contain none of these probes.

Full-solution test-enabled Release build: **236,030 ms**, zero warnings/errors;
receipt `07e1568192a6e8c25c2d01369d2c7c58f82894d5fcbc8bedf71d8e6c953cf67b`,
snapshot `71cf0be399fa0a06a45529a4a01eb4ddd39c6d22d825b8925a6a46efb2b7b926`.
Source/resource policy: **197/197**, zero skips, 18.0852845 seconds. The first
policy invocation rejected unsupported Pester `-Show` before executing tests;
the corrected invocation uses installed Pester 3.4 `-Quiet` and retains its XML.
Policy overlaps the build, not runtime. `tested.patch` and source hashes preserve
the tested delta before final evidence-only documentation updates.

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter C0_CurlEntryLookupTruth -RunId c0-curl-connect-probe-lookup-release-20260906
```

Exact result/trace/fingerprint/policy copies have source hashes. The ordered Curl/C0
projection binds raw metric bytes/rows/hash and selected lines/hash. Build-log
normalization is explicit. Raw evidence stays in the governed run; its 135 disk
audit findings are retained. No build or other test overlaps runtime. No real
server data or normal settings changed. C0 remains open.

API semantics: [Windows nonblocking connect](https://learn.microsoft.com/en-us/windows/win32/api/winsock2/nf-winsock2-connect)
and [select completion](https://learn.microsoft.com/en-us/windows/win32/api/winsock2/nf-winsock2-select).
