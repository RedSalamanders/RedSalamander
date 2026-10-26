# C0 broad Release: both connection probes succeed after each failure

Native suite **15 passed / 1 failed / 1 skipped**, 75.5 seconds. Copy completes
all **46 scenarios** at **132/142** checks; lookup is not reached after failure.
The separate [same-build lookup](../2026-09-06_c0_curl_connect_probe_lookup_release/README.md)
passes 252/252. Both wide copies publish and byte-check all 4,097 files: parallel
6.639306 s, serial 21.340659 s. These incidental timings do not prove an improvement.

Six captured Curl 7 / Windows 10048 failures each have a successful bind-only
observation and two successful fresh-listener connects immediately afterward on
the same thread. In original row order:

| Failure | Implicit local port / us | Explicit local port / us |
|---|---|---|
| 1 | 49156 / 16394 | 49156 / 14323 |
| 2 | 49179 / 14141 | 49179 / 1707 |
| 3 | 49152 / 23667 | 64781 / 1741 |
| 4 | 49192 / 21518 | 49192 / 1605 |
| 5 | 49197 / 25970 | 49197 / 16140 |
| 6 | 49192 / 1595 | 49192 / 11610 |

All 12 connects report stage `connected`, S_OK. Each uses a fresh owned listener;
the numeric ports are historical and already released. Sequential reuse of a
number is not concurrent ownership or an identification of the original tuple.
The first three Curl rows have unavailable local/remote ports and response zero;
the latter three report control-port pairs and response 227. These easy-handle
values do not identify the failed data socket.

All four accepted-depth Move scenarios fail their success/content/receipt oracle.
MoveItem serial has one publication and returns partial-copy 0x8007012B; the other
three report zero publications and 0x800704C9. Serial depth-129 CopyItem/CopyItems
return connection refusal rather than the expected depth refusal, with unchanged
contents and truthful no-commit receipts. Do not collapse this result into the
earlier four-zero-publication failure or claim all provider feedback is dishonest.

The result gives no evidence that explicit binding repairs the failure: both
connection modes already work to new peers. It also does not rule out transient or
endpoint-specific pressure. Next inspect the actual Curl socket creation/closure
and endpoint identity under the broad order; do not add a speculative production
bind/reuse option, repeat unknown mutations, or change machine-wide settings.
No build, policy run or OS socket-table sampler overlaps runtime. The diagnostic
itself adds the recorded sequential work on failures; preserve that caveat.

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -TestRoot 'C:\RedSalamander.Perf' -CaseFilter C0_ -RunId c0-curl-connect-probe-broad-release-20260906
```

Exact results/traces/fingerprints retain failures and skip reasons. The complete
ordered Curl/C0 projection binds raw bytes/rows/hash and selected lines/hash. Raw
data stays in the governed run, with its 136 disk-audit findings retained. C0 and
final Full qualification remain open.
