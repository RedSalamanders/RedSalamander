# Exclusive ownership — broad Release and failed TCP pairs

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -TestRoot C:\RedSalamander.Perf -CaseFilter C0_ -RunId c0-curl-exclusive-broad-release-20260906`.
Same tested source/receipt as the sibling lookup archive. **15/1/1**, Copy
**134/142**, all 46 scenarios, wrapper 82.9 s. Broad lookup is not reached.

Four depth128 Move variants fail before publication with OS 10048. Actual socket
diagnostics identify source-control pairs (local → remote): serial MoveItem
49152 → 57692; parallel MoveItem 49152 → 49449; serial MoveItems 49163 → 58276;
parallel MoveItems 49196 → 53482. Four bind and eight fresh-connect probes succeed.
The success-only Move truth checks fail; this alone is not a dishonest-error claim.

`failed-socket-table.json` preserves a read-only query of only those exact IPv4
loopback pairs: all four have a TIME_WAIT row. Queries ran sequentially from
22:44:34.6167486Z to 22:45:48.8637200Z, after the failures, not at connect time.
All owners are PID zero and are **not attributed to the test process**. This
correlation supports investigating retained/churning cursor connections; it is
not a proof of global exhaustion or the originating process. No socket was
terminated, reused or reconfigured by the observer.

Both 4097-file wide copies pass all existing contents/receipt/listener guards,
parallel 7426246 us and serial 23920633 us, with zero pending retirements. Compared
with the preceding trace-only 6463328/20703186 us run, this is not a throughput
improvement. The previous retirement failure stays archived; one absence does
not prove its cause fixed. Exclusive ownership is not sufficient to repair C0.

Exact results, numeric snapshot and all selected metrics retain source hashes
and original order. The preceding red baseline is committed in `2283f0cf`.
146 retained-run disk-audit findings remain visible. See the main sibling README
for build warning, source identity and the next production lifetime step.
