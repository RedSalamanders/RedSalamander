# Actual Curl socket trace — broad Release remains red

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -TestRoot C:\RedSalamander.Perf -CaseFilter C0_ -RunId c0-curl-socket-trace-broad-release-20260906`.
Same receipt/source as the sibling lookup archive, which retains the tested patch.
Result: **15 passed / 1 failed / 1 not reached**, Copy **132/142**, 46 scenarios,
wrapper 76.1 s. The broad lookup is explicitly not reached, not passed.

Five ConnectionFailure rows report OS 10048. Seven numeric socket rows correlate
in original order (local → remote; same thread/scenario):

| Scenario | Owned source / destination control | Failed local → remote | Observation |
|---|---|---|---|
| depth128 MoveItem serial | 56296 / 56314 | 49154 → 49153, twice | Other port after passive reply 227 |
| depth128 MoveItem parallel | 49428 / 49708 | 49337 → 49428 | Source control |
| depth128 MoveItems serial | 49708 / 50637 | 49284 → 49708 | Source control |
| depth128 MoveItems parallel | 50637 / 51272 | 49152 → 50637 | Source control |
| depth129 CopyItem serial | 58410 / 58424 | 49292 → 49291, twice | Other port after passive reply 227 |

The duplicate socket rows are not extra operation attempts. Five fresh bind and
ten fresh implicit/explicit connect probes succeed. These observations do not
prove the original tuple was available or that explicit client bind would fix it.
Serial depth128 Move publishes once then returns partial-copy 0x8007012B; the other
three Move variants publish nothing. The Move success-only truth oracle failing
does not independently prove dishonest partial-result feedback.

Parallel/serial wide copies both publish all 4097 files with S_OK and exact data,
in 6463328 / 20703186 us. **Serial still fails** its fixture accounting guard:
destination retires one pending passive listener (3 creations, 12290 reuses).
Do not call this a green wide run. The prior Debug timeout is also still open.

Next: test exclusive ownership for the fake control/data listeners, then repeat
this exact broad workload; do not change production socket/retry policy or loosen
the listener cap/deadline. [Microsoft's socket guidance](https://learn.microsoft.com/en-us/windows/win32/winsock/using-so-reuseaddr-and-so-exclusiveaddruse)
requires setting SO_EXCLUSIVEADDRUSE before bind for server ownership. This is a
fixture correction hypothesis, not an established cause of these failures.

Exact results and all selected rows are preserved with source hashes/order.
142 retained-run disk-audit findings remain visible. No live-provider writes.
