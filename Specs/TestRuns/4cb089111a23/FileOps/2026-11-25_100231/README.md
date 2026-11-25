# I66 Move source retention after metadata loss

The `FileOps_ReviewRemediationContracts` case uses a Local file with
`Zone.Identifier` and a Dummy final writer that proves the content digest.
It verifies the exact source bytes and stream, the committed destination
bytes, a localized metadata warning and the typed **Published + Retained**
result.

| x64 Debug observation | Source | Case run | Result |
|---|---|---|---|
| Baseline | `8ddd99570` | `20261125T094531Z-51276-4977f25a0eed42afa0549f61765a1446` | The old contract asked one metadata-loss question after the destination commit. The test preserved that failure. |
| Candidate | `6fefa8172` | `20261125T100156Z-3756-b8be69a3db7947c7a2e77153d66a0ab8` | Setup, the case and cleanup passed. The destination stayed verified and the source kept its stream without a second decision. |

Both builds emitted zero warnings and errors and both runs used no desktop
activation. The candidate artifact receipt is
`ff881106d2b06b27d4a2d1fd07c3098236b71dff57f0d87a48c6d3e96bd4f9b5`.

Candidate perf rows 3192 through 3195 record the 67-byte stream, one late
retention event for the Mark of the Web mask `0x2`, and successful provider
digest verification of the same 67 bytes. The comparison measures the late
decision and exact outcome. The runs used different local test volumes, so
the raw event timings describe their individual observations.

The original baseline `artifacts/selftest/last_run/perf/perf_metrics.jsonl`
has 3,430 rows and 1,275,994 bytes. SHA-256 is
`5e1eecb192d0c9649bc342ebb18918e9337b3abb897398986b1cb55b4ed0da70`.
The candidate file has 46,297 rows and 16,796,374 bytes. SHA-256 is
`924d51486b1a0dc374a3778434622b9f647c27e9ac86ac1ee746880b60e8f35d`.

Pre-publication risk consolidation and Skip item remain separate work in I66.
