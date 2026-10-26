# I26 ordinary-action repair — focused candidate passed

Same-machine, x64 Debug, native 144-DPI comparison against
[settled baseline v2](../2026-09-19_102600-i26-actions-baseline-v2/README.md).
The only production change gives ordinary task actions their own 24-DIP row instead of deriving
their height from absent conflict rows. This does not implement the broader adaptive redesign.

The build passed with zero warnings/errors. Setup, `Phase6_PopupSmokeResizeAndPause` and Cleanup
all passed (3/3). The case checks positive bounds and matching hit targets after layout settles,
then bulk pause, narrow resize, resume with observed progress, and cancellation. During discovery
the eligible primary action is Skip; afterward it is Pause.

- Runner: `20260919T102714Z-38180-1be50e724bfc4970a5b39cd378281b47`.
- Receipt: `015eba17988d2ec8de138e64a973891a3825165cd3d33dea9f6fa14f67cb3f71`.
- Source snapshot: `5ec74db20db50bacd61c7b1fcae88b14d1e509d0e5428154700e107bc806c5e4`.
- Fixture SHA-256, identical to baseline: `91364B97CF190E44C930ED1C68E99C6832135EF243F8A1A346513DAB7A82AE58`.
- Toolchain identity, identical to baseline: `ec868e067972f0bd3f7d27e3d565da1e7fc42e307eac23af2de0c6400f25e7db`.

The command/profile is unchanged from baseline. Four warm-up redraws precede five batches of 16:

| Redraw batch, microseconds | Baseline | Candidate |
| --- | ---: | ---: |
| 1 | 166907 | 133352 |
| 2 | 141477 | 132817 |
| 3 | 159969 | 133785 |
| 4 | 151136 | 132308 |
| 5 | 130296 | 135274 |
| Median | 151136 | 133352 |

No increase was observed in this synchronous redraw metric. The lower median is not a claim about
display FPS, input latency, resource retention or the completed redesign; those gates remain open.
The candidate's full 10,764,336-byte/28,840-row stream stays under the authorized local evidence root.
`perf-curation.json` binds its SHA-256, byte/row counts and exact retained row indices. The curated
JSONL preserves all five claim-bearing rows in original order. No other metric is claimed here.

The separate sandbox audit reports 225 unrelated pre-existing run directories. The runner returned
success, but this record does not call that audit clean or constitute Full repository closeout.
