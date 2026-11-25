# I66 credential-wait progress

The same machine ran the x64 Debug File Operations cadence case without desktop activation.
Both runs passed setup, the case and cleanup. The baseline runtime predates credential-wait
projection. Its Stopping witness runs before cadence sampling. The candidate adds the credential
observer and held Local Copy witness.

| Variant | Tested head | Case archive | Verified build receipt |
|---|---|---|---|
| Baseline | `44f3b0262` | `20261125T080213Z-21792-269813d51151413b9fa4f631a7861284` | `29cab84ad3eae7c7da9442677f93d3f3afb1d011f0fe56f8655b4f8a5c188e55` |
| Candidate | `2968d4493` | `20261125T074040Z-49896-b85ccfe75a6f459589f1369af95d62a9` | `8c36561a7f9707153146bda5b37819f3b06baa06f496c28b0c5a5fdcf5d42266` |

| Observation | Baseline | Candidate |
|---|---:|---:|
| Synthetic running timer ticks / rate updates | 11 / 11 | 11 / 11 |
| Synthetic completed idle timer ticks / rate updates | 0 / 0 | 0 / 0 |
| Paused engine timer ticks / rate updates | 2 / 2 | 2 / 2 |
| Running engine timer ticks / rate updates | 11 / 31 | 11 / 22 |
| Running engine UI CPU time | 296,875 µs | 281,250 µs |
| Synthetic retained observer rows | 46 | 46 |

The candidate's held Copy recorded one entry and one exit for credential wait, then restarted
with one rate sample. Its separately checked active peer retained its ETA. These single
observations were gathered during parallel validation work. Allocation counts, per-sample
UpdateRates cost and live authentication remain open.

The archived JSONL retains FileOps.PopupCadence events in original order.

| Variant | Full original rows | Full original bytes | Full original SHA-256 |
|---|---:|---:|---|
| baseline | 7708 | 2754944 | `c5fa5d1e0d66f12ab2096e38b892df9964c3cd8ad87fa7d0541277265b2927cd` |
| candidate | 7520 | 2733755 | `f50efbeef32258b76a28caa660ff5ceef9b7ea4c1f042e92085da42bf9cdc2e6` |
