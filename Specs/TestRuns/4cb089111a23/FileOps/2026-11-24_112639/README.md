# I66 Local Recycle ingress comparison

The x64 Debug host at f5aa393f0 used attested receipt
9fb00c721b37714ca6f6f3dff919e0ddb520d01a822043a34ecd0954a3ba6804.
The no-activation FileOps_ReviewRemediationContracts selection passed all three cases in run
20261124T102043Z-60344-b9aaba78a56e498da31ba7e65186af3b. Repository coverage remained NOT_EVALUATED.

The same process used four equivalent owned sets of twelve 4 KiB same-parent files.
The explicit per-item executor is the baseline. Each ingress issued one start question.
The three accepted sets were recycled and their test journal items purged. The canceled set stayed in place.

| Selection | Shell PerformOperations calls | Admission to settlement | Process handles before and after | Process private bytes before and after |
|---|---:|---:|---|---|
| Explicit per-item baseline | 12 | 297.178 ms | 2721 to 2791 | 386457600 to 379215872 |
| Find | 1 | 110.988 ms | 2789 to 2840 | 379215872 to 380989440 |
| Viewer Cancel | 0 | 5.893 ms | 2840 to 2840 | 380989440 to 381005824 |
| Viewer | 1 | 132.269 ms | 2834 to 2840 | 381005824 to 381546496 |

The exact call counts characterize the admitted routes. The timing and process snapshots are a single
paired observation. Process resources include other retained host and Shell state.

The marked sandbox holds the full raw input at
evidence/completed-runs/20261124T102043Z-60344-b9aaba78a56e498da31ba7e65186af3b/artifacts/selftest/last_run/perf/perf_metrics.jsonl.
Its SHA-256 is 4f8bac636a93987854ebcd88d6f030078bcac8f72d764a0dfe49c2e6ad5ebde0 and it contains 16081473 bytes across 45597 rows.
The 26 Recycle rows from original lines 1852 through 2774 are retained in their original order.
