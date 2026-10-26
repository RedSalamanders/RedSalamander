# Preferences host retention

The exact-case regression `cmd_preferences_dialog_repeated_close_releases_hosts`
fails before the teardown correction: the first close leaves 10 attached native
hosts and 10 thread graphics attachments, against baselines of 7. After the
correction all 48 cycles return both counters to 7. The test preserves the settings
artifacts and in-memory settings snapshot.

`retention.json` preserves all 49 measured rows in order, source trace paths and
SHA256 digests. The two results files preserve the complete per-case outcomes.
Build logs and raw traces remain under `Z:\RedSalamander.Perf\evidence\I19-20260909`
and the run paths recorded in the receipt. These are x64 Debug diagnostic runs
against DxUi 3dae073. The Windows subclass now forwards WM_NCDESTROY through the
captured procedure; the page host reaches its existing drain and detach path.

This establishes bounded host lifetime, not a paired throughput or process-memory
budget. The failing run stops after its first cycle; elapsed times are diagnostic.
The Preferences family, Fresh Full verdict and other profiles remain separate gates.
