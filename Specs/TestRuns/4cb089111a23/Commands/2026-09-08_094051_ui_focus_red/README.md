# Deterministic UI focus regressions: RED

Run `20260908T094051Z-82264-2e1ce9c3c0b0448f87588b9a0e7794e1`, x64 Debug; receipt `081788386765d45180f9168154555a6c568e65238635af899ba1f6242fba44e7`.
Command: `.\Tools\Run-AllTests.ps1 -Suite Commands -CaseFilter cmd_reliability_ -TestRoot Z:\RedSalamander.Perf`.
Source: `892f9235` plus regression fixtures and the one-shot popup-destruction seam only; neither fix was applied.

Zero passes, two expected failures, zero skips. Keyboard native focus is displaced onto the category host while logical search focus remains; the old wait wrongly accepts that live HWND. Popup teardown dispatches its queued restore reentrantly, then clears retiring native focus; the old implementation leaves focus null after Escape, matching the archived BR-7 failure signature. Its immediate focus-follows-pointer predecessor also executes inside this controlled case.

The two ordered metrics are compacted from the original stream with digests and counts in perf-compaction.json. They are correctness/wait-bound witnesses, not a general throughput comparison. No raw trace has been reconstructed.
