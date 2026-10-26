# Atlas blocking Commands hang snapshot

- Captured: 2026-08-04 12:37:46 +02:00
- Blocking command: `.\Tools\Run-AllTests.ps1 -SkipBuild`
- Runner: PID 4744, started 2026-08-04 11:57:34 +02:00
- Child: PID 10884, `.build\x64\Debug\RedSalamander.exe --commands-selftest`, started 2026-08-04 12:01:07 +02:00
- Last trace write: 2026-08-04 12:30:35 +02:00
- Last started case: `cmd_pane_itemProperties_window_live_dx_interaction`
- Observed state: child reported responsive, its main window title was `Properties`, and the trace remained unchanged for more than seven minutes. `Run-AllTests.ps1` waits without a per-suite timeout, so the process retained `.build/artifact-operation.lock` and prevented the required Atlas Full run.
- Disposition: evidence was copied before terminating only the hung Commands child. Its existing runner retained the lock, proceeded through normal failure reporting and remaining-suite cleanup, and Atlas queued a clean `Tools/Run-AllTests.ps1 -Suite Full` invocation through the normal artifact lock.

Evidence:

- `commands-trace.txt` — SHA-256 `75810E54429DA3C7C980A8F706137CE3DB097AB47246206CE48CC3C3A97990AA`
- `commands-results.partial.json` — SHA-256 `11AAC01DCAF9FAEAA68D1688DFFB1323287BEA63CD8BFBBBBC0D0ADEE8FD0AF9`
- `plugin_config_live_step.txt` — last auxiliary live-step marker retained with the run

This snapshot records an infrastructure/selftest hang encountered while waiting for Atlas closeout verification. It does not establish a product requirement or a passing test result.

## Clean Full-suite reproduction

The exclusive Full run `20260804T105536Z-26856-ffa2b9e880f2497789895129e0c57a5c` rebuilt the complete Debug x64 solution in 5m43s with zero warnings and zero errors, then reproduced the same stall after 701 successful Commands returns. At 2026-08-04 13:38:01 +02:00, PID 6656 remained responsive with main-window title `Properties`, while the trace had been unchanged at the start of `cmd_pane_itemProperties_window_live_dx_interaction` since 13:34:49.

- `commands-trace.clean-full.txt` — SHA-256 `BF1D8EA7F6C8F7CB8B62CBB32284F1F6A1A0BF8F29E413B1D2A812A0A55D3E44`
- `commands-results.clean-full.partial.json` — SHA-256 `A800EEB2DFDAD08CBE18702D3CF44FFF8F8FCA97DFE078964DAE97C16920F9B7`
- `msbuild-20260804_125544_998.log` — SHA-256 `5620E2E68432435A6E480651178CEBF1DF2382A548196BE4C6BE833CD826DB15`

This second reproduction rules out the earlier concurrent rebuild as the sole cause. It is evidence of a current selftest/UIA hang, not yet evidence that the normal Item Properties product flow is defective.

## Stream Remove follow-on reproduction

The mutex-isolated Full run
`20260804T133910Z-29896-9e3819a148b84d5b9d7628b8a0899c3f`
rebuilt Debug x64 with zero warnings and zero errors and proved the corrected
`cmd_pane_itemProperties_window_live_dx_interaction` case passes in broad
order. The immediately following
`cmd_pane_itemProperties_window_streams_can_remove` case then stopped after
its direct `InvokeVisibleDescendantByName(...)` call. PID 48056 remained
responsive while the trace was unchanged from 16:17:10 through the bounded
observation window; only that Commands child was terminated so the runner
could record the failure and continue cleanup.

This was the same owning-thread UIA hazard at a second call site. The Remove
button invoke now also runs through `RunDialogUiaTaskWithMessagePump(...)`,
and the source-contract test covers both Item Properties functions.

## Root cause and focused passing proof

The selftest performed UI Automation client calls on the same UI thread that owned the Item Properties provider. Provider work marshaled back to that blocked owner thread, so the first live pattern query could deadlock. `RunDialogUiaTaskWithMessagePump(...)` now uses the shared bounded `IUIAutomation2` timeout, COM cancellation and join path, and all three Item Properties UIA operations run on that worker while the owning thread pumps messages. No product Item Properties behavior changed.

The required recovery build `build.ps1 -Rebuild -MaxCpuCount 1` completed in 20m43s with zero warnings and zero errors. The exact former hang then passed in run `20260804T121751Z-46268-66e42a6507a64c39ac85d996a832c3df`: 1 passed, 0 failed, 0 skipped, 9.320s native case duration. The trace proves both initial and reopened pattern reads, file-name lookup and close invocation completed before a clean shutdown.

- `commands-trace.focused-passing.txt` — SHA-256 `5C4980A1887A1ABD70F0B06C24C3A9D9B4069B5536DC1E41D1415245489F7769`
- `commands-results.focused-passing.json` — SHA-256 `176B02005E65E8E710BB6CFE6B07F6EC822F4B607E197393125EECB9522CF4EC`
- `msbuild-20260804_135701_198.log` — SHA-256 `4A818BF75C9930613190616183D9882BBA6AEF09A2E8CBAFDDAAE9D920BE889A`
