# Preferences teardown qualification

The two teardown defects are corrected: the wheel-routing subclass forwards
WM_NCDESTROY to its saved procedure, and pane models/delegates disconnect their
borrowed controls before the shared page-host tree is destroyed. The earlier
failing host-count witness and intermediate General-only result remain in
[the first archive](../2026-09-09_152332-DxUiPreferences/README.md).

The expanded regression opens/closes Preferences 48 times, rotating through all
fourteen categories. Every close returns native hosts and thread graphics
attachments to their baselines of 7. `retention.json` retains every measured row
and the raw trace hash; elapsed times are diagnostics, not throughput acceptance.
`all-pages-results.json` records the exact case outcome.

The complete Preferences family passes 176 cases with zero failures and one
foreground-dependent skip in 304.352 seconds. `family-results.json` retains the
skip reason and every case. These are x64 Debug diagnostic runs against DxUi
3dae073; the full solution build had zero warnings/errors. Focused harness tests
also pass 202 cases under PowerShell 7. Fresh Full and the remaining native profiles
are separate gates and remain open.

The first-chance crash evidence for the pre-fix pane destruction order remains in
`Z:\RedSalamander.Perf\evidence\I19-20260909\preferences-family-crash.log`.
All generated raw artifacts stay within the governed test root.
