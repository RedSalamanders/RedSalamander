# File Operations Popup DxUi Host — Release

- Date: 2026-07-22
- Build: test-enabled x64 Release
- Source sandbox run: `20260722T030755Z-58672-91cf0a0db98e469892dd3832d2a3250e`
- Result: 3 passed, 0 failed, 0 skipped
- Command: `.\Tools\Run-AllTests.ps1 -SkipBuild -Suite Commands -Configuration Release -CaseFilter 'cmd_pane_fileops_conflict_prompt_metadata_and_actions,cmd_pane_fileops_popup_presentation_settings_and_taskbar,cmd_pane_fileops_completed_group_and_navigation' -TimeoutMultiplier 2`

The cohort covers Keep Both and conflict metadata/action layout; same-HWND hosted-control/UIA parity, RangeValue and Invoke patterns, keyboard focus, accessibility notifications, normal/reduced motion, minimize/timer/restore, focus-preserving density rebuilds, and persistence; plus completed grouping/navigation.

| Metric | Detail | Samples | Min | Median | p95 | Max |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `FileOps.Popup.HostedControls.SyncUs` | all | 172 | 97 us | 189 us | 991 us | 3,646 us |
| `FileOps.Popup.HostedControls.SyncUs` | rebuild | 25 | 419 us | 734 us | 1,150 us | 2,553 us |
| `FileOps.Popup.HostedControls.SyncUs` | update | 147 | 97 us | 178 us | 446 us | 3,646 us |
| `FileOps.Popup.Render.TotalUs` | all | 172 | 5,547 us | 44,650 us | 48,776 us | 66,818 us |

This is a candidate-only measurement for the new hosted-control path, not an improvement claim. The raw local JSONL contained 172 synchronized popup frames and was 2+ MiB; this curated archive retains the complete case result and compact distribution required to review the claim.
