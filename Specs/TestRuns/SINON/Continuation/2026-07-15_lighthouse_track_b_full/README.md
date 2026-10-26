# Lighthouse Track B — Full-suite closeout

- Date: 2026-07-15
- Branch: `codex/lighthouse-settings-recovery`
- Full command: `./Tools/Run-AllTests.ps1 -Suite Full -SkipBuild`
- Full run id: `20260715T073255Z-75940-7a01dc26022d475da152a1264c3ac70d`
- Duration: 2,335,619 ms (38m 56s)
- Observed result: 1,146 passed, 6 failed, 53 skipped

No SettingsStore or SettingsHotReload case failed. The Full manifest reported:

- `Commands`: one failure, `cmd_pane_command_line_insertion_and_execute`; its focused rerun passed 1/1.
- `ViewerPETests`: outer harness exit 1; a sequential rerun reproduced exit 1 after its inner ViewerText run reported passed.
- `LocalizationTests`: access-violation exit `-1073741819`; a sequential rerun reproduced it before output.
- `RedConfigureTests`: access-violation exit `-1073741819`; a sequential rerun reproduced it before output.
- `RedSalamanderMonitorEtwLatency`: exit 1, matching the pre-Track-B Full baseline blocker.
- `ToolsPesterTests`: three stale inventory assertions. Track B corrected Commands from 699 to 704 static
  registrations, from 800 to 805 runner-listed cases, and the Settings family from 92 to 97. The complete
  tooling Pester suite then passed 262/262.

Track B gates after reconciliation:

- Debug x64 RedSalamander build: 0 warnings, 0 errors.
- Resource localization contract: 5/5.
- Test inventory contract: 5/5.
- Commands `settings_store_`: 16/16 (final run id
  `20260715T082211Z-60700-cced28aa7ac14680874b43d979fb4036`).
- Tools Pester: 262/262.

Evidence limitation: the unified runner's stale-run cleanup removed the raw Full run directory when the
focused Commands rerun began, before it was copied into this durable archive. The exact manifest totals,
run id, duration, six failure records, and classification (`UNCLASSIFIED_FAILURE`) were read before cleanup
and are recorded above. The focused SettingsStore aggregate is preserved alongside this note.
