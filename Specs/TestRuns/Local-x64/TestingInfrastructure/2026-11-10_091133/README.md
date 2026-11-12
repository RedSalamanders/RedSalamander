# I61 local diagnostic and fixture qualification

Date: 2026-11-10. Profile: native x64 Release, test-enabled build. Scope: the
[organization acceleration plan](https://github.com/RedSalamanders/.github/blob/main/Specs/Plans/WIP/TestingAndIntegrationAcceleration_2026-11-09.md).
This is local functional/diagnostic evidence, not a performance comparison or hosted/native ARM64 acceptance.

## Implemented behavior

Terminal input-routing/lifecycle assertions use the shared stderr reporter, including
formatted and non-terminated string views. Rejected runner acceptance retains reason
codes, named unallowed skips and earlier diagnostics. The image stale-generation
reservation hook waits for explicit harness release rather than a 250 ms window;
its existing observation/cleanup assertions remain and a delayed-observer assertion
proves the reservation is held. The 30-second hook watchdog exists only in test builds.

## Executed checks

- The actual runner publication/plan/evidence controls passed 123 checks with zero
  failures or skips. Initial attempts using unsupported Pester assertion syntax or
  missing validation-profile setup are retained separately; neither is relabelled.
- The exact non-activating image resource/export case passed twice, including the
  delayed observer, stale reservation cleanup and long-path export. The pre-format
  focused image hashes and native exits are in the retained driver record below.
- Canonical formatting, inventory, specifications, static adapter and changelog
  checks passed. The subsequent broad run rebuilt the final formatted sources with
  zero accepted warnings/errors.
- `Test-Changes.ps1 -Mode PrePush -Configuration Release -Platform x64 -TestRoot
  Z:/RedSalamander.Perf` completed with 1,538 reported passed, zero failed and 30
  declared skips. All 1,243 Pester assertions passed with zero skips. Full plugin
  contracts passed, including all 162 Terminal checks. The runner archived the
  complete run and returned exit 2, `NOT_EVALUATED`, because four focus-taking
  entries remain pending: Commands, FileOperations, ViewerPETests.Interactive and
  ViewerSqliteTests.Interactive. This does not establish repository Full qualification.

## Retained local evidence

- Broad run: `Z:/RedSalamander.Perf/evidence/completed-runs/20261110T091133Z-80272-7a0d26771200472a88278596369b3097`.
- Broad driver log: `Z:/RedSalamander.Perf/evidence/i61-resume-20261110/root-diagnostics-prepush.log`.
- Focused runner controls: `Z:/RedSalamander.Perf/runs/20261110T085955Z-74296-63c6530aeb264cb18e9042b7fba3f3b3/scratch/I61Diagnostics/Pester`.
- Focused image driver: `Z:/RedSalamander.Perf/runs/20261110T090456Z-80644-3fdab3a21bb248df9427c3b5f54b29b1/scratch/I61ImageReservation/Handshake/diagnostic.json`.

The earlier broad run failed Terminal counts, missed the image reservation and
rejected two skipped baseline-data audits. Restoring the pinned audit data yielded
32/0/0; no skip allowlist was relaxed. A later full plugin diagnostic crashed during
Microsoft Drive self-tests; matching binaries, symbols and source patch were preserved,
and the owned-child dump-capture attempt subsequently passed. The current broad plugin
pass does not prove that historical intermittent crash is fixed. Hosted ARM64 desktop
retirement remains a separate unresolved gate, whose inner errors are exposed by #211.
