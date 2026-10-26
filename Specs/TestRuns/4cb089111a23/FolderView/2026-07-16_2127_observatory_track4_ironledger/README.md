# Observatory Track 4 / IronLedger FolderView Integrity Evidence

Date: 2026-07-16
Machine profile: `4cb089111a23`
CPU: AMD Ryzen 9 9950X3D 16-Core Processor
OS: Windows 11 Pro 10.0.26200
Working-tree base: `f4e0c8c3bed8`
Configuration: Debug x64

## Scenario

Close the FolderView data-integrity defects owned by IronLedger and Observatory Track 4: bounded private/
`CF_HDROP` parsing, actual hovered/background destination resolution, self/same-parent/descendant rejection,
settled destination gates, explicit local-source provider routing for external imports, honest asynchronous MOVE
reporting, gesture target stability, bounded enumeration allocation hints, stale-hover rendering safety, checked
clipboard construction, and verified move-paste clipboard invalidation.

The payload ceilings are 65,536 paths, 32,768 UTF-16 characters per path, and 16 Mi UTF-16 characters in
aggregate. The regression corpus includes an internal header with a hostile maximum count and a `CF_HDROP`
whose file offset lies beyond its HGLOBAL.

## Commands and results

```powershell
.\build.ps1 -ProjectName RedSalamander -Configuration Debug -Platform x64
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_drop_integrity_guards --selftest-timeout-multiplier=2
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_drop_rejects_malformed_payloads --selftest-timeout-multiplier=2
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=folderView_pointer_targets_and_stale_hover_are_safe --selftest-timeout-multiplier=2
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=cmd_pane_clipboardPaste_uses_preferred_move_effect --selftest-timeout-multiplier=2
.\.build\x64\Debug\RedSalamander.exe --commands-selftest --selftest-case=Commands_FileOpsDragDropMissingCallbackRejectsDirectFallback --selftest-timeout-multiplier=2
Invoke-Pester -Script .\Tools\Tests\TestHarnessSourceContracts.Tests.ps1 -PassThru
```

- Focused Debug build: passed, 0 warnings / 0 errors; log `.build/logs/msbuild-20260716_212337_249.log`.
- Focused Commands regressions: 5 / 5 processes exited 0.
- Source contracts: 134 passed / 0 failed / 0 skipped.
- No broad Full suite was repeated. This is the bounded closeout gate requested for convergence after the prior
  aggregate Observatory work; the focused build, runtime cases, and repository contract suite cover the touched
  surfaces without spending another cycle on unrelated Full-suite failures.

## Performance interpretation

No throughput or latency improvement is claimed. The new work is constant-time structural validation plus bounded
allocation before existing operation dispatch. The archived values are safety ceilings and observable dispatch
counts, not benchmark claims. Malformed inputs and stale destinations enqueue zero file-operation tasks.
