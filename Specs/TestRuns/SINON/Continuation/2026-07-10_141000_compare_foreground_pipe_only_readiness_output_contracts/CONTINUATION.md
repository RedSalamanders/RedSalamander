# Continuation Evidence: Compare Foreground Pipe-Only Readiness Output Contracts

Captured after commit `efef2ac4a` made foreground search readiness pipe-only. This is not final proof.

Command evidence:
- Wrapper log: `.build/codex-runs/compare-full-pipe-only-readiness-20260710_140415.out.log`
- Run root: `C:\RSPerf\runs\20260710T120416Z-95744-ad6b10c7344f404188edc08f890808c3`

Result:
- CompareDirectories: 218 passed / 5 failed / 30 skipped
- Disk audit: 0 unexpected entries

Failures:
- `search_service_sqlite_idle_maintenance_queue_and_completion`
- `search_service_sqlite_invalid_store_falls_back_live_scan`
- `search_service_sqlite_status_reports_pending_legacy_import`
- `search_service_sqlite_startup_warms_overridden_roots`
- `search_service_foreground_logs_request_status`

Working theory:
- Removing the hidden `ForegroundSearchServiceProcess::WaitUntilReady` status request is correct, but several capture-output tests were implicitly relying on that hidden request to generate foreground dashboard/log lines or to consume a max-request budget. They now need explicit status/query interactions and request budgets that match the behavior they assert.
