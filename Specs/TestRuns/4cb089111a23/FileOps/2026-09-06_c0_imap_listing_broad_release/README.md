# IMAP checkpoint — surrounding C0 Release coverage

18 passed, zero failures/skips. Native duration: 106.8 seconds; runner wall time:
113.1 seconds. No timeout multiplier, retries or failure reclassification.

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Release -SkipBuild -CaseFilter C0_ -TestRoot C:\RedSalamander.Perf -RunId c0-imap-listing-broad-release-20260906
```

Uses the exact [focused Release build](../2026-09-06_c0_imap_listing_candidate_release/README.md)
and the same tested source snapshot as Debug. The trace retains every selected
case and its nested assertion count, including IMAP 70/70, Move 247/247,
Copy 142/142, lookup 294/294 and size 85/85.

Metric selection retains raw digest and original line positions. The runner
reported 201 historical disk-audit issues; no cleanup was performed to hide them.
This is affected C0 coverage, not a new final Fresh Full or live-server result.
