# IMAP checkpoint — surrounding C0 Debug coverage

18 passed, zero failures/skips. Native duration: 260.1 seconds; runner wall time:
266.8 seconds. No timeout multiplier, retries or failure reclassification.

```powershell
.\Tools\Run-AllTests.ps1 -Suite FileOps -Configuration Debug -SkipBuild -CaseFilter C0_ -TestRoot C:\RedSalamander.Perf -RunId c0-imap-listing-broad-debug-20260906
```

Uses the exact [focused Debug build and tested inputs](../2026-09-06_c0_imap_listing_candidate_debug/README.md).
Trace retains Delete guards 29/29, lost-reply 33/33, partial-tree 37/37,
late-listing 49/49, Delete traversal 127/127, size 85/85, Move 247/247,
Copy 142/142, lookup 294/294 and IMAP 70/70, plus selected host cases.
These nested counts are not additional top-level runner cases.

The ordered selected metric stream retains its raw digest and original line
positions. The runner retained 198 disk-audit issues. No files were removed to
suppress them. This is affected C0 coverage, not Fresh Full or live-server proof.
