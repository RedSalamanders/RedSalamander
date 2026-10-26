# Native Copy traversal baseline — interrupted, not passing

Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_CurlCopyTraversalTruth -FailFast -RunId c0-curl-copy-traversal-red-debug-20260905`.

Production is unchanged from `60012f9a`; the new 139-assertion fixture and registration
are the candidate. Build completed in 02:55 with zero warnings/errors. Receipt:
`f042ac9fefa670009ffbe504ce7da8d4459725bc8482366715a5cd0d5fb23c67`.
Workspace snapshot: `b07882ccebb341fba99584471535ca2298b1c74e2e67d974de575ee39673b1f0`.

The owned process (PID 3284) terminated after 34 completed scenario records and
the Begin record for shape 9, singular Copy at concurrency 1 (accepted depth 128).
There is no FileOps results.json or usable stack in this run. The wrapper correctly
reports missing results and an unclassified failure. Do not infer the exception
code or fabricate verdicts for the remaining depth cases.

Observed results before termination:

- Shallow/literal controls passed across all four Copy routes.
- Malformed and overlong LIST data produced FILE_NOT_FOUND, not their precise
  INVALID_DATA / BUFFER_OVERFLOW cause, in all four routes each.
- Out-of-root entries changed destination contents in all four routes and returned
  ERROR_BAD_NET_RESP rather than rejecting the invalid child name.
- First-LIST cancellation returned in about one second. Serial late cancellation
  retained the earlier published file; both parallel variants reached the later
  wait before publishing it. Late access refusals preserved the earlier file.
- The 4,097-file serial Copy reached its 120-second operation deadline after 3,119
  publications; parallel Copy returned E_FAIL in 333,515 us after 113 publications.
  Neither is a passing wide traversal result.
- Depth-128 singular Copy began but produced no terminal record. Deep recursion is
  a suspect, not a proven crash diagnosis without a stack.

The archive retains the original complete perf stream and both traces, plus all
wrapper failure reasons. Raw evidence remains in the named run beneath
`C:/RedSalamander.Perf`; no legacy/user-owned runtime directory was inspected.
Disk audit has 66 retained-run issues; no data was deleted to hide them.

Checkpoint checks: 197/197 source/localization Pester checks; spec inventory zero
blocking findings; explicit archive validation five files; `git diff --check` clean.

Next: replace recursive Copy with the existing strict cursor and iterative bounded
batches, qualify the wide failure separately if it remains, then rerun the entire
new gate and existing C0/ordinary FTP gates. C0 and the plan remain open.
