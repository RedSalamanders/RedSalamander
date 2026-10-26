# File Operations ASan baseline diagnosis

Eight isolated x64 ASan probes compare master `0dd0bd6d` with frozen migration
`a20f6680`. They do not replace the failed broad run or establish paired performance
non-regression. No hosted CI ran. Original suite/runner/aggregate results and inert
drivers are retained, including baseline precondition failures.

- Master reproduces R0d's two-second capability benchmark failure. Its Release
  performance contract and the adjacent R0e Debug diagnostic profile are distinct.
- Master reproduces the Curl wide-copy failure: 141/142 assertions, with the serial
  4,097-file case timing out at 180 seconds after 3,690 publications. The migration
  broad run times out after 3,280 publications. Both results preserve truthful
  timeout completion; neither qualifies the full corpus.
- Master queue-after and bandwidth cases pass in isolation. Migration queue-after
  also passes in isolation, preserving its earlier broad-process zero-wait failure.
- The first master MTP probe fails setup on an existing journal path (0x800700B7).
  A fresh process-local LOCALAPPDATA root admits the fixture and passes. Migration's
  already-isolated runner also passes. This does not reproduce or waive the earlier
  broad-run discovery/cancellation overlap failure.

The queue-after correction holds the Copy at the existing publication pause until
its Delete actually waits or completes explicit cancellation. R0d keeps its Release
limit and adopts R0e's five-second Debug diagnostic ceiling. The Curl fixture keeps
all 4,097 files and uses 360 seconds only in ASan; ordinary deadlines/cancel bounds
remain. Candidate focused and Full validation are pending.

`case-comparison.json` records each exact case result. `retained-files.json` binds
original small receipts; metric excerpts preserve original row bytes with separate
full-source hashes and selection rules. Source comparison includes canonical
formatting hashes where original formatting differed.
