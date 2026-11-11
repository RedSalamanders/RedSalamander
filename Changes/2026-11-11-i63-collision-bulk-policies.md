- A file collision can now be answered with one of two task-wide choices under `More...`: `Skip all
  with the same size and date` and `Replace all older`. They run an interrupted copy or move again
  without a question for every file already there.
  - Each is offered only when the prompt's own pair matches it and the prompt offers Apply to all;
    choosing it answers the current file and every later matching file of the task. `Replace all older`
    replaces only when the incoming file is newer by more than the two-second timestamp tolerance.
  - They cover plain regular-file collisions only: a read-only file, a link, a folder, a type mismatch,
    the one-hour class and unknown facts still ask. They are consulted before the Apply-to-all answers,
    and `Skip all with the same size and date` and a remembered `Delete source` for all replace each
    other (the later one wins).
  - A policy skip is an ordinary user skip (a Move keeps its source) and opens no content; a policy
    replacement takes the route of a single `Replace`. The card's Result line adds `N already there`
    and `N replaced because older`; the task records one Information note per policy.
  - New actions `ConflictAction::SkipSameSizeAndDate` and `ReplaceOlder` (appended, existing values
    unchanged), two task counters, four strings in five resource files, policy-matrix coverage and a
    new self-test in `FileOps_ReviewRemediationContracts`.
  - Not part of this change: choosing a policy before the task starts (D-13), `Skip` and `Keep both`
    as policy values, a task-wide `Delete source` beyond B4, and item-scoped policy receipts for I18.
