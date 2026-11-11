- Queued file operations now start in the order of their cards. A task takes its queue place when it is
  admitted and keeps it while Preparing, so a later task that prepared faster no longer starts first. A
  task is passed only while it waits for the user's answer on its own card or is queued behind a named
  task, and it starts next once that reason clears; a task that is cancelled or fails before starting
  gives its place up at once. Active cards are ordered by one key, so a card no longer moves when its
  task starts. A waiting card states its place (`Next in the queue`, `Position N in the queue`), and Move
  Up or Move Down is offered only for a move the engine performs, never across a task queued behind a
  named task. Seven engine scenarios of the start order run inside `FileOps_IdentityInterlockWaiting`
  and pass; the popup cases (position text, move edges, card order) wait for an agreed desktop time.
