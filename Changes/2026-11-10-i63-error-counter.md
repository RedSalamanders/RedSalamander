- One failed item no longer reads `Errors: 2`. The counters on a task card count problems: the
  trailing `task.result` Error is not counted when an Error of the same task already explains the
  failure. The record is still written to Issues, the log and the export, and it counts when it is
  the only Error. The two visual-gallery fixtures of the failed Move card use `Errors: 1`.
