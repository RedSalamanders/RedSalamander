- Make the SearchService contention fixture acquire its owned pipe across listener
  recreation within the existing deadline, retain operation timing assertions, and
  add a deterministic bounded missing-listener control.
