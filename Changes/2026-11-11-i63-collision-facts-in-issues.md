- The Issues pane now shows, for a file that already existed at the destination, the size and date of both
  files and how they compare, in the words the question used.
  - The prompt observation keeps a typed `ConflictObservedComparison` (both items' size and last-write
    time and the engine's comparison) for a collision of two regular files whose facts were read.
  - `BuildConflictObservedTexts` reuses the shared prompt text builder; the detail adds the summary line
    to the Problem and each item's line under its Source and Destination path.
  - Not done: the finished card of a single-item task (B3.3's second half), sizes on provider routes
    (B3.2; the provider issue callback carries no sizes), and facts for items a remembered answer or a
    policy decided.
