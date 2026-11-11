- Staging files created on Windows paths now end in a real extension: `<name>.<32 hex>[_<stream>].rs-tmp`
  for incomplete content (host bridge stage, Local copy, link and directory stages, the writer's
  overwrite stage) and `<name>.<32 hex>.rs-bak` for a name that may hold an original (Local overwrite
  backup, case-only rename hop). The bridge stage name is now built by the shared bounded-name helper,
  so its leaf stays within 255 characters for a long final name. Leftover recognition is simpler: it
  accepts only the names the current build writes (owner decision D-17), so a leftover written by an
  earlier build is an ordinary file, and a name that only ends in `.rs-tmp` or `.rs-bak` stays ordinary.
  Remote, S3 and MTP names are unchanged.
