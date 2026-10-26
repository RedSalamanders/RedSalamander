# C0 Curl directory-size truth — red baseline and pause checkpoint

- Starting implementation: `78a61e7b`. This checkpoint changes tests, host test
  registration, and documentation only; production `GetDirectorySize` is unchanged.
- Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_CurlDirectorySizeTruth -FailFast -RunId c0-curl-size-truth-red-debug-20260905`.
- Debug build: 2:55 (175,937 ms), zero warnings/errors, tests enabled. Receipt
  `14b2a7e8a1f6835646b4b4c1b0cfce9115a26ecbb8bc56c95eaf80adf3136eff`;
  source snapshot `bf597239d957b222401ae850b1232f0f1efb43377a47443df439d8ae53ab4df6`.
- Harness: two setup/cleanup passes, one failed behavior case, zero skips;
  12.5 seconds overall, 10.3 seconds for the behavior case.
- Native checks: **28 passed, 33 failed**, across 20 scenarios. All setup checks
  pass and all final fake-server health checks are successful.

Seven scenarios have correct totals/status/content: known file, empty directory,
immediate children, 4,097 files, depth 80, checked byte overflow, and literal arrow
name. All 20 fail the feedback contract: directory scans check cancellation before
the first progress report, and successful final reports are not followed by the
required cancellation check. These related assertions are not 33 independent bugs.

The other thirteen scenarios expose concrete state/result gaps:

- Unknown listed sizes produce zero-byte success for both file roots and children,
  even when the server supports a targeted size probe; genuinely unavailable sizes
  also produce false success instead of unsupported-root/partial-child results.
- Depth 129 succeeds rather than enforcing the shared explicit traversal ceiling.
- Pre-cancel still contacts the server. Flat-list cancellation at a progress report
  is ignored: all 4,097 files finish with `S_OK`.
- A final callback returning `E_ABORT` replaces the original root access denial.
- A denied child stops sibling traversal (3 bytes, one file, two directories) rather
  than returning partial totals after the remaining sibling (6 bytes, two files,
  three directories).
- Malformed, out-of-root, and overlong LIST rows are accepted as successful totals.
- Cancellation during an observed blocked LIST does not return within two seconds.
  The owned emergency server stop is needed; the call returns `ERROR_SEM_TIMEOUT`,
  not cancellation. The 2,005,805-us wait measurement is a failed bound, not successful
  cancellation latency; overall call time is 7,103,065 us including emergency drain.

The baseline wide/deep durations are 53,548 / 269,338 us. These are unbounded,
buffered/recursive implementation timings, not bounded-memory qualification or a
throughput target. The repaired candidate must preserve data and honest feedback
before any latency comparison is meaningful. No Release candidate or full C0 rerun
has been made for this checkpoint. The plan remains WIP.

The five raw files beside this README are exact copies from the owned run beneath
`C:\RedSalamander.Perf`; the complete 123,559-byte perf stream is retained. The runner
records 51 retained-run disk-audit issues. No existing run, user `last_run/`, or other
user data was inspected, removed, or staged. Resume instructions are in the active
reliable-core plan's C0 pause checkpoint.

Checkpoint checks: source/localization Pester 197/197 (16.72 seconds), spec inventory
zero blockers, explicit six-file archive validation passed, and patch whitespace
checks passed. All five raw copies match their originals by SHA-256. These checks
validate the saved checkpoint, not the intentionally failing production behavior.
