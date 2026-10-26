# Observatory Track 2 IMAP Identity/Delete Safety Evidence

Date: 2026-07-16
Machine profile: `4cb089111a23`
Working-tree base: `f4e0c8c3bed8`
Configurations: Debug x64 correctness/runtime wiring; test-enabled Release x64 performance baseline

## Scenario

Protect mailbox listing, message open, and single-message delete after a UIDVALIDITY epoch change. Protect a
mailbox that already contains an unrelated `\Deleted` message when UIDPLUS is absent or `UID EXPUNGE` fails.

The user-visible I/O cost is deliberately constant with mailbox size:

- listing: one STATUS UIDVALIDITY command;
- fetch: one STATUS UIDVALIDITY command before `UID FETCH`;
- delete: one STATUS UIDVALIDITY command plus one CAPABILITY command before the mark/UID-EXPUNGE transaction.

Existing `filesystem.imap.status_mailbox_us` instrumentation measures the STATUS round trip.
`filesystem.imap.capability_us` now measures CAPABILITY latency, response bytes, UIDPLUS availability, and HRESULT.

## Commands and results

```powershell
.\build.ps1 -ProjectName FileSystemCurlTests -Configuration Debug
.\.build\x64\Debug\FileSystemCurlTests.exe
.\build.ps1 -ProjectName FileSystemCurl -Configuration Debug
Invoke-Pester -Script Tools/Tests/TestHarnessSourceContracts.Tests.ps1 -PassThru
.\build.ps1 -ProjectName FileSystemCurlTests -Configuration Release
.\.build\x64\Release\FileSystemCurlTests.exe
.\.build\x64\Release\FileSystemCurlTests.exe --perf
```

- Debug FileSystemCurlTests build: 0 warnings / 0 errors.
- Debug FileSystemCurl plugin build: 0 warnings / 0 errors.
- Fake-mailbox/identity/capability executable: passed in Debug and Release.
- Source contracts: 131 passed / 0 failed / 0 skipped.
- Release CPU/parser baseline is preserved in `results.json`, `trace.txt`, and `perf_metrics.jsonl`.

## Before/after interpretation

The 2026-05-14 same-machine Debug diagnostic is the only earlier comparable harness snapshot. Against that run,
the qualified-identity parser changed from 111,134 us to 135,570 us for 500,000 iterations (+22.0%, about
0.271 us per current parse), while filename construction changed from 2,016,913 us to 2,061,549 us (+2.2%).
The parser scenario is intentionally stricter and does more work; this is a directional diagnostic, not a Release
latency claim. Network round trips dominate real IMAP operations and require a scripted/live server for a wall-clock
comparison.

The final test-enabled Release baseline records 10,933 us for 500,000 qualified parses and 125,722 us for 500,000
qualified filename builds. Security command overhead is one command at each named boundary and does not scale with
mailbox message count. No performance improvement is claimed.

## Caveat

No deterministic network-capable IMAP server fixture or credentials are available in this workspace, so live
`filesystem.imap.status_mailbox_us` and `filesystem.imap.capability_us` samples are not part of this archive.
Correctness and command sequencing are deterministic in the fake mailbox; this archive establishes the Release CPU
baseline and the constant command-count contract. A future scripted-server integration lane should capture the live
round-trip metrics without weakening the fail-closed semantics.
