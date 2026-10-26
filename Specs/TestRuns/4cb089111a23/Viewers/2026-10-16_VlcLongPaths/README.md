# VLC long paths and failure messages

Focused x64 Debug runtime evidence, 2026-10-16, machine `4cb089111a23`.
VLC 3.0.24 x64 was installed and Windows `LongPathsEnabled` remained `0`.
`provenance.json` binds the source files and actual executable/DLL bytes used.

All three cases ran with `--no-activate` under the profile artifact lease:

| Case | Passed assertions | Failed | Skipped |
| --- | ---: | ---: | ---: |
| Long paths and error messages | 125 | 0 | 0 |
| HUD, loading, fault injection, cleanup and configuration boundaries | 121 | 0 | 0 |
| Persisted volume and mute configuration | 30 | 0 | 0 |

The long-path case decoded six two-second silent WAV fixtures: standalone paths
of exactly 259, 260 and 275 characters; a 275-character embedded-preview path;
an already extended spelling; and a forward-slash spelling. All used a Unicode
directory containing literal spaces, `%` and `#`. Positive decoded duration was
required, not merely successful media/player allocation or an asynchronous play
return. The user's media was not used as test data.

Missing long files, invalid/non-local paths, directories, missing installations,
and an installation-shaped directory with an invalid DLL proved the actual
heading/body/link policy. File failures separately required native errors 2 and
123. The existing fault-injection case also verified a neutral loader-submit
error without installation advice, stale-result rejection, asynchronous close,
retained drawables, and module lifetime through forced parent destruction.

The 114 raw `viewer.vlc.*` rows include intentional load/cleanup delays and
failure injection. The largest synchronous Close sample was 17,042 us; the
largest cleanup sample was 305,672 us. These are scenario observations, not a
before/after throughput comparison or a claim of repository-wide qualification.

`failed-fixture-r1.*` preserves the initial failed attempt: its fixture searched
top-level windows for an embedded child. The test was corrected to use its owned
parent's child lookup, then rerun as `vlc-paths-20261016-102108-r2`. Standalone
long-path decoding had already passed in that failed run. Only its failed trace
and terminal result are promoted; no timing comparison relies on that run.

The separate localization/source-contract results contain 217 passing checks.
The explicit TestRuns archive validator and inventory validator are required
alongside this archive. Broader affected-suite results are separate evidence;
this focused archive does not authorize Full or cross-profile reuse.

The first broader Affected attempt, `20261016T083814Z-89160-8cf74e0bcff04fcd8f79b30d335df41b`,
was blocked when a concurrent Release build held the repository/platform shared
dependency lock. Build-tooling fixtures and subsequent profile re-attestation
were rejected by that lock. `failed-build-coordination-trace.txt` retains one
original MSBuild fixture log; `failed-build-coordination-state.json` is the
unfinished runner state, not a successful terminal result. A fresh sequential
attempt is required; none of this attempt is promoted as passing coverage.
