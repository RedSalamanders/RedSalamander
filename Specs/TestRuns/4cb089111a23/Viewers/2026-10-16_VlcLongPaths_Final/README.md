# VLC long paths: final focused evidence

2026-10-16, x64 Debug runtime, VLC 3.0.24 x64, machine `4cb089111a23`.
Windows `LongPathsEnabled` remained `0`. All cases used `--no-activate`, a
profile artifact lease, and fixtures beneath the marked `C:\RedSalamander.Perf`.
The user's media was not opened or copied by these tests.

| Case | Passed assertions | Failed | Skipped |
| --- | ---: | ---: | ---: |
| Long paths and error messages | 149 | 0 | 0 |
| HUD, loader faults, stale results, cleanup and unload | 121 | 0 | 0 |
| Configuration persistence and bounds | 30 | 0 | 0 |

Seven actual two-second WAV decodes cover standalone paths of exactly 259, 260
and 275 characters, embedded preview at 275 characters, an extended spelling,
a forward-slash spelling, and a symbolic link to the long-path media. The
directory includes Unicode, spaces, `%` and `#`. The test requires decoded
duration, not merely successful player allocation or an asynchronous play call.

The missing-file and invalid-path cases require native errors 2 and 123
respectively. Directory and broken-link cases require neutral file errors;
only a missing installation shows installation advice and a download link.
A directory shaped like an installation but containing an invalid DLL retains
its specific loader error without installation advice. Guarded relative link
targets stay in the same marked sandbox; both link scenarios executed here.

`provenance.json` records the actual source and executable/DLL hashes used by
run `vlc-paths-final-r4-20261016`. `results.json` and the three assertion traces
are focused native evidence, not a Full repository verdict or reusable evidence
for another artifact/profile. The source/localization/sandbox guard trace has
217 passing checks. `release-build.json` separately binds the successful x64
Release ViewerVLC build; Release does not expose the Debug observation hooks.

The 122 raw `viewer.vlc.*` rows include deliberate load/cleanup delays and
failure injection. The largest synchronous Close sample was 7,697 us; the
largest cleanup sample was 307,513 us. These observations protect the existing
asynchronous contract. No before/after throughput or causal improvement is
claimed.

Earlier failed fixture attempts, the build-lock collision, and the MSVC internal
compiler-error log are retained in sibling `2026-10-16_VlcLongPaths`. The link
fixture was corrected to use the guard's supported relative-target form; the
guard was not relaxed. The compiler retry used fewer MSBuild workers and passed;
that single retry is not a diagnosis or a general compiler-workaround claim.

The canonical x64 Debug Affected run
`20261016T090333Z-77908-4d0fd0cf952647a6bc263bb649339d21` executed and promoted
all 15 selected entries: 1,211 recorded cases passed, zero failed, and 26 were
conditionally skipped (external services/devices and journal-currentness
capabilities). The ViewerPE noninteractive entry executed this regression again.
`affected-results.json`, its legacy detail companion, the plan, and run state
retain the aggregate verdicts and skip reasons. Its Debug build receipt is
`3a8e06714a7cbee1aabc5c94054ff86b789899232270bb7a50042032ffbc584b`.

The change verdict is `PASSED`; the repository verdict is `NOT_EVALUATED` because
coverage is filtered/Affected. Commands, FileOperations, ViewerPE interactive,
ViewerSqlite interactive, and Monitor ETW latency remain `PENDING_INTERACTIVE`
until the person agrees to desktop input. The outer shell reported exit 1;
this is not recorded as a successful Full invocation. The v2 aggregate and
entry promotions, rather than the legacy zero exit field, define the scope.

After promotion, the three owned focused scratch runs were removed with the
canonical guarded cleanup. The before/after disk audits retain three remaining
unattributed or older directories, which were left untouched. This archive
does not claim a clean whole-sandbox audit.

Validate this folder with `Tools/Test-TestRunArchive.ps1 -RunPath` and the
archive inventory gate.
