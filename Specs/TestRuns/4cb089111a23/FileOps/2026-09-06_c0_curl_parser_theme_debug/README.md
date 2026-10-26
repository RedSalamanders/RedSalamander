# C0 reusable strict UTF converter: Debug standalone validation

Native command: `DxUiTests.exe --suite=Theme --no-activate --perf-jsonl=<owned run path>`.
The exact executable, arguments, receipt, complete runtime-closure hashes and
elapsed 111 ms are in `theme-validation.json`. Execution used the repository's
profile-scoped artifact lease and contained-process helper, with temporary paths
under `C:\RedSalamander.Perf\runs\c0-curl-parser-theme-debug-20260906`.

The native test exited zero, printed `All DxUi tests passed.`, and emitted both
Theme start/completion markers. The existing UTF-policy test exercises caller
storage reuse, Unicode including supplementary characters, embedded NUL, invalid
and truncated UTF-8, empty input, and recovery after failure.

The one-off wrapper incorrectly looked for markers only in stdout; DxUiTests.cpp
emits them on stderr. Its original negative classification is retained unchanged.
`theme-reviewed.json` binds that record and both raw streams, and records the
correct native PASS. This is an observer correction, not a rerun or test change.
The wrapper's output projection also showed null values because it selected an
ordered dictionary as an object; the persisted record contains the real values.
Subsequent standalone invocations inspect both streams and project an object.

`file-provenance.json` binds every exact raw artifact. This focused Theme run
does not qualify all DxUi suites or the repository. The candidate patch/hashes
are in `../2026-09-06_c0_curl_parser_lookup_debug/`.
