# C0 reusable strict UTF converter: Release standalone validation

Passed the native `DxUiTests.exe --suite=Theme --no-activate` run in 90 ms,
exit zero, with both Theme start/completion markers. `theme-validation.json`
retains exact arguments, receipt, source snapshot and the complete runtime-closure
hashes. Execution used the normal profile artifact lease and contained process;
scratch and raw evidence remain inside the named owned test run.

The existing UTF-policy test checks caller storage reuse, ASCII and Unicode,
embedded NUL, malformed/truncated input, valid empty input and subsequent recovery.
The observer checks both stdout and stderr, matching the executable's logging
contract. All raw files are copied exactly and bound by `file-provenance.json`.

This is focused shared-converter coverage, not all DxUi suites or Fresh Full.
The exact candidate patch/hashes are retained in the adjacent Debug lookup archive;
both builds bind source snapshot
`b59e77718ac62ce6212984e1fa92efc59f2aa92283b4b9f1fd64b123fa44dbf4`.
