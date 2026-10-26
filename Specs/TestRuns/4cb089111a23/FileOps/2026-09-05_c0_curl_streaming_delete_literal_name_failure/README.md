# C0 bounded Curl Delete — literal command-path failure

Starting checkpoint `5132d866` plus bounded Delete, exact-root error handling,
and corrected peer-disconnect fixture behavior. Command:
`Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_ -FailFast -RunId c0-curl-streaming-delete-qualified-debug-20260905`.

Build: 2:43, zero warnings/errors. Receipt
`dedc28727d4821f53e320165fd23da87284915be3ec8e919459322979c870805`.
Harness: 12 passed, 1 failed, zero skips. Partial-tree checks 37/37 and
late-listing checks 49/49 now pass. Bounded traversal: 89 passed / 8 failed.

All four wide cases delete 4,097 leaves (0.447–0.850 seconds); all four 80-level
cases pass (0.362–0.411 seconds). Depth overflow and malformed/out-of-root/overlong
records fail explicitly with preserved contents and no-commit receipts. Link
entries are removed without touching their target sibling.

The four literal-arrow-name cases each fail both final-state and success-receipt
assertions, with zero deleted leaves. The cursor retains the complete name, but
RemotePathForCommand incorrectly reused URL-escaped text in a literal FTP command.
The server returns failure for the percent-encoded spelling. The correction must
use the captured literal base path and protocol-specific operand encoding, preserve
literal percent sequences, and reject command delimiters/NUL/invalid Unicode.
SFTP/SCP operands require their own libcurl argument quoting, not URL encoding.

All results/reasons, traces and the complete ordered perf stream are retained.
The run's diagnostic name does not make this a qualification pass. Disk-audit
limitations remain in the aggregate, and earlier evidence has not been removed.
