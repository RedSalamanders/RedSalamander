# First ASan repair checks

Source 4eba261d9122d16aa86762f48307b905716f2e38, x64 ASan Debug, receipt
de4dcccc353d3bcd95264b923a4aee440e4270b8708faefadeac6941279495e7.
The solution built without warnings/errors. These are focused checks only.

Search passed the owned status gate and both original default-store startup cases,
but the original single-query case failed. UI passed twelve cases but still failed
inspection return with correct focused identity and empty default/cancel bindings.
The native trace retains the exact witnesses. Retention passed its original measured
1,000-command/private-8-MiB guard after the added ASan cold burst, with the additional
cold/measured/combined live-allocation and cold HWND/GDI/USER requirements intact.
This measurement correction is not a production memory improvement.

Three breadcrumb repeats passed without clearing sequence replay protection. Five
selection-shape target repeats passed, plus their ten setup/cleanup cases, with exact
4/1 active native transfer counts, bytes, contents, sources and discovery guards.
The 100 ms close assertion passed the UI predecessor sequence; this does not establish
the exact cause of the older isolated overrun. Failed broad receipts remain failed.

The failed qualification and complete remaining-probes states, native results and
claim-bearing native UI trace are retained byte-for-byte with manifest digests.
Full current-source validation, positive no-wait regression and migration remain open.
