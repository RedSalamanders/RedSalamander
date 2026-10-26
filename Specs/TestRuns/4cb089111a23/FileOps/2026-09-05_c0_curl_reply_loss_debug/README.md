# C0 Curl lost mutation reply — candidate Debug

- Source: `58f22d0e` plus the C0 working changes, one-attempt/no-body quote repair,
  native lost-reply checks and host Delete feedback witness.
- Command: `Tools/Run-AllTests.ps1 -Suite FileOps -Configuration Debug -CaseFilter C0_ -FailFast -RunId c0-curl-reply-loss-candidate-debug-20260905`.
- Accepted build receipt: `2a098a511b31fc8738401678dcbd0ebdbaa800f2f52b052bbd8aebad72e2c2ca`;
  zero warnings/errors; build 2:48.
- Harness: 9 passed, 0 failed/skipped; total native harness duration 7,872 ms.
  Native Curl delete guards: 29/29. Native lost-reply checks: 33/33.

The same-machine [pre-fix run](../2026-09-05_c0_curl_reply_loss_red_debug/README.md)
replayed all four lost-reply mutations and destroyed or moved the replacement.
Here each mutation is sent exactly once; replacement and sibling bytes survive,
Rename retains the original bytes at its destination, and one uncertain callback
is emitted. Ordinary acknowledgement controls remain successful. Every scenario
also proves zero incidental NLST commands.

Lost-reply durations (DeleteItem / DeleteItems / RenameItem / RenameItems):
3,566 / 3,616 / 5,613 / 5,174 us, versus approximately 219–221 ms with replay.
The extra retry delay disappeared; these loopback fault timings are not a general
throughput or real-server latency claim. The no-NLST assertion and host witness
were added after the pre-fix run, so total suite times are not directly comparable.

The host Delete witness completed in 63 ms from task submission: one request,
one backend commit, preserved replacement/sibling, immutable Indeterminate result,
`ERROR_IO_INCOMPLETE`, localized unknown summary and zero conflict prompts.
Google Drive response-loss and Local publication/ABI guards also remain green.

Evidence is copied intact from
`C:\RedSalamander.Perf\runs\c0-curl-reply-loss-candidate-debug-20260905\artifacts\selftest\last_run`.
The runner reported 34 retained-run disk-audit issues; no old data was removed.
Release, native route/walker completeness and C1–C6 remain separate gates. This
focused pass is not final Fresh Full or live FTP/SFTP/SCP qualification.
